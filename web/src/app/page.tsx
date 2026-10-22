"use client";

import Link from "next/link";
import { useRouter } from "next/navigation";
import { useEffect, useState } from "react";
import { authentication } from "@microsoft/teams-js";
import { getRequestedModule, getSsoToken, initTeams, isPreview } from "@/lib/teams";

type ModuleKey = "leave" | "expense";

type ModuleSummary = {
  key: ModuleKey;
  pendingCount: number;
  error?: string;
};

type PendingItem = {
  module: ModuleKey;
  id: string;
  subject: string;
  employeeCode: string;
  employeeName: string;
  submittedDate: string;
  detail: string;
};

type HomeSummary = {
  displayName: string;
  mail: string;
  modules: ModuleSummary[];
  pending: PendingItem[];
};

/** Everything the home screen knows about a module besides its route. Order here is tile order. */
const MODULES: Record<ModuleKey, { newLabel: string; title: string; description: string; icon: string; short: string }> = {
  leave: {
    newLabel: "New leave request",
    title: "Leave request",
    description: "Holiday, medical, bonus days — request time off and track your balance.",
    icon: "🏖️",
    short: "Leave",
  },
  expense: {
    newLabel: "New expense report",
    title: "Expense report",
    description: "Report business expenses with receipts, or a cash advance.",
    icon: "🧾",
    short: "Expense",
  },
};
const MODULE_ORDER: ModuleKey[] = ["leave", "expense"];

const cardClass = "rounded-xl border border-line bg-surface p-4";

function initials(name: string): string {
  const parts = name.trim().split(/\s+/);
  return ((parts[0]?.[0] ?? "") + (parts[1]?.[0] ?? "")).toUpperCase() || "?";
}

function formatDate(dateString: string): string {
  if (!dateString) return "";
  return new Date(dateString).toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" });
}

async function fetchSummary(): Promise<HomeSummary> {
  if (isPreview) return (await import("@/lib/preview")).previewApiFetch<HomeSummary>("/home/GetHomeSummary");
  const ssoToken = await getSsoToken();
  const response = await fetch(`${process.env.NEXT_PUBLIC_API_BASE_URL}/home/GetHomeSummary`, {
    headers: { "X-Teams-Sso-Token": ssoToken },
  });
  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    throw new Error(body.error ?? `Request failed with status ${response.status}`);
  }
  return (await response.json()) as HomeSummary;
}

/**
 * The request-type picker: one tile per module this tenant is provisioned
 * for, plus a single combined "awaiting your approval" list across all of
 * them. Opening a pending item lands on that module's own Pending Approvals
 * tab, where the actual approve/decline happens — this screen only routes.
 */
export default function Home() {
  const router = useRouter();
  const [status, setStatus] = useState<"loading" | "ready" | "error">("loading");
  const [summary, setSummary] = useState<HomeSummary | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  useEffect(() => {
    async function load() {
      try {
        await initTeams();

        // Notification emails deep-link straight into a module (see the API's getTeamsDeepLink).
        const requested = await getRequestedModule();
        if (requested === "leave" || requested === "expense") {
          router.replace(`/${requested}`);
          return;
        }

        try {
          setSummary(await fetchSummary());
        } catch (error) {
          const message = error instanceof Error ? error.message : "Unknown error";
          if (!message.includes("AADSTS65001") && !message.includes("consent")) {
            throw error;
          }
          // Same incremental-consent fallback as the module pages.
          await authentication.authenticate({
            url: `${window.location.origin}/auth-start`,
            width: 600,
            height: 535,
          });
          setSummary(await fetchSummary());
        }
        setStatus("ready");
      } catch (error) {
        setErrorMessage(error instanceof Error ? error.message : "Unknown error");
        setStatus("error");
      }
    }

    void load();
  }, [router]);

  if (status === "loading") {
    return (
      <main className="flex min-h-screen flex-col items-center justify-center gap-3 bg-page p-6 font-sans text-ink sm:p-8">
        <p className="text-sm text-ink-2">Connecting…</p>
      </main>
    );
  }

  if (status === "error" || !summary) {
    return (
      <main className="flex min-h-screen flex-col items-center justify-center gap-5 bg-page p-6 font-sans text-ink sm:p-8">
        <div className={`${cardClass} max-w-md`}>
          <p className="text-sm font-semibold text-critical">Loading error</p>
          <p className="mt-1 text-sm text-ink-2">{errorMessage}</p>
        </div>
      </main>
    );
  }

  const enabled = MODULE_ORDER.map((key) => summary.modules.find((m) => m.key === key)).filter(
    (m): m is ModuleSummary => Boolean(m)
  );
  const totalPending = summary.pending.length;

  return (
    <main className="flex min-h-screen flex-col items-center justify-start gap-5 bg-page p-6 font-sans text-ink sm:p-8">
      <div className="flex w-full max-w-4xl flex-col gap-6">
        {/* Identity strip */}
        <div className="flex items-center justify-between gap-3">
          <div className="flex min-w-0 items-center gap-2.5">
            <span className="flex h-9 w-9 flex-none items-center justify-center rounded-full bg-accent text-[13px] font-bold text-accent-ink">
              {initials(summary.displayName)}
            </span>
            <div className="min-w-0 leading-tight">
              <div className="truncate text-[14.5px] font-semibold text-ink">{summary.displayName}</div>
              <div className="truncate text-xs text-ink-2">{summary.mail}</div>
            </div>
          </div>
          <div className="hidden text-[11px] font-bold tracking-wider text-ink-3 uppercase sm:block">
            Internal processes{isPreview && " · preview"}
          </div>
        </div>

        {/* Request-type tiles */}
        <section className="flex flex-col gap-3">
          <h1 className="text-lg font-bold text-ink">What would you like to do?</h1>
          {enabled.length === 0 ? (
            <div className={cardClass}>
              <p className="text-sm text-ink-2">No modules are set up for your organization yet — contact IT ASSIST.</p>
            </div>
          ) : (
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
              {enabled.map((m) => {
                const meta = MODULES[m.key];
                return (
                  <div key={m.key} className={`${cardClass} flex flex-col gap-3`}>
                    <Link
                      href={`/${m.key}`}
                      className="group flex items-start gap-3 focus:outline-none focus-visible:ring-2 focus-visible:ring-accent/40"
                    >
                      <span className="flex h-11 w-11 flex-none items-center justify-center rounded-lg bg-accent-soft text-2xl" aria-hidden>
                        {meta.icon}
                      </span>
                      <span className="flex min-w-0 flex-1 flex-col gap-1">
                        <span className="flex items-center justify-between gap-2">
                          <span className="text-[15px] font-semibold text-ink group-hover:text-accent">{meta.title}</span>
                          {m.pendingCount > 0 && (
                            <span className="inline-flex flex-none items-center gap-1.5 rounded-full bg-warning-soft px-2 py-0.5 text-[10px] font-bold tracking-wide text-warning uppercase">
                              <span className="h-1.5 w-1.5 rounded-full bg-warning" />
                              {m.pendingCount} to approve
                            </span>
                          )}
                        </span>
                        <span className="text-sm text-ink-2">{meta.description}</span>
                        {m.error && <span className="text-xs text-critical">Couldn&apos;t load: {m.error}</span>}
                      </span>
                    </Link>
                    <div className="flex items-center justify-between gap-2">
                      <Link
                        href={`/${m.key}?new=1`}
                        className="rounded-lg bg-accent px-3 py-1.5 text-sm font-semibold text-accent-ink hover:brightness-110"
                      >
                        + {meta.newLabel}
                      </Link>
                      <Link href={`/${m.key}`} className="text-sm font-semibold text-accent hover:underline">
                        Open
                      </Link>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </section>

        {/* Combined approvals inbox */}
        <section className="flex flex-col gap-3">
          <div className="flex items-baseline justify-between gap-2">
            <h2 className="text-sm font-bold text-ink">Awaiting your approval</h2>
            <span className="text-xs text-ink-3">{totalPending === 0 ? "Nothing pending" : `${totalPending} total · oldest first`}</span>
          </div>
          {totalPending === 0 ? (
            <div className={cardClass}>
              <p className="text-sm text-ink-2">You&apos;re all caught up — nothing is waiting on you.</p>
            </div>
          ) : (
            <ul className="flex flex-col gap-2">
              {summary.pending.map((item) => (
                <li key={`${item.module}-${item.id}`}>
                  <Link
                    href={`/${item.module}?tab=pendingApprovals`}
                    className="flex items-center gap-3 rounded-lg border border-line border-l-[3px] bg-surface p-3 hover:bg-surface-2"
                    style={{ borderLeftColor: "var(--warning)" }}
                  >
                    <span className="w-16 flex-none rounded-full bg-accent-soft px-2 py-0.5 text-center text-[10px] font-bold tracking-wide text-accent uppercase">
                      {MODULES[item.module].short}
                    </span>
                    <span className="min-w-0 flex-1">
                      <span className="block truncate text-sm font-semibold text-ink">
                        <span className="font-medium text-ink-2">{item.employeeName} — </span>
                        {item.subject}
                      </span>
                      <span className="block truncate text-xs text-ink-2">
                        {item.detail} · submitted {formatDate(item.submittedDate)}
                      </span>
                    </span>
                    <span className="flex-none text-sm font-semibold text-accent">Review →</span>
                  </Link>
                </li>
              ))}
            </ul>
          )}
        </section>
      </div>
    </main>
  );
}
