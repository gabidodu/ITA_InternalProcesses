"use client";

import { useEffect, useMemo, useState } from "react";
import { app, authentication } from "@microsoft/teams-js";
import { getSsoToken, initTeams, isPreview } from "@/lib/teams";
import { SearchableSelect } from "@/components/SearchableSelect";
import { DatePicker } from "@/components/DatePicker";

type ExpenseLine = {
  id: string;
  lineNumber: number;
  expenseDate: string;
  categoryId: number;
  categoryName: string;
  vendor: string;
  description: string;
  amount: number;
  currency: string;
  exchangeRate: number;
};

type ExpenseAttachment = {
  id: string;
  name: string;
  webUrl: string;
  documentType: string;
  expenseLineId: number | null;
  uploadedBy: string;
  uploadedDate: string;
};

type ExpenseReport = {
  id: string;
  subject: string;
  requestDate: string;
  totalAmount: number;
  currency: string;
  status: string;
  currentUser: string;
  currentUserName: string;
  employeeCode: string;
  employeeName: string;
  comments: string;
  lines: ExpenseLine[];
  attachments: ExpenseAttachment[];
};

type ExpenseContext = {
  displayName: string;
  mail: string;
  myReports: ExpenseReport[];
  pendingApprovals: ExpenseReport[];
  isHrAdmin: boolean;
  isFinance: boolean;
};

type ExpenseCategory = {
  id: number;
  name: string;
  nameRo: string;
  code: string;
};

type Employee = {
  mail: string;
  displayName: string;
};

type HistoryEntry = {
  timestamp: string;
  actor: string;
  actorName: string;
  action: string;
  details: string;
};

type LineDraft = {
  // Present when editing an existing line during an Edit & Resubmit (its
  // DE_ExpenseLines item id) — kept lines are PATCHed in place server-side so
  // their existing attachment stays linked; absent for a brand-new line.
  id?: string;
  // The line's already-uploaded attachment, if any, when editing — shown as
  // "Existing: <name>" instead of an empty file picker. Attaching a new file
  // alongside it adds a second attachment rather than replacing it.
  existingAttachment?: ExpenseAttachment;
  expenseDate: string;
  categoryId: string;
  vendor: string;
  description: string;
  amount: string;
  currency: string;
  exchangeRate: string;
  documentType: "Receipt" | "Invoice" | "Other";
  file: File | null;
};

type Tab = "myReports" | "pendingApprovals" | "allReports" | "settings";
type Tone = "good" | "warning" | "critical" | "neutral";

const MANAGER_TOKEN = "{manager}";
const GM_TOKEN = "{GM}";
const DEFAULT_CURRENCY = "RON";
const CURRENCY_OPTIONS = ["RON", "EUR", "USD"];
const MONTH_NAMES = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];
const EXPENSE_STATUSES: { value: string; label: string }[] = [
  { value: "PENDING", label: "Pending" },
  { value: "APPROVED", label: "Approved" },
  { value: "DECLINED", label: "Declined" },
  { value: "CANCELLED", label: "Cancelled" },
];
// Line-level controls get a smaller, denser variant of the standard field
// styling (see fieldClass below) — the 3-row expense-line layout packs a lot
// into a small space, so keeping this text at the normal size felt heavy.
const fieldClassSm =
  "rounded-lg border border-line bg-surface-2 p-1.5 text-xs text-ink focus:outline-none focus:ring-2 focus:ring-accent/40 disabled:cursor-not-allowed disabled:opacity-50";
// Category code for a cash advance (seeded as "ADV" in DE_ExpenseCategory) — an
// advance is just a sum handed over, not a purchase, so the receipt/vendor/
// document-type fields describing what was bought don't apply to it.
const CASH_ADVANCE_CATEGORY_CODE = "ADV";
const btnGhostSm =
  "rounded-lg border border-line px-2 py-1 text-xs font-semibold text-ink-2 hover:bg-surface-2 hover:text-ink disabled:opacity-50";

const TONE_SOFT: Record<Tone, string> = {
  good: "bg-good-soft text-good",
  warning: "bg-warning-soft text-warning",
  critical: "bg-critical-soft text-critical",
  neutral: "bg-neutral-soft text-neutral",
};

function toneForStatus(status: string): Tone {
  if (status === "APPROVED") return "good";
  if (status === "PENDING") return "warning";
  if (status === "DECLINED") return "critical";
  return "neutral";
}

function initials(name: string): string {
  const parts = name.trim().split(/\s+/);
  return ((parts[0]?.[0] ?? "") + (parts[1]?.[0] ?? "")).toUpperCase() || "?";
}

function isCancellable(status: string): boolean {
  return status === "PENDING";
}

function labelForChainEntry(entry: string, nameByMail: Map<string, string>): string {
  if (entry === MANAGER_TOKEN) return "Direct Manager";
  if (entry === GM_TOKEN) return "GM";
  return nameByMail.get(entry.toLowerCase()) ?? entry;
}

const HISTORY_ACTION_LABELS: Record<string, string> = {
  Created: "Created",
  Approved: "Approved",
  ApprovedFinal: "Approved",
  Declined: "Declined",
  Resubmitted: "Resubmitted",
  Cancelled: "Cancelled",
};

function historyActionLabel(action: string): string {
  return HISTORY_ACTION_LABELS[action] ?? action;
}

// A small, self-contained download glyph — no library, no emoji (per house
// style), just an inline SVG so the button reads as an action rather than a
// line of plain text.
function DownloadIcon() {
  return (
    <svg viewBox="0 0 20 20" fill="currentColor" className="h-4 w-4" aria-hidden="true">
      <path d="M10 3a1 1 0 0 1 1 1v6.586l2.293-2.293a1 1 0 1 1 1.414 1.414l-4 4a1 1 0 0 1-1.414 0l-4-4a1 1 0 1 1 1.414-1.414L9 10.586V4a1 1 0 0 1 1-1Z" />
      <path d="M4 14a1 1 0 0 1 1 1v1h10v-1a1 1 0 1 1 2 0v2a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1v-2a1 1 0 0 1 1-1Z" />
    </svg>
  );
}

// Triggers a real browser save-to-disk for a blob/object URL — deliberately
// NOT window.open or an in-page <img>/<iframe>: this app's Teams-desktop
// testing surfaced that the embedded webview there can neither navigate to a
// blob: URL (window.open) nor reliably render a PDF inline (blank iframe, no
// error). A programmatic download only ever asks the browser's own download
// manager to save bytes, which doesn't depend on either of those.
function triggerBrowserDownload(url: string, filename: string) {
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

type ConfirmDialogState = { message: string; onConfirm: () => void } | null;

function formatDate(dateString: string): string {
  if (!dateString) return "";
  const d = new Date(dateString);
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function formatDateTime(dateString: string): string {
  if (!dateString) return "";
  const d = new Date(dateString);
  const hh = String(d.getHours()).padStart(2, "0");
  const min = String(d.getMinutes()).padStart(2, "0");
  return `${formatDate(dateString)} ${hh}:${min}`;
}

function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      // Strip the "data:<mime>;base64," prefix — the API wants raw base64.
      resolve(result.slice(result.indexOf(",") + 1));
    };
    reader.onerror = () => reject(reader.error ?? new Error("Failed to read file"));
    reader.readAsDataURL(file);
  });
}

// A cold-started replica's "web" container (nginx) comes up a few seconds
// before its sibling "api" container (the Azure Functions host) finishes
// loading — see web/nginx.conf's /healthz + the Container App readinessProbe
// this pairs with. That probe keeps most traffic away from that window, but
// a request can still land in it, so the very first data load retries
// through it instead of surfacing a raw 502 to the user.
function isRetryableLoadError(error: unknown): boolean {
  const message = error instanceof Error ? error.message : "";
  return message.includes("Request failed with status 502");
}

async function withRetry<T>(fn: () => Promise<T>, attempts = 5, delayMs = 1500): Promise<T> {
  for (let attempt = 1; attempt <= attempts; attempt++) {
    try {
      return await fn();
    } catch (error) {
      if (attempt === attempts || !isRetryableLoadError(error)) throw error;
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    }
  }
  throw new Error("unreachable");
}

// Small inline spinner — no library, matching the house style of DownloadIcon
// above (a plain SVG, not an emoji).
function Spinner() {
  return (
    <svg viewBox="0 0 24 24" fill="none" className="h-5 w-5 animate-spin text-accent" aria-hidden="true">
      <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="3" className="opacity-25" />
      <path d="M21 12a9 9 0 0 0-9-9" stroke="currentColor" strokeWidth="3" strokeLinecap="round" />
    </svg>
  );
}

function emptyLineDraft(): LineDraft {
  return {
    expenseDate: "",
    categoryId: "",
    vendor: "",
    description: "",
    amount: "",
    currency: DEFAULT_CURRENCY,
    exchangeRate: "",
    documentType: "Receipt",
    file: null,
  };
}

async function apiFetch<T>(path: string, ssoToken: string, init?: RequestInit): Promise<T> {
  if (isPreview) return (await import("@/lib/preview")).previewApiFetch<T>(`/expense${path}`, init);
  const apiBaseUrl = process.env.NEXT_PUBLIC_API_BASE_URL;

  const response = await fetch(`${apiBaseUrl}/expense${path}`, {
    ...init,
    headers: {
      // Note: Azure Static Web Apps' managed API proxy overwrites the standard
      // "Authorization" header for its own edge-to-function auth, so the Teams
      // SSO token is passed via a custom header instead.
      "X-Teams-Sso-Token": ssoToken,
      ...init?.headers,
    },
  });

  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    throw new Error(body.error ?? `Request failed with status ${response.status}`);
  }

  return (await response.json()) as T;
}

// Shared control classes — the "ledger" visual language, matching the sibling
// LeaveRequestForTeams app: surface cards on a slightly cool page plane, a
// single deep-indigo accent for interaction, status conveyed by tone (color +
// label), never color alone.
const fieldClass =
  "rounded-lg border border-line bg-surface-2 p-2 text-sm text-ink focus:outline-none focus:ring-2 focus:ring-accent/40";
const labelClass = "flex flex-col gap-1 text-xs font-semibold text-ink-2";
const cardClass = "rounded-xl border border-line bg-surface p-4";
const btnPrimary = "rounded-lg bg-accent px-3 py-1.5 text-sm font-semibold text-accent-ink hover:brightness-110 disabled:opacity-50";
const btnGhost =
  "rounded-lg border border-line px-3 py-1.5 text-sm font-semibold text-ink-2 hover:bg-surface-2 hover:text-ink disabled:opacity-50";
const btnDanger = "rounded-lg border border-critical/40 px-3 py-1.5 text-sm font-semibold text-critical hover:bg-critical-soft disabled:opacity-50";
const linkBtn = "text-sm font-semibold text-accent hover:underline";

export default function ExpensePage() {
  const [status, setStatus] = useState<"loading" | "ready" | "error">("loading");
  const [context, setContext] = useState<ExpenseContext | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [categories, setCategories] = useState<ExpenseCategory[]>([]);
  const [employees, setEmployees] = useState<Employee[]>([]);
  const nameByMail = useMemo(() => new Map(employees.map((e) => [e.mail.toLowerCase(), e.displayName])), [employees]);
  const categoryOptions = useMemo(() => categories.map((c) => ({ value: String(c.id), label: c.name || c.nameRo })), [categories]);
  const categoryById = useMemo(() => new Map(categories.map((c) => [String(c.id), c])), [categories]);
  const [activeTab, setActiveTab] = useState<Tab>("myReports");

  const [showNewReportForm, setShowNewReportForm] = useState(false);
  const [editingReportId, setEditingReportId] = useState<string | null>(null);
  const [newSubject, setNewSubject] = useState("");
  const [newComments, setNewComments] = useState("");
  const [newLines, setNewLines] = useState<LineDraft[]>([emptyLineDraft()]);
  const [submitting, setSubmitting] = useState(false);
  const [submitError, setSubmitError] = useState<string | null>(null);

  const [respondingId, setRespondingId] = useState<string | null>(null);
  const [respondError, setRespondError] = useState<string | null>(null);
  const [respondComments, setRespondComments] = useState<Record<string, string>>({});

  const [cancellingId, setCancellingId] = useState<string | null>(null);
  const [cancelError, setCancelError] = useState<string | null>(null);

  const [myReportsStatusFilter, setMyReportsStatusFilter] = useState<string[]>(["PENDING", "APPROVED"]); // [] = all statuses

  // Custom confirm modal — NOT window.confirm(): Teams tabs run in a
  // sandboxed iframe that (at least for this app) silently blocks native
  // confirm()/alert()/prompt() dialogs, so window.confirm() just returns
  // false with nothing ever shown on screen — exactly why Cancel appeared to
  // do nothing earlier. This is a real, in-page modal instead.
  const [confirmDialog, setConfirmDialog] = useState<ConfirmDialogState>(null);
  function requestConfirm(message: string, onConfirm: () => void) {
    setConfirmDialog({ message, onConfirm });
  }

  const [showChainEditor, setShowChainEditor] = useState(false);
  const [chainEmployeeSelection, setChainEmployeeSelection] = useState(""); // "" = standard/default chain
  const [chainDraft, setChainDraft] = useState<string[]>([]);
  const [chainIsOverride, setChainIsOverride] = useState(true);
  const [employeesWithOverride, setEmployeesWithOverride] = useState<string[]>([]);
  const [chainAddSelection, setChainAddSelection] = useState("");
  const [chainLoading, setChainLoading] = useState(false);
  const [chainSubmitting, setChainSubmitting] = useState(false);
  const [chainError, setChainError] = useState<string | null>(null);

  const [showFinanceEditor, setShowFinanceEditor] = useState(false);
  const [financeDraft, setFinanceDraft] = useState<string[]>([]);
  const [financeAddSelection, setFinanceAddSelection] = useState("");
  const [financeLoading, setFinanceLoading] = useState(false);
  const [financeSubmitting, setFinanceSubmitting] = useState(false);
  const [financeError, setFinanceError] = useState<string | null>(null);

  const ALL_REPORTS_PAGE_SIZE = 50;
  const [allReports, setAllReports] = useState<ExpenseReport[]>([]);
  const [allReportsLoading, setAllReportsLoading] = useState(false);
  const [allReportsLoadingMore, setAllReportsLoadingMore] = useState(false);
  const [allReportsError, setAllReportsError] = useState<string | null>(null);
  const [allReportsHasMore, setAllReportsHasMore] = useState(false);
  const [allReportsMatchedTotal, setAllReportsMatchedTotal] = useState(0);
  const [allReportsMatchedAmount, setAllReportsMatchedAmount] = useState(0);
  const [allReportsYear, setAllReportsYear] = useState(new Date().getFullYear());
  const [allReportsEmployee, setAllReportsEmployee] = useState(""); // "" = every employee
  const [allReportsMonth, setAllReportsMonth] = useState(""); // "" = every month

  const [refreshing, setRefreshing] = useState(false);
  const [lastRefreshedAt, setLastRefreshedAt] = useState<Date | null>(null);

  async function fetchExpenseContext(): Promise<ExpenseContext> {
    const ssoToken = await getSsoToken();
    return apiFetch<ExpenseContext>("/GetMyExpenseContext", ssoToken);
  }

  async function fetchCategories(): Promise<ExpenseCategory[]> {
    const ssoToken = await getSsoToken();
    const { categories: list } = await apiFetch<{ categories: ExpenseCategory[] }>("/GetExpenseCategories", ssoToken);
    return list;
  }

  async function fetchEmployees(): Promise<Employee[]> {
    const ssoToken = await getSsoToken();
    const { employees: list } = await apiFetch<{ employees: Employee[] }>("/GetPickableEmployees", ssoToken);
    return list;
  }

  useEffect(() => {
    // The home screen's "Awaiting your approval" rows link here with ?tab=pendingApprovals.
    if (new URLSearchParams(window.location.search).get("tab") === "pendingApprovals") {
      setActiveTab("pendingApprovals");
    }
    // The home screen's "New report" button links here with ?new=1 — open the form straight away.
    if (new URLSearchParams(window.location.search).get("new") === "1") {
      setShowNewReportForm(true);
    }

    async function signInAndLoad() {
      try {
        await initTeams();

        try {
          setContext(await withRetry(fetchExpenseContext));
        } catch (error) {
          const message = error instanceof Error ? error.message : "Unknown error";
          if (!message.includes("AADSTS65001") && !message.includes("consent")) {
            throw error;
          }

          // The user hasn't yet consented to a downstream Graph scope our API
          // needs (e.g. Sites.Selected). Prompt for it via a popup, then retry.
          await authentication.authenticate({
            url: `${window.location.origin}/auth-start`,
            width: 600,
            height: 535,
          });

          setContext(await withRetry(fetchExpenseContext));
        }

        setCategories(await withRetry(fetchCategories));
        setEmployees(await withRetry(fetchEmployees));
        setStatus("ready");
      } catch (error) {
        setErrorMessage(error instanceof Error ? error.message : "Unknown error");
        setStatus("error");
      }
    }

    void signInAndLoad();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  async function handleRefresh() {
    setRefreshing(true);
    try {
      setContext(await fetchExpenseContext());
      setLastRefreshedAt(new Date());
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setRefreshing(false);
    }
  }

  // Background auto-refresh every 4 minutes, in addition to the manual button.
  useEffect(() => {
    if (status !== "ready") return;
    const intervalId = setInterval(() => {
      void handleRefresh();
    }, 4 * 60 * 1000);
    return () => clearInterval(intervalId);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [status]);

  function resetNewReportForm() {
    setEditingReportId(null);
    setNewSubject("");
    setNewComments("");
    setNewLines([emptyLineDraft()]);
    setSubmitError(null);
  }

  function openEditForm(report: ExpenseReport) {
    setEditingReportId(report.id);
    setNewSubject(report.subject);
    setNewComments(report.comments);
    setNewLines(
      report.lines.map((line) => ({
        id: line.id,
        existingAttachment: report.attachments.find((a) => a.expenseLineId === Number(line.id)),
        expenseDate: line.expenseDate.slice(0, 10),
        categoryId: String(line.categoryId),
        vendor: line.vendor,
        description: line.description,
        amount: String(line.amount),
        currency: line.currency,
        exchangeRate: line.currency === DEFAULT_CURRENCY ? "" : String(line.exchangeRate),
        documentType: "Receipt",
        file: null,
      }))
    );
    setSubmitError(null);
    setShowNewReportForm(true);
  }

  function updateLine(index: number, patch: Partial<LineDraft>) {
    setNewLines((prev) => prev.map((line, i) => (i === index ? { ...line, ...patch } : line)));
  }

  function addLine() {
    setNewLines((prev) => [...prev, emptyLineDraft()]);
  }

  function removeLine(index: number) {
    setNewLines((prev) => (prev.length > 1 ? prev.filter((_, i) => i !== index) : prev));
  }

  // Amounts are only summed within a single currency — mixing currencies in
  // one report has no honest single "total" without an FX rate, so we group
  // The report's own total is always in DEFAULT_CURRENCY (RON) — a line in a
  // different currency contributes via its own exchange rate, so mixed
  // currencies within one report are fine.
  const newLinesTotal = useMemo(
    () =>
      newLines.reduce((sum, line) => {
        const amt = Number(line.amount) || 0;
        const rate = line.currency === DEFAULT_CURRENCY ? 1 : Number(line.exchangeRate) || 0;
        return sum + amt * rate;
      }, 0),
    [newLines]
  );

  async function handleSubmitNewReport(e: React.FormEvent) {
    e.preventDefault();
    setSubmitError(null);

    if (!newSubject.trim()) {
      setSubmitError("Subject is required.");
      return;
    }
    for (let i = 0; i < newLines.length; i++) {
      const line = newLines[i];
      if (!line.expenseDate || !line.categoryId || !(Number(line.amount) > 0)) {
        setSubmitError(`Line ${i + 1}: fill in the date, category, and a positive amount.`);
        return;
      }
      if (line.currency !== DEFAULT_CURRENCY && !(Number(line.exchangeRate) > 0)) {
        setSubmitError(`Line ${i + 1}: enter the exchange rate to ${DEFAULT_CURRENCY} for this currency.`);
        return;
      }
    }

    setSubmitting(true);
    try {
      const ssoToken = await getSsoToken();
      const created = await apiFetch<{ id: string; lines: { id: string; lineNumber: number }[] }>(
        editingReportId ? "/ResubmitExpenseReport" : "/CreateExpenseReport",
        ssoToken,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            ...(editingReportId ? { id: editingReportId } : {}),
            subject: newSubject,
            comments: newComments || undefined,
            lines: newLines.map((line) => ({
              ...(line.id ? { id: line.id } : {}),
              expenseDate: line.expenseDate,
              categoryId: Number(line.categoryId),
              vendor: line.vendor || undefined,
              description: line.description,
              amount: Number(line.amount),
              currency: line.currency,
              exchangeRate: line.currency === DEFAULT_CURRENCY ? 1 : Number(line.exchangeRate),
            })),
          }),
        }
      );

      // Upload each line's attached file (if any) now that we have real line ids.
      const uploadErrors: string[] = [];
      for (let i = 0; i < newLines.length; i++) {
        const draft = newLines[i];
        if (!draft.file) continue;
        const createdLine = created.lines[i];
        try {
          const fileBase64 = await fileToBase64(draft.file);
          const uploadSsoToken = await getSsoToken();
          await apiFetch("/UploadExpenseAttachment", uploadSsoToken, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              expenseReportId: Number(created.id),
              expenseLineId: createdLine ? Number(createdLine.id) : undefined,
              documentType: draft.documentType,
              fileName: draft.file.name,
              contentType: draft.file.type,
              fileBase64,
            }),
          });
        } catch (uploadError) {
          // The report itself was created successfully — don't roll anything
          // back; the user can retry the upload later from the report's own
          // detail view. Collected rather than set immediately: resetting the
          // form below (on the success path) also clears submitError, which
          // would silently wipe out exactly this message before anyone could
          // read it — so a failure here must skip that reset instead.
          uploadErrors.push(`Line ${i + 1}: ${uploadError instanceof Error ? uploadError.message : "unknown error"}`);
        }
      }

      await handleRefresh();
      if (uploadErrors.length > 0) {
        setSubmitError(
          `The report was ${editingReportId ? "resubmitted" : "created"}, but ${uploadErrors.length} file upload${uploadErrors.length === 1 ? "" : "s"} failed — ${uploadErrors.join("; ")}`
        );
      } else {
        resetNewReportForm();
        setShowNewReportForm(false);
      }
    } catch (error) {
      setSubmitError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setSubmitting(false);
    }
  }

  async function handleRespond(id: string, decision: "approved" | "declined") {
    setRespondingId(id);
    setRespondError(null);
    try {
      const ssoToken = await getSsoToken();
      await apiFetch("/RespondToExpenseReport", ssoToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id, decision, comment: respondComments[id] || undefined }),
      });
      await handleRefresh();
    } catch (error) {
      setRespondError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setRespondingId(null);
    }
  }

  async function handleCancel(id: string) {
    setCancellingId(id);
    setCancelError(null);
    try {
      const ssoToken = await getSsoToken();
      await apiFetch("/CancelExpenseReport", ssoToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id }),
      });
      await handleRefresh();
    } catch (error) {
      setCancelError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setCancellingId(null);
    }
  }

  async function loadChainFor(employeeCode: string) {
    setChainLoading(true);
    setChainError(null);
    try {
      const ssoToken = await getSsoToken();
      const path = employeeCode ? `/GetApprovalFlow?employeeCode=${encodeURIComponent(employeeCode)}` : "/GetApprovalFlow";
      const { chain, isOverride, employeesWithOverrides } = await apiFetch<{
        chain: string[];
        isOverride: boolean;
        employeesWithOverrides: string[];
      }>(path, ssoToken);
      setChainDraft(chain);
      setChainIsOverride(isOverride);
      setEmployeesWithOverride(employeesWithOverrides);
    } catch (error) {
      setChainError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setChainLoading(false);
    }
  }

  function handleOpenChainEditor() {
    setShowChainEditor(true);
    setChainEmployeeSelection("");
    void loadChainFor("");
  }

  async function handleSaveChain() {
    setChainSubmitting(true);
    setChainError(null);
    try {
      const ssoToken = await getSsoToken();
      const { chain, isOverride } = await apiFetch<{ chain: string[]; isOverride: boolean }>("/SetApprovalFlow", ssoToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ chain: chainDraft, employeeCode: chainEmployeeSelection || undefined }),
      });
      setChainDraft(chain);
      setChainIsOverride(isOverride);
    } catch (error) {
      setChainError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setChainSubmitting(false);
    }
  }

  async function loadFinanceUsers() {
    setFinanceLoading(true);
    setFinanceError(null);
    try {
      const ssoToken = await getSsoToken();
      const { emails } = await apiFetch<{ emails: string[] }>("/GetFinanceUsers", ssoToken);
      setFinanceDraft(emails);
    } catch (error) {
      setFinanceError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setFinanceLoading(false);
    }
  }

  function handleOpenFinanceEditor() {
    setShowFinanceEditor(true);
    void loadFinanceUsers();
  }

  async function handleSaveFinance() {
    setFinanceSubmitting(true);
    setFinanceError(null);
    try {
      const ssoToken = await getSsoToken();
      const { emails } = await apiFetch<{ emails: string[] }>("/SetFinanceUsers", ssoToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ emails: financeDraft }),
      });
      setFinanceDraft(emails);
    } catch (error) {
      setFinanceError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setFinanceSubmitting(false);
    }
  }

  async function loadAllReports(reset: boolean) {
    reset ? setAllReportsLoading(true) : setAllReportsLoadingMore(true);
    setAllReportsError(null);
    try {
      const ssoToken = await getSsoToken();
      const params = new URLSearchParams({
        year: String(allReportsYear),
        offset: String(reset ? 0 : allReports.length),
        limit: String(ALL_REPORTS_PAGE_SIZE),
      });
      if (allReportsEmployee) params.set("employeeCode", allReportsEmployee);
      if (allReportsMonth) params.set("month", allReportsMonth);
      const { reports, total, totalAmount, hasMore } = await apiFetch<{
        year: number;
        reports: ExpenseReport[];
        total: number;
        totalAmount: number;
        hasMore: boolean;
      }>(`/GetAllExpenseReports?${params}`, ssoToken);
      setAllReports((prev) => (reset ? reports : [...prev, ...reports]));
      setAllReportsMatchedTotal(total);
      setAllReportsMatchedAmount(totalAmount);
      setAllReportsHasMore(hasMore);
    } catch (error) {
      setAllReportsError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setAllReportsLoading(false);
      setAllReportsLoadingMore(false);
    }
  }

  // Reload from the first page whenever the tab is opened or a filter changes.
  useEffect(() => {
    if (activeTab !== "allReports") return;
    void loadAllReports(true);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeTab, allReportsYear, allReportsEmployee, allReportsMonth]);

  if (status === "loading") {
    return (
      <main className="flex min-h-screen flex-col items-center justify-center gap-3 bg-page p-6 font-sans text-ink sm:p-8">
        <Spinner />
        <p className="text-sm text-ink-2">Loading…</p>
      </main>
    );
  }

  if (status === "error") {
    return (
      <main className="flex min-h-screen flex-col items-center justify-center gap-5 bg-page p-6 font-sans text-ink sm:p-8">
        <div className={`${cardClass} max-w-md`}>
          <p className="text-sm font-semibold text-critical">Loading error</p>
          <p className="mt-1 text-sm text-ink-2">{errorMessage}</p>
        </div>
      </main>
    );
  }

  if (!context) return null;

  return (
    <main className="flex min-h-screen flex-col items-center justify-start gap-5 bg-page p-6 font-sans text-ink sm:p-8">
      <div className="flex w-full max-w-4xl flex-col gap-4">
        <header className="flex items-center justify-between">
          <div>
            <h1 className="text-lg font-bold text-ink">Expense Reports</h1>
            <p className="text-xs text-ink-2">{context.displayName}</p>
          </div>
          <div className="flex items-center gap-2">
            {lastRefreshedAt && (
              <span className="text-[11px] text-ink-3">Updated {lastRefreshedAt.toLocaleTimeString()}</span>
            )}
            <button type="button" className={btnGhost} onClick={handleRefresh} disabled={refreshing}>
              {refreshing ? "Refreshing…" : "Refresh"}
            </button>
            <button
              type="button"
              className={btnPrimary}
              onClick={() => {
                resetNewReportForm();
                setShowNewReportForm((v) => !v);
              }}
            >
              {showNewReportForm ? "Cancel" : "New Report"}
            </button>
          </div>
        </header>

        {showNewReportForm && (
          <form onSubmit={handleSubmitNewReport} className={`${cardClass} flex flex-col gap-3`}>
            <h2 className="text-sm font-bold text-ink">{editingReportId ? "Edit Expense Report" : "New Expense Report"}</h2>
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
              <label className={labelClass}>
                Subject
                <input
                  type="text"
                  required
                  value={newSubject}
                  onChange={(e) => setNewSubject(e.target.value)}
                  className={fieldClass}
                  placeholder="e.g. Client visit, March 12-14"
                />
              </label>
              <label className={labelClass}>
                Comments (optional)
                <input type="text" value={newComments} onChange={(e) => setNewComments(e.target.value)} className={fieldClass} />
              </label>
            </div>

            <div className="flex flex-col gap-2">
              <div className="flex items-center justify-between">
                <h3 className="text-xs font-semibold text-ink-2">Expense lines</h3>
                <span className="text-xs font-semibold text-ink-2">
                  Total: {newLinesTotal.toFixed(2)} {DEFAULT_CURRENCY}
                </span>
              </div>
              {newLines.map((line, index) => {
                const isCashAdvance = categoryById.get(line.categoryId)?.code === CASH_ADVANCE_CATEGORY_CODE;
                return (
                <div key={index} className="flex flex-col gap-1.5 rounded-lg border border-line p-2">
                  {/* Row 1: date, category, document type */}
                  <div className="grid grid-cols-1 gap-1.5 sm:grid-cols-3">
                    <label className={labelClass}>
                      Date
                      <DatePicker value={line.expenseDate} onChange={(v) => updateLine(index, { expenseDate: v })} className={fieldClassSm} required />
                    </label>
                    <label className={labelClass}>
                      Category
                      <SearchableSelect
                        options={categoryOptions}
                        value={line.categoryId}
                        onChange={(v) => {
                          const patch: Partial<LineDraft> = { categoryId: v };
                          // A cash advance isn't a purchase — clear the fields that
                          // describe one so stale values from a prior category don't
                          // get submitted alongside it while they're disabled.
                          if (categoryById.get(v)?.code === CASH_ADVANCE_CATEGORY_CODE) {
                            patch.documentType = "Other";
                            patch.vendor = "";
                            patch.description = "";
                            patch.file = null;
                          }
                          updateLine(index, patch);
                        }}
                        placeholder="Choose…"
                        className={fieldClassSm}
                        required
                      />
                    </label>
                    <label className={labelClass}>
                      Document type
                      <select
                        value={line.documentType}
                        onChange={(e) => updateLine(index, { documentType: e.target.value as LineDraft["documentType"] })}
                        className={fieldClassSm}
                        disabled={isCashAdvance}
                      >
                        <option value="Receipt">Receipt</option>
                        <option value="Invoice">Invoice</option>
                        <option value="Other">Other</option>
                      </select>
                    </label>
                  </div>

                  {/* Row 2: vendor, amount, currency, exchange rate (only for a non-RON currency) */}
                  <div className={`grid grid-cols-1 gap-1.5 ${line.currency === DEFAULT_CURRENCY ? "sm:grid-cols-3" : "sm:grid-cols-4"}`}>
                    <label className={labelClass}>
                      Vendor
                      <input
                        type="text"
                        value={line.vendor}
                        onChange={(e) => updateLine(index, { vendor: e.target.value })}
                        className={fieldClassSm}
                        disabled={isCashAdvance}
                      />
                    </label>
                    <label className={labelClass}>
                      Amount
                      <input
                        type="number"
                        min="0.01"
                        step="0.01"
                        required
                        value={line.amount}
                        onChange={(e) => updateLine(index, { amount: e.target.value })}
                        className={fieldClassSm}
                      />
                    </label>
                    <label className={labelClass}>
                      Currency
                      <select
                        value={line.currency}
                        onChange={(e) => updateLine(index, { currency: e.target.value })}
                        className={fieldClassSm}
                      >
                        {CURRENCY_OPTIONS.map((c) => (
                          <option key={c} value={c}>
                            {c}
                          </option>
                        ))}
                      </select>
                    </label>
                    {line.currency !== DEFAULT_CURRENCY && (
                      <label className={labelClass}>
                        Rate to {DEFAULT_CURRENCY}
                        <input
                          type="number"
                          min="0.0001"
                          step="0.0001"
                          required
                          value={line.exchangeRate}
                          onChange={(e) => updateLine(index, { exchangeRate: e.target.value })}
                          placeholder="e.g. 5.07"
                          className={fieldClassSm}
                        />
                      </label>
                    )}
                  </div>

                  {/* Row 3: description (generous width), add file, remove line */}
                  <div className="grid grid-cols-1 gap-1.5 sm:grid-cols-6">
                    <label className={`${labelClass} sm:col-span-3`}>
                      Description (optional)
                      <input
                        type="text"
                        value={line.description}
                        onChange={(e) => updateLine(index, { description: e.target.value })}
                        className={fieldClassSm}
                        disabled={isCashAdvance}
                      />
                    </label>
                    <div className={`${labelClass} sm:col-span-2`}>
                      <span>Receipt/invoice (optional, max 4 MB)</span>
                      <div className="flex items-center gap-2">
                        <label
                          className={`${btnGhostSm} inline-flex w-fit items-center ${isCashAdvance ? "cursor-not-allowed opacity-50" : "cursor-pointer"}`}
                        >
                          {line.existingAttachment ? "Add another file" : "Add file"}
                          <input
                            type="file"
                            // MIME type, not the ".pdf" extension — some embedded browser
                            // shells (Teams desktop's webview included) filter file pickers
                            // unreliably when an accept list mixes a wildcard MIME type with
                            // an extension, hiding PDFs from the picker entirely.
                            accept="image/*,application/pdf"
                            onChange={(e) => updateLine(index, { file: e.target.files?.[0] ?? null })}
                            className="hidden"
                            disabled={isCashAdvance}
                          />
                        </label>
                        <span className="truncate text-xs text-ink-2">
                          {line.file
                            ? line.file.name
                            : line.existingAttachment
                              ? `Existing: ${line.existingAttachment.name}`
                              : "No file selected"}
                        </span>
                        {line.file && (
                          <button
                            type="button"
                            className={`${linkBtn} inline-flex items-center gap-1`}
                            title="Download to review before submitting"
                            onClick={() => {
                              // A local, not-yet-uploaded File — createObjectURL needs no server
                              // round-trip. Triggers a real save-to-disk rather than an in-page
                              // preview or window.open, neither of which this app's Teams-desktop
                              // testing found reliable (see triggerBrowserDownload's own comment).
                              if (line.file) triggerBrowserDownload(URL.createObjectURL(line.file), line.file.name);
                            }}
                          >
                            <DownloadIcon />
                          </button>
                        )}
                      </div>
                    </div>
                    <div className="flex items-end sm:col-span-1">
                      <button type="button" className={btnGhostSm} onClick={() => removeLine(index)} disabled={newLines.length === 1}>
                        Remove line
                      </button>
                    </div>
                  </div>
                </div>
                );
              })}
              <button type="button" className={linkBtn} onClick={addLine}>
                + Add line
              </button>
            </div>

            {submitError && <p className="text-sm text-critical">{submitError}</p>}
            <div className="flex justify-end gap-2">
              <button type="submit" className={btnPrimary} disabled={submitting}>
                {submitting ? "Submitting…" : editingReportId ? "Resubmit report" : "Submit report"}
              </button>
            </div>
          </form>
        )}

        {!showNewReportForm && (
        <>
        <nav className="flex flex-wrap gap-2">
          <button
            type="button"
            className={`rounded-lg px-3 py-1.5 text-sm font-semibold ${activeTab === "myReports" ? "bg-accent text-accent-ink" : "bg-surface-2 text-ink-2"}`}
            onClick={() => setActiveTab("myReports")}
          >
            My Reports ({context.myReports.length})
          </button>
          <button
            type="button"
            className={`rounded-lg px-3 py-1.5 text-sm font-semibold ${activeTab === "pendingApprovals" ? "bg-accent text-accent-ink" : "bg-surface-2 text-ink-2"}`}
            onClick={() => setActiveTab("pendingApprovals")}
          >
            Pending Approvals ({context.pendingApprovals.length})
          </button>
          {(context.isHrAdmin || context.isFinance) && (
            <button
              type="button"
              className={`rounded-lg px-3 py-1.5 text-sm font-semibold ${activeTab === "allReports" ? "bg-accent text-accent-ink" : "bg-surface-2 text-ink-2"}`}
              onClick={() => setActiveTab("allReports")}
            >
              Reports
            </button>
          )}
          {context.isHrAdmin && (
            <button
              type="button"
              className={`rounded-lg px-3 py-1.5 text-sm font-semibold ${activeTab === "settings" ? "bg-accent text-accent-ink" : "bg-surface-2 text-ink-2"}`}
              onClick={() => setActiveTab("settings")}
            >
              Settings
            </button>
          )}
        </nav>

        {activeTab === "myReports" && (
          <section className="flex flex-col gap-3">
            <div className={labelClass}>
              Status (select one or more; none selected = all)
              <div className="flex flex-wrap gap-2">
                {EXPENSE_STATUSES.map((s) => (
                  <button
                    key={s.value}
                    type="button"
                    className={`rounded-lg px-3 py-1.5 text-sm font-semibold ${
                      myReportsStatusFilter.includes(s.value) ? "bg-accent text-accent-ink" : "bg-surface-2 text-ink-2"
                    }`}
                    onClick={() =>
                      setMyReportsStatusFilter((prev) =>
                        prev.includes(s.value) ? prev.filter((v) => v !== s.value) : [...prev, s.value]
                      )
                    }
                  >
                    {s.label}
                  </button>
                ))}
                {myReportsStatusFilter.length > 0 && (
                  <button type="button" className={btnGhost} onClick={() => setMyReportsStatusFilter([])}>
                    Clear
                  </button>
                )}
              </div>
            </div>
            {(() => {
              const filtered =
                myReportsStatusFilter.length > 0
                  ? context.myReports.filter((r) => myReportsStatusFilter.includes(r.status))
                  : context.myReports;
              return (
                <>
                  {filtered.length === 0 && (
                    <p className="text-sm text-ink-2">
                      {context.myReports.length === 0 ? "You don't have any expense reports yet." : "No reports match this status."}
                    </p>
                  )}
                  {filtered.map((r) => (
                    <ExpenseReportCard
                      key={r.id}
                      report={r}
                      showCancel={isCancellable(r.status)}
                      onCancel={() => handleCancel(r.id)}
                      cancelling={cancellingId === r.id}
                      showEdit={r.status === "DECLINED"}
                      onEdit={() => openEditForm(r)}
                      requestConfirm={requestConfirm}
                    />
                  ))}
                </>
              );
            })()}
            {cancelError && <p className="text-sm text-critical">{cancelError}</p>}
          </section>
        )}

        {activeTab === "pendingApprovals" && (
          <section className="flex flex-col gap-3">
            {context.pendingApprovals.length === 0 && <p className="text-sm text-ink-2">No reports awaiting your approval.</p>}
            {context.pendingApprovals.map((r) => (
              <div key={r.id} className={cardClass}>
                <ExpenseReportCard
                  report={r}
                  showCancel={false}
                  onCancel={() => {}}
                  cancelling={false}
                  showEdit={false}
                  onEdit={() => {}}
                  requestConfirm={requestConfirm}
                />
                <div className="mt-3 flex flex-col gap-2 border-t border-line pt-3">
                  <input
                    type="text"
                    placeholder="Comment (optional)"
                    value={respondComments[r.id] ?? ""}
                    onChange={(e) => setRespondComments((prev) => ({ ...prev, [r.id]: e.target.value }))}
                    className={fieldClass}
                  />
                  <div className="flex gap-2">
                    <button type="button" className={btnPrimary} disabled={respondingId === r.id} onClick={() => handleRespond(r.id, "approved")}>
                      Approve
                    </button>
                    <button
                      type="button"
                      className={btnDanger}
                      disabled={respondingId === r.id}
                      onClick={() => requestConfirm("Decline this expense report?", () => handleRespond(r.id, "declined"))}
                    >
                      Decline
                    </button>
                  </div>
                </div>
              </div>
            ))}
            {respondError && <p className="text-sm text-critical">{respondError}</p>}
          </section>
        )}

        {activeTab === "allReports" && (context.isHrAdmin || context.isFinance) && (
          <section className="flex flex-col gap-3">
            <div className={`${cardClass} flex flex-wrap items-end gap-3`}>
              <label className={labelClass}>
                Year
                <select value={allReportsYear} onChange={(e) => setAllReportsYear(Number(e.target.value))} className={fieldClass}>
                  {Array.from({ length: 5 }, (_, i) => new Date().getFullYear() - i).map((y) => (
                    <option key={y} value={y}>
                      {y}
                    </option>
                  ))}
                </select>
              </label>
              <label className={labelClass}>
                Month
                <select value={allReportsMonth} onChange={(e) => setAllReportsMonth(e.target.value)} className={fieldClass}>
                  <option value="">All months</option>
                  {MONTH_NAMES.map((name, i) => (
                    <option key={name} value={i + 1}>
                      {name}
                    </option>
                  ))}
                </select>
              </label>
              <label className={`${labelClass} min-w-[14rem] flex-1`}>
                Employee
                <SearchableSelect
                  options={[{ value: "", label: "All employees" }, ...employees.map((e) => ({ value: e.mail, label: e.displayName }))]}
                  value={allReportsEmployee}
                  onChange={setAllReportsEmployee}
                  placeholder="Search employee…"
                  className={fieldClass}
                />
              </label>
              <span className="text-sm font-semibold text-ink-2">
                Total: {allReportsMatchedAmount.toFixed(2)} {DEFAULT_CURRENCY} ({allReportsMatchedTotal} report
                {allReportsMatchedTotal === 1 ? "" : "s"})
              </span>
            </div>

            {allReportsLoading && <p className="text-sm text-ink-2">Loading…</p>}
            {allReportsError && <p className="text-sm text-critical">{allReportsError}</p>}
            {!allReportsLoading && allReports.length === 0 && (
              <p className="text-sm text-ink-2">No reports match these filters.</p>
            )}
            {allReports.map((r) => (
              <ExpenseReportCard
                key={r.id}
                report={r}
                showCancel={false}
                onCancel={() => {}}
                cancelling={false}
                showEdit={false}
                onEdit={() => {}}
                requestConfirm={requestConfirm}
              />
            ))}
            {allReportsHasMore && (
              <button
                type="button"
                className={btnGhost}
                onClick={() => void loadAllReports(false)}
                disabled={allReportsLoadingMore}
              >
                {allReportsLoadingMore ? "Loading…" : "Load more"}
              </button>
            )}
          </section>
        )}

        {activeTab === "settings" && context.isHrAdmin && (
          <>
          <section className={`${cardClass} flex flex-col gap-3`}>
            <div className="flex items-center justify-between">
              <h2 className="text-sm font-bold text-ink">Approval chain</h2>
              {!showChainEditor && (
                <button type="button" className={btnGhost} onClick={handleOpenChainEditor}>
                  Edit
                </button>
              )}
            </div>
            {showChainEditor && (
              <div className="flex flex-col gap-2">
                <label className={labelClass}>
                  Employee (blank = standard chain)
                  <SearchableSelect
                    options={[{ value: "", label: "— Standard chain —" }, ...employees.map((e) => ({ value: e.mail, label: e.displayName }))]}
                    value={chainEmployeeSelection}
                    onChange={(v) => {
                      setChainEmployeeSelection(v);
                      void loadChainFor(v);
                    }}
                    placeholder="Search employee…"
                    className={fieldClass}
                  />
                </label>
                {chainLoading ? (
                  <p className="text-sm text-ink-2">Loading…</p>
                ) : (
                  <>
                    <p className="text-xs text-ink-3">
                      {chainIsOverride ? "Custom chain for this employee." : "Inherits the standard chain."}
                    </p>
                    <ul className="flex flex-col gap-1">
                      {chainDraft.map((entry, i) => (
                        <li key={i} className="flex items-center justify-between rounded-lg border border-line bg-surface-2 px-3 py-1.5 text-sm">
                          <span>
                            {i + 1}. {labelForChainEntry(entry, nameByMail)}
                          </span>
                          <button
                            type="button"
                            className="text-xs font-semibold text-critical hover:underline"
                            onClick={() => setChainDraft((prev) => prev.filter((_, idx) => idx !== i))}
                          >
                            Remove
                          </button>
                        </li>
                      ))}
                      {chainDraft.length === 0 && <li className="text-sm text-ink-3">No approvers.</li>}
                    </ul>
                    <div className="flex gap-2">
                      <SearchableSelect
                        options={employees.map((e) => ({ value: e.mail, label: e.displayName }))}
                        value={chainAddSelection}
                        onChange={setChainAddSelection}
                        placeholder="Add approver…"
                        className={fieldClass}
                        wrapperClassName="flex-1"
                      />
                      <button
                        type="button"
                        className={btnGhost}
                        disabled={!chainAddSelection}
                        onClick={() => {
                          if (chainAddSelection && !chainDraft.includes(chainAddSelection)) {
                            setChainDraft((prev) => [...prev, chainAddSelection]);
                          }
                          setChainAddSelection("");
                        }}
                      >
                        Add
                      </button>
                    </div>
                    {chainError && <p className="text-sm text-critical">{chainError}</p>}
                    <div className="flex justify-end gap-2">
                      <button type="button" className={btnGhost} onClick={() => setShowChainEditor(false)}>
                        Close
                      </button>
                      <button type="button" className={btnPrimary} disabled={chainSubmitting} onClick={handleSaveChain}>
                        {chainSubmitting ? "Saving…" : "Save"}
                      </button>
                    </div>
                  </>
                )}
              </div>
            )}
          </section>

          <section className={`${cardClass} flex flex-col gap-3`}>
            <div className="flex items-center justify-between">
              <div>
                <h2 className="text-sm font-bold text-ink">Finance access</h2>
                <p className="text-xs text-ink-2">These people have full visibility into every expense report, including receipts/invoices.</p>
              </div>
              {!showFinanceEditor && (
                <button type="button" className={btnGhost} onClick={handleOpenFinanceEditor}>
                  Edit
                </button>
              )}
            </div>
            {showFinanceEditor && (
              <div className="flex flex-col gap-2">
                {financeLoading ? (
                  <p className="text-sm text-ink-2">Loading…</p>
                ) : (
                  <>
                    <ul className="flex flex-col gap-1">
                      {financeDraft.map((email, i) => (
                        <li key={email} className="flex items-center justify-between rounded-lg border border-line bg-surface-2 px-3 py-1.5 text-sm">
                          <span>{nameByMail.get(email.toLowerCase()) ?? email}</span>
                          <button
                            type="button"
                            className="text-xs font-semibold text-critical hover:underline"
                            onClick={() => setFinanceDraft((prev) => prev.filter((_, idx) => idx !== i))}
                          >
                            Remove
                          </button>
                        </li>
                      ))}
                      {financeDraft.length === 0 && <li className="text-sm text-ink-3">No one has Finance access yet.</li>}
                    </ul>
                    <div className="flex gap-2">
                      <SearchableSelect
                        options={employees.map((e) => ({ value: e.mail, label: e.displayName }))}
                        value={financeAddSelection}
                        onChange={setFinanceAddSelection}
                        placeholder="Add person…"
                        className={fieldClass}
                        wrapperClassName="flex-1"
                      />
                      <button
                        type="button"
                        className={btnGhost}
                        disabled={!financeAddSelection}
                        onClick={() => {
                          if (financeAddSelection && !financeDraft.includes(financeAddSelection)) {
                            setFinanceDraft((prev) => [...prev, financeAddSelection]);
                          }
                          setFinanceAddSelection("");
                        }}
                      >
                        Add
                      </button>
                    </div>
                    {financeError && <p className="text-sm text-critical">{financeError}</p>}
                    <div className="flex justify-end gap-2">
                      <button type="button" className={btnGhost} onClick={() => setShowFinanceEditor(false)}>
                        Close
                      </button>
                      <button type="button" className={btnPrimary} disabled={financeSubmitting} onClick={handleSaveFinance}>
                        {financeSubmitting ? "Saving…" : "Save"}
                      </button>
                    </div>
                  </>
                )}
              </div>
            )}
          </section>
          </>
        )}
        </>
        )}
      </div>

      {confirmDialog && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4">
          <div className={`${cardClass} w-full max-w-sm`}>
            <p className="text-sm text-ink">{confirmDialog.message}</p>
            <div className="mt-4 flex justify-end gap-2">
              <button type="button" className={btnGhost} onClick={() => setConfirmDialog(null)}>
                Cancel
              </button>
              <button
                type="button"
                className={btnDanger}
                onClick={() => {
                  confirmDialog.onConfirm();
                  setConfirmDialog(null);
                }}
              >
                Confirm
              </button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}

function ExpenseReportCard({
  report,
  showCancel,
  onCancel,
  cancelling,
  showEdit,
  onEdit,
  requestConfirm,
}: {
  report: ExpenseReport;
  showCancel: boolean;
  onCancel: () => void;
  cancelling: boolean;
  showEdit: boolean;
  onEdit: () => void;
  requestConfirm: (message: string, onConfirm: () => void) => void;
}) {
  const tone = toneForStatus(report.status);
  const [showHistory, setShowHistory] = useState(false);
  const [history, setHistory] = useState<HistoryEntry[] | null>(null);
  const [historyLoading, setHistoryLoading] = useState(false);
  const [historyError, setHistoryError] = useState<string | null>(null);
  const [openingAttachmentId, setOpeningAttachmentId] = useState<string | null>(null);
  const [attachmentError, setAttachmentError] = useState<string | null>(null);

  // Fetched (not a plain <a href={webUrl}>) because the viewer may have no
  // direct SharePoint permission on AVI_Site itself — this app otherwise
  // never asks anyone to browse SharePoint directly for the download itself,
  // only to call Graph via the OBO-exchanged token, which this same fetch
  // reuses. No in-page preview attempt anymore (dropped both window.open and
  // an <img>/<iframe> modal — this app's own Teams-desktop testing found
  // neither reliable there: window.open can't navigate to a blob: URL at
  // all, and an iframe rendered a PDF blank with no error). Instead: a real
  // browser download (works regardless of any inline viewer), plus a
  // best-effort microsoftTeams.app.openLink on the SharePoint webUrl to also
  // try opening it in the user's actual default browser — that call is a
  // silent no-op if it fails, never blocking the download that already
  // succeeded.
  async function handleOpenAttachment(attachment: ExpenseAttachment) {
    setOpeningAttachmentId(attachment.id);
    setAttachmentError(null);
    try {
      const ssoToken = await getSsoToken();
      const apiBaseUrl = process.env.NEXT_PUBLIC_API_BASE_URL;
      const response = await fetch(`${apiBaseUrl}/expense/GetExpenseAttachment?id=${encodeURIComponent(attachment.id)}`, {
        headers: { "X-Teams-Sso-Token": ssoToken },
      });
      if (!response.ok) {
        const body = await response.json().catch(() => ({}));
        throw new Error(body.error ?? `Request failed with status ${response.status}`);
      }
      const blob = await response.blob();
      const blobUrl = URL.createObjectURL(blob);
      triggerBrowserDownload(blobUrl, attachment.name);
      if (attachment.webUrl) {
        app.openLink(attachment.webUrl).catch(() => {
          // Best-effort — the download above already succeeded regardless.
        });
      }
    } catch (error) {
      setAttachmentError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setOpeningAttachmentId(null);
    }
  }

  async function handleToggleHistory() {
    const next = !showHistory;
    setShowHistory(next);
    if (!next || history !== null) return;
    setHistoryLoading(true);
    setHistoryError(null);
    try {
      const ssoToken = await getSsoToken();
      const { history: entries } = await apiFetch<{ history: HistoryEntry[] }>(
        `/GetExpenseReportHistory?id=${encodeURIComponent(report.id)}`,
        ssoToken
      );
      setHistory(entries);
    } catch (error) {
      setHistoryError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setHistoryLoading(false);
    }
  }

  return (
    <div className={cardClass}>
      <div className="flex items-start justify-between gap-3">
        <div className="flex items-center gap-2">
          <span className="flex h-8 w-8 items-center justify-center rounded-full bg-accent-soft text-xs font-bold text-accent">
            {initials(report.employeeName || report.employeeCode)}
          </span>
          <div>
            <p className="text-sm font-semibold text-ink">{report.subject}</p>
            <p className="text-xs text-ink-2">
              {report.employeeName || report.employeeCode} · {formatDate(report.requestDate)}
            </p>
          </div>
        </div>
        <span className={`rounded-full px-2 py-0.5 text-xs font-semibold ${TONE_SOFT[tone]}`}>{report.status}</span>
      </div>

      <div className="mt-3 overflow-x-auto">
        <table className="w-full text-xs">
          <thead>
            <tr className="text-left text-ink-3">
              <th className="pb-1 pr-2 font-semibold">Date</th>
              <th className="pb-1 pr-2 font-semibold">Category</th>
              <th className="pb-1 pr-2 font-semibold">Vendor</th>
              <th className="pb-1 pr-2 font-semibold">Description</th>
              <th className="pb-1 pr-2 text-right font-semibold">Amount</th>
              <th className="pb-1 font-semibold">Document</th>
            </tr>
          </thead>
          <tbody>
            {report.lines.map((line) => {
              const lineAttachments = report.attachments.filter((a) => a.expenseLineId === Number(line.id));
              return (
                <tr key={line.id} className="border-t border-line">
                  <td className="py-1 pr-2 text-ink-2">{formatDate(line.expenseDate)}</td>
                  <td className="py-1 pr-2 text-ink-2">{line.categoryName}</td>
                  <td className="py-1 pr-2 text-ink-2">{line.vendor}</td>
                  <td className="py-1 pr-2 text-ink">{line.description}</td>
                  <td className="py-1 pr-2 text-right font-semibold text-ink">
                    {line.amount.toFixed(2)} {line.currency}
                  </td>
                  <td className="py-1">
                    {lineAttachments.length > 0 ? (
                      <div className="flex items-center gap-1.5">
                        {lineAttachments.map((attachment) => (
                          <button
                            key={attachment.id}
                            type="button"
                            className={`${linkBtn} inline-flex items-center gap-1`}
                            disabled={openingAttachmentId === attachment.id}
                            title={`Download ${attachment.name} (and try opening in your default browser)`}
                            onClick={() => handleOpenAttachment(attachment)}
                          >
                            <DownloadIcon />
                            {openingAttachmentId === attachment.id && <span>…</span>}
                          </button>
                        ))}
                      </div>
                    ) : (
                      <span className="text-ink-3">—</span>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {attachmentError && <p className="mt-1 text-xs text-critical">{attachmentError}</p>}

      <div className="mt-3 flex flex-wrap items-center justify-between gap-2 border-t border-line pt-2 text-xs text-ink-2">
        <div className="flex flex-wrap items-center gap-3">
          <span className="font-semibold text-ink">
            Total: {report.totalAmount.toFixed(2)} {report.currency}
          </span>
          {report.status === "PENDING" && report.currentUserName && <span>Awaiting: {report.currentUserName}</span>}
        </div>
        {/* Cancel (when shown) always sits left of History, so History is
            reliably the rightmost action regardless of which other buttons
            a given tab renders for this card. */}
        <div className="flex items-center gap-2">
          {showEdit && (
            <button type="button" className={btnGhost} onClick={onEdit}>
              Edit &amp; Resubmit
            </button>
          )}
          {showCancel && (
            <button
              type="button"
              className={btnGhost}
              disabled={cancelling}
              onClick={() => requestConfirm("Cancel this expense report?", () => onCancel())}
            >
              {cancelling ? "Cancelling…" : "Cancel"}
            </button>
          )}
          <button type="button" className={btnGhost} onClick={handleToggleHistory}>
            {showHistory ? "Hide history" : "History"}
          </button>
        </div>
      </div>

      {showHistory && (
        <div className="mt-2 flex flex-col gap-1 border-t border-line pt-2 text-xs text-ink-2">
          {historyLoading && <p>Loading…</p>}
          {historyError && <p className="text-critical">{historyError}</p>}
          {history && history.length === 0 && <p>No history yet.</p>}
          {history?.map((h, i) => (
            <p key={i}>
              <span className="font-semibold text-ink">{formatDateTime(h.timestamp)}</span> — {historyActionLabel(h.action)} by {h.actorName}
              {h.details ? `. Comment: ${h.details}` : ""}
            </p>
          ))}
        </div>
      )}
    </div>
  );
}
