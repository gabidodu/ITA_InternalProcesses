"use client";

import { useEffect, useMemo, useState } from "react";
import { authentication } from "@microsoft/teams-js";
import { getSsoToken, initTeams, isPreview } from "@/lib/teams";
import { SearchableSelect } from "@/components/SearchableSelect";
import { DatePicker } from "@/components/DatePicker";

type LeaveRequest = {
  id: string;
  subject: string;
  startDate: string;
  endDate: string;
  takenDays: number;
  status: string;
  currentUser: string;
  currentUserName: string;
};

type ParticipatedRequest = LeaveRequest & {
  employeeCode: string;
  employeeName: string;
};

type LeaveBalance = {
  requestType: number;
  actionTypeName: string;
  allocDays: number;
  remainingDays: number;
};

type UsedThisYear = {
  requestType: number;
  actionTypeName: string;
  takenDaysThisYear: number;
};

type LeaveContext = {
  displayName: string;
  mail: string;
  myRequests: LeaveRequest[];
  pendingApprovals: ParticipatedRequest[];
  balances: LeaveBalance[];
  usedThisYear: UsedThisYear[];
  canAccessSettings: boolean;
  canViewReports: boolean;
};

type ActionType = {
  id: number;
  name: string;
  nameRo: string;
  code: string;
};

type Employee = {
  mail: string;
  displayName: string;
};

type YearlyAllocationRow = {
  employeeCode: string;
  employeeFullName: string;
  isOnboarded: boolean;
  allocatedThisYear: number | null;
  bonusAllocatedThisYear: number | null;
};

type Holiday = {
  id: string;
  name: string;
  date: string;
  year: number;
  type: string;
};

type HistoryEntry = {
  timestamp: string;
  actor: string;
  actorName: string;
  action: string;
  details: string;
};

type ReportCurrentBalance = {
  requestType: number;
  actionTypeName: string;
  allocDays: number;
  takenDays: number;
  remainingDays: number;
};

type ReportYearlyBreakdown = {
  year: number;
  holidayAllocatedThisYear: number | null;
  takenByType: { actionTypeName: string; days: number }[];
  totalTaken: number;
};

type EmailStatus = { ok: boolean; text: string } | null;

type ReportHistoryEntry = {
  id: string;
  actionTypeName: string;
  requestSubject: string;
  startDate: string;
  endDate: string;
  takenDays: number;
  status: string;
  year: number;
};

type EmployeeLeaveReport = {
  employeeCode: string;
  employeeFullName: string;
  department?: string;
  jobTitle?: string;
  currentBalances: ReportCurrentBalance[];
  yearlyBreakdown: ReportYearlyBreakdown[];
  history: ReportHistoryEntry[];
};

type Tab = "myRequests" | "pendingApprovals" | "participated" | "reports" | "hr";
type Tone = "good" | "warning" | "critical" | "neutral";

const MANAGER_TOKEN = "{manager}";
const GM_TOKEN = "{GM}";

const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

const TONE_BG: Record<Tone, string> = {
  good: "bg-good",
  warning: "bg-warning",
  critical: "bg-critical",
  neutral: "bg-neutral",
};
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
  return status !== "DECLINED" && status !== "CANCELLED";
}

function labelForChainEntry(entry: string, nameByMail: Map<string, string>): string {
  if (entry === MANAGER_TOKEN) return "Direct Manager";
  if (entry === GM_TOKEN) return "GM";
  return nameByMail.get(entry.toLowerCase()) ?? entry;
}

function formatDate(dateString: string): string {
  const d = new Date(dateString);
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

// Some LR_HolidaySetup rows (2022+) store their date as local-midnight
// converted to UTC (e.g. "2026-08-14T21:00:00Z" for the Aug 15 holiday) —
// a raw `.slice(0, 10)` on that string reads the UTC calendar day, one day
// EARLIER than intended. Deriving the date via the browser's local-timezone
// getters (same as formatDate above) resolves correctly regardless of which
// era's storage convention a given row uses, as long as the browser is in
// the same timezone the dates were entered in (true for this company).
function toLocalIsoDate(dateString: string): string {
  const d = new Date(dateString);
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// Container Apps runs with minReplicas: 0 — the first request(s) after an idle
// period can hit a 502 while the api sidecar is still booting (nginx is up
// before the Functions host is). Only safe to auto-retry idempotent (GET)
// calls, since a 502 here means the request never reached the function at all.
const COLD_START_RETRY_DELAYS_MS = [1000, 2000];

async function apiFetch<T>(path: string, ssoToken: string, init?: RequestInit): Promise<T> {
  if (isPreview) return (await import("@/lib/preview")).previewApiFetch<T>(`/leave${path}`, init);
  const apiBaseUrl = process.env.NEXT_PUBLIC_API_BASE_URL;
  const isIdempotent = !init?.method || init.method === "GET";

  let response: Response;
  let attempt = 0;
  for (;;) {
    response = await fetch(`${apiBaseUrl}/leave${path}`, {
      ...init,
      headers: {
        // Note: Azure Static Web Apps' managed API proxy overwrites the standard
        // "Authorization" header for its own edge-to-function auth, so the Teams
        // SSO token is passed via a custom header instead.
        "X-Teams-Sso-Token": ssoToken,
        ...init?.headers,
      },
    });

    const isColdStartStatus = response.status === 502 || response.status === 503;
    if (!isColdStartStatus || !isIdempotent || attempt >= COLD_START_RETRY_DELAYS_MS.length) break;

    await sleep(COLD_START_RETRY_DELAYS_MS[attempt]);
    attempt += 1;
  }

  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    throw new Error(body.error ?? `Request failed with status ${response.status}`);
  }

  return (await response.json()) as T;
}

// Shared control classes — the "ledger" visual language: surface cards on a
// slightly cool page plane, a single deep-indigo accent for interaction, and
// status conveyed by tone (color + label + dot), never color alone.
const fieldClass =
  "rounded-lg border border-line bg-surface-2 p-2 text-sm text-ink focus:outline-none focus:ring-2 focus:ring-accent/40";
const labelClass = "flex flex-col gap-1 text-xs font-semibold text-ink-2";
const cardClass = "rounded-xl border border-line bg-surface p-4";
const btnPrimary =
  "rounded-lg bg-accent px-3 py-1.5 text-sm font-semibold text-accent-ink hover:brightness-110 disabled:opacity-50";
const btnGhost =
  "select-none rounded-lg border border-line px-3 py-1.5 text-sm font-semibold text-ink-2 hover:bg-surface-2 hover:text-ink disabled:opacity-50";
const linkBtn = "text-sm font-semibold text-accent hover:underline";

export default function LeavePage() {
  const [status, setStatus] = useState<"loading" | "ready" | "error">("loading");
  const [context, setContext] = useState<LeaveContext | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [actionTypes, setActionTypes] = useState<ActionType[]>([]);
  const [employees, setEmployees] = useState<Employee[]>([]);
  const nameByMail = useMemo(
    () => new Map(employees.map((e) => [e.mail.toLowerCase(), e.displayName])),
    [employees]
  );
  const [activeTab, setActiveTab] = useState<Tab>("myRequests");
  const [myRequestsStatusFilter, setMyRequestsStatusFilter] = useState<string>("APPROVED");
  const [participatedStatusFilter, setParticipatedStatusFilter] = useState<string>("APPROVED");

  const [showNewRequestForm, setShowNewRequestForm] = useState(false);
  const [newRequestType, setNewRequestType] = useState("");
  const [newRequestSubject, setNewRequestSubject] = useState("");
  const [newRequestStartDate, setNewRequestStartDate] = useState("");
  const [newRequestEndDate, setNewRequestEndDate] = useState("");
  const [newRequestComments, setNewRequestComments] = useState("");
  const [newRequestStandIn, setNewRequestStandIn] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [submitError, setSubmitError] = useState<string | null>(null);
  const [submitEmailStatus, setSubmitEmailStatus] = useState<EmailStatus>(null);

  const [respondingId, setRespondingId] = useState<string | null>(null);
  const [respondError, setRespondError] = useState<string | null>(null);
  const [respondEmailStatus, setRespondEmailStatus] = useState<EmailStatus>(null);
  const [respondComments, setRespondComments] = useState<Record<string, string>>({});
  const [bulkSelectMode, setBulkSelectMode] = useState(false);
  const [selectedApprovalIds, setSelectedApprovalIds] = useState<Set<string>>(new Set());
  const [bulkApproving, setBulkApproving] = useState(false);
  const [bulkProgress, setBulkProgress] = useState<{ done: number; total: number } | null>(null);
  const [cancellingId, setCancellingId] = useState<string | null>(null);
  const [cancelError, setCancelError] = useState<string | null>(null);
  const [cancelEmailStatus, setCancelEmailStatus] = useState<EmailStatus>(null);

  const [expandedHistoryId, setExpandedHistoryId] = useState<string | null>(null);
  const [historyByRequest, setHistoryByRequest] = useState<Record<string, HistoryEntry[]>>({});
  const [historyLoadingId, setHistoryLoadingId] = useState<string | null>(null);
  const [historyError, setHistoryError] = useState<string | null>(null);

  const [participatedRequests, setParticipatedRequests] = useState<ParticipatedRequest[]>([]);
  const [participatedLoaded, setParticipatedLoaded] = useState(false);
  const [participatedLoading, setParticipatedLoading] = useState(false);
  const [participatedError, setParticipatedError] = useState<string | null>(null);

  const [reportableEmployees, setReportableEmployees] = useState<Employee[]>([]);
  const [reportableLoaded, setReportableLoaded] = useState(false);
  const [reportEmployeeCode, setReportEmployeeCode] = useState("");
  const [report, setReport] = useState<EmployeeLeaveReport | null>(null);
  const [reportLoading, setReportLoading] = useState(false);
  const [reportError, setReportError] = useState<string | null>(null);
  const [reportStatusFilter, setReportStatusFilter] = useState<string[]>(["PENDING", "APPROVED"]);
  const [reportYearFilter, setReportYearFilter] = useState<number | "all">(new Date().getFullYear());
  const [reportMonthFilter, setReportMonthFilter] = useState<number | "all">("all"); // 0-11, matches Date#getMonth

  const [showChainEditor, setShowChainEditor] = useState(false);
  const [chainEmployeeSelection, setChainEmployeeSelection] = useState(""); // "" = standard/default chain
  const [chainDraft, setChainDraft] = useState<string[]>([]);
  const [chainIsOverride, setChainIsOverride] = useState(true);
  const [employeesWithOverride, setEmployeesWithOverride] = useState<string[]>([]);
  const [chainAddSelection, setChainAddSelection] = useState("");
  const [chainLoading, setChainLoading] = useState(false);
  const [chainSubmitting, setChainSubmitting] = useState(false);
  const [chainError, setChainError] = useState<string | null>(null);


  const [showAllocations, setShowAllocations] = useState(false);
  const [allocationsLoading, setAllocationsLoading] = useState(false);
  const [allocationYear, setAllocationYear] = useState(new Date().getFullYear());
  const [allocationRows, setAllocationRows] = useState<YearlyAllocationRow[]>([]);
  const [allocationError, setAllocationError] = useState<string | null>(null);
  const [bulkAllocating, setBulkAllocating] = useState(false);
  const [bulkAllocateSummary, setBulkAllocateSummary] = useState<string | null>(null);
  const [editingAllocationCode, setEditingAllocationCode] = useState<string | null>(null);
  const [editAllocationValue, setEditAllocationValue] = useState("");
  const [savingAllocationCode, setSavingAllocationCode] = useState<string | null>(null);
  const [onboardingCode, setOnboardingCode] = useState<string | null>(null);
  const [editingBonusCode, setEditingBonusCode] = useState<string | null>(null);
  const [editBonusValue, setEditBonusValue] = useState("");
  const [savingBonusCode, setSavingBonusCode] = useState<string | null>(null);

  const [showHolidays, setShowHolidays] = useState(false);
  const [holidaysLoading, setHolidaysLoading] = useState(false);
  const [holidays, setHolidays] = useState<Holiday[]>([]);
  // Days the New Request date pickers should refuse: legal holidays, plus every day
  // already covered by one of the user's own PENDING/APPROVED requests. `myRequests`
  // is windowed to last/current/next year (same as everywhere else in this app) — a
  // request just outside that window won't be pre-blocked here, but CreateLeaveRequest
  // still rejects the overlap server-side, so this is a UX nicety, not the real guard.
  const blockedDateReasons = useMemo(() => {
    const map = new Map<string, string>();
    holidays.forEach((h) => {
      map.set(toLocalIsoDate(h.date), `Legal holiday: ${h.name}`);
    });
    (context?.myRequests ?? [])
      .filter((r) => r.status === "PENDING" || r.status === "APPROVED")
      .forEach((r) => {
        const day = new Date(`${r.startDate.slice(0, 10)}T00:00:00`);
        const end = new Date(`${r.endDate.slice(0, 10)}T00:00:00`);
        while (day <= end) {
          const iso = `${day.getFullYear()}-${String(day.getMonth() + 1).padStart(2, "0")}-${String(day.getDate()).padStart(2, "0")}`;
          if (!map.has(iso)) map.set(iso, "Already requested");
          day.setDate(day.getDate() + 1);
        }
      });
    return map;
  }, [holidays, context]);
  const [showHolidayForm, setShowHolidayForm] = useState(false);
  const [holidayName, setHolidayName] = useState("");
  const [holidayDate, setHolidayDate] = useState("");
  const [holidaySubmitting, setHolidaySubmitting] = useState(false);
  const [holidayError, setHolidayError] = useState<string | null>(null);
  const [deletingHolidayId, setDeletingHolidayId] = useState<string | null>(null);

  const [refreshing, setRefreshing] = useState(false);
  const [lastRefreshedAt, setLastRefreshedAt] = useState<Date | null>(null);

  async function fetchLeaveContext(): Promise<LeaveContext> {
    const ssoToken = await getSsoToken();
    return apiFetch<LeaveContext>("/GetMyLeaveContext", ssoToken);
  }

  async function fetchParticipated() {
    setParticipatedLoading(true);
    setParticipatedError(null);
    try {
      const ssoToken = await getSsoToken();
      const { requests } = await apiFetch<{ requests: ParticipatedRequest[] }>("/GetParticipatedRequests", ssoToken);
      setParticipatedRequests(requests);
      setParticipatedLoaded(true);
    } catch (error) {
      setParticipatedError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setParticipatedLoading(false);
    }
  }

  async function fetchReportableEmployees() {
    try {
      const ssoToken = await getSsoToken();
      const { employees: list } = await apiFetch<{ employees: Employee[] }>("/GetReportableEmployees", ssoToken);
      setReportableEmployees(list);
      setReportableLoaded(true);
    } catch (error) {
      setReportError(error instanceof Error ? error.message : "Unknown error");
    }
  }

  async function handlePickReportEmployee(employeeCode: string) {
    setReportEmployeeCode(employeeCode);
    setReport(null);
    setReportYearFilter("all");
    setReportMonthFilter("all");
    if (!employeeCode) return;

    setReportLoading(true);
    setReportError(null);
    try {
      const ssoToken = await getSsoToken();
      const data = await apiFetch<EmployeeLeaveReport>(
        `/GetEmployeeLeaveReport?employeeCode=${encodeURIComponent(employeeCode)}`,
        ssoToken
      );
      setReport(data);
    } catch (error) {
      setReportError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setReportLoading(false);
    }
  }

  function toggleReportStatus(status: string) {
    setReportStatusFilter((prev) => (prev.includes(status) ? prev.filter((s) => s !== status) : [...prev, status]));
  }

  function getFilteredReportHistory(): ReportHistoryEntry[] {
    if (!report) return [];
    return report.history.filter(
      (h) =>
        reportStatusFilter.includes(h.status) &&
        (reportYearFilter === "all" || h.year === reportYearFilter) &&
        (reportMonthFilter === "all" || new Date(h.startDate).getMonth() === reportMonthFilter)
    );
  }

  function getReportAvailableYears(): number[] {
    if (!report) return [];
    return Array.from(new Set(report.history.map((h) => h.year))).sort((a, b) => b - a);
  }

  function handleExportReportCsv() {
    if (!report) return;

    const lines: string[] = [];
    lines.push(`Leave report for,${report.employeeFullName}`);
    lines.push("");
    lines.push("Current balances");
    lines.push("Type,Allocated,Taken,Remaining");
    report.currentBalances.forEach((b) => {
      lines.push(`${b.actionTypeName},${b.allocDays},${b.takenDays},${b.remainingDays}`);
    });
    lines.push("");
    lines.push("Yearly breakdown");
    lines.push("Year,Holiday Allocated,Total Taken,Breakdown");
    report.yearlyBreakdown.forEach((y) => {
      const breakdown = y.takenByType.map((t) => `${t.actionTypeName} ${t.days}`).join("; ");
      lines.push(`${y.year},${y.holidayAllocatedThisYear ?? ""},${y.totalTaken},"${breakdown}"`);
    });
    lines.push("");
    lines.push(`History (status: ${reportStatusFilter.join(", ") || "none selected"})`);
    lines.push("Type,Subject,Start,End,Days,Status");
    getFilteredReportHistory().forEach((h) => {
      lines.push(
        `${h.actionTypeName},"${h.requestSubject}",${formatDate(h.startDate)},${formatDate(h.endDate)},${h.takenDays},${h.status}`
      );
    });

    const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `leave-report-${report.employeeCode}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  }

  function handleSelectTab(tab: Tab) {
    setActiveTab(tab);
    if (tab === "participated" && !participatedLoaded && !participatedLoading) {
      void fetchParticipated();
    }
    if (tab === "reports" && !reportableLoaded) {
      void fetchReportableEmployees();
    }
  }

  // Re-fetches the base context plus whatever the active tab already has
  // loaded, so a manual click or the periodic timer below always brings back
  // fresh balances/approvals/etc. without losing the user's current view.
  async function handleRefresh() {
    if (refreshing) return;
    setRefreshing(true);
    try {
      setContext(await fetchLeaveContext());
      if (activeTab === "participated" && participatedLoaded) {
        await fetchParticipated();
      }
      if (activeTab === "reports" && reportEmployeeCode) {
        await handlePickReportEmployee(reportEmployeeCode);
      }
      setLastRefreshedAt(new Date());
    } catch (error) {
      // Best-effort — a failed background refresh shouldn't replace what's already on screen with an error.
      console.error("Refresh failed", error);
    } finally {
      setRefreshing(false);
    }
  }

  useEffect(() => {
    // The home screen's "Awaiting your approval" rows link here with ?tab=pendingApprovals.
    if (new URLSearchParams(window.location.search).get("tab") === "pendingApprovals") {
      setActiveTab("pendingApprovals");
    }
    // The home screen's "New request" button links here with ?new=1 — open the form straight away.
    if (new URLSearchParams(window.location.search).get("new") === "1") {
      setShowNewRequestForm(true);
    }

    async function fetchActionTypes(): Promise<ActionType[]> {
      const ssoToken = await getSsoToken();
      const { actionTypes } = await apiFetch<{ actionTypes: ActionType[] }>("/GetActionTypes", ssoToken);
      return actionTypes;
    }

    async function fetchEmployees(): Promise<Employee[]> {
      const ssoToken = await getSsoToken();
      const { employees } = await apiFetch<{ employees: Employee[] }>("/GetPickableEmployees", ssoToken);
      return employees;
    }

    async function fetchHolidaysList(): Promise<Holiday[]> {
      const ssoToken = await getSsoToken();
      const { holidays: list } = await apiFetch<{ holidays: Holiday[] }>("/GetHolidays", ssoToken);
      return list;
    }

    async function signInAndLoad() {
      try {
        await initTeams();

        try {
          setContext(await fetchLeaveContext());
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

          setContext(await fetchLeaveContext());
        }

        setActionTypes(await fetchActionTypes());
        setEmployees(await fetchEmployees());
        setHolidays(await fetchHolidaysList());
        setStatus("ready");
      } catch (error) {
        setErrorMessage(error instanceof Error ? error.message : "Unknown error");
        setStatus("error");
      }
    }

    void signInAndLoad();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Background auto-refresh every 4 minutes, in addition to the manual button —
  // re-created whenever the active tab (or its loaded state) changes so it always
  // refreshes whatever's actually on screen, not a stale snapshot of it.
  useEffect(() => {
    if (status !== "ready") return;
    const intervalId = setInterval(() => {
      void handleRefresh();
    }, 4 * 60 * 1000);
    return () => clearInterval(intervalId);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [status, activeTab, participatedLoaded, reportEmployeeCode]);

  function resetNewRequestForm() {
    setNewRequestType("");
    setNewRequestSubject("");
    setNewRequestStartDate("");
    setNewRequestEndDate("");
    setNewRequestComments("");
    setNewRequestStandIn("");
    setSubmitError(null);
  }

  function describeNonWorkingDay(iso: string): string | null {
    if (!iso) return null;
    const reason = blockedDateReasons.get(iso);
    if (reason) return reason.toLowerCase();
    const day = new Date(`${iso}T00:00:00`).getDay();
    if (day === 0 || day === 6) return "a weekend";
    return null;
  }

  async function handleSubmitNewRequest(e: React.FormEvent) {
    e.preventDefault();
    setSubmitError(null);
    setSubmitEmailStatus(null);

    const startIssue = describeNonWorkingDay(newRequestStartDate);
    if (startIssue) {
      setSubmitError(`Start date can't be ${startIssue}.`);
      return;
    }
    const endIssue = describeNonWorkingDay(newRequestEndDate);
    if (endIssue) {
      setSubmitError(`End date can't be ${endIssue}.`);
      return;
    }

    setSubmitting(true);
    try {
      const ssoToken = await getSsoToken();
      const { emailSent, currentUser } = await apiFetch<{ emailSent: boolean; currentUser: string }>(
        "/CreateLeaveRequest",
        ssoToken,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            requestType: Number(newRequestType),
            requestSubject: newRequestSubject,
            startDate: newRequestStartDate,
            endDate: newRequestEndDate,
            comments: newRequestComments,
            standInEmployee: newRequestStandIn || undefined,
          }),
        }
      );
      setSubmitEmailStatus(
        emailSent
          ? { ok: true, text: `Email sent to ${currentUser}.` }
          : { ok: false, text: `Request submitted, but the notification email to ${currentUser} failed to send.` }
      );

      setContext(await fetchLeaveContext());
      resetNewRequestForm();
      setShowNewRequestForm(false);
    } catch (error) {
      setSubmitError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setSubmitting(false);
    }
  }

  async function handleRespond(id: string, decision: "approved" | "declined") {
    setRespondingId(id);
    setRespondError(null);
    setRespondEmailStatus(null);

    try {
      const ssoToken = await getSsoToken();
      const { emailSent, currentUser } = await apiFetch<{ emailSent: boolean; currentUser: string }>(
        "/RespondToLeaveRequest",
        ssoToken,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id, decision, comment: respondComments[id] || undefined }),
        }
      );
      setRespondEmailStatus(
        emailSent
          ? { ok: true, text: `Email sent to ${currentUser}.` }
          : { ok: false, text: `${decision === "approved" ? "Approved" : "Declined"}, but the notification email to ${currentUser} failed to send.` }
      );

      // Drop it from Pending Approvals immediately — don't wait on the refetch below,
      // which can lag a beat behind the write and leave an already-processed request
      // visible for a moment.
      setContext((prev) => (prev ? { ...prev, pendingApprovals: prev.pendingApprovals.filter((r) => r.id !== id) } : prev));
      setContext(await fetchLeaveContext());
      if (participatedLoaded) await fetchParticipated();
      setRespondComments((prev) => {
        const next = { ...prev };
        delete next[id];
        return next;
      });
      if (historyByRequest[id]) {
        setHistoryByRequest((prev) => {
          const next = { ...prev };
          delete next[id];
          return next;
        });
      }
    } catch (error) {
      setRespondError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setRespondingId(null);
    }
  }

  function toggleApprovalSelected(id: string) {
    setSelectedApprovalIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  // "Select all" enters bulk mode with every pending request already ticked — the
  // approver unticks the few they don't want, rather than ticking each one by hand.
  function toggleBulkSelectMode() {
    if (bulkSelectMode) {
      setBulkSelectMode(false);
      setSelectedApprovalIds(new Set());
    } else {
      setBulkSelectMode(true);
      setSelectedApprovalIds(new Set(context?.pendingApprovals.map((r) => r.id) ?? []));
    }
  }

  async function handleBulkApprove() {
    const ids = Array.from(selectedApprovalIds);
    if (ids.length === 0) return;

    setBulkApproving(true);
    setRespondError(null);
    setRespondEmailStatus(null);

    let failures = 0;
    setBulkProgress({ done: 0, total: ids.length });
    for (const [index, id] of ids.entries()) {
      try {
        const ssoToken = await getSsoToken();
        await apiFetch("/RespondToLeaveRequest", ssoToken, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id, decision: "approved" }),
        });
        // Drop each one from the list as soon as it's done — each approval can take a few
        // seconds (PDF + emails on a final step), so waiting for the whole batch makes the
        // app look frozen.
        setContext((prev) => (prev ? { ...prev, pendingApprovals: prev.pendingApprovals.filter((r) => r.id !== id) } : prev));
        setSelectedApprovalIds((prev) => {
          const next = new Set(prev);
          next.delete(id);
          return next;
        });
      } catch {
        failures++;
      }
      setBulkProgress({ done: index + 1, total: ids.length });
    }
    setBulkProgress(null);

    setRespondEmailStatus(
      failures === 0
        ? { ok: true, text: `Approved ${ids.length} request(s).` }
        : { ok: false, text: `Approved ${ids.length - failures} of ${ids.length} — ${failures} failed, check them individually.` }
    );

    setSelectedApprovalIds(new Set());
    setBulkSelectMode(false);
    setContext(await fetchLeaveContext());
    if (participatedLoaded) await fetchParticipated();
    setBulkApproving(false);
  }

  async function handleToggleHistory(id: string) {
    if (expandedHistoryId === id) {
      setExpandedHistoryId(null);
      return;
    }
    setExpandedHistoryId(id);
    setHistoryError(null);
    if (historyByRequest[id]) return;

    setHistoryLoadingId(id);
    try {
      const ssoToken = await getSsoToken();
      const { history } = await apiFetch<{ history: HistoryEntry[] }>(
        `/GetRequestHistory?id=${encodeURIComponent(id)}`,
        ssoToken
      );
      setHistoryByRequest((prev) => ({ ...prev, [id]: history }));
    } catch (error) {
      setHistoryError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setHistoryLoadingId(null);
    }
  }

  async function handleCancel(id: string) {
    setCancellingId(id);
    setCancelError(null);
    setCancelEmailStatus(null);

    try {
      const ssoToken = await getSsoToken();
      const { emailSent } = await apiFetch<{ emailSent: boolean }>("/CancelLeaveRequest", ssoToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id }),
      });
      setCancelEmailStatus(
        emailSent
          ? { ok: true, text: "Cancellation email sent." }
          : { ok: false, text: "Request cancelled, but the notification email failed to send." }
      );

      setContext(await fetchLeaveContext());
      if (participatedLoaded) await fetchParticipated();
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

  async function handleOpenChainEditor() {
    const next = !showChainEditor;
    setShowChainEditor(next);
    if (!next) return;
    setChainEmployeeSelection("");
    await loadChainFor("");
  }

  async function handleChangeChainEmployee(value: string) {
    setChainEmployeeSelection(value);
    await loadChainFor(value);
  }

  function handleAddToChain() {
    if (!chainAddSelection || chainDraft.includes(chainAddSelection)) {
      return;
    }
    setChainDraft((prev) => [...prev, chainAddSelection]);
    setChainAddSelection("");
  }

  function handleRemoveFromChain(email: string) {
    setChainDraft((prev) => prev.filter((e) => e !== email));
  }

  async function handleSaveChain() {
    setChainSubmitting(true);
    setChainError(null);
    try {
      const ssoToken = await getSsoToken();
      const { chain, isOverride } = await apiFetch<{ chain: string[]; isOverride: boolean }>(
        "/SetApprovalFlow",
        ssoToken,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ chain: chainDraft, employeeCode: chainEmployeeSelection || undefined }),
        }
      );
      setChainDraft(chain);
      setChainIsOverride(isOverride);
    } catch (error) {
      setChainError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setChainSubmitting(false);
    }
  }

  async function loadAllocations(year: number) {
    setAllocationsLoading(true);
    setAllocationError(null);
    try {
      const ssoToken = await getSsoToken();
      const { employees: rows } = await apiFetch<{ year: number; employees: YearlyAllocationRow[] }>(
        `/GetYearlyAllocations?year=${year}`,
        ssoToken
      );
      setAllocationRows(rows);
    } catch (error) {
      setAllocationError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setAllocationsLoading(false);
    }
  }

  async function handleOpenAllocations() {
    const next = !showAllocations;
    setShowAllocations(next);
    if (!next) return;
    setBulkAllocateSummary(null);
    await loadAllocations(allocationYear);
  }

  async function handleChangeAllocationYear(delta: number) {
    const nextYear = allocationYear + delta;
    setAllocationYear(nextYear);
    setBulkAllocateSummary(null);
    setEditingAllocationCode(null);
    await loadAllocations(nextYear);
  }

  async function handleBulkAllocate() {
    setBulkAllocating(true);
    setAllocationError(null);
    setBulkAllocateSummary(null);
    try {
      const ssoToken = await getSsoToken();
      const result = await apiFetch<{
        defaultDays: number;
        allocatedCount: number;
        skippedAlreadyAllocated: number;
        skippedNotOnboarded: number;
      }>("/BulkAllocateLeaveDays", ssoToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ year: allocationYear }),
      });
      setBulkAllocateSummary(
        `Granted ${result.defaultDays} days to ${result.allocatedCount} employee(s). ` +
          `Skipped ${result.skippedAlreadyAllocated} already allocated, ${result.skippedNotOnboarded} not yet onboarded.`
      );
      await loadAllocations(allocationYear);
    } catch (error) {
      setAllocationError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setBulkAllocating(false);
    }
  }

  function handleStartEditAllocation(row: YearlyAllocationRow) {
    setEditingAllocationCode(row.employeeCode);
    setEditAllocationValue(String(row.allocatedThisYear ?? 0));
  }

  async function handleSaveAllocation(employeeCode: string) {
    setSavingAllocationCode(employeeCode);
    setAllocationError(null);
    try {
      const ssoToken = await getSsoToken();
      await apiFetch("/SetYearlyAllocation", ssoToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          employeeCode,
          year: allocationYear,
          holidayDaysAllocatedThisYear: Number(editAllocationValue),
        }),
      });
      setEditingAllocationCode(null);
      await loadAllocations(allocationYear);
    } catch (error) {
      setAllocationError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setSavingAllocationCode(null);
    }
  }

  async function handleOnboardEmployee(employeeCode: string) {
    setOnboardingCode(employeeCode);
    setAllocationError(null);
    try {
      const ssoToken = await getSsoToken();
      await apiFetch("/OnboardEmployee", ssoToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ employeeCode, year: allocationYear }),
      });
      await loadAllocations(allocationYear);
    } catch (error) {
      setAllocationError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setOnboardingCode(null);
    }
  }

  function handleStartEditBonus(row: YearlyAllocationRow) {
    setEditingBonusCode(row.employeeCode);
    setEditBonusValue(String(row.bonusAllocatedThisYear ?? 0));
  }

  async function handleSaveBonus(employeeCode: string) {
    setSavingBonusCode(employeeCode);
    setAllocationError(null);
    try {
      const ssoToken = await getSsoToken();
      await apiFetch("/SetYearlyAllocation", ssoToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          employeeCode,
          year: allocationYear,
          bonusDaysAllocatedThisYear: Number(editBonusValue),
        }),
      });
      setEditingBonusCode(null);
      await loadAllocations(allocationYear);
    } catch (error) {
      setAllocationError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setSavingBonusCode(null);
    }
  }

  async function handleOpenHolidays() {
    const next = !showHolidays;
    setShowHolidays(next);
    if (!next) return;

    setHolidaysLoading(true);
    try {
      const ssoToken = await getSsoToken();
      const { holidays: list } = await apiFetch<{ holidays: Holiday[] }>("/GetHolidays", ssoToken);
      setHolidays(list);
    } catch (error) {
      setHolidayError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setHolidaysLoading(false);
    }
  }

  async function handleAddHoliday(e: React.FormEvent) {
    e.preventDefault();
    setHolidaySubmitting(true);
    setHolidayError(null);

    try {
      const ssoToken = await getSsoToken();
      const created = await apiFetch<Holiday>("/AddHoliday", ssoToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: holidayName, date: holidayDate }),
      });

      setHolidays((prev) => [created, ...prev].sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()));
      setHolidayName("");
      setHolidayDate("");
      setShowHolidayForm(false);
    } catch (error) {
      setHolidayError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setHolidaySubmitting(false);
    }
  }

  async function handleDeleteHoliday(id: string) {
    setDeletingHolidayId(id);
    setHolidayError(null);

    try {
      const ssoToken = await getSsoToken();
      await apiFetch("/DeleteHoliday", ssoToken, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id }),
      });

      setHolidays((prev) => prev.filter((h) => h.id !== id));
    } catch (error) {
      setHolidayError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setDeletingHolidayId(null);
    }
  }

  function renderRequestLine(
    lr: LeaveRequest | ParticipatedRequest,
    opts: { showApproveDecline: boolean; showCancel: boolean; showSelectCheckbox?: boolean }
  ) {
    const owner = "employeeName" in lr ? lr.employeeName : null;
    const tone = toneForStatus(lr.status);
    return (
      <li
        key={lr.id}
        className="rounded-lg border border-line border-l-[3px] bg-surface p-3"
        style={{ borderLeftColor: `var(--${tone === "neutral" ? "neutral" : tone})` }}
      >
        <div className="flex items-baseline justify-between gap-2">
          {opts.showSelectCheckbox ? (
            <label className="flex select-none items-baseline gap-2 py-1 text-sm font-semibold text-ink">
              <input
                type="checkbox"
                checked={selectedApprovalIds.has(lr.id)}
                onChange={() => toggleApprovalSelected(lr.id)}
                className="h-4 w-4 flex-none accent-accent"
              />
              {owner && <span className="font-medium text-ink-2">{owner} — </span>}
              {lr.subject}
            </label>
          ) : (
            <div className="flex items-baseline gap-2 text-sm font-semibold text-ink">
              {owner && <span className="font-medium text-ink-2">{owner} — </span>}
              {lr.subject}
            </div>
          )}
          <span
            className={`inline-flex flex-none items-center gap-1.5 rounded-full px-2 py-0.5 text-[10px] font-bold tracking-wide uppercase ${TONE_SOFT[tone]}`}
          >
            <span className={`h-1.5 w-1.5 rounded-full ${TONE_BG[tone]}`} />
            {lr.status}
          </span>
        </div>

        <div className="mt-1 flex flex-wrap items-center gap-x-2 gap-y-0.5 text-xs text-ink-2">
          <span className="font-mono tabular-nums text-ink">
            {formatDate(lr.startDate)} – {formatDate(lr.endDate)}
          </span>
          <span className="text-line">·</span>
          <span className="font-mono tabular-nums">{lr.takenDays} days</span>
          {lr.currentUserName && (
            <>
              <span className="text-line">·</span>
              <span className="text-ink-3">
                with <span className="font-semibold text-ink-2">{lr.currentUserName}</span>
              </span>
            </>
          )}
        </div>

        {opts.showApproveDecline && (
          <textarea
            value={respondComments[lr.id] ?? ""}
            onChange={(e) => setRespondComments((prev) => ({ ...prev, [lr.id]: e.target.value }))}
            placeholder="Comment (optional)"
            rows={2}
            className={`mt-2 w-full ${fieldClass}`}
          />
        )}

        <div className="mt-2 flex flex-wrap items-center gap-2">
          {opts.showApproveDecline && (
            <>
              <button
                type="button"
                disabled={respondingId === lr.id}
                onClick={() => handleRespond(lr.id, "approved")}
                className="rounded-lg bg-good px-3 py-1 text-xs font-bold text-white hover:brightness-110 disabled:opacity-50"
              >
                Approve
              </button>
              <button
                type="button"
                disabled={respondingId === lr.id}
                onClick={() => handleRespond(lr.id, "declined")}
                className="rounded-lg border border-critical/40 px-3 py-1 text-xs font-bold text-critical hover:bg-critical-soft disabled:opacity-50"
              >
                Decline
              </button>
            </>
          )}
          {opts.showCancel && isCancellable(lr.status) && (
            <button
              type="button"
              disabled={cancellingId === lr.id}
              onClick={() => handleCancel(lr.id)}
              className="rounded-lg border border-line px-3 py-1 text-xs font-semibold text-ink-2 hover:bg-surface-2 disabled:opacity-50"
            >
              Cancel
            </button>
          )}
          <button type="button" onClick={() => handleToggleHistory(lr.id)} className={linkBtn + " ml-auto"}>
            {expandedHistoryId === lr.id ? "Hide history" : "History"}
          </button>
        </div>

        {expandedHistoryId === lr.id && (
          <div className="mt-3 border-t border-dashed border-line pt-3">
            {historyLoadingId === lr.id ? (
              <p className="text-xs text-ink-2">Loading…</p>
            ) : historyError ? (
              <p className="text-xs text-critical">{historyError}</p>
            ) : (historyByRequest[lr.id]?.length ?? 0) === 0 ? (
              <p className="text-xs text-ink-2">No actions recorded yet.</p>
            ) : (
              <ul className="flex flex-col gap-2.5">
                {historyByRequest[lr.id].map((h, i, arr) => (
                  <li key={i} className="flex gap-2.5">
                    <div className="flex w-2 flex-none flex-col items-center pt-0.5">
                      <span className="h-1.5 w-1.5 flex-none rounded-full bg-accent" />
                      {i < arr.length - 1 && <span className="mt-1 w-px flex-1 bg-line" />}
                    </div>
                    <div className="pb-0.5">
                      <div className="font-mono text-[11px] tabular-nums text-ink-3">{formatDate(h.timestamp)}</div>
                      <div className="text-xs font-semibold text-ink">{h.actorName}</div>
                      <div className="text-xs text-ink-2">{h.details}</div>
                    </div>
                  </li>
                ))}
              </ul>
            )}
          </div>
        )}
      </li>
    );
  }

  return (
    <main className="flex min-h-screen flex-col items-center justify-start gap-5 bg-page p-6 font-sans text-ink sm:p-8">
      {status === "loading" && <p className="mt-10 text-sm text-ink-2">Connecting…</p>}

      {status === "error" && (
        <p className="mt-10 max-w-md text-center text-sm text-critical">Authentication failed: {errorMessage}</p>
      )}

      {status === "ready" && context && (
        <div className="flex w-full max-w-4xl flex-col gap-5">
          {/* Identity strip */}
          <div className="flex items-center justify-between gap-3">
            <div className="flex min-w-0 items-center gap-2.5">
              <span className="flex h-9 w-9 flex-none items-center justify-center rounded-full bg-accent text-[13px] font-bold text-accent-ink">
                {initials(context.displayName)}
              </span>
              <div className="min-w-0 leading-tight">
                <div className="truncate text-[14.5px] font-semibold text-ink">{context.displayName}</div>
                <div className="truncate text-xs text-ink-2">{context.mail}</div>
              </div>
            </div>
            <div className="flex items-center gap-2.5">
              <button
                type="button"
                onClick={() => void handleRefresh()}
                disabled={refreshing}
                title={lastRefreshedAt ? `Last refreshed ${lastRefreshedAt.toLocaleTimeString()}` : "Refresh"}
                className={`${btnGhost} flex items-center gap-1 px-2 py-1 text-xs`}
              >
                <span className={refreshing ? "inline-block animate-spin" : "inline-block"}>↻</span>
                {refreshing ? "Refreshing…" : "Refresh"}
              </button>
              <div className="text-[11px] font-bold tracking-wider text-ink-3 uppercase">Leave</div>
            </div>
          </div>

          {/* Segmented tabs */}
          <div className="flex gap-0.5 overflow-x-auto rounded-xl border border-line bg-surface-2 p-[3px]">
            {(
              [
                ["myRequests", "My Requests"],
                ["pendingApprovals", "Pending Approvals"],
                ["participated", "Managed Requests"],
              ] as const
            ).map(([tab, label]) => (
              <button
                key={tab}
                type="button"
                onClick={() => handleSelectTab(tab)}
                aria-selected={activeTab === tab}
                className={`flex items-center gap-1.5 rounded-lg px-3 py-2 text-[13px] font-semibold whitespace-nowrap transition-colors ${
                  activeTab === tab ? "bg-surface text-ink shadow-sm" : "text-ink-2 hover:text-ink"
                }`}
              >
                {label}
                {tab === "pendingApprovals" && context.pendingApprovals.length > 0 && (
                  <span className="rounded-full bg-warning-soft px-1.5 py-0.5 font-mono text-[10px] tabular-nums text-warning">
                    {context.pendingApprovals.length}
                  </span>
                )}
              </button>
            ))}
            {context.canViewReports && (
              <button
                type="button"
                onClick={() => handleSelectTab("reports")}
                aria-selected={activeTab === "reports"}
                className={`rounded-lg px-3 py-2 text-[13px] font-semibold whitespace-nowrap transition-colors ${
                  activeTab === "reports" ? "bg-surface text-ink shadow-sm" : "text-ink-2 hover:text-ink"
                }`}
              >
                Reports
              </button>
            )}
            {context.canAccessSettings && (
              <button
                type="button"
                onClick={() => handleSelectTab("hr")}
                aria-selected={activeTab === "hr"}
                className={`rounded-lg px-3 py-2 text-[13px] font-semibold whitespace-nowrap transition-colors ${
                  activeTab === "hr" ? "bg-surface text-ink shadow-sm" : "text-ink-2 hover:text-ink"
                }`}
              >
                Settings
              </button>
            )}
          </div>

          {activeTab === "myRequests" && (
            <>
              {(context.balances.length > 0 || context.usedThisYear.length > 0) && (
                <div className="flex flex-col gap-2.5">
                  <div className="px-0.5 text-[11px] font-bold tracking-wider text-ink-3 uppercase">Balance</div>
                  {context.balances.length > 0 && (
                    <div className="grid grid-cols-2 gap-2.5">
                      {context.balances.map((b, i) => (
                        <div
                          key={b.requestType}
                          className={`relative overflow-hidden rounded-lg border border-line bg-surface p-3 before:absolute before:inset-x-0 before:top-0 before:h-[3px] ${
                            i === 0 ? "before:bg-accent" : "before:bg-[#6b7cb0]"
                          }`}
                        >
                          <div className="font-mono text-2xl leading-none font-bold tracking-tight tabular-nums text-ink">
                            {b.remainingDays}
                            <small className="ml-0.5 text-[13px] font-medium text-ink-3">days</small>
                          </div>
                          <div className="mt-1 text-xs text-ink-2">{b.actionTypeName} remaining</div>
                        </div>
                      ))}
                    </div>
                  )}
                  {context.usedThisYear.length > 0 && (
                    <div className="grid grid-cols-2 gap-2.5">
                      {context.usedThisYear.map((u) => (
                        <div
                          key={u.requestType}
                          className="relative overflow-hidden rounded-lg border border-line bg-surface p-3 before:absolute before:inset-x-0 before:top-0 before:h-[3px] before:bg-ink-3"
                        >
                          <div className="font-mono text-2xl leading-none font-bold tracking-tight tabular-nums text-ink">
                            {u.takenDaysThisYear}
                            <small className="ml-0.5 text-[13px] font-medium text-ink-3">days</small>
                          </div>
                          <div className="mt-1 text-xs text-ink-2">{u.actionTypeName}, this year</div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}

              <div className="flex items-center justify-between gap-3">
                <label className="flex items-center gap-2 text-xs font-semibold text-ink-2">
                  Status
                  <select
                    value={myRequestsStatusFilter}
                    onChange={(e) => setMyRequestsStatusFilter(e.target.value)}
                    className={fieldClass}
                  >
                    <option value="ALL">All</option>
                    <option value="PENDING">Pending</option>
                    <option value="APPROVED">Approved</option>
                    <option value="DECLINED">Declined</option>
                    <option value="CANCELLED">Cancelled</option>
                  </select>
                </label>
                <button type="button" onClick={() => setShowNewRequestForm((v) => !v)} className={btnPrimary}>
                  {showNewRequestForm ? "Cancel" : "+ New Request"}
                </button>
              </div>

              {showNewRequestForm && (
                <form onSubmit={handleSubmitNewRequest} className={`${cardClass} flex flex-col gap-3`}>
                  <div className="px-0.5 text-[11px] font-bold tracking-wider text-ink-3 uppercase">New Request</div>
                  <div className="flex gap-3">
                    <label className={`flex-1 ${labelClass}`}>
                      Type
                      <select
                        required
                        value={newRequestType}
                        onChange={(e) => setNewRequestType(e.target.value)}
                        className={fieldClass}
                      >
                        <option value="" disabled>
                          Select leave type
                        </option>
                        {actionTypes.map((t) => (
                          <option key={t.id} value={t.id}>
                            {t.name}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label className={`flex-1 ${labelClass}`}>
                      Stand-in (optional)
                      <SearchableSelect
                        value={newRequestStandIn}
                        onChange={setNewRequestStandIn}
                        options={[
                          { value: "", label: "No stand-in" },
                          ...employees.map((emp) => ({ value: emp.mail, label: emp.displayName })),
                        ]}
                        className={fieldClass}
                      />
                    </label>
                  </div>

                  <label className={labelClass}>
                    Subject
                    <input
                      required
                      type="text"
                      value={newRequestSubject}
                      onChange={(e) => setNewRequestSubject(e.target.value)}
                      className={fieldClass}
                    />
                  </label>

                  <div className="flex gap-3">
                    <label className={`flex-1 ${labelClass}`}>
                      Start date
                      <DatePicker
                        required
                        value={newRequestStartDate}
                        onChange={setNewRequestStartDate}
                        disabledDates={blockedDateReasons}
                        className={`${fieldClass} font-mono`}
                      />
                    </label>
                    <label className={`flex-1 ${labelClass}`}>
                      End date
                      <DatePicker
                        required
                        value={newRequestEndDate}
                        onChange={setNewRequestEndDate}
                        disabledDates={blockedDateReasons}
                        className={`${fieldClass} font-mono`}
                      />
                    </label>
                  </div>

                  <label className={labelClass}>
                    Comments (optional)
                    <textarea
                      value={newRequestComments}
                      onChange={(e) => setNewRequestComments(e.target.value)}
                      className={fieldClass}
                      rows={2}
                    />
                  </label>

                  {submitError && <p className="text-xs text-critical">{submitError}</p>}

                  <button type="submit" disabled={submitting} className={`${btnPrimary} self-start`}>
                    {submitting ? "Submitting…" : "Submit request"}
                  </button>
                </form>
              )}

              {submitEmailStatus && (
                <p className={`text-xs ${submitEmailStatus.ok ? "text-good" : "text-critical"}`}>{submitEmailStatus.text}</p>
              )}
              {cancelError && <p className="text-xs text-critical">{cancelError}</p>}
              {cancelEmailStatus && (
                <p className={`text-xs ${cancelEmailStatus.ok ? "text-good" : "text-critical"}`}>{cancelEmailStatus.text}</p>
              )}

              {(() => {
                const filtered =
                  myRequestsStatusFilter === "ALL"
                    ? context.myRequests
                    : context.myRequests.filter((lr) => lr.status === myRequestsStatusFilter);
                return filtered.length === 0 ? (
                  <p className="text-sm text-ink-2">
                    {myRequestsStatusFilter === "ALL" ? "You have no leave requests." : "No requests match this status."}
                  </p>
                ) : (
                  <ul className="flex flex-col gap-2.5">
                    {filtered.map((lr) => renderRequestLine(lr, { showApproveDecline: false, showCancel: true }))}
                  </ul>
                );
              })()}
            </>
          )}

          {activeTab === "pendingApprovals" && (
            <>
              {respondError && <p className="text-xs text-critical">{respondError}</p>}
              {respondEmailStatus && (
                <p className={`text-xs ${respondEmailStatus.ok ? "text-good" : "text-critical"}`}>{respondEmailStatus.text}</p>
              )}

              {context.pendingApprovals.length > 0 && (
                <div className="flex items-center justify-between gap-2">
                  <button type="button" onClick={toggleBulkSelectMode} disabled={bulkApproving} className={btnGhost}>
                    {bulkSelectMode ? "Cancel selection" : "Select all"}
                  </button>
                  {bulkSelectMode && (
                    <button
                      type="button"
                      onClick={handleBulkApprove}
                      disabled={selectedApprovalIds.size === 0 || bulkApproving}
                      className="rounded-lg bg-good px-3 py-1.5 text-xs font-bold text-white hover:brightness-110 disabled:opacity-50"
                    >
                      {bulkApproving
                        ? `Approving… ${bulkProgress?.done ?? 0}/${bulkProgress?.total ?? 0}`
                        : `Approve selected (${selectedApprovalIds.size})`}
                    </button>
                  )}
                </div>
              )}

              {context.pendingApprovals.length === 0 ? (
                <p className="text-sm text-ink-2">You have no pending approvals.</p>
              ) : (
                <ul className="flex flex-col gap-2.5">
                  {context.pendingApprovals.map((lr) =>
                    renderRequestLine(lr, { showApproveDecline: true, showCancel: false, showSelectCheckbox: bulkSelectMode })
                  )}
                </ul>
              )}
            </>
          )}

          {activeTab === "participated" && (
            <>
              {cancelError && <p className="text-xs text-critical">{cancelError}</p>}

              <label className="flex items-center gap-2 text-xs font-semibold text-ink-2">
                Status
                <select
                  value={participatedStatusFilter}
                  onChange={(e) => setParticipatedStatusFilter(e.target.value)}
                  className={fieldClass}
                >
                  <option value="ALL">All</option>
                  <option value="PENDING">Pending</option>
                  <option value="APPROVED">Approved</option>
                  <option value="DECLINED">Declined</option>
                  <option value="CANCELLED">Cancelled</option>
                </select>
              </label>

              {participatedLoading ? (
                <p className="text-sm text-ink-2">Loading…</p>
              ) : participatedError ? (
                <p className="text-xs text-critical">{participatedError}</p>
              ) : (
                (() => {
                  const filtered =
                    participatedStatusFilter === "ALL"
                      ? participatedRequests
                      : participatedRequests.filter((lr) => lr.status === participatedStatusFilter);
                  return filtered.length === 0 ? (
                    <p className="text-sm text-ink-2">
                      {participatedStatusFilter === "ALL"
                        ? "You haven't managed any requests yet."
                        : "No requests match this status."}
                    </p>
                  ) : (
                    <ul className="flex flex-col gap-2.5">
                      {filtered.map((lr) => renderRequestLine(lr, { showApproveDecline: false, showCancel: true }))}
                    </ul>
                  );
                })()
              )}
            </>
          )}

          {activeTab === "reports" && context.canViewReports && (
            <div className={`${cardClass} flex flex-col gap-3`}>
              <label className={labelClass}>
                Employee
                <SearchableSelect
                  value={reportEmployeeCode}
                  onChange={handlePickReportEmployee}
                  placeholder="Select employee"
                  options={reportableEmployees.map((emp) => ({ value: emp.mail, label: emp.displayName }))}
                  className={fieldClass}
                />
              </label>

              {reportError && <p className="text-xs text-critical">{reportError}</p>}

              {reportLoading ? (
                <p className="text-sm text-ink-2">Loading…</p>
              ) : report ? (
                <>
                  <div className="flex items-center justify-between gap-2">
                    <div className="min-w-0">
                      <div className="truncate text-sm font-semibold text-ink">{report.employeeFullName}</div>
                      {(report.department || report.jobTitle) && (
                        <div className="truncate text-xs text-ink-2">
                          {[report.jobTitle, report.department].filter(Boolean).join(" — ")}
                        </div>
                      )}
                    </div>
                    <button type="button" onClick={handleExportReportCsv} className={`${btnGhost} flex-none`}>
                      Export CSV
                    </button>
                  </div>

                  {report.currentBalances.length > 0 && (
                    <div className="grid grid-cols-2 gap-2.5">
                      {report.currentBalances.map((b, i) => (
                        <div
                          key={b.requestType}
                          className={`relative overflow-hidden rounded-lg border border-line bg-surface p-3 before:absolute before:inset-x-0 before:top-0 before:h-[3px] ${
                            i === 0 ? "before:bg-accent" : "before:bg-[#6b7cb0]"
                          }`}
                        >
                          <div className="font-mono text-2xl leading-none font-bold tracking-tight tabular-nums text-ink">
                            {b.remainingDays}
                            <small className="ml-0.5 text-[13px] font-medium text-ink-3">days</small>
                          </div>
                          <div className="mt-1 text-xs text-ink-2">{b.actionTypeName} remaining</div>
                        </div>
                      ))}
                    </div>
                  )}

                  {report.yearlyBreakdown.length > 0 && (
                    <div className="flex flex-col gap-1.5">
                      <div className="px-0.5 text-[11px] font-bold tracking-wider text-ink-3 uppercase">By year</div>
                      <div className="overflow-x-auto rounded-lg border border-line">
                        <table className="w-full text-xs">
                          <thead>
                            <tr className="border-b border-line bg-surface-2 text-left text-ink-2">
                              <th className="px-2.5 py-1.5 font-semibold">Year</th>
                              <th className="px-2.5 py-1.5 font-semibold">Holiday allocated</th>
                              <th className="px-2.5 py-1.5 font-semibold">Total taken</th>
                              <th className="px-2.5 py-1.5 font-semibold">Breakdown</th>
                            </tr>
                          </thead>
                          <tbody>
                            {report.yearlyBreakdown.map((y) => (
                              <tr key={y.year} className="border-b border-line last:border-0">
                                <td className="px-2.5 py-1.5 font-mono tabular-nums text-ink">{y.year}</td>
                                <td className="px-2.5 py-1.5 font-mono tabular-nums text-ink">
                                  {y.holidayAllocatedThisYear ?? "—"}
                                </td>
                                <td className="px-2.5 py-1.5 font-mono tabular-nums text-ink">{y.totalTaken}</td>
                                <td className="px-2.5 py-1.5 text-ink-2">
                                  {y.takenByType.map((t) => `${t.actionTypeName} ${t.days}`).join(", ") || "—"}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  <div className="flex flex-col gap-1.5">
                    <div className="flex flex-wrap items-center justify-between gap-2">
                      <div className="px-0.5 text-[11px] font-bold tracking-wider text-ink-3 uppercase">History</div>
                      <div className="flex flex-wrap items-center gap-2">
                        <label className="flex items-center gap-1.5 text-xs font-semibold text-ink-2">
                          Year
                          <select
                            value={reportYearFilter}
                            onChange={(e) => setReportYearFilter(e.target.value === "all" ? "all" : Number(e.target.value))}
                            className={fieldClass}
                          >
                            <option value="all">All</option>
                            {getReportAvailableYears().map((y) => (
                              <option key={y} value={y}>
                                {y}
                              </option>
                            ))}
                          </select>
                        </label>
                        <label className="flex items-center gap-1.5 text-xs font-semibold text-ink-2">
                          Month
                          <select
                            value={reportMonthFilter}
                            onChange={(e) => setReportMonthFilter(e.target.value === "all" ? "all" : Number(e.target.value))}
                            className={fieldClass}
                          >
                            <option value="all">All</option>
                            {MONTH_NAMES.map((name, i) => (
                              <option key={name} value={i}>
                                {name}
                              </option>
                            ))}
                          </select>
                        </label>
                      </div>
                    </div>
                    <div className="flex flex-wrap items-center justify-end gap-1">
                      <div className="flex flex-wrap gap-1">
                        {(["PENDING", "APPROVED", "DECLINED", "CANCELLED"] as const).map((status) => (
                          <button
                            key={status}
                            type="button"
                            onClick={() => toggleReportStatus(status)}
                            className={`rounded-full border px-2 py-0.5 text-[10px] font-bold tracking-wide uppercase transition-colors ${
                              reportStatusFilter.includes(status)
                                ? `border-transparent ${TONE_SOFT[toneForStatus(status)]}`
                                : "border-line text-ink-3 hover:text-ink"
                            }`}
                          >
                            {status}
                          </button>
                        ))}
                      </div>
                    </div>
                    {(() => {
                      const filteredHistory = getFilteredReportHistory();
                      return filteredHistory.length === 0 ? (
                        <p className="text-sm text-ink-2">
                          {reportStatusFilter.length === 0
                            ? "Select at least one status above to show requests."
                            : "No leave requests match the selected statuses."}
                        </p>
                      ) : (
                        <ul className="flex max-h-72 flex-col gap-1.5 overflow-y-auto">
                          {filteredHistory.map((h) => {
                            const tone = toneForStatus(h.status);
                            return (
                              <li
                                key={h.id}
                                className="rounded-lg border border-line border-l-[3px] bg-surface p-2.5 text-sm"
                                style={{ borderLeftColor: `var(--${tone === "neutral" ? "neutral" : tone})` }}
                              >
                                <div className="flex items-baseline justify-between gap-2">
                                  <div className="font-semibold text-ink">{h.requestSubject}</div>
                                  <span
                                    className={`inline-flex flex-none items-center gap-1.5 rounded-full px-2 py-0.5 text-[10px] font-bold tracking-wide uppercase ${TONE_SOFT[tone]}`}
                                  >
                                    <span className={`h-1.5 w-1.5 rounded-full ${TONE_BG[tone]}`} />
                                    {h.status}
                                  </span>
                                </div>
                                <div className="mt-1 flex flex-wrap items-center gap-x-2 gap-y-0.5 text-xs text-ink-2">
                                  <span className="font-mono tabular-nums text-ink">
                                    {formatDate(h.startDate)} – {formatDate(h.endDate)}
                                  </span>
                                  <span className="text-line">·</span>
                                  <span className="font-mono tabular-nums">{h.takenDays} days</span>
                                  <span className="text-line">·</span>
                                  <span>{h.actionTypeName}</span>
                                </div>

                                {h.id.startsWith("archive-") ? (
                                  <p className="mt-2 text-xs text-ink-3">Archived before this app existed — no approval trail available.</p>
                                ) : (
                                  <>
                                    <div className="mt-2 flex items-center justify-end">
                                      <button type="button" onClick={() => handleToggleHistory(h.id)} className={linkBtn}>
                                        {expandedHistoryId === h.id ? "Hide details" : "Details"}
                                      </button>
                                    </div>
                                    {expandedHistoryId === h.id && (
                                      <div className="mt-2 border-t border-dashed border-line pt-2.5">
                                        {historyLoadingId === h.id ? (
                                          <p className="text-xs text-ink-2">Loading…</p>
                                        ) : historyError ? (
                                          <p className="text-xs text-critical">{historyError}</p>
                                        ) : (historyByRequest[h.id]?.length ?? 0) === 0 ? (
                                          <p className="text-xs text-ink-2">No actions recorded yet.</p>
                                        ) : (
                                          <ul className="flex flex-col gap-2.5">
                                            {historyByRequest[h.id].map((entry, i, arr) => (
                                              <li key={i} className="flex gap-2.5">
                                                <div className="flex w-2 flex-none flex-col items-center pt-0.5">
                                                  <span className="h-1.5 w-1.5 flex-none rounded-full bg-accent" />
                                                  {i < arr.length - 1 && <span className="mt-1 w-px flex-1 bg-line" />}
                                                </div>
                                                <div className="pb-0.5">
                                                  <div className="font-mono text-[11px] tabular-nums text-ink-3">
                                                    {formatDate(entry.timestamp)}
                                                  </div>
                                                  <div className="text-xs font-semibold text-ink">{entry.actorName}</div>
                                                  <div className="text-xs text-ink-2">{entry.details}</div>
                                                </div>
                                              </li>
                                            ))}
                                          </ul>
                                        )}
                                      </div>
                                    )}
                                  </>
                                )}
                              </li>
                            );
                          })}
                        </ul>
                      );
                    })()}
                  </div>
                </>
              ) : (
                <p className="text-sm text-ink-2">Select an employee to view their leave report.</p>
              )}
            </div>
          )}

          {activeTab === "hr" && context.canAccessSettings && (
            <div className={`${cardClass} flex flex-col gap-3`}>
              <div className="flex items-center justify-between">
                <p className="text-[11px] font-bold tracking-wider text-ink-3 uppercase">Approval Chain</p>
                <button type="button" onClick={handleOpenChainEditor} className={btnGhost}>
                  {showChainEditor ? "Hide" : "Show"}
                </button>
              </div>

              {showChainEditor && (
                <>
                  <label className={labelClass}>
                    Edit chain for
                    <SearchableSelect
                      value={chainEmployeeSelection}
                      onChange={handleChangeChainEmployee}
                      options={[
                        { value: "", label: "Default chain (all employees)" },
                        ...employees.map((emp) => ({
                          value: emp.mail,
                          label: employeesWithOverride.includes(emp.mail.toLowerCase())
                            ? `${emp.displayName} — custom chain`
                            : emp.displayName,
                        })),
                      ]}
                      className={fieldClass}
                    />
                  </label>

                  {employeesWithOverride.length > 0 && (
                    <p className="text-xs text-ink-3">
                      Custom chains:{" "}
                      {employeesWithOverride
                        .map((mail) => nameByMail.get(mail) ?? mail)
                        .sort((a, b) => a.localeCompare(b))
                        .join(", ")}
                    </p>
                  )}

                  {chainEmployeeSelection && (
                    <p className="text-xs text-ink-3">
                      {chainIsOverride
                        ? "This employee has their own chain — remove every approver below and save to revert them to the default chain."
                        : "This employee uses the default chain — save to create one just for them."}
                    </p>
                  )}

                  {chainLoading ? (
                    <p className="text-sm text-ink-2">Loading…</p>
                  ) : (
                    <>
                      {chainDraft.length === 0 ? (
                        <p className="text-sm text-ink-2">No approvers added.</p>
                      ) : (
                        <ol className="flex flex-col gap-1">
                          {chainDraft.map((email, i) => (
                            <li
                              key={email}
                              className="flex items-center justify-between rounded-lg border border-line bg-surface-2 px-2.5 py-1.5 text-sm"
                            >
                              <span className="text-ink">
                                <span className="font-mono text-ink-3">{i + 1}.</span> {labelForChainEntry(email, nameByMail)}
                              </span>
                              <button
                                type="button"
                                onClick={() => handleRemoveFromChain(email)}
                                className="text-xs font-semibold text-critical hover:underline"
                              >
                                Remove
                              </button>
                            </li>
                          ))}
                        </ol>
                      )}

                      <div className="flex gap-2">
                        <SearchableSelect
                          value={chainAddSelection}
                          onChange={setChainAddSelection}
                          placeholder="Add approver…"
                          options={[
                            { value: MANAGER_TOKEN, label: "Direct Manager" },
                            { value: GM_TOKEN, label: "GM" },
                            ...employees.map((emp) => ({ value: emp.mail, label: emp.displayName })),
                          ]}
                          className={fieldClass}
                          wrapperClassName="flex-1"
                        />
                        <button type="button" onClick={handleAddToChain} className={btnGhost}>
                          Add
                        </button>
                      </div>

                      {chainError && <p className="text-xs text-critical">{chainError}</p>}

                      <button type="button" onClick={handleSaveChain} disabled={chainSubmitting} className={`${btnPrimary} self-start`}>
                        {chainSubmitting ? "Saving…" : "Save"}
                      </button>
                    </>
                  )}
                </>
              )}
            </div>
          )}


          {activeTab === "hr" && context.canAccessSettings && (
            <div className={`${cardClass} flex flex-col gap-3`}>
              <div className="flex items-center justify-between">
                <p className="text-[11px] font-bold tracking-wider text-ink-3 uppercase">Leave-day Allocations</p>
                <button type="button" onClick={handleOpenAllocations} className={btnGhost}>
                  {showAllocations ? "Hide" : "Show"}
                </button>
              </div>

              {showAllocations && (
                <>
                  <div className="flex items-center justify-between gap-2">
                    <div className="flex items-center gap-2">
                      <button type="button" onClick={() => handleChangeAllocationYear(-1)} className={btnGhost}>
                        ←
                      </button>
                      <span className="font-mono text-sm font-semibold text-ink tabular-nums">{allocationYear}</span>
                      <button type="button" onClick={() => handleChangeAllocationYear(1)} className={btnGhost}>
                        →
                      </button>
                    </div>
                    <button type="button" onClick={handleBulkAllocate} disabled={bulkAllocating} className={btnPrimary}>
                      {bulkAllocating ? "Allocating…" : `Allocate for ${allocationYear}`}
                    </button>
                  </div>

                  {bulkAllocateSummary && <p className="text-xs text-good">{bulkAllocateSummary}</p>}
                  {allocationError && <p className="text-xs text-critical">{allocationError}</p>}

                  {allocationsLoading ? (
                    <p className="text-sm text-ink-2">Loading…</p>
                  ) : allocationRows.length === 0 ? (
                    <p className="text-sm text-ink-2">No employees found.</p>
                  ) : (
                    <ul className="flex max-h-80 flex-col gap-1.5 overflow-y-auto">
                      {allocationRows.map((row) => (
                        <li
                          key={row.employeeCode}
                          className="flex flex-col gap-1 rounded-lg border border-line bg-surface-2 px-2.5 py-2 text-sm"
                        >
                          <div className="flex items-start justify-between gap-2">
                            <span className="font-semibold text-ink">{row.employeeFullName}</span>
                            {!row.isOnboarded ? (
                              <button
                                type="button"
                                onClick={() => handleOnboardEmployee(row.employeeCode)}
                                disabled={onboardingCode === row.employeeCode}
                                className={linkBtn}
                              >
                                {onboardingCode === row.employeeCode ? "Onboarding…" : "Onboard"}
                              </button>
                            ) : (
                              <div className="flex flex-col items-end gap-1">
                                <div className="flex items-center gap-2">
                                  <span className="text-xs text-ink-2">
                                    Holiday:{" "}
                                    <span className="font-mono tabular-nums text-ink">{row.allocatedThisYear ?? "—"}</span>
                                    {` in ${allocationYear}`}
                                  </span>
                                  {editingAllocationCode === row.employeeCode ? (
                                    <div className="flex items-center gap-2">
                                      <input
                                        type="number"
                                        min={0}
                                        value={editAllocationValue}
                                        onChange={(e) => setEditAllocationValue(e.target.value)}
                                        className={`${fieldClass} w-16 font-mono`}
                                      />
                                      <button
                                        type="button"
                                        onClick={() => handleSaveAllocation(row.employeeCode)}
                                        disabled={savingAllocationCode === row.employeeCode}
                                        className={linkBtn}
                                      >
                                        {savingAllocationCode === row.employeeCode ? "Saving…" : "Save"}
                                      </button>
                                      <button type="button" onClick={() => setEditingAllocationCode(null)} className={linkBtn}>
                                        Cancel
                                      </button>
                                    </div>
                                  ) : (
                                    <button type="button" onClick={() => handleStartEditAllocation(row)} className={linkBtn}>
                                      Edit
                                    </button>
                                  )}
                                </div>

                                <div className="flex items-center gap-2">
                                  <span className="text-xs text-ink-2">
                                    Bonus:{" "}
                                    <span className="font-mono tabular-nums text-ink">{row.bonusAllocatedThisYear ?? "—"}</span>
                                    {` in ${allocationYear}`}
                                  </span>
                                  {editingBonusCode === row.employeeCode ? (
                                    <div className="flex items-center gap-2">
                                      <input
                                        type="number"
                                        min={0}
                                        value={editBonusValue}
                                        onChange={(e) => setEditBonusValue(e.target.value)}
                                        className={`${fieldClass} w-16 font-mono`}
                                      />
                                      <button
                                        type="button"
                                        onClick={() => handleSaveBonus(row.employeeCode)}
                                        disabled={savingBonusCode === row.employeeCode}
                                        className={linkBtn}
                                      >
                                        {savingBonusCode === row.employeeCode ? "Saving…" : "Save"}
                                      </button>
                                      <button type="button" onClick={() => setEditingBonusCode(null)} className={linkBtn}>
                                        Cancel
                                      </button>
                                    </div>
                                  ) : (
                                    <button type="button" onClick={() => handleStartEditBonus(row)} className={linkBtn}>
                                      Edit
                                    </button>
                                  )}
                                </div>
                              </div>
                            )}
                          </div>

                          {!row.isOnboarded && <span className="text-xs text-ink-3">Not yet onboarded</span>}
                        </li>
                      ))}
                    </ul>
                  )}
                </>
              )}
            </div>
          )}

          {activeTab === "hr" && context.canAccessSettings && (
            <div className={`${cardClass} flex flex-col gap-3`}>
              <div className="flex items-center justify-between">
                <p className="text-[11px] font-bold tracking-wider text-ink-3 uppercase">
                  Legal Holidays <span className="font-normal normal-case text-ink-3">(last / current / next year)</span>
                </p>
                <button type="button" onClick={handleOpenHolidays} className={btnGhost}>
                  {showHolidays ? "Hide" : "Show"}
                </button>
              </div>

              {showHolidays && (
                <>
                  {holidaysLoading ? (
                    <p className="text-sm text-ink-2">Loading…</p>
                  ) : (
                    <>
                      <button
                        type="button"
                        onClick={() => setShowHolidayForm((v) => !v)}
                        className={`${btnGhost} self-start`}
                      >
                        {showHolidayForm ? "Cancel" : "Add holiday"}
                      </button>

                      {showHolidayForm && (
                        <form onSubmit={handleAddHoliday} className="flex items-end gap-2 rounded-lg border border-line bg-surface-2 p-3">
                          <label className={`flex-1 ${labelClass}`}>
                            Name
                            <input
                              required
                              type="text"
                              value={holidayName}
                              onChange={(e) => setHolidayName(e.target.value)}
                              className={fieldClass}
                            />
                          </label>
                          <label className={labelClass}>
                            Date
                            <input
                              required
                              type="date"
                              value={holidayDate}
                              onChange={(e) => setHolidayDate(e.target.value)}
                              className={`${fieldClass} font-mono`}
                            />
                          </label>
                          <button type="submit" disabled={holidaySubmitting} className={btnPrimary}>
                            {holidaySubmitting ? "…" : "Add"}
                          </button>
                        </form>
                      )}

                      {holidayError && <p className="text-xs text-critical">{holidayError}</p>}

                      <ul className="flex max-h-64 flex-col gap-1 overflow-y-auto">
                        {holidays.map((h) => (
                          <li
                            key={h.id}
                            className="flex items-center justify-between rounded-lg border border-line bg-surface-2 px-2.5 py-1.5 text-sm"
                          >
                            <span className="text-ink-2">
                              {h.name} — <span className="font-mono tabular-nums text-ink">{formatDate(h.date)}</span>
                            </span>
                            <button
                              type="button"
                              disabled={deletingHolidayId === h.id}
                              onClick={() => handleDeleteHoliday(h.id)}
                              className="text-xs font-semibold text-critical hover:underline disabled:opacity-50"
                            >
                              Delete
                            </button>
                          </li>
                        ))}
                      </ul>
                    </>
                  )}
                </>
              )}
            </div>
          )}
        </div>
      )}

      {/* Full-screen blocker while a bulk approve runs — each final-step approval takes a
          few seconds, so without it the app looks frozen and invites double-clicks. */}
      {bulkApproving && bulkProgress && (
        <div
          role="dialog"
          aria-modal="true"
          aria-live="polite"
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 px-4"
        >
          <div className="w-full max-w-xs rounded-xl border border-line bg-surface p-5 text-center shadow-lg">
            <div className="mx-auto mb-3 h-8 w-8 animate-spin rounded-full border-4 border-line border-t-accent" />
            <p className="text-sm font-semibold text-ink">Work in progress…</p>
            <p className="mt-1 text-sm text-ink-2">
              {bulkProgress.done} of {bulkProgress.total} processed
            </p>
            <div className="mt-3 h-1.5 overflow-hidden rounded-full bg-surface-2">
              <div
                className="h-full bg-accent transition-all"
                style={{ width: `${(bulkProgress.done / bulkProgress.total) * 100}%` }}
              />
            </div>
          </div>
        </div>
      )}
    </main>
  );
}
