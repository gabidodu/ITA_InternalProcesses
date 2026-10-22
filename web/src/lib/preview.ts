/**
 * Canned API responses for local preview mode (see isPreview in
 * lib/teams.ts) — enough sample data for the home screen and each module's
 * main tabs to render, so the layout and navigation can be looked at in a
 * plain browser with no Teams, SSO or SharePoint. Writes (POST) pretend to
 * succeed and change nothing. Never bundled into a real build's behavior:
 * every caller checks isPreview first.
 */

const ME = { displayName: "Gabi Balaban", mail: "gabi.balaban@itassist.eu" };

const EMPLOYEES = [
  { mail: "gabi.balaban@itassist.eu", displayName: "Gabi Balaban" },
  { mail: "ana.popescu@itassist.eu", displayName: "Ana Popescu" },
  { mail: "mihai.ionescu@itassist.eu", displayName: "Mihai Ionescu" },
  { mail: "elena.dumitru@itassist.eu", displayName: "Elena Dumitru" },
];

function isoDaysFromNow(days: number): string {
  const d = new Date();
  d.setUTCHours(0, 0, 0, 0);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString();
}

const year = new Date().getFullYear();

const leaveContext = {
  ...ME,
  myRequests: [
    { id: "101", subject: "Concediu de odihnă", startDate: isoDaysFromNow(14), endDate: isoDaysFromNow(18), takenDays: 5, status: "PENDING", currentUser: "elena.dumitru@itassist.eu", currentUserName: "Elena Dumitru" },
    { id: "87", subject: "Zi liberă", startDate: isoDaysFromNow(-30), endDate: isoDaysFromNow(-30), takenDays: 1, status: "APPROVED", currentUser: "", currentUserName: "" },
  ],
  pendingApprovals: [
    { id: "112", subject: "Concediu de odihnă", startDate: isoDaysFromNow(7), endDate: isoDaysFromNow(11), takenDays: 5, status: "PENDING", currentUser: ME.mail, currentUserName: ME.displayName, employeeCode: "ana.popescu@itassist.eu", employeeName: "Ana Popescu" },
    { id: "115", subject: "Medical", startDate: isoDaysFromNow(2), endDate: isoDaysFromNow(3), takenDays: 2, status: "PENDING", currentUser: ME.mail, currentUserName: ME.displayName, employeeCode: "mihai.ionescu@itassist.eu", employeeName: "Mihai Ionescu" },
  ],
  balances: [
    { requestType: 1, actionTypeName: "Holiday", allocDays: 21, remainingDays: 9 },
    { requestType: 2, actionTypeName: "Bonus Days", allocDays: 2, remainingDays: 2 },
  ],
  usedThisYear: [
    { requestType: 3, actionTypeName: "Medical", takenDaysThisYear: 3 },
    { requestType: 4, actionTypeName: "Unpaid", takenDaysThisYear: 0 },
  ],
  canAccessSettings: true,
  canViewReports: true,
};

function expenseReport(id: string, subject: string, daysAgo: number, total: number, status: string, employee: (typeof EMPLOYEES)[number], currentUser = "") {
  return {
    id,
    subject,
    requestDate: isoDaysFromNow(-daysAgo),
    totalAmount: total,
    currency: "RON",
    status,
    currentUser,
    currentUserName: EMPLOYEES.find((e) => e.mail === currentUser)?.displayName ?? "",
    employeeCode: employee.mail,
    employeeName: employee.displayName,
    comments: "",
    lines: [
      { id: `${id}-1`, lineNumber: 1, expenseDate: isoDaysFromNow(-daysAgo - 2), categoryId: 1, categoryName: "Transport", vendor: "OMV", description: "Carburant deplasare client", amount: total * 0.6, currency: "RON", exchangeRate: 1 },
      { id: `${id}-2`, lineNumber: 2, expenseDate: isoDaysFromNow(-daysAgo - 1), categoryId: 2, categoryName: "Accommodation", vendor: "Hotel Central", description: "Cazare 1 noapte", amount: total * 0.4, currency: "RON", exchangeRate: 1 },
    ],
    attachments: [],
  };
}

const expenseContext = {
  ...ME,
  myReports: [
    expenseReport("41", "Deplasare Cluj", 5, 1240, "PENDING", EMPLOYEES[0], "elena.dumitru@itassist.eu"),
    expenseReport("33", "Deplasare Iași", 40, 860.5, "APPROVED", EMPLOYEES[0]),
  ],
  pendingApprovals: [expenseReport("44", "Training Microsoft", 3, 2310, "PENDING", EMPLOYEES[2], ME.mail)],
  isHrAdmin: true,
  isFinance: true,
};

const RESPONSES: Record<string, unknown> = {
  "/home/GetHomeSummary": {
    ...ME,
    modules: [
      { key: "leave", pendingCount: leaveContext.pendingApprovals.length },
      { key: "expense", pendingCount: expenseContext.pendingApprovals.length },
    ],
    pending: [
      ...leaveContext.pendingApprovals.map((r) => ({
        module: "leave",
        id: r.id,
        subject: r.subject,
        employeeCode: r.employeeCode,
        employeeName: r.employeeName,
        submittedDate: isoDaysFromNow(-2),
        detail: `${r.takenDays} days`,
      })),
      ...expenseContext.pendingApprovals.map((r) => ({
        module: "expense",
        id: r.id,
        subject: r.subject,
        employeeCode: r.employeeCode,
        employeeName: r.employeeName,
        submittedDate: r.requestDate,
        detail: `${r.totalAmount.toFixed(2)} ${r.currency}`,
      })),
    ].sort((a, b) => new Date(a.submittedDate).getTime() - new Date(b.submittedDate).getTime()),
  },

  "/leave/GetMyLeaveContext": leaveContext,
  "/leave/GetActionTypes": {
    actionTypes: [
      { id: 1, name: "Holiday", nameRo: "Concediu de odihnă", code: "CO" },
      { id: 2, name: "Bonus Days", nameRo: "Zile bonus", code: "BON" },
      { id: 3, name: "Medical", nameRo: "Concediu medical", code: "CM" },
      { id: 4, name: "Unpaid", nameRo: "Concediu fără plată", code: "CFP" },
    ],
  },
  "/leave/GetPickableEmployees": { employees: EMPLOYEES },
  "/leave/GetReportableEmployees": { employees: EMPLOYEES },
  "/leave/GetHolidays": {
    holidays: [
      { id: "1", name: "Ziua Națională", date: `${year}-12-01T00:00:00Z`, year, type: "Legal" },
      { id: "2", name: "Crăciunul", date: `${year}-12-25T00:00:00Z`, year, type: "Legal" },
    ],
  },
  "/leave/GetParticipatedRequests": { requests: [] },
  "/leave/GetRequestHistory": { history: [] },
  "/leave/GetApprovalFlow": { chain: ["{manager}", "{GM}"], isOverride: false, employeesWithOverrides: [] },
  "/leave/GetYearlyAllocations": {
    year,
    employees: EMPLOYEES.map((e) => ({ employeeCode: e.mail, employeeFullName: e.displayName, isOnboarded: true, allocatedThisYear: 21, bonusAllocatedThisYear: 2 })),
  },

  "/expense/GetMyExpenseContext": expenseContext,
  "/expense/GetExpenseCategories": {
    categories: [
      { id: 1, name: "Transport", nameRo: "Transport", code: "TRA" },
      { id: 2, name: "Accommodation", nameRo: "Cazare", code: "CAZ" },
      { id: 3, name: "Meals", nameRo: "Masă", code: "MAS" },
      { id: 4, name: "Cash Advance", nameRo: "Avans", code: "ADV" },
    ],
  },
  "/expense/GetPickableEmployees": { employees: EMPLOYEES },
  "/expense/GetApprovalFlow": { chain: ["{manager}", "{GM}"], isOverride: false, employeesWithOverrides: [] },
  "/expense/GetFinanceUsers": { emails: ["elena.dumitru@itassist.eu"] },
  "/expense/GetExpenseReportHistory": { history: [] },
  "/expense/GetAllExpenseReports": {
    year,
    reports: [...expenseContext.myReports, ...expenseContext.pendingApprovals],
    total: 3,
    totalAmount: 4410.5,
    hasMore: false,
  },
};

export async function previewApiFetch<T>(path: string, init?: RequestInit): Promise<T> {
  await new Promise((resolve) => setTimeout(resolve, 300));
  const route = path.split("?")[0];
  if (init?.method && init.method !== "GET") {
    // Every write in the real API answers with at least these; nothing is stored.
    return { emailSent: true, currentUser: ME.mail, chain: [], isOverride: false, emails: [], id: "0", lines: [] } as T;
  }
  const response = RESPONSES[route];
  if (response === undefined) {
    throw new Error(`Preview mode has no sample data for ${route}.`);
  }
  return structuredClone(response) as T;
}
