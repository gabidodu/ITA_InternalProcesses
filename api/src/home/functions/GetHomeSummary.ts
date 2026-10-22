import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getExpenseConfig, getTenantConfig, getTenantModules, ModuleKey } from "../../shared/lib/tenantConfig";
import { getAllListItems } from "../../shared/lib/graphPaging";
import { getSetupRows } from "../../shared/lib/setupBase";
import { getRelevantYears } from "../../shared/lib/dates";
import { buildDisplayNameMap, getEmployeeDirectory } from "../../shared/lib/employees";

type ListItem = { id: string; fields: Record<string, string> };

/** One row of the home screen's combined "awaiting my approval" list — just enough to show it and jump into its module. */
type PendingItem = {
    module: ModuleKey;
    id: string;
    subject: string;
    employeeCode: string;
    employeeName: string;
    /** When it was submitted (leave: RequestDate, expense: RequestDate) — the list is sorted oldest first, so whatever has waited longest is on top. */
    submittedDate: string;
    /** Short module-specific summary, e.g. "12–16 May · 5 days" or "1,240.00 RON". */
    detail: string;
};

type ModuleSummary = {
    key: ModuleKey;
    pendingCount: number;
    /** Set when this module's lists couldn't be read — the tile still shows, with a warning, instead of the whole home screen failing. */
    error?: string;
};

function isPendingFor(item: ListItem, principal: string): boolean {
    return item.fields.Status === "PENDING" && item.fields.CurrentUser?.toLowerCase() === principal.toLowerCase();
}

function formatShortDate(iso: string): string {
    return new Date(iso).toLocaleDateString("en-GB", { day: "numeric", month: "short", timeZone: "Europe/Bucharest" });
}

async function loadLeavePending(graphClient: Client, tenantId: string, principal: string): Promise<PendingItem[]> {
    const { siteId, leaveRequestsListId, setupListId } = await getTenantConfig(tenantId);
    const items = await getAllListItems(graphClient, siteId, leaveRequestsListId);
    // Same recent-years window as GetMyLeaveContext's pendingApprovals, so the home badge/list
    // never counts a request the Leave module's own Pending Approvals tab wouldn't show.
    const relevantYears = getRelevantYears();
    const pending = items.filter((item) => isPendingFor(item, principal) && relevantYears.has(Number(item.fields.Year)));
    if (pending.length === 0) return [];

    const nameByMail = buildDisplayNameMap(await getEmployeeDirectory(graphClient, await getSetupRows(graphClient, siteId, setupListId)));
    return pending.map((item) => {
        const days = Number(item.fields.TakenDays) || 0;
        return {
            module: "leave",
            id: item.id,
            subject: item.fields.RequestSubject,
            employeeCode: item.fields.EmployeeCode,
            employeeName: nameByMail.get(item.fields.EmployeeCode?.toLowerCase()) || item.fields.EmployeeCode,
            submittedDate: item.fields.RequestDate || item.fields.StartDate,
            detail: `${formatShortDate(item.fields.StartDate)} – ${formatShortDate(item.fields.EndDate)} · ${days} ${days === 1 ? "day" : "days"}`,
        };
    });
}

async function loadExpensePending(graphClient: Client, tenantId: string, principal: string): Promise<PendingItem[]> {
    const { siteId, expenseReportsListId, setupListId } = await getExpenseConfig(tenantId);
    // Report headers only — unlike GetMyExpenseContext, no lines/attachments are needed just to list and count them.
    const items = await getAllListItems(graphClient, siteId, expenseReportsListId);
    const pending = items.filter((item) => isPendingFor(item, principal));
    if (pending.length === 0) return [];

    const nameByMail = buildDisplayNameMap(await getEmployeeDirectory(graphClient, await getSetupRows(graphClient, siteId, setupListId)));
    return pending.map((item) => ({
        module: "expense",
        id: item.id,
        subject: item.fields.RequestSubject,
        employeeCode: item.fields.EmployeeCode,
        employeeName: nameByMail.get(item.fields.EmployeeCode?.toLowerCase()) || item.fields.EmployeeCode,
        submittedDate: item.fields.RequestDate,
        detail: `${(Number(item.fields.TotalAmount) || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ${item.fields.Currency ?? ""}`.trim(),
    }));
}

/**
 * Everything the home (request-type picker) screen needs in one call: which
 * modules this tenant is provisioned for (only those get a tile), and one
 * combined list of everything currently awaiting the signed-in user's
 * approval across all of them. Each module is loaded independently — one
 * module's lists being unreachable degrades to an error on that tile, never
 * a failed home screen.
 */
export async function GetHomeSummary(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, [
            "Sites.Selected",
            "GroupMember.Read.All",
            "User.ReadBasic.All",
        ]);
        const graphClient = Client.init({ authProvider: (done) => done(null, graphToken) });

        const [me, moduleKeys] = await Promise.all([
            graphClient.api("/me").select("displayName,mail,userPrincipalName").get(),
            // An unprovisioned tenant gets the "no modules set up" home screen, not an error.
            getTenantModules(tenantId).catch((error: unknown) => {
                if (error instanceof Error && error.message.startsWith("No tenant configuration")) return [] as ModuleKey[];
                throw error;
            }),
        ]);
        const principal: string = me.mail ?? me.userPrincipalName;

        const loaders: Record<ModuleKey, () => Promise<PendingItem[]>> = {
            leave: () => loadLeavePending(graphClient, tenantId, principal),
            expense: () => loadExpensePending(graphClient, tenantId, principal),
        };

        const results = await Promise.allSettled(moduleKeys.map((key) => loaders[key]()));

        const modules: ModuleSummary[] = [];
        const pending: PendingItem[] = [];
        results.forEach((result, i) => {
            const key = moduleKeys[i];
            if (result.status === "fulfilled") {
                modules.push({ key, pendingCount: result.value.length });
                pending.push(...result.value);
            } else {
                context.warn(`GetHomeSummary: ${key} pending load failed`, result.reason);
                modules.push({
                    key,
                    pendingCount: 0,
                    error: result.reason instanceof Error ? result.reason.message : "Could not load this module.",
                });
            }
        });

        pending.sort((a, b) => new Date(a.submittedDate).getTime() - new Date(b.submittedDate).getTime());

        return {
            jsonBody: {
                displayName: me.displayName,
                mail: principal,
                modules,
                pending,
            },
        };
    } catch (error) {
        context.error("GetHomeSummary failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Home_GetHomeSummary", {
    route: "home/GetHomeSummary",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetHomeSummary,
});
