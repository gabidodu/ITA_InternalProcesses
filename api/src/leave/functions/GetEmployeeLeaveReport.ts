import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows, canAccessSettings } from "../lib/setup";
import { getEmployeeDirectory, buildDisplayNameMap, isDirectManagerOf } from "../../shared/lib/employees";
import { getAllListItems } from "../../shared/lib/graphPaging";
import { isBlockingStatus } from "../lib/requestStatus";

type CurrentBalance = {
    requestType: number;
    actionTypeName: string;
    allocDays: number;
    takenDays: number;
    remainingDays: number;
};

type YearlyBreakdown = {
    year: number;
    holidayAllocatedThisYear: number | null;
    takenByType: { actionTypeName: string; days: number }[];
    totalTaken: number;
};

type HistoryEntry = {
    id: string;
    actionTypeName: string;
    requestSubject: string;
    startDate: string;
    endDate: string;
    takenDays: number;
    status: string;
    year: number;
};

/**
 * Returns a compact leave-history report for one employee: current cumulative
 * balances (Holiday/Bonus), a per-year breakdown (Holiday's per-year grant
 * alongside days taken by type that year, computed from actually-approved
 * requests since LR_LeaveStatus itself is cumulative-forever and can't answer
 * "this year" on its own), and the full itemized request history. Viewable by
 * HR/Admin/GM (canAccessSettings) or the employee's own direct manager
 * (live Entra ID lookup) — nobody else.
 */
export async function GetEmployeeLeaveReport(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const employeeCode = request.query.get("employeeCode");
        if (!employeeCode) {
            throw new Error("employeeCode query parameter is required.");
        }

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, [
            "Sites.Selected",
            "GroupMember.Read.All",
            "User.ReadBasic.All",
            "User.Read.All",
        ]);
        const { siteId, leaveRequestsListId, leaveRequestsArchiveListId, leaveStatusListId, leaveSetupListId, actionTypeListId, setupListId } =
            await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const isAuthorized =
            canAccessSettings(setupRows, currentUserPrincipal) ||
            (await isDirectManagerOf(graphClient, currentUserPrincipal, employeeCode));
        if (!isAuthorized) {
            return {
                status: 403,
                jsonBody: { error: "Only HR, Admin, GM, or this employee's own manager can view this report." },
            };
        }

        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const nameByMail = buildDisplayNameMap(employees);
        const employeeFullName = nameByMail.get(employeeCode.toLowerCase()) ?? employeeCode;

        let department: string | undefined;
        let jobTitle: string | undefined;
        try {
            const profile = await graphClient
                .api(`/users/${encodeURIComponent(employeeCode)}`)
                .select("department,jobTitle")
                .get();
            department = profile.department;
            jobTitle = profile.jobTitle;
        } catch (profileError) {
            context.warn("Failed to fetch employee AD profile for report (non-blocking)", profileError);
        }

        const actionTypeItems = await getAllListItems(graphClient, siteId, actionTypeListId);
        const actionTypeById = new Map<number, string>();
        actionTypeItems.forEach((item: { fields: Record<string, string> }) => {
            actionTypeById.set(Number(item.fields.IDActionType), item.fields.ActionTypeName);
        });

        const leaveStatusItems = await getAllListItems(graphClient, siteId, leaveStatusListId);
        const currentBalances: CurrentBalance[] = leaveStatusItems
            .filter(
                (item: { fields: Record<string, string> }) =>
                    item.fields.EmployeeCode?.toLowerCase() === employeeCode.toLowerCase() &&
                    (item.fields.ActionTypeName === "Holiday" || item.fields.ActionTypeName === "Bonus Days")
            )
            .map((item: { fields: Record<string, string> }) => {
                const allocDays = Number(item.fields.AllocDays) || 0;
                const takenDays = Number(item.fields.TakenDays) || 0;
                return {
                    requestType: Number(item.fields.RequestType),
                    actionTypeName: item.fields.ActionTypeName,
                    allocDays,
                    takenDays,
                    remainingDays: allocDays - takenDays,
                };
            })
            .sort((a: CurrentBalance, b: CurrentBalance) => a.requestType - b.requestType);

        const leaveSetupItems = await getAllListItems(graphClient, siteId, leaveSetupListId);
        const holidayAllocByYear = new Map<number, number>();
        leaveSetupItems
            .filter((item: { fields: Record<string, string> }) => item.fields.EmployeeCode?.toLowerCase() === employeeCode.toLowerCase())
            .forEach((item: { fields: Record<string, string> }) => {
                const raw = item.fields.HolidayDaysAllocatedThisYear;
                if (raw !== undefined && raw !== null && raw !== ("" as unknown)) {
                    holidayAllocByYear.set(Number(item.fields.Year), Number(raw));
                }
            });

        const leaveRequestItems = await getAllListItems(graphClient, siteId, leaveRequestsListId);

        // Historical requests some tenants (e.g. IT ASSIST's own, migrated from the
        // legacy site) split into a separate archive list before this app existed.
        // Only the report merges it in — live balance checks elsewhere only care
        // about currently-open requests, which are never archived.
        const archiveItems = leaveRequestsArchiveListId
            ? (await getAllListItems(graphClient, siteId, leaveRequestsArchiveListId)).map(
                  (item: { id: string; fields: Record<string, string> }) => ({ ...item, id: `archive-${item.id}` })
              )
            : [];

        const employeeRequests = [...leaveRequestItems, ...archiveItems].filter(
            (item: { fields: Record<string, string> }) => item.fields.EmployeeCode?.toLowerCase() === employeeCode.toLowerCase()
        );

        const history: HistoryEntry[] = employeeRequests
            .map((item: { id: string; fields: Record<string, string> }) => ({
                id: item.id,
                actionTypeName: actionTypeById.get(Number(item.fields.RequestType)) ?? "Unknown",
                requestSubject: item.fields.RequestSubject,
                startDate: item.fields.StartDate,
                endDate: item.fields.EndDate,
                takenDays: Number(item.fields.TakenDays) || 0,
                status: item.fields.Status,
                year: Number(item.fields.Year),
            }))
            .sort((a: HistoryEntry, b: HistoryEntry) => new Date(b.startDate).getTime() - new Date(a.startDate).getTime());

        const takenByYearAndType = new Map<number, Map<string, number>>();
        employeeRequests
            .filter((item: { fields: Record<string, string> }) => isBlockingStatus(item.fields.Status))
            .forEach((item: { fields: Record<string, string> }) => {
                const year = Number(item.fields.Year);
                const actionTypeName = actionTypeById.get(Number(item.fields.RequestType)) ?? "Unknown";
                const takenDays = Number(item.fields.TakenDays) || 0;
                if (!takenByYearAndType.has(year)) {
                    takenByYearAndType.set(year, new Map());
                }
                const byType = takenByYearAndType.get(year)!;
                byType.set(actionTypeName, (byType.get(actionTypeName) ?? 0) + takenDays);
            });

        const allYears = new Set<number>([...holidayAllocByYear.keys(), ...takenByYearAndType.keys()]);
        const yearlyBreakdown: YearlyBreakdown[] = Array.from(allYears)
            .sort((a, b) => b - a)
            .map((year) => {
                const byType = takenByYearAndType.get(year);
                const takenByType = byType
                    ? Array.from(byType.entries())
                          .map(([actionTypeName, days]) => ({ actionTypeName, days }))
                          .sort((a, b) => b.days - a.days)
                    : [];
                return {
                    year,
                    holidayAllocatedThisYear: holidayAllocByYear.has(year) ? holidayAllocByYear.get(year)! : null,
                    takenByType,
                    totalTaken: takenByType.reduce((sum, t) => sum + t.days, 0),
                };
            });

        return {
            jsonBody: {
                employeeCode,
                employeeFullName,
                department,
                jobTitle,
                currentBalances,
                yearlyBreakdown,
                history,
            },
        };
    } catch (error) {
        context.error("GetEmployeeLeaveReport failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_GetEmployeeLeaveReport", {
    route: "leave/GetEmployeeLeaveReport",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetEmployeeLeaveReport,
});
