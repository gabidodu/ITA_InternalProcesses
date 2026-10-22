import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows, canAccessSettings } from "../lib/setup";
import { buildDisplayNameMap, getEmployeeDirectory, getMyDirectReports } from "../../shared/lib/employees";
import { getRelevantYears } from "../../shared/lib/dates";
import { getAllListItems } from "../../shared/lib/graphPaging";
import { isBlockingStatus } from "../lib/requestStatus";

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

/**
 * Returns the signed-in user's profile plus two distinct request lists:
 * "myRequests" (everything they themselves submitted, by EmployeeCode,
 * regardless of status or who currently holds it) and "pendingApprovals"
 * (requests currently awaiting their action, by CurrentUser + Status ===
 * PENDING). These are genuinely different things — see [[project-leave-teams-app]]
 * for why conflating them (the old single-list behavior) was a quirk.
 */
export async function GetMyLeaveContext(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, [
            "Sites.Selected",
            "GroupMember.Read.All",
            "User.ReadBasic.All",
            "User.Read.All",
        ]);
        const { siteId, leaveRequestsListId, leaveStatusListId, actionTypeListId, setupListId } =
            await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("displayName,mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const nameByMail = buildDisplayNameMap(employees);
        const resolveName = (mail: string | undefined) =>
            (mail && nameByMail.get(mail.toLowerCase())) || mail || "";

        const leaveRequestItems = await getAllListItems(graphClient, siteId, leaveRequestsListId);

        const toLeaveRequest = (item: { id: string; fields: Record<string, string> }): LeaveRequest => ({
            id: item.id,
            subject: item.fields.RequestSubject,
            startDate: item.fields.StartDate,
            endDate: item.fields.EndDate,
            takenDays: Number(item.fields.TakenDays),
            status: item.fields.Status,
            currentUser: item.fields.CurrentUser,
            currentUserName: resolveName(item.fields.CurrentUser),
        });

        const byStartDateDesc = (a: LeaveRequest, b: LeaveRequest) =>
            new Date(b.startDate).getTime() - new Date(a.startDate).getTime();
        const relevantYears = getRelevantYears();

        const myRequests: LeaveRequest[] = leaveRequestItems
            .filter(
                (item: { fields: Record<string, string> }) =>
                    item.fields.EmployeeCode?.toLowerCase() === currentUserPrincipal?.toLowerCase() &&
                    relevantYears.has(Number(item.fields.Year))
            )
            .map(toLeaveRequest)
            .sort(byStartDateDesc);

        // Unlike myRequests, these belong to someone else — the approver needs to see whose.
        const pendingApprovals: (LeaveRequest & { employeeCode: string; employeeName: string })[] = leaveRequestItems
            .filter(
                (item: { fields: Record<string, string> }) =>
                    item.fields.CurrentUser?.toLowerCase() === currentUserPrincipal?.toLowerCase() &&
                    item.fields.Status === "PENDING" &&
                    relevantYears.has(Number(item.fields.Year))
            )
            .map((item: { id: string; fields: Record<string, string> }) => ({
                ...toLeaveRequest(item),
                employeeCode: item.fields.EmployeeCode,
                employeeName: resolveName(item.fields.EmployeeCode),
            }))
            .sort(byStartDateDesc);

        // Holiday/Bonus have a settable max, so their remaining balance comes
        // from LR_LeaveStatus (AllocDays - TakenDays). Medical/Unpaid have no
        // max — instead we show how many days were used THIS YEAR, computed
        // from actual approved requests (LR_LeaveStatus is cumulative-forever,
        // not per-year, so it can't answer "this year" on its own).
        const actionTypeItems = await getAllListItems(graphClient, siteId, actionTypeListId);
        const actionTypeById = new Map<number, string>();
        actionTypeItems.forEach((item: { fields: Record<string, string> }) => {
            actionTypeById.set(Number(item.fields.IDActionType), item.fields.ActionTypeCode);
        });

        const leaveStatusItems = await getAllListItems(graphClient, siteId, leaveStatusListId);

        const balances: LeaveBalance[] = leaveStatusItems
            .filter(
                (item: { fields: Record<string, string> }) =>
                    item.fields.EmployeeCode?.toLowerCase() === currentUserPrincipal?.toLowerCase() &&
                    (item.fields.ActionTypeName === "Holiday" || item.fields.ActionTypeName === "Bonus Days")
            )
            .map((item: { fields: Record<string, string> }) => {
                const allocDays = Number(item.fields.AllocDays) || 0;
                const takenDays = Number(item.fields.TakenDays) || 0;
                return {
                    requestType: Number(item.fields.RequestType),
                    actionTypeName: item.fields.ActionTypeName,
                    allocDays,
                    remainingDays: allocDays - takenDays,
                };
            })
            .sort((a: LeaveBalance, b: LeaveBalance) => a.requestType - b.requestType);

        const currentYear = new Date().getFullYear();
        const usedThisYearTotals = new Map<number, { actionTypeName: string; total: number }>();
        leaveRequestItems
            .filter(
                (item: { fields: Record<string, string> }) =>
                    item.fields.EmployeeCode?.toLowerCase() === currentUserPrincipal?.toLowerCase() &&
                    isBlockingStatus(item.fields.Status) &&
                    Number(item.fields.Year) === currentYear &&
                    ["CM", "CFP"].includes(actionTypeById.get(Number(item.fields.RequestType)) ?? "")
            )
            .forEach((item: { fields: Record<string, string> }) => {
                const requestType = Number(item.fields.RequestType);
                const existing = usedThisYearTotals.get(requestType);
                const actionTypeName =
                    actionTypeById.get(requestType) === "CM" ? "Medical Leave" : "Unpaid";
                usedThisYearTotals.set(requestType, {
                    actionTypeName,
                    total: (existing?.total ?? 0) + (Number(item.fields.TakenDays) || 0),
                });
            });
        const usedThisYear: UsedThisYear[] = Array.from(usedThisYearTotals.entries())
            .map(([requestType, v]) => ({ requestType, actionTypeName: v.actionTypeName, takenDaysThisYear: v.total }))
            .sort((a, b) => a.requestType - b.requestType);

        const isSettingsAdmin = canAccessSettings(setupRows, currentUserPrincipal);
        let canViewReports = isSettingsAdmin;
        if (!canViewReports) {
            try {
                canViewReports = (await getMyDirectReports(graphClient)).length > 0;
            } catch (directReportsError) {
                context.warn("getMyDirectReports failed (non-blocking, treated as no reports)", directReportsError);
            }
        }

        return {
            jsonBody: {
                displayName: me.displayName,
                mail: currentUserPrincipal,
                myRequests,
                pendingApprovals,
                balances,
                usedThisYear,
                canAccessSettings: isSettingsAdmin,
                canViewReports,
            },
        };
    } catch (error) {
        context.error("GetMyLeaveContext failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_GetMyLeaveContext", {
    route: "leave/GetMyLeaveContext",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetMyLeaveContext,
});
