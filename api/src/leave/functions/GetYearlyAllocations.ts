import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows, canAccessSettings } from "../lib/setup";
import { getEmployeeDirectory } from "../../shared/lib/employees";
import { getActionTypeIdByCode } from "../lib/actionTypes";
import { getAllListItems } from "../../shared/lib/graphPaging";

type YearlyAllocationRow = {
    employeeCode: string;
    employeeFullName: string;
    isOnboarded: boolean;
    allocatedThisYear: number | null;
    bonusAllocatedThisYear: number | null;
};

/**
 * Returns, for a given year, every AD-group employee's allocation picture:
 * how many Holiday days (LR_LeaveSetup.HolidayDaysAllocatedThisYear) and
 * Bonus days (LR_LeaveSetup.BonusDays) they were granted for that specific
 * year — null if not yet set for that year. Both are per-year grant records,
 * same shape; the cumulative running total each feeds into (LR_LeaveStatus.
 * AllocDays) is bookkeeping for the balance check in CreateLeaveRequest, not
 * shown here. Powers the unified HR "employees + allocations" view.
 * Restricted to HR admins.
 */
export async function GetYearlyAllocations(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const yearParam = request.query.get("year");
        const year = yearParam ? Number(yearParam) : new Date().getFullYear();
        if (!Number.isFinite(year)) {
            throw new Error("year must be a number.");
        }

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, [
            "Sites.Selected",
            "GroupMember.Read.All",
            "User.ReadBasic.All",
        ]);
        const { siteId, leaveSetupListId, leaveStatusListId, actionTypeListId, setupListId } =
            await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        if (!canAccessSettings(setupRows, currentUserPrincipal)) {
            return { status: 403, jsonBody: { error: "Only HR, Admin, or GM can view leave-day allocations." } };
        }

        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const holidayTypeId = await getActionTypeIdByCode(graphClient, siteId, actionTypeListId, "CO");

        const leaveSetupItems = await getAllListItems(graphClient, siteId, leaveSetupListId);
        const allocatedThisYearByEmployee = new Map<string, number>();
        const bonusAllocatedThisYearByEmployee = new Map<string, number>();
        leaveSetupItems
            .filter((item: { fields: Record<string, string> }) => Number(item.fields.Year) === year)
            .forEach((item: { fields: Record<string, string> }) => {
                const key = item.fields.EmployeeCode?.toLowerCase();
                const holidayRaw = item.fields.HolidayDaysAllocatedThisYear;
                if (holidayRaw !== undefined && holidayRaw !== null && holidayRaw !== ("" as unknown)) {
                    allocatedThisYearByEmployee.set(key, Number(holidayRaw));
                }
                const bonusRaw = item.fields.BonusDays;
                if (bonusRaw !== undefined && bonusRaw !== null && bonusRaw !== ("" as unknown)) {
                    bonusAllocatedThisYearByEmployee.set(key, Number(bonusRaw));
                }
            });

        // Onboarded = has a Holiday LR_LeaveStatus row at all — the underlying
        // cumulative AllocDays value itself isn't shown, only used elsewhere
        // (balance checks) via adjustAllocDays.
        const leaveStatusItems = await getAllListItems(graphClient, siteId, leaveStatusListId);
        const onboardedEmployeeCodes = new Set<string>(
            leaveStatusItems
                .filter((item: { fields: Record<string, string> }) => Number(item.fields.RequestType) === Number(holidayTypeId))
                .map((item: { fields: Record<string, string> }) => item.fields.EmployeeCode?.toLowerCase())
        );

        const rows: YearlyAllocationRow[] = employees
            .map((emp) => {
                const key = emp.mail.toLowerCase();
                return {
                    employeeCode: emp.mail,
                    employeeFullName: emp.displayName,
                    isOnboarded: onboardedEmployeeCodes.has(key),
                    allocatedThisYear: allocatedThisYearByEmployee.has(key) ? allocatedThisYearByEmployee.get(key)! : null,
                    bonusAllocatedThisYear: bonusAllocatedThisYearByEmployee.has(key)
                        ? bonusAllocatedThisYearByEmployee.get(key)!
                        : null,
                };
            })
            .sort((a, b) => a.employeeFullName.localeCompare(b.employeeFullName));

        return { jsonBody: { year, employees: rows } };
    } catch (error) {
        context.error("GetYearlyAllocations failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_GetYearlyAllocations", {
    route: "leave/GetYearlyAllocations",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetYearlyAllocations,
});
