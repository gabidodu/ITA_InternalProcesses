import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows, canAccessSettings } from "../lib/setup";
import { getAllListItems } from "../../shared/lib/graphPaging";

type EmployeeRoster = {
    employeeCode: string;
    employeeFullName: string;
    employeeStartLeave: string | null;
    holidayAllocDays: number;
    holidayTakenDays: number;
    bonusAllocDays: number;
    bonusTakenDays: number;
};

/**
 * Returns the employee roster for HR: current balance (AllocDays/TakenDays)
 * for Holiday and Bonus — the only two leave types with a settable maximum,
 * per [[project-leave-teams-app]] — merged with this year's LR_LeaveSetup
 * metadata (full name, start-leave date) where present. Restricted to HR admins.
 */
export async function GetEmployeeSetups(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected"]);
        const { siteId, leaveSetupListId, leaveStatusListId, actionTypeListId, setupListId } =
            await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        if (!canAccessSettings(setupRows, currentUserPrincipal)) {
            return { status: 403, jsonBody: { error: "Only HR, Admin, or GM can view the employee roster." } };
        }

        const currentYear = new Date().getFullYear();

        const actionTypeItems = await getAllListItems(graphClient, siteId, actionTypeListId);
        const holidayTypeId = actionTypeItems.find(
            (i: { fields: Record<string, string> }) => i.fields.ActionTypeCode === "CO"
        )?.fields.IDActionType;
        const bonusTypeId = actionTypeItems.find(
            (i: { fields: Record<string, string> }) => i.fields.ActionTypeCode === "CB"
        )?.fields.IDActionType;

        const leaveStatusItems = await getAllListItems(graphClient, siteId, leaveStatusListId);

        const leaveSetupItems = await getAllListItems(graphClient, siteId, leaveSetupListId);
        const metadataByEmployee = new Map<string, { employeeFullName: string; employeeStartLeave: string | null }>();
        leaveSetupItems
            .filter((item: { fields: Record<string, string> }) => Number(item.fields.Year) === currentYear)
            .forEach((item: { fields: Record<string, string> }) => {
                metadataByEmployee.set(item.fields.EmployeeCode?.toLowerCase(), {
                    employeeFullName: item.fields.EmployeeFullName,
                    employeeStartLeave: item.fields.EmployeeStartLeave ?? null,
                });
            });

        const rosterByEmployee = new Map<string, EmployeeRoster>();
        leaveStatusItems.forEach((item: { fields: Record<string, string> }) => {
            const code = item.fields.EmployeeCode;
            if (!code) return;
            const key = code.toLowerCase();
            if (!rosterByEmployee.has(key)) {
                const meta = metadataByEmployee.get(key);
                rosterByEmployee.set(key, {
                    employeeCode: code,
                    employeeFullName: meta?.employeeFullName ?? "",
                    employeeStartLeave: meta?.employeeStartLeave ?? null,
                    holidayAllocDays: 0,
                    holidayTakenDays: 0,
                    bonusAllocDays: 0,
                    bonusTakenDays: 0,
                });
            }
            const roster = rosterByEmployee.get(key)!;
            if (holidayTypeId && Number(item.fields.RequestType) === Number(holidayTypeId)) {
                roster.holidayAllocDays = Number(item.fields.AllocDays) || 0;
                roster.holidayTakenDays = Number(item.fields.TakenDays) || 0;
            }
            if (bonusTypeId && Number(item.fields.RequestType) === Number(bonusTypeId)) {
                roster.bonusAllocDays = Number(item.fields.AllocDays) || 0;
                roster.bonusTakenDays = Number(item.fields.TakenDays) || 0;
            }
        });

        const employees = Array.from(rosterByEmployee.values()).sort((a, b) =>
            (a.employeeFullName || a.employeeCode).localeCompare(b.employeeFullName || b.employeeCode)
        );

        const definedThisYearCodes = Array.from(metadataByEmployee.keys());

        return { jsonBody: { year: currentYear, employees, definedThisYearCodes } };
    } catch (error) {
        context.error("GetEmployeeSetups failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_GetEmployeeSetups", {
    route: "leave/GetEmployeeSetups",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetEmployeeSetups,
});
