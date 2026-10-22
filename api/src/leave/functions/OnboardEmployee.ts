import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows, canAccessSettings, getDefaultAnnualAllocDays } from "../lib/setup";
import { getEmployeeDirectory } from "../../shared/lib/employees";
import { logAuditEntry } from "../../shared/lib/auditLog";
import { getAllListItems } from "../../shared/lib/graphPaging";
import { createLeaveStatusRow } from "../lib/leaveStatus";

type OnboardEmployeeBody = {
    employeeCode: string;
    year: number;
};

function parseAndValidateBody(raw: unknown): OnboardEmployeeBody {
    const body = raw as Partial<OnboardEmployeeBody> | null;
    if (!body?.employeeCode || typeof body.employeeCode !== "string") {
        throw new Error("employeeCode is required.");
    }
    if (typeof body.year !== "number") {
        throw new Error("year is required and must be a number.");
    }
    return { employeeCode: body.employeeCode, year: body.year };
}

/**
 * Onboards a genuinely new employee (one with no LR_LeaveStatus rows at all
 * yet) directly from the "Leave-day Allocations" view: seeds all LR_ActionType
 * rows (Holiday gets the default annual grant for the given year, Bonus and
 * the rest start at 0), and the year's LR_LeaveSetup metadata row (full name
 * pulled from the Entra ID directory — no manual entry needed). Restricted to
 * HR admins; logged to LR_AuditLog.
 */
export async function OnboardEmployee(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const body = parseAndValidateBody(await request.json());

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, [
            "Sites.Selected",
            "GroupMember.Read.All",
            "User.ReadBasic.All",
        ]);
        const { siteId, leaveSetupListId, leaveStatusListId, actionTypeListId, setupListId, auditLogListId } =
            await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("displayName,mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;
        const currentUserName: string = me.displayName ?? currentUserPrincipal;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        if (!canAccessSettings(setupRows, currentUserPrincipal)) {
            return { status: 403, jsonBody: { error: "Only HR, Admin, or GM can onboard employees." } };
        }

        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const employee = employees.find((e) => e.mail.toLowerCase() === body.employeeCode.toLowerCase());
        if (!employee) {
            throw new Error("Employee not found in the configured Entra ID group.");
        }

        const existingStatusItems = await getAllListItems(graphClient, siteId, leaveStatusListId);
        const alreadyOnboarded = existingStatusItems.some(
            (item: { fields: Record<string, string> }) => item.fields.EmployeeCode?.toLowerCase() === body.employeeCode.toLowerCase()
        );
        if (alreadyOnboarded) {
            return { status: 400, jsonBody: { error: "This employee is already onboarded." } };
        }

        const defaultDays = getDefaultAnnualAllocDays(setupRows);

        const actionTypeItems = await getAllListItems(graphClient, siteId, actionTypeListId);

        await Promise.all(
            actionTypeItems.map((item: { fields: Record<string, string> }) => {
                const allocDays = item.fields.ActionTypeCode === "CO" ? defaultDays : 0;
                return createLeaveStatusRow(tenantId, siteId, leaveStatusListId, {
                    EmployeeCode: employee.mail,
                    RequestType: Number(item.fields.IDActionType),
                    ActionTypeName: item.fields.ActionTypeName,
                    AllocDays: allocDays,
                    TakenDays: 0,
                });
            })
        );

        const leaveSetupItems = await getAllListItems(graphClient, siteId, leaveSetupListId);
        const existingSetupRow = leaveSetupItems.find(
            (item: { fields: Record<string, string> }) =>
                item.fields.EmployeeCode?.toLowerCase() === body.employeeCode.toLowerCase() &&
                Number(item.fields.Year) === body.year
        );
        const setupFields = {
            EmployeeCode: employee.mail,
            EmployeeFullName: employee.displayName,
            Year: body.year,
            HolidayDaysAllocatedThisYear: defaultDays,
        };
        if (existingSetupRow) {
            await graphClient
                .api(`/sites/${siteId}/lists/${leaveSetupListId}/items/${existingSetupRow.id}/fields`)
                .patch(setupFields);
        } else {
            await graphClient.api(`/sites/${siteId}/lists/${leaveSetupListId}/items`).post({ fields: setupFields });
        }

        try {
            await logAuditEntry(graphClient, siteId, auditLogListId, {
                actor: currentUserPrincipal,
                action: "EmployeeCreated",
                employeeCode: employee.mail,
                details: `${currentUserName} onboarded ${employee.displayName}, allocated ${defaultDays} Holiday days for ${body.year}.`,
            });
        } catch (auditError) {
            context.warn("logAuditEntry failed (non-blocking)", auditError);
        }

        return { jsonBody: { employeeCode: employee.mail, year: body.year, holidayAllocDays: defaultDays } };
    } catch (error) {
        context.error("OnboardEmployee failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_OnboardEmployee", {
    route: "leave/OnboardEmployee",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: OnboardEmployee,
});
