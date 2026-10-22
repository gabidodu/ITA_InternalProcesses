import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows, canAccessSettings } from "../lib/setup";
import { logAuditEntry } from "../../shared/lib/auditLog";
import { getAllListItems } from "../../shared/lib/graphPaging";
import { createLeaveStatusRow, setAllocDays } from "../lib/leaveStatus";

type EmployeeSetupBody = {
    employeeCode: string;
    employeeFullName: string;
    holidayAllocDays: number;
    bonusAllocDays: number;
    employeeStartLeave?: string;
};

function parseAndValidateBody(raw: unknown): EmployeeSetupBody {
    const body = raw as Partial<EmployeeSetupBody> | null;
    if (!body?.employeeCode || typeof body.employeeCode !== "string") {
        throw new Error("employeeCode is required.");
    }
    if (!body.employeeFullName || typeof body.employeeFullName !== "string") {
        throw new Error("employeeFullName is required.");
    }
    if (typeof body.holidayAllocDays !== "number" || body.holidayAllocDays < 0) {
        throw new Error("holidayAllocDays must be a non-negative number.");
    }
    if (typeof body.bonusAllocDays !== "number" || body.bonusAllocDays < 0) {
        throw new Error("bonusAllocDays must be a non-negative number.");
    }
    return {
        employeeCode: body.employeeCode,
        employeeFullName: body.employeeFullName,
        holidayAllocDays: body.holidayAllocDays,
        bonusAllocDays: body.bonusAllocDays,
        employeeStartLeave: body.employeeStartLeave,
    };
}

/**
 * Creates or updates an employee's current leave balance — directly setting
 * AllocDays on their LR_LeaveStatus rows for Holiday and Bonus (the only two
 * types with a settable maximum; medical/unpaid have none, per user decision).
 * Also upserts this year's LR_LeaveSetup row for name/start-date metadata.
 * For a genuinely new employee (no LR_LeaveStatus rows at all yet), seeds all
 * LR_ActionType rows (0 AllocDays for types with no max). Restricted to HR
 * admins; every change is written to LR_AuditLog.
 */
export async function CreateOrUpdateEmployeeSetup(
    request: HttpRequest,
    context: InvocationContext
): Promise<HttpResponseInit> {
    try {
        const body = parseAndValidateBody(await request.json());

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected"]);
        const { siteId, leaveSetupListId, leaveStatusListId, actionTypeListId, setupListId, auditLogListId } =
            await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        if (!canAccessSettings(setupRows, currentUserPrincipal)) {
            return { status: 403, jsonBody: { error: "Only HR, Admin, or GM can add or edit employees." } };
        }

        const currentYear = new Date().getFullYear();

        // --- LR_LeaveSetup metadata (name / start-leave date), this year only ---
        const existingSetupItems = await getAllListItems(graphClient, siteId, leaveSetupListId);
        const thisYearSetupRow = existingSetupItems.find(
            (item: { fields: Record<string, string> }) =>
                item.fields.EmployeeCode?.toLowerCase() === body.employeeCode.toLowerCase() &&
                Number(item.fields.Year) === currentYear
        );
        const setupFields = {
            EmployeeCode: body.employeeCode,
            EmployeeFullName: body.employeeFullName,
            Year: currentYear,
            ...(body.employeeStartLeave ? { EmployeeStartLeave: body.employeeStartLeave } : {}),
        };
        if (thisYearSetupRow) {
            await graphClient
                .api(`/sites/${siteId}/lists/${leaveSetupListId}/items/${thisYearSetupRow.id}/fields`)
                .patch(setupFields);
        } else {
            await graphClient.api(`/sites/${siteId}/lists/${leaveSetupListId}/items`).post({ fields: setupFields });
        }

        // --- LR_LeaveStatus balance (Holiday/Bonus AllocDays) ---
        const existingStatusItems = await getAllListItems(graphClient, siteId, leaveStatusListId);
        const employeeStatusRows = existingStatusItems.filter(
            (item: { fields: Record<string, string> }) =>
                item.fields.EmployeeCode?.toLowerCase() === body.employeeCode.toLowerCase()
        );
        const isNewEmployee = employeeStatusRows.length === 0;

        let oldHolidayAlloc = 0;
        let oldBonusAlloc = 0;

        if (isNewEmployee) {
            const actionTypeItems = await getAllListItems(graphClient, siteId, actionTypeListId);

            await Promise.all(
                actionTypeItems.map((item: { fields: Record<string, string> }) => {
                    const allocDays =
                        item.fields.ActionTypeCode === "CO"
                            ? body.holidayAllocDays
                            : item.fields.ActionTypeCode === "CB"
                              ? body.bonusAllocDays
                              : 0;
                    return createLeaveStatusRow(tenantId, siteId, leaveStatusListId, {
                        EmployeeCode: body.employeeCode,
                        RequestType: Number(item.fields.IDActionType),
                        ActionTypeName: item.fields.ActionTypeName,
                        AllocDays: allocDays,
                        TakenDays: 0,
                    });
                })
            );
        } else {
            const holidayRow = employeeStatusRows.find(
                (item: { fields: Record<string, string> }) => item.fields.ActionTypeName === "Holiday"
            );
            const bonusRow = employeeStatusRows.find(
                (item: { fields: Record<string, string> }) => item.fields.ActionTypeName === "Bonus Days"
            );
            oldHolidayAlloc = holidayRow ? Number(holidayRow.fields.AllocDays) || 0 : 0;
            oldBonusAlloc = bonusRow ? Number(bonusRow.fields.AllocDays) || 0 : 0;

            if (holidayRow) {
                await setAllocDays(tenantId, siteId, leaveStatusListId, holidayRow.id, body.holidayAllocDays);
            }
            if (bonusRow) {
                await setAllocDays(tenantId, siteId, leaveStatusListId, bonusRow.id, body.bonusAllocDays);
            }
        }

        try {
            await logAuditEntry(graphClient, siteId, auditLogListId, {
                actor: currentUserPrincipal,
                action: isNewEmployee ? "EmployeeCreated" : "BalanceAdjusted",
                employeeCode: body.employeeCode,
                details: isNewEmployee
                    ? `New employee seeded: Holiday=${body.holidayAllocDays}, Bonus=${body.bonusAllocDays}.`
                    : `Holiday AllocDays ${oldHolidayAlloc} -> ${body.holidayAllocDays}; Bonus AllocDays ${oldBonusAlloc} -> ${body.bonusAllocDays}.`,
            });
        } catch (auditError) {
            context.warn("logAuditEntry failed (non-blocking)", auditError);
        }

        return { jsonBody: { employeeCode: body.employeeCode, isNewEmployee } };
    } catch (error) {
        context.error("CreateOrUpdateEmployeeSetup failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_CreateOrUpdateEmployeeSetup", {
    route: "leave/CreateOrUpdateEmployeeSetup",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: CreateOrUpdateEmployeeSetup,
});
