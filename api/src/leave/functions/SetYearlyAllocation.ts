import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows, canAccessSettings } from "../lib/setup";
import { getEmployeeDirectory, buildDisplayNameMap } from "../../shared/lib/employees";
import { getActionTypeIdByCode } from "../lib/actionTypes";
import { adjustAllocDays } from "../lib/leaveStatus";
import { logAuditEntry } from "../../shared/lib/auditLog";
import { getAllListItems } from "../../shared/lib/graphPaging";

type SetYearlyAllocationBody = {
    employeeCode: string;
    year: number;
    holidayDaysAllocatedThisYear?: number;
    bonusDaysAllocatedThisYear?: number;
};

function parseAndValidateBody(raw: unknown): SetYearlyAllocationBody {
    const body = raw as Partial<SetYearlyAllocationBody> | null;
    if (!body?.employeeCode || typeof body.employeeCode !== "string") {
        throw new Error("employeeCode is required.");
    }
    if (typeof body.year !== "number") {
        throw new Error("year is required and must be a number.");
    }
    const hasHoliday = body.holidayDaysAllocatedThisYear !== undefined;
    const hasBonus = body.bonusDaysAllocatedThisYear !== undefined;
    if (!hasHoliday && !hasBonus) {
        throw new Error("Provide holidayDaysAllocatedThisYear and/or bonusDaysAllocatedThisYear.");
    }
    if (hasHoliday && (typeof body.holidayDaysAllocatedThisYear !== "number" || body.holidayDaysAllocatedThisYear! < 0)) {
        throw new Error("holidayDaysAllocatedThisYear must be a non-negative number.");
    }
    if (hasBonus && (typeof body.bonusDaysAllocatedThisYear !== "number" || body.bonusDaysAllocatedThisYear! < 0)) {
        throw new Error("bonusDaysAllocatedThisYear must be a non-negative number.");
    }
    return {
        employeeCode: body.employeeCode,
        year: body.year,
        holidayDaysAllocatedThisYear: body.holidayDaysAllocatedThisYear,
        bonusDaysAllocatedThisYear: body.bonusDaysAllocatedThisYear,
    };
}

/**
 * Edits a single employee's Holiday and/or Bonus allocation for a specific
 * year (LR_LeaveSetup.HolidayDaysAllocatedThisYear / .BonusDays), adjusting
 * the matching cumulative AllocDays total on LR_LeaveStatus by the delta
 * between the old and new value — so editing a past or current year's grant
 * actually moves the real running total (used for the balance check in
 * CreateLeaveRequest), not just the per-year record. Both leave types use
 * the exact same per-year-grant-feeds-a-cumulative-total shape; real
 * production data confirms Bonus was tracked this way historically too
 * (LR_LeaveSetup.BonusDays), same as Holiday. Restricted to HR admins, every
 * change is written to LR_AuditLog.
 */
export async function SetYearlyAllocation(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
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
            return { status: 403, jsonBody: { error: "Only HR, Admin, or GM can edit leave-day allocations." } };
        }

        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const nameByMail = buildDisplayNameMap(employees);
        const employeeName = nameByMail.get(body.employeeCode.toLowerCase()) ?? body.employeeCode;

        const leaveSetupItems = await getAllListItems(graphClient, siteId, leaveSetupListId);
        const existingRow = leaveSetupItems.find(
            (item: { fields: Record<string, string> }) =>
                item.fields.EmployeeCode?.toLowerCase() === body.employeeCode.toLowerCase() &&
                Number(item.fields.Year) === body.year
        );

        const setupFields: Record<string, number> = {};
        const auditParts: string[] = [];

        if (body.holidayDaysAllocatedThisYear !== undefined) {
            const oldValue =
                existingRow?.fields.HolidayDaysAllocatedThisYear !== undefined &&
                existingRow?.fields.HolidayDaysAllocatedThisYear !== ("" as unknown)
                    ? Number(existingRow.fields.HolidayDaysAllocatedThisYear)
                    : 0;
            const delta = body.holidayDaysAllocatedThisYear - oldValue;
            setupFields.HolidayDaysAllocatedThisYear = body.holidayDaysAllocatedThisYear;
            if (delta !== 0) {
                const holidayTypeId = await getActionTypeIdByCode(graphClient, siteId, actionTypeListId, "CO");
                if (!holidayTypeId) {
                    throw new Error("Holiday action type (CO) not found in LR_ActionType.");
                }
                await adjustAllocDays(graphClient, tenantId, siteId, leaveStatusListId, body.employeeCode, holidayTypeId, delta);
            }
            auditParts.push(`Holiday ${oldValue} -> ${body.holidayDaysAllocatedThisYear} (total adjusted by ${delta >= 0 ? "+" : ""}${delta})`);
        }

        if (body.bonusDaysAllocatedThisYear !== undefined) {
            const oldValue =
                existingRow?.fields.BonusDays !== undefined && existingRow?.fields.BonusDays !== ("" as unknown)
                    ? Number(existingRow.fields.BonusDays)
                    : 0;
            const delta = body.bonusDaysAllocatedThisYear - oldValue;
            setupFields.BonusDays = body.bonusDaysAllocatedThisYear;
            if (delta !== 0) {
                const bonusTypeId = await getActionTypeIdByCode(graphClient, siteId, actionTypeListId, "CB");
                if (!bonusTypeId) {
                    throw new Error("Bonus action type (CB) not found in LR_ActionType.");
                }
                await adjustAllocDays(graphClient, tenantId, siteId, leaveStatusListId, body.employeeCode, bonusTypeId, delta);
            }
            auditParts.push(`Bonus ${oldValue} -> ${body.bonusDaysAllocatedThisYear} (total adjusted by ${delta >= 0 ? "+" : ""}${delta})`);
        }

        if (existingRow) {
            await graphClient
                .api(`/sites/${siteId}/lists/${leaveSetupListId}/items/${existingRow.id}/fields`)
                .patch(setupFields);
        } else {
            await graphClient.api(`/sites/${siteId}/lists/${leaveSetupListId}/items`).post({
                fields: {
                    EmployeeCode: body.employeeCode,
                    EmployeeFullName: employeeName,
                    Year: body.year,
                    ...setupFields,
                },
            });
        }

        try {
            await logAuditEntry(graphClient, siteId, auditLogListId, {
                actor: currentUserPrincipal,
                action: "AllocationEdited",
                employeeCode: body.employeeCode,
                details: `${currentUserName} set ${employeeName}'s ${body.year} allocation: ${auditParts.join("; ")}.`,
            });
        } catch (auditError) {
            context.warn("logAuditEntry failed (non-blocking)", auditError);
        }

        return {
            jsonBody: {
                employeeCode: body.employeeCode,
                year: body.year,
                holidayDaysAllocatedThisYear: body.holidayDaysAllocatedThisYear,
                bonusDaysAllocatedThisYear: body.bonusDaysAllocatedThisYear,
            },
        };
    } catch (error) {
        context.error("SetYearlyAllocation failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_SetYearlyAllocation", {
    route: "leave/SetYearlyAllocation",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: SetYearlyAllocation,
});
