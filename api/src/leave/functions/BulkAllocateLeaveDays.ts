import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows, canAccessSettings, getDefaultAnnualAllocDays } from "../lib/setup";
import { getEmployeeDirectory } from "../../shared/lib/employees";
import { getActionTypeIdByCode } from "../lib/actionTypes";
import { adjustAllocDays } from "../lib/leaveStatus";
import { logAuditEntry } from "../../shared/lib/auditLog";
import { getAllListItems } from "../../shared/lib/graphPaging";

type BulkAllocateBody = {
    year: number;
};

function parseAndValidateBody(raw: unknown): BulkAllocateBody {
    const body = raw as Partial<BulkAllocateBody> | null;
    if (typeof body?.year !== "number") {
        throw new Error("year is required and must be a number.");
    }
    return { year: body.year };
}

/**
 * Grants the default number of Holiday days (LR_Setup Setting/defaultAnnualAllocDays,
 * 21 unless HR has changed it) for a given year, to every AD-group employee who
 * (a) has already been onboarded (has an LR_LeaveStatus Holiday row at all) and
 * (b) doesn't already have a Holiday allocation recorded for that year
 * (LR_LeaveSetup.HolidayDaysAllocatedThisYear). Never overwrites an existing
 * grant — use SetYearlyAllocation to edit one after the fact. Each grant both
 * writes/updates the per-year LR_LeaveSetup record and adds the days to the
 * employee's real cumulative AllocDays total on LR_LeaveStatus. Restricted to
 * HR admins; every grant is written to LR_AuditLog.
 */
export async function BulkAllocateLeaveDays(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
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
            return { status: 403, jsonBody: { error: "Only HR, Admin, or GM can bulk-allocate leave days." } };
        }

        const defaultDays = getDefaultAnnualAllocDays(setupRows);
        const employees = await getEmployeeDirectory(graphClient, setupRows);

        const holidayTypeId = await getActionTypeIdByCode(graphClient, siteId, actionTypeListId, "CO");
        if (!holidayTypeId) {
            throw new Error("Holiday action type (CO) not found in LR_ActionType.");
        }

        const leaveStatusItems = await getAllListItems(graphClient, siteId, leaveStatusListId);
        const onboardedEmployeeCodes = new Set<string>(
            leaveStatusItems
                .filter((item: { fields: Record<string, string> }) => Number(item.fields.RequestType) === holidayTypeId)
                .map((item: { fields: Record<string, string> }) => item.fields.EmployeeCode?.toLowerCase())
        );

        const leaveSetupItems = await getAllListItems(graphClient, siteId, leaveSetupListId);
        const existingRowForYear = new Map<string, { id: string; fields: Record<string, string> }>();
        leaveSetupItems
            .filter((item: { fields: Record<string, string> }) => Number(item.fields.Year) === body.year)
            .forEach((item: { id: string; fields: Record<string, string> }) => {
                existingRowForYear.set(item.fields.EmployeeCode?.toLowerCase(), item);
            });

        const allocatedEmployees: string[] = [];
        let skippedAlreadyAllocated = 0;
        let skippedNotOnboarded = 0;

        for (const emp of employees) {
            const key = emp.mail.toLowerCase();

            if (!onboardedEmployeeCodes.has(key)) {
                skippedNotOnboarded++;
                continue;
            }

            const existingRow = existingRowForYear.get(key);
            const alreadyAllocated =
                existingRow?.fields.HolidayDaysAllocatedThisYear !== undefined &&
                existingRow?.fields.HolidayDaysAllocatedThisYear !== ("" as unknown);
            if (alreadyAllocated) {
                skippedAlreadyAllocated++;
                continue;
            }

            if (existingRow) {
                await graphClient
                    .api(`/sites/${siteId}/lists/${leaveSetupListId}/items/${existingRow.id}/fields`)
                    .patch({ HolidayDaysAllocatedThisYear: defaultDays });
            } else {
                await graphClient.api(`/sites/${siteId}/lists/${leaveSetupListId}/items`).post({
                    fields: {
                        EmployeeCode: emp.mail,
                        EmployeeFullName: emp.displayName,
                        Year: body.year,
                        HolidayDaysAllocatedThisYear: defaultDays,
                    },
                });
            }

            await adjustAllocDays(graphClient, tenantId, siteId, leaveStatusListId, emp.mail, holidayTypeId, defaultDays);
            allocatedEmployees.push(emp.displayName);

            try {
                await logAuditEntry(graphClient, siteId, auditLogListId, {
                    actor: currentUserPrincipal,
                    action: "AllocationBulkGranted",
                    employeeCode: emp.mail,
                    details: `${currentUserName} granted ${emp.displayName} ${defaultDays} Holiday days for ${body.year}.`,
                });
            } catch (auditError) {
                context.warn(`logAuditEntry failed for ${emp.mail} (non-blocking)`, auditError);
            }
        }

        return {
            jsonBody: {
                year: body.year,
                defaultDays,
                allocatedCount: allocatedEmployees.length,
                allocatedEmployees,
                skippedAlreadyAllocated,
                skippedNotOnboarded,
            },
        };
    } catch (error) {
        context.error("BulkAllocateLeaveDays failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_BulkAllocateLeaveDays", {
    route: "leave/BulkAllocateLeaveDays",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: BulkAllocateLeaveDays,
});
