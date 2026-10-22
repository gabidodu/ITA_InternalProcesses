import { Client } from "@microsoft/microsoft-graph-client";
import { getAllListItems } from "../../shared/lib/graphPaging";
import { getAppOnlyGraphToken } from "../../shared/auth/appOnly";

/**
 * LR_LeaveStatus is the balance ledger — unlike LR_LeaveRequests, regular
 * employees are deliberately NOT given write access to it (mirrors the
 * legacy app's own "service user" convention). Reads still go through the
 * caller's own delegated token (employees can see balances); every write
 * goes through this app-only client instead, so no employee's own SharePoint
 * permissions need to include write access to their own balance row — which
 * would also mean write access to everyone else's.
 */
async function getServiceGraphClient(tenantId: string): Promise<Client> {
    const token = await getAppOnlyGraphToken(tenantId);
    return Client.init({ authProvider: (done) => done(null, token) });
}

/**
 * Adjusts (adds deltaDays to, can be negative) the TakenDays of an employee's
 * LR_LeaveStatus row for a given leave type. Used to keep the balance ledger
 * accurate when a request is approved (+takenDays) or a previously-approved
 * request is cancelled (-takenDays, rectification). No-ops if the employee
 * has no row for that leave type yet.
 */
export async function adjustTakenDays(
    graphClient: Client,
    tenantId: string,
    siteId: string,
    leaveStatusListId: string,
    employeeCode: string,
    requestType: number,
    deltaDays: number
): Promise<void> {
    const items = await getAllListItems(graphClient, siteId, leaveStatusListId);

    const row = items.find(
        (item: { fields: Record<string, string> }) =>
            item.fields.EmployeeCode?.toLowerCase() === employeeCode.toLowerCase() &&
            Number(item.fields.RequestType) === requestType
    );
    if (!row) {
        return;
    }

    const currentTaken = Number(row.fields.TakenDays) || 0;
    const serviceClient = await getServiceGraphClient(tenantId);
    await serviceClient
        .api(`/sites/${siteId}/lists/${leaveStatusListId}/items/${row.id}/fields`)
        .patch({ TakenDays: Math.max(0, currentTaken + deltaDays) });
}

/**
 * Adjusts (adds deltaDays to, can be negative) the AllocDays of an employee's
 * LR_LeaveStatus row for a given leave type — the cumulative-since-forever
 * allocation total. Used by the yearly bulk-allocation feature so a year's
 * granted days move the real running total, not just a per-year record.
 * No-ops if the employee has no row for that leave type yet.
 */
export async function adjustAllocDays(
    graphClient: Client,
    tenantId: string,
    siteId: string,
    leaveStatusListId: string,
    employeeCode: string,
    requestType: number,
    deltaDays: number
): Promise<void> {
    const items = await getAllListItems(graphClient, siteId, leaveStatusListId);

    const row = items.find(
        (item: { fields: Record<string, string> }) =>
            item.fields.EmployeeCode?.toLowerCase() === employeeCode.toLowerCase() &&
            Number(item.fields.RequestType) === requestType
    );
    if (!row) {
        return;
    }

    const currentAlloc = Number(row.fields.AllocDays) || 0;
    const serviceClient = await getServiceGraphClient(tenantId);
    await serviceClient
        .api(`/sites/${siteId}/lists/${leaveStatusListId}/items/${row.id}/fields`)
        .patch({ AllocDays: Math.max(0, currentAlloc + deltaDays) });
}

/**
 * Creates a new LR_LeaveStatus row (e.g. seeding a new employee's balance)
 * using the service (app-only) identity — same rationale as the adjust*
 * functions above: employees never need write access to this list directly.
 */
export async function createLeaveStatusRow(
    tenantId: string,
    siteId: string,
    leaveStatusListId: string,
    fields: Record<string, unknown>
): Promise<void> {
    const serviceClient = await getServiceGraphClient(tenantId);
    await serviceClient.api(`/sites/${siteId}/lists/${leaveStatusListId}/items`).post({ fields });
}

/**
 * Directly sets (not deltas) AllocDays on an existing LR_LeaveStatus row,
 * using the service (app-only) identity. Used when HR sets an absolute
 * balance value rather than adjusting it by a delta.
 */
export async function setAllocDays(
    tenantId: string,
    siteId: string,
    leaveStatusListId: string,
    itemId: string,
    allocDays: number
): Promise<void> {
    const serviceClient = await getServiceGraphClient(tenantId);
    await serviceClient
        .api(`/sites/${siteId}/lists/${leaveStatusListId}/items/${itemId}/fields`)
        .patch({ AllocDays: allocDays });
}
