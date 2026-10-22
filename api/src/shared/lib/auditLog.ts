import { Client } from "@microsoft/microsoft-graph-client";
import { getAllListItems } from "./graphPaging";

export type AuditEntry = {
    actor: string;
    action: string;
    employeeCode?: string;
    requestId?: string;
    details: string;
};

/**
 * Appends an entry to the module's audit log (LR_AuditLog / DE_AuditLog) — every balance change, approval/decline,
 * and cancellation goes through here. Best-effort: a logging failure must
 * never block the actual action it's describing, so callers should wrap
 * this in try/catch rather than let it fail the request.
 */
export async function logAuditEntry(
    graphClient: Client,
    siteId: string,
    auditLogListId: string,
    entry: AuditEntry
): Promise<void> {
    await graphClient.api(`/sites/${siteId}/lists/${auditLogListId}/items`).post({
        fields: {
            Timestamp: new Date().toISOString(),
            Actor: entry.actor,
            Action: entry.action,
            EmployeeCode: entry.employeeCode ?? "",
            RequestId: entry.requestId ?? "",
            Details: entry.details,
        },
    });
}

/**
 * The latest decision this actor recorded on a request (Approved / ApprovedFinal /
 * Declined), or undefined if they never acted on it. Lets an action endpoint tell a
 * repeat of an already-applied click (Outlook retrying a slow first call, or a double
 * click) apart from a genuinely unauthorized one.
 */
export async function getActorDecision(
    graphClient: Client,
    siteId: string,
    auditLogListId: string,
    requestId: string,
    actor: string
): Promise<string | undefined> {
    const items = await getAllListItems(graphClient, siteId, auditLogListId);
    const latest = items
        .filter(
            (item: { fields: Record<string, string> }) =>
                item.fields.RequestId === requestId &&
                item.fields.Actor?.toLowerCase() === actor.toLowerCase() &&
                ["Approved", "ApprovedFinal", "Declined"].includes(item.fields.Action)
        )
        .sort(
            (a: { fields: Record<string, string> }, b: { fields: Record<string, string> }) =>
                new Date(b.fields.Timestamp).getTime() - new Date(a.fields.Timestamp).getTime()
        )[0];
    return latest?.fields.Action;
}

export type ApprovalStep = {
    actor: string;
    timestamp: string;
};

/**
 * Returns each approval "signature" for a request — one entry per
 * Approved/ApprovedFinal audit-log action, in chronological order. Used to
 * render the approval-signatures section of the generated leave-request PDF.
 */
export async function getApprovalSteps(
    graphClient: Client,
    siteId: string,
    auditLogListId: string,
    requestId: string
): Promise<ApprovalStep[]> {
    const items = await getAllListItems(graphClient, siteId, auditLogListId);

    return items
        .filter(
            (item: { fields: Record<string, string> }) =>
                item.fields.RequestId === requestId && (item.fields.Action === "Approved" || item.fields.Action === "ApprovedFinal")
        )
        .map((item: { fields: Record<string, string> }) => ({
            actor: item.fields.Actor,
            timestamp: item.fields.Timestamp,
        }))
        .sort((a: ApprovalStep, b: ApprovalStep) => new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime());
}

/**
 * Whether `actorPrincipal` was ever an approver on this specific request (an
 * "Approved"/"ApprovedFinal" audit-log entry for it) — used by the cancel
 * endpoints to allow a past approver to cancel, since approval history isn't
 * kept on the request item itself.
 */
export async function wasApproverOn(
    graphClient: Client,
    siteId: string,
    auditLogListId: string,
    requestId: string,
    actorPrincipal: string
): Promise<boolean> {
    const items = await getAllListItems(graphClient, siteId, auditLogListId);
    return items.some(
        (item: { fields: Record<string, string> }) =>
            item.fields.RequestId === requestId &&
            item.fields.Actor?.toLowerCase() === actorPrincipal?.toLowerCase() &&
            (item.fields.Action === "Approved" || item.fields.Action === "ApprovedFinal")
    );
}
