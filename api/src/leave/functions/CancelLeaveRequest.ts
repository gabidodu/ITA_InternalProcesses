import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { getHrAdminEmails, getSetupRows, canAccessSettings } from "../lib/setup";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { sendLeaveActivityNotification } from "../lib/activityNotification";
import { sendLeaveEmailNotification } from "../lib/emailNotification";
import { adjustTakenDays } from "../lib/leaveStatus";
import { logAuditEntry } from "../../shared/lib/auditLog";
import { buildDisplayNameMap, getEmployeeDirectory } from "../../shared/lib/employees";
import { getTeamsDeepLink } from "../../shared/lib/teamsLink";
import { deleteLeaveCalendarEvent } from "../lib/calendarEvent";
import { getAllListItems } from "../../shared/lib/graphPaging";

type CancelLeaveRequestBody = {
    id: string;
};

function parseAndValidateBody(raw: unknown): CancelLeaveRequestBody {
    const body = raw as Partial<CancelLeaveRequestBody> | null;
    if (!body?.id || typeof body.id !== "string") {
        throw new Error("id is required.");
    }
    return { id: body.id };
}

/**
 * Cancels a leave request. Allowed for: the original requester, any HR/Admin/
 * GM (canAccessSettings), or any manager who was ever an approver on this
 * specific request (checked against LR_AuditLog "Approved"/"ApprovedFinal"
 * entries, since we don't keep approval history on the request item itself).
 * For the requester/participant, cancellation is only allowed up to 15 days
 * after the request's EndDate, regardless of status (PENDING or APPROVED) —
 * HR/Admin/GM are exempt from this time window and can cancel anytime. Days
 * are reserved at request-creation time (not at approval — see
 * CreateLeaveRequest), so any request reaching this point (PENDING or
 * APPROVED) always has reserved days to release: rectifies the balance by
 * subtracting the request's TakenDays back out of LR_LeaveStatus. Logged to
 * LR_AuditLog.
 */
export async function CancelLeaveRequest(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const tabBaseUrl = process.env.TAB_BASE_URL;
        const teamsAppId = process.env.TEAMS_APP_ID;
        if (!tabBaseUrl || !teamsAppId) {
            throw new Error("TAB_BASE_URL / TEAMS_APP_ID app settings are not fully configured.");
        }
        const teamsDeepLink = getTeamsDeepLink(teamsAppId, tabBaseUrl, "leave");

        const body = parseAndValidateBody(await request.json());

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, [
            "Sites.Selected",
            "TeamsActivity.Send",
            "GroupMember.Read.All",
            "User.ReadBasic.All",
        ]);
        const { siteId, leaveRequestsListId, leaveStatusListId, setupListId, auditLogListId } =
            await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("displayName,mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;
        const currentUserName: string = me.displayName ?? currentUserPrincipal;

        const item = await graphClient
            .api(`/sites/${siteId}/lists/${leaveRequestsListId}/items/${body.id}`)
            .expand("fields")
            .get();
        const fields = item.fields as Record<string, string>;

        if (fields.Status === "DECLINED" || fields.Status === "CANCELLED") {
            return { status: 400, jsonBody: { error: "This request is already resolved and cannot be cancelled." } };
        }
        // Guards against this racing an Approve/Decline on the same item — e.g. from the app
        // or an Outlook email — read off the same stale state. The conditional PATCH below
        // only lets the first writer through, since it changes the item's eTag.
        const itemETag: string | undefined = item.eTag;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const nameByMail = buildDisplayNameMap(employees);
        const resolveName = (mail: string | undefined) => (mail && nameByMail.get(mail.toLowerCase())) || mail || "";

        const isRequester = fields.EmployeeCode?.toLowerCase() === currentUserPrincipal?.toLowerCase();
        const isSettingsAdmin = canAccessSettings(setupRows, currentUserPrincipal);

        let wasParticipant = false;
        if (!isRequester && !isSettingsAdmin) {
            const auditItems = await getAllListItems(graphClient, siteId, auditLogListId);
            wasParticipant = auditItems.some(
                (auditItem: { fields: Record<string, string> }) =>
                    auditItem.fields.RequestId === body.id &&
                    auditItem.fields.Actor?.toLowerCase() === currentUserPrincipal?.toLowerCase() &&
                    (auditItem.fields.Action === "Approved" || auditItem.fields.Action === "ApprovedFinal")
            );
        }

        if (!isRequester && !isSettingsAdmin && !wasParticipant) {
            return {
                status: 403,
                jsonBody: { error: "Only the requester, HR/Admin/GM, or a manager who approved a step can cancel this request." },
            };
        }

        // HR/Admin/GM can cancel anytime; everyone else only within 15 days after the request's EndDate.
        if (!isSettingsAdmin && fields.EndDate) {
            const cancelCutoff = new Date(fields.EndDate);
            cancelCutoff.setDate(cancelCutoff.getDate() + 15);
            if (new Date() > cancelCutoff) {
                return {
                    status: 400,
                    jsonBody: { error: "This request can no longer be cancelled — more than 15 days have passed since its end date." },
                };
            }
        }

        const previousStatus = fields.Status;
        const requester = fields.EmployeeCode;
        const requesterName = resolveName(requester);

        try {
            await graphClient
                .api(`/sites/${siteId}/lists/${leaveRequestsListId}/items/${body.id}/fields`)
                .header("If-Match", itemETag ?? "*")
                .patch({
                    Status: "CANCELLED",
                    CurrentUser: requester,
                    ApprovalFLow: "",
                });
        } catch (patchError) {
            if ((patchError as { statusCode?: number })?.statusCode === 412) {
                return {
                    status: 409,
                    jsonBody: { error: "This request was just actioned elsewhere (e.g. an approval/decline). Please refresh." },
                };
            }
            throw patchError;
        }

        await adjustTakenDays(
            graphClient,
            tenantId,
            siteId,
            leaveStatusListId,
            requester,
            Number(fields.RequestType),
            -(Number(fields.TakenDays) || 0)
        );

        if (fields.CalendarEventId) {
            try {
                await deleteLeaveCalendarEvent(tenantId, fields.CalendarEventId);
            } catch (calendarError) {
                context.warn("deleteLeaveCalendarEvent failed (non-blocking)", calendarError);
            }
        }

        try {
            await logAuditEntry(graphClient, siteId, auditLogListId, {
                actor: currentUserPrincipal,
                action: "Cancelled",
                employeeCode: requester,
                requestId: body.id,
                details: `"${fields.RequestSubject}" cancelled by ${currentUserName} (was ${previousStatus}). Balance rectified: -${fields.TakenDays} days.`,
            });
        } catch (auditError) {
            context.warn("logAuditEntry failed (non-blocking)", auditError);
        }

        const notifyRecipients = new Set<string>([requester, ...getHrAdminEmails(setupRows)]);
        if (fields.CurrentUser && fields.CurrentUser.toLowerCase() !== requester?.toLowerCase()) {
            notifyRecipients.add(fields.CurrentUser);
        }
        const notifyMessage = `${requesterName}'s leave request "${fields.RequestSubject}" was cancelled.`;
        const subject = `Leave Request - ${requesterName}`;

        // The one recipient who most needed to hear about this: whoever still held the
        // request (no longer needs to review it) — falls back to the requester if there
        // was no pending approver. Surfaced to the acting user, same as the other actions.
        const primaryRecipient = (fields.CurrentUser || requester)?.toLowerCase();
        let primaryEmailSent = false;

        await Promise.all(
            Array.from(notifyRecipients).flatMap((recipient) => [
                sendLeaveActivityNotification(graphClient, recipient, notifyMessage, tabBaseUrl).catch(
                    (notificationError) =>
                        context.warn(`sendLeaveActivityNotification to ${recipient} failed (non-blocking)`, notificationError)
                ),
                sendLeaveEmailNotification(tenantId, recipient, subject, notifyMessage, teamsDeepLink)
                    .then(() => {
                        if (recipient.toLowerCase() === primaryRecipient) {
                            primaryEmailSent = true;
                        }
                    })
                    .catch((emailError) =>
                        context.warn(`sendLeaveEmailNotification to ${recipient} failed (non-blocking)`, emailError)
                    ),
            ])
        );

        return { jsonBody: { id: body.id, status: "CANCELLED", emailSent: primaryEmailSent } };
    } catch (error) {
        context.error("CancelLeaveRequest failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_CancelLeaveRequest", {
    route: "leave/CancelLeaveRequest",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: CancelLeaveRequest,
});
