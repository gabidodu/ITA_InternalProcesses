import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows } from "../lib/setup";
import { sendLeaveActivityNotification } from "../lib/activityNotification";
import { sendLeaveEmailNotification, sendLeaveEmailWithAttachment, sendApprovalActionEmail } from "../lib/emailNotification";
import { adjustTakenDays } from "../lib/leaveStatus";
import { logAuditEntry, getApprovalSteps } from "../../shared/lib/auditLog";
import { buildDisplayNameMap, getEmployeeDirectory } from "../../shared/lib/employees";
import { generateLeaveRequestPdf } from "../lib/pdfGenerator";
import { saveApprovedLeaveRequestPdf } from "../lib/approvalArchive";
import { getTeamsDeepLink } from "../../shared/lib/teamsLink";
import { deleteLeaveCalendarEvent, markLeaveCalendarEventApproved } from "../lib/calendarEvent";
import { buildApprovalCardScriptTag } from "../lib/actionableMessage";
import { formatDate } from "../../shared/lib/dates";
import { getAllListItems } from "../../shared/lib/graphPaging";

type Decision = "approved" | "declined";

type RespondToLeaveRequestBody = {
    id: string;
    decision: Decision;
    comment?: string;
};

function parseAndValidateBody(raw: unknown): RespondToLeaveRequestBody {
    const body = raw as Partial<RespondToLeaveRequestBody> | null;
    if (!body?.id || typeof body.id !== "string") {
        throw new Error("id is required.");
    }
    if (body.decision !== "approved" && body.decision !== "declined") {
        throw new Error('decision must be "approved" or "declined".');
    }
    return { id: body.id, decision: body.decision, comment: body.comment?.trim() || undefined };
}

/**
 * Approves or declines a leave request as the current approver. Advances the
 * remaining approval chain (stored in ApprovalFLow) to the next approver, or
 * closes out the request (back to the original requester) once the chain is
 * exhausted or on a decline. Days are reserved on LR_LeaveStatus at request
 * creation time (see CreateLeaveRequest), not here — final approval needs no
 * balance change, but a decline releases the reservation since the leave
 * never happens (rectified again on a later cancel — see CancelLeaveRequest).
 * Every step is logged to LR_AuditLog.
 */
export async function RespondToLeaveRequest(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
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
            "User.Read.All",
        ]);
        const { siteId, leaveRequestsListId, leaveStatusListId, actionTypeListId, setupListId, auditLogListId, approvalsLibraryId } =
            await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("displayName,mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;
        const currentUserName: string = me.displayName ?? currentUserPrincipal;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const nameByMail = buildDisplayNameMap(employees);
        const resolveName = (mail: string | undefined) => (mail && nameByMail.get(mail.toLowerCase())) || mail || "";

        const item = await graphClient
            .api(`/sites/${siteId}/lists/${leaveRequestsListId}/items/${body.id}`)
            .expand("fields")
            .get();

        const fields = item.fields as Record<string, string>;
        if (fields.CurrentUser?.toLowerCase() !== currentUserPrincipal?.toLowerCase()) {
            return {
                status: 403,
                jsonBody: { error: "This request is not currently awaiting your action." },
            };
        }
        // Guards against the same step being actioned twice at once — e.g. the app and a
        // stale Outlook Approve/Decline email racing each other. Both would pass the
        // CurrentUser check above off the same read; the conditional PATCH below only lets
        // the first one through, since it changes the item's eTag.
        const itemETag: string | undefined = item.eTag;

        const requester = fields.EmployeeCode;
        const requesterName = resolveName(requester);
        const remainingChain = (fields.ApprovalFLow ?? "")
            .split(";")
            .map((s) => s.trim())
            .filter(Boolean);

        const periodDetail = { label: "Period", value: `${formatDate(fields.StartDate)} – ${formatDate(fields.EndDate)}` };
        const daysDetail = { label: "Days", value: String(Number(fields.TakenDays) || 0) };

        let newStatus: string;
        let newCurrentUser: string;
        let newRemainingChain: string[];
        let notifyRecipients: string[];
        let notifyMessage: string;
        let notifyDetails: { label: string; value: string }[];
        // Set only for the mid-chain case, where `notifyMessage` above is addressed to the
        // next approver — the requester still needs their own distinct status update, since
        // they're not otherwise in `notifyRecipients` for that branch.
        let requesterUpdateMessage: string | undefined;
        let requesterUpdateDetails: { label: string; value: string }[] | undefined;

        if (body.decision === "declined") {
            newStatus = "DECLINED";
            newCurrentUser = requester;
            newRemainingChain = [];
            notifyMessage = `Your leave request "${fields.RequestSubject}" was declined.`;
            notifyRecipients = [requester];
            notifyDetails = [periodDetail, daysDetail, { label: "Declined by", value: currentUserName }];
        } else if (remainingChain.length > 0) {
            const [nextApprover, ...rest] = remainingChain;
            newStatus = "PENDING";
            newCurrentUser = nextApprover;
            newRemainingChain = rest;
            notifyMessage = `${requesterName} needs your approval for a leave request.`;
            notifyRecipients = [nextApprover];
            notifyDetails = [periodDetail, daysDetail, { label: "Requested by", value: requesterName }];
            requesterUpdateMessage = `Your leave request "${fields.RequestSubject}" was approved by ${currentUserName} and is now awaiting approval from ${resolveName(nextApprover)}.`;
            requesterUpdateDetails = [
                periodDetail,
                daysDetail,
                { label: "Approved by", value: currentUserName },
                { label: "Next approver", value: resolveName(nextApprover) },
            ];
        } else {
            newStatus = "APPROVED";
            newCurrentUser = requester;
            newRemainingChain = [];
            notifyMessage = `Your leave request "${fields.RequestSubject}" was approved.`;
            notifyRecipients = [requester];
            notifyDetails = [periodDetail, daysDetail, { label: "Approved by", value: currentUserName }];
        }

        // Surface the approver's comment (if any) in the notification email's details
        // table too, not just the audit log — both the acting recipient's email and, on
        // a mid-chain approval, the requester's separate status-update email.
        if (body.comment) {
            notifyDetails = [...notifyDetails, { label: "Comment", value: body.comment }];
            if (requesterUpdateDetails) {
                requesterUpdateDetails = [...requesterUpdateDetails, { label: "Comment", value: body.comment }];
            }
        }

        // Dedupe case-insensitively — the requester can end up listed twice in some
        // branches (e.g. also being the acting approver on a test account).
        notifyRecipients = Array.from(new Map(notifyRecipients.map((r) => [r.toLowerCase(), r])).values());

        try {
            await graphClient
                .api(`/sites/${siteId}/lists/${leaveRequestsListId}/items/${body.id}/fields`)
                .header("If-Match", itemETag ?? "*")
                .patch({
                    Status: newStatus,
                    CurrentUser: newCurrentUser,
                    ApprovalFLow: newRemainingChain.join(";"),
                });
        } catch (patchError) {
            if ((patchError as { statusCode?: number })?.statusCode === 412) {
                return {
                    status: 409,
                    jsonBody: { error: "This request was just actioned elsewhere (e.g. from an email or another device). Please refresh." },
                };
            }
            throw patchError;
        }

        // Days are reserved at CreateLeaveRequest time, not here — final approval needs no
        // balance change. A decline releases the reservation, since the leave never happens.
        if (body.decision === "declined") {
            await adjustTakenDays(
                graphClient,
                tenantId,
                siteId,
                leaveStatusListId,
                requester,
                Number(fields.RequestType),
                -(Number(fields.TakenDays) || 0)
            );
        }

        if (fields.CalendarEventId) {
            try {
                if (body.decision === "declined") {
                    await deleteLeaveCalendarEvent(tenantId, fields.CalendarEventId);
                } else if (newStatus === "APPROVED") {
                    await markLeaveCalendarEventApproved(tenantId, fields.CalendarEventId, requesterName);
                }
            } catch (calendarError) {
                context.warn("Calendar event update failed (non-blocking)", calendarError);
            }
        }

        try {
            await logAuditEntry(graphClient, siteId, auditLogListId, {
                actor: currentUserPrincipal,
                action: body.decision === "declined" ? "Declined" : newStatus === "APPROVED" ? "ApprovedFinal" : "Approved",
                employeeCode: requester,
                requestId: body.id,
                details:
                    `"${fields.RequestSubject}" — ${body.decision} by ${currentUserName}. New status: ${newStatus}.` +
                    (body.comment ? ` Comment: ${body.comment}` : ""),
            });
        } catch (auditError) {
            context.warn("logAuditEntry failed (non-blocking)", auditError);
        }

        if (newStatus === "APPROVED") {
            try {
                const actionTypeItems = await getAllListItems(graphClient, siteId, actionTypeListId);
                const actionTypeRow = actionTypeItems.find(
                    (item: { fields: Record<string, string> }) => Number(item.fields.IDActionType) === Number(fields.RequestType)
                );
                const actionTypeName = actionTypeRow?.fields.ActionTypeNameRO || actionTypeRow?.fields.ActionTypeName || "concediu";
                const actionTypeNameEn =
                    actionTypeRow?.fields.ActionTypeNameRO && actionTypeRow?.fields.ActionTypeName
                        ? actionTypeRow.fields.ActionTypeName
                        : undefined;

                let requesterProfile: { department?: string; jobTitle?: string; companyName?: string } = {};
                try {
                    requesterProfile = await graphClient
                        .api(`/users/${encodeURIComponent(requester)}`)
                        .select("department,jobTitle,companyName")
                        .get();
                } catch (profileError) {
                    context.warn("Failed to fetch requester AD profile for PDF (non-blocking)", profileError);
                }

                const approvalSteps = await getApprovalSteps(graphClient, siteId, auditLogListId, body.id);
                const signatures = approvalSteps.map((step) => ({
                    actorName: resolveName(step.actor),
                    timestamp: step.timestamp,
                }));

                const pdfBuffer = await generateLeaveRequestPdf({
                    requestId: body.id,
                    requestDate: fields.RequestDate,
                    employeeFullName: requesterName,
                    department: requesterProfile.department,
                    jobTitle: requesterProfile.jobTitle,
                    companyName: requesterProfile.companyName,
                    actionTypeName,
                    actionTypeNameEn,
                    startDate: fields.StartDate,
                    endDate: fields.EndDate,
                    takenDays: Number(fields.TakenDays) || 0,
                    requestSubject: fields.RequestSubject,
                    comments: fields.Comments || undefined,
                    standInName: fields.StandInEmployee ? resolveName(fields.StandInEmployee) : undefined,
                    signatures,
                });

                const pdfRecipients = new Set<string>([requester]);
                const pdfMessage = `Your leave request "${fields.RequestSubject}" was approved. The completed request form is attached.`;
                const pdfSubject = `Leave Request - ${requesterName}`;
                const pdfFileName = `LeaveRequest_${requester.split("@")[0]}_${body.id}.pdf`;
                const pdfDetails = [periodDetail, daysDetail, { label: "Approved by", value: currentUserName }];

                await Promise.all(
                    Array.from(pdfRecipients).map((recipient) =>
                        sendLeaveEmailWithAttachment(
                            tenantId,
                            recipient,
                            pdfSubject,
                            pdfMessage,
                            teamsDeepLink,
                            {
                                name: pdfFileName,
                                contentType: "application/pdf",
                                contentBytes: pdfBuffer.toString("base64"),
                            },
                            pdfDetails
                        ).catch((pdfEmailError) =>
                            context.warn(`sendLeaveEmailWithAttachment to ${recipient} failed (non-blocking)`, pdfEmailError)
                        )
                    )
                );

                // Own try/catch, independent of the email send above — a SharePoint
                // hiccup here shouldn't be blamed as an email failure or vice versa.
                // Skipped entirely (not an error) for a tenant not yet provisioned
                // with LR_Approvals — see tenantConfig.ts's approvalsLibraryId note.
                if (approvalsLibraryId) {
                    try {
                        await saveApprovedLeaveRequestPdf(graphClient, siteId, approvalsLibraryId, {
                            employeeFullName: requesterName,
                            employeeCode: requester,
                            requestId: body.id,
                            fileName: pdfFileName,
                            pdfBuffer,
                        });
                    } catch (archiveError) {
                        context.warn("saveApprovedLeaveRequestPdf failed (non-blocking)", archiveError);
                    }
                }
            } catch (pdfError) {
                context.warn("PDF generation/send failed (non-blocking)", pdfError);
            }
        }

        // Only the mid-chain "next approver" email is an approval-needed notification —
        // that's the one that gets Approve/Decline Actionable Message buttons.
        const nextApproverCardScriptTag =
            newStatus === "PENDING"
                ? buildApprovalCardScriptTag({
                      requesterName,
                      requestSubject: fields.RequestSubject,
                      startDateLabel: formatDate(fields.StartDate),
                      endDateLabel: formatDate(fields.EndDate),
                      takenDays: Number(fields.TakenDays) || 0,
                      requestId: body.id,
                      tabBaseUrl,
                      teamsDeepLink,
                  })
                : null;

        const subject = `Leave Request - ${requesterName}`;
        // Whether the primary notification — the one addressed to whoever needs to see
        // this decision next (the next approver mid-chain, or the requester once the
        // request is fully resolved) — actually went out. Surfaced to the acting user so
        // a silent email failure (best-effort everywhere else) isn't invisible to them.
        let primaryEmailSent = false;
        await Promise.all([
            ...notifyRecipients.flatMap((recipient) => [
                sendLeaveActivityNotification(graphClient, recipient, notifyMessage, tabBaseUrl).catch(
                    (notificationError) =>
                        context.warn(`sendLeaveActivityNotification to ${recipient} failed (non-blocking)`, notificationError)
                ),
                (newStatus === "PENDING" && recipient === newCurrentUser && nextApproverCardScriptTag
                    ? sendApprovalActionEmail(
                          tenantId,
                          recipient,
                          subject,
                          notifyMessage,
                          teamsDeepLink,
                          nextApproverCardScriptTag,
                          notifyDetails
                      )
                    : sendLeaveEmailNotification(tenantId, recipient, subject, notifyMessage, teamsDeepLink, notifyDetails)
                )
                    .then(() => {
                        if (recipient.toLowerCase() === newCurrentUser?.toLowerCase()) {
                            primaryEmailSent = true;
                        }
                    })
                    .catch((emailError: unknown) =>
                        context.warn(`email to ${recipient} failed (non-blocking)`, emailError)
                    ),
            ]),
            ...(requesterUpdateMessage
                ? [
                      sendLeaveActivityNotification(graphClient, requester, requesterUpdateMessage, tabBaseUrl).catch(
                          (notificationError) =>
                              context.warn("sendLeaveActivityNotification to requester failed (non-blocking)", notificationError)
                      ),
                      sendLeaveEmailNotification(
                          tenantId,
                          requester,
                          subject,
                          requesterUpdateMessage,
                          teamsDeepLink,
                          requesterUpdateDetails
                      ).catch((emailError) =>
                          context.warn("sendLeaveEmailNotification to requester failed (non-blocking)", emailError)
                      ),
                  ]
                : []),
        ]);

        return { jsonBody: { id: body.id, status: newStatus, currentUser: newCurrentUser, emailSent: primaryEmailSent } };
    } catch (error) {
        context.error("RespondToLeaveRequest failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_RespondToLeaveRequest", {
    route: "leave/RespondToLeaveRequest",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: RespondToLeaveRequest,
});
