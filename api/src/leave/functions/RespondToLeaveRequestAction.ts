import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { getAppOnlyGraphToken } from "../../shared/auth/appOnly";
import { validateActionableMessageToken } from "../../shared/auth/actionableMessageToken";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows } from "../lib/setup";
import { sendLeaveEmailNotification, sendLeaveEmailWithAttachment, sendApprovalActionEmail } from "../lib/emailNotification";
import { adjustTakenDays } from "../lib/leaveStatus";
import { logAuditEntry, getApprovalSteps, getActorDecision } from "../../shared/lib/auditLog";
import { buildDisplayNameMap, getEmployeeDirectory } from "../../shared/lib/employees";
import { generateLeaveRequestPdf } from "../lib/pdfGenerator";
import { saveApprovedLeaveRequestPdf } from "../lib/approvalArchive";
import { getTeamsDeepLink } from "../../shared/lib/teamsLink";
import { deleteLeaveCalendarEvent, markLeaveCalendarEventApproved } from "../lib/calendarEvent";
import { buildApprovalCardScriptTag, buildActionRecordedCard } from "../lib/actionableMessage";
import { formatDate } from "../../shared/lib/dates";
import { getAllListItems } from "../../shared/lib/graphPaging";

type Decision = "approved" | "declined";

// Graph's error for a failed If-Match precondition on a SharePoint list item
// PATCH doesn't reliably expose statusCode in this SDK version (seen live in
// the sibling ExpenseReportForTeams app: the request genuinely 412'd with the
// property unpopulated) — so also match Graph's own wording.
function isEtagConflictError(error: unknown): boolean {
    const err = error as { statusCode?: number; message?: string };
    return err?.statusCode === 412 || /eTag mismatch|resource has changed/i.test(err?.message ?? "");
}

/**
 * Approve/Decline target for Outlook Actionable Messages' Action.Http
 * buttons — the email-native counterpart to RespondToLeaveRequest (which
 * runs inside the Teams tab via delegated OBO). No Teams SSO context exists
 * from an email click, so this endpoint is public and instead authenticates
 * the caller via the Entra ID bearer token Microsoft's Actions service
 * attaches to the request (see auth/actionableMessageToken.ts); the acting
 * user's identity comes from that token, not a header we control. Uses
 * app-only Graph throughout (no delegated/OBO token is available here) — this
 * app's application permissions already cover SharePoint, Mail.Send and
 * Calendars.ReadWrite, so only the AD-profile lookup (for the approval PDF)
 * and the Teams activity-feed notification degrade gracefully (both need
 * User.Read.All / TeamsActivity.Send as Application permissions, which
 * aren't granted — both are already wrapped in non-blocking try/catch or
 * simply omitted here).
 *
 * Deliberately duplicates RespondToLeaveRequest's decision/transition logic
 * rather than sharing it: the two entry points have fundamentally different
 * trust models, and RespondToLeaveRequest is already live in production — not
 * worth the regression risk of refactoring it to share code with a brand-new,
 * unverified entry point.
 */
export async function RespondToLeaveRequestAction(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const tabBaseUrl = process.env.TAB_BASE_URL;
        const teamsAppId = process.env.TEAMS_APP_ID;
        if (!tabBaseUrl || !teamsAppId) {
            throw new Error("TAB_BASE_URL / TEAMS_APP_ID app settings are not fully configured.");
        }
        const teamsDeepLink = getTeamsDeepLink(teamsAppId, tabBaseUrl, "leave");

        const requestId = request.query.get("id");
        const decisionParam = request.query.get("decision");
        if (!requestId) {
            return { status: 400, jsonBody: { error: "id query parameter is required." } };
        }
        if (decisionParam !== "approved" && decisionParam !== "declined") {
            return { status: 400, jsonBody: { error: 'decision must be "approved" or "declined".' } };
        }
        const decision: Decision = decisionParam;
        // The card's Approve/Decline Action.Http sends the comment Input.Text's raw value
        // as the entire POST body (see actionableMessage.ts) — not JSON, since Outlook's
        // input substitution doesn't escape the value, so wrapping it in hand-built JSON
        // risks an invalid payload if the comment contains a quote.
        const comment = (await request.text())?.trim() || undefined;

        // Validated first, before anything else: this is the only source of a tenant id on
        // this entry point (no Teams SSO context to get one from), and every subsequent
        // step — which client's SharePoint config to use, which tenant to mint an app-only
        // token against — depends on it.
        let actingUserPrincipal: string;
        let tenantId: string;
        try {
            // The card's Action.Http sets Authorization to "" so Outlook sends its real
            // bearer token via Action-Authorization instead — Azure Static Web Apps'
            // managed-API proxy overwrites the standard Authorization header with its own
            // internal token before it reaches this code (same gotcha as Teams SSO
            // earlier in this project), so reading "authorization" here would always see
            // the wrong token.
            const validated = await validateActionableMessageToken(request.headers.get("action-authorization"));
            actingUserPrincipal = validated.actingUser;
            tenantId = validated.tenantId;
        } catch (authError) {
            context.warn("Actionable Message token validation failed", authError);
            return { status: 401, jsonBody: { error: "Invalid or missing action token." } };
        }

        const graphToken = await getAppOnlyGraphToken(tenantId);
        const graphClient = Client.init({ authProvider: (done) => done(null, graphToken) });
        const { siteId, leaveRequestsListId, leaveStatusListId, actionTypeListId, setupListId, auditLogListId, approvalsLibraryId } =
            await getTenantConfig(tenantId);

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const nameByMail = buildDisplayNameMap(employees);
        const resolveName = (mail: string | undefined) => (mail && nameByMail.get(mail.toLowerCase())) || mail || "";

        const item = await graphClient
            .api(`/sites/${siteId}/lists/${leaveRequestsListId}/items/${requestId}`)
            .expand("fields")
            .get();

        const fields = item.fields as Record<string, string>;
        if (fields.CurrentUser?.toLowerCase() !== actingUserPrincipal.toLowerCase()) {
            // A repeat of a click that already went through — Outlook retries when the first
            // call is slow (e.g. a cold start from zero replicas), and that retry used to
            // surface as a 403 even though the approval had succeeded. Answer it with the
            // same "recorded" card instead, which also removes the buttons.
            const priorDecision = await getActorDecision(
                graphClient,
                siteId,
                auditLogListId,
                requestId,
                actingUserPrincipal
            ).catch(() => undefined);
            if (priorDecision) {
                const statusNote =
                    fields.Status === "PENDING"
                        ? ` It's now awaiting approval from ${resolveName(fields.CurrentUser)}.`
                        : ` Its status is now ${fields.Status}.`;
                return {
                    status: 200,
                    headers: { "CARD-UPDATE-IN-BODY": "true" },
                    jsonBody: buildActionRecordedCard(
                        `You already ${priorDecision === "Declined" ? "declined" : "approved"} this leave request.${statusNote}`
                    ),
                };
            }
            // Answered with a refresh card rather than a bare 403: Outlook only shows a
            // generic "remote endpoint returned an error" toast for non-2xx responses,
            // and the card also removes buttons that can no longer do anything.
            return {
                status: 200,
                headers: { "CARD-UPDATE-IN-BODY": "true" },
                jsonBody: buildActionRecordedCard(
                    fields.Status === "PENDING"
                        ? `This leave request is awaiting approval from ${resolveName(fields.CurrentUser)}, not from you.`
                        : `This leave request is no longer awaiting approval (current status: ${fields.Status}).`,
                    "No action needed from you."
                ),
            };
        }
        // Guards against the same step being actioned twice at once — e.g. this email and
        // the app (RespondToLeaveRequest) racing each other. Both would pass the CurrentUser
        // check above off the same read; the conditional PATCH below only lets the first one
        // through, since it changes the item's eTag.
        const itemETag: string | undefined = item.eTag;

        const actingUserName = resolveName(actingUserPrincipal);
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
        let requesterUpdateMessage: string | undefined;
        let requesterUpdateDetails: { label: string; value: string }[] | undefined;

        if (decision === "declined") {
            newStatus = "DECLINED";
            newCurrentUser = requester;
            newRemainingChain = [];
            notifyMessage = `Your leave request "${fields.RequestSubject}" was declined.`;
            notifyRecipients = [requester];
            notifyDetails = [periodDetail, daysDetail, { label: "Declined by", value: actingUserName }];
        } else if (remainingChain.length > 0) {
            const [nextApprover, ...rest] = remainingChain;
            newStatus = "PENDING";
            newCurrentUser = nextApprover;
            newRemainingChain = rest;
            notifyMessage = `${requesterName} needs your approval for a leave request.`;
            notifyRecipients = [nextApprover];
            notifyDetails = [periodDetail, daysDetail, { label: "Requested by", value: requesterName }];
            requesterUpdateMessage = `Your leave request "${fields.RequestSubject}" was approved by ${actingUserName} and is now awaiting approval from ${resolveName(nextApprover)}.`;
            requesterUpdateDetails = [
                periodDetail,
                daysDetail,
                { label: "Approved by", value: actingUserName },
                { label: "Next approver", value: resolveName(nextApprover) },
            ];
        } else {
            newStatus = "APPROVED";
            newCurrentUser = requester;
            newRemainingChain = [];
            notifyMessage = `Your leave request "${fields.RequestSubject}" was approved.`;
            notifyRecipients = [requester];
            notifyDetails = [periodDetail, daysDetail, { label: "Approved by", value: actingUserName }];
        }

        if (comment) {
            notifyDetails = [...notifyDetails, { label: "Comment", value: comment }];
            if (requesterUpdateDetails) {
                requesterUpdateDetails = [...requesterUpdateDetails, { label: "Comment", value: comment }];
            }
        }

        // Dedupe case-insensitively — the requester can end up listed twice in some
        // branches (e.g. also being the acting approver on a test account).
        notifyRecipients = Array.from(new Map(notifyRecipients.map((r) => [r.toLowerCase(), r])).values());

        try {
            await graphClient
                .api(`/sites/${siteId}/lists/${leaveRequestsListId}/items/${requestId}/fields`)
                .header("If-Match", itemETag ?? "*")
                .patch({
                    Status: newStatus,
                    CurrentUser: newCurrentUser,
                    ApprovalFLow: newRemainingChain.join(";"),
                });
        } catch (patchError) {
            if (isEtagConflictError(patchError)) {
                return {
                    status: 200,
                    headers: { "CARD-UPDATE-IN-BODY": "true" },
                    jsonBody: buildActionRecordedCard(
                        "Someone already responded to this leave request (from the app or another device) — this click wasn't applied.",
                        "This request was already answered."
                    ),
                };
            }
            throw patchError;
        }

        // Days are reserved at CreateLeaveRequest time, not here — final approval needs no
        // balance change. A decline releases the reservation, since the leave never happens.
        if (decision === "declined") {
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
                if (decision === "declined") {
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
                actor: actingUserPrincipal,
                action: decision === "declined" ? "Declined" : newStatus === "APPROVED" ? "ApprovedFinal" : "Approved",
                employeeCode: requester,
                requestId,
                details:
                    `"${fields.RequestSubject}" — ${decision} by ${actingUserName} via Outlook actionable message. New status: ${newStatus}.` +
                    (comment ? ` Comment: ${comment}` : ""),
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
                    context.warn(
                        "Failed to fetch requester AD profile for PDF (non-blocking — needs User.Read.All as an Application permission, not yet granted)",
                        profileError
                    );
                }

                const approvalSteps = await getApprovalSteps(graphClient, siteId, auditLogListId, requestId);
                const signatures = approvalSteps.map((step) => ({
                    actorName: resolveName(step.actor),
                    timestamp: step.timestamp,
                }));

                const pdfBuffer = await generateLeaveRequestPdf({
                    requestId,
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
                const pdfFileName = `LeaveRequest_${requester.split("@")[0]}_${requestId}.pdf`;
                const pdfDetails = [periodDetail, daysDetail, { label: "Approved by", value: actingUserName }];

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

                // Own try/catch, independent of the email send above — see the
                // matching comment in RespondToLeaveRequest.ts.
                if (approvalsLibraryId) {
                    try {
                        await saveApprovedLeaveRequestPdf(graphClient, siteId, approvalsLibraryId, {
                            employeeFullName: requesterName,
                            employeeCode: requester,
                            requestId,
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
                      requestId,
                      tabBaseUrl,
                      teamsDeepLink,
                  })
                : null;

        const subject = `Leave Request - ${requesterName}`;
        await Promise.all([
            ...notifyRecipients.map((recipient) =>
                newStatus === "PENDING" && recipient === newCurrentUser && nextApproverCardScriptTag
                    ? sendApprovalActionEmail(
                          tenantId,
                          recipient,
                          subject,
                          notifyMessage,
                          teamsDeepLink,
                          nextApproverCardScriptTag,
                          notifyDetails
                      ).catch((emailError: unknown) =>
                          context.warn(`sendApprovalActionEmail to ${recipient} failed (non-blocking)`, emailError)
                      )
                    : sendLeaveEmailNotification(tenantId, recipient, subject, notifyMessage, teamsDeepLink, notifyDetails).catch(
                          (emailError) => context.warn(`sendLeaveEmailNotification to ${recipient} failed (non-blocking)`, emailError)
                      )
            ),
            ...(requesterUpdateMessage
                ? [
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

        // A "refresh card" response — Outlook replaces the Approve/Decline card in the
        // email with this one (CARD-UPDATE-IN-BODY: true tells it to parse the body as
        // a new card), permanently removing the buttons so the same click can't be
        // submitted twice.
        const recordedMessage =
            decision === "declined"
                ? "You declined this leave request."
                : newStatus === "APPROVED"
                  ? "You approved this leave request. It is now fully approved."
                  : `You approved this leave request. It's now awaiting approval from ${resolveName(newCurrentUser)}.`;
        return {
            status: 200,
            headers: { "CARD-UPDATE-IN-BODY": "true" },
            jsonBody: buildActionRecordedCard(recordedMessage),
        };
    } catch (error) {
        context.error("RespondToLeaveRequestAction failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_RespondToLeaveRequestAction", {
    route: "leave/RespondToLeaveRequestAction",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: RespondToLeaveRequestAction,
});
