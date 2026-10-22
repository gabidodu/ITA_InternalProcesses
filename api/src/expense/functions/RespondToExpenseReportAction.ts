import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { getAppOnlyGraphToken } from "../../shared/auth/appOnly";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";
import { validateActionableMessageToken } from "../../shared/auth/actionableMessageToken";
import { getSetupRows } from "../lib/setup";
import { sendExpenseEmailNotification, sendApprovalActionEmail } from "../lib/emailNotification";
import { logAuditEntry } from "../../shared/lib/auditLog";
import { buildDisplayNameMap, getEmployeeDirectory } from "../../shared/lib/employees";
import { getTeamsDeepLink } from "../../shared/lib/teamsLink";
import { buildApprovalCardScriptTag, buildActionRecordedCard } from "../lib/actionableMessage";

type Decision = "approved" | "declined";

// Graph's error shape for a failed If-Match precondition on a SharePoint list
// item PATCH doesn't reliably expose a numeric statusCode on the thrown error
// in this SDK version — checked live: the actual property came back
// unpopulated even though the request genuinely 412'd — so fall back to
// matching Graph's own wording ("The resource has changed since the caller
// last read it; usually an eTag mismatch") rather than trust statusCode alone.
function isEtagConflictError(error: unknown): boolean {
    const err = error as { statusCode?: number; message?: string };
    return err?.statusCode === 412 || /eTag mismatch|resource has changed/i.test(err?.message ?? "");
}

/**
 * Approve/Decline target for Outlook Actionable Messages' Action.Http
 * buttons — the email-native counterpart to RespondToExpenseReport (which
 * runs inside the Teams tab via delegated OBO). No Teams SSO context exists
 * from an email click, so this endpoint is public and instead authenticates
 * the caller via the Entra ID bearer token Microsoft's Actions service
 * attaches to the request (see auth/actionableMessageToken.ts); the acting
 * user's identity comes from that token, not a header we control. Uses
 * app-only Graph throughout (no delegated/OBO token is available here).
 *
 * Deliberately duplicates RespondToExpenseReport's decision/transition logic
 * rather than sharing it — same reasoning as the sibling LeaveRequestForTeams
 * app's RespondToLeaveRequestAction: the two entry points have fundamentally
 * different trust models, not worth the regression risk of coupling them.
 * Approve/decline is the only decision reachable this way — editing and
 * resubmitting a declined report (ResubmitExpenseReport.ts) has no
 * email-actionable counterpart, by design (needs the full form).
 */
export async function RespondToExpenseReportAction(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const tabBaseUrl = process.env.TAB_BASE_URL;
        const teamsAppId = process.env.TEAMS_APP_ID;
        if (!tabBaseUrl || !teamsAppId) {
            throw new Error("TAB_BASE_URL / TEAMS_APP_ID are not configured.");
        }
        const teamsDeepLink = getTeamsDeepLink(teamsAppId, tabBaseUrl, "expense");

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
        // as the entire POST body (see lib/actionableMessage.ts) — not JSON, since Outlook's
        // input substitution doesn't escape the value.
        const comment = (await request.text())?.trim() || undefined;

        // Validated first, before anything else: the action token is the only source of
        // a tenant id on this endpoint — which client's SharePoint config to use, and
        // which tenant to mint the app-only Graph token for.
        let actingUserPrincipal: string;
        let tenantId: string;
        try {
            // The card's Action.Http sets Authorization to "" so Outlook sends its real
            // bearer token via Action-Authorization instead — Azure Static Web Apps'
            // managed-API proxy overwrites the standard Authorization header with its own
            // internal token before it reaches this code (same gotcha as Teams SSO).
            ({ actingUser: actingUserPrincipal, tenantId } = await validateActionableMessageToken(
                request.headers.get("action-authorization")
            ));
        } catch (authError) {
            context.warn("Actionable Message token validation failed", authError);
            return { status: 401, jsonBody: { error: "Invalid or missing action token." } };
        }

        const { siteId, expenseReportsListId, setupListId, auditLogListId } = await getExpenseConfig(tenantId);
        const graphToken = await getAppOnlyGraphToken(tenantId);
        const graphClient = Client.init({ authProvider: (done) => done(null, graphToken) });

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const nameByMail = buildDisplayNameMap(employees);
        const resolveName = (mail: string | undefined) => (mail && nameByMail.get(mail.toLowerCase())) || mail || "";

        const item = await graphClient.api(`/sites/${siteId}/lists/${expenseReportsListId}/items/${requestId}`).expand("fields").get();
        const fields = item.fields as Record<string, string>;
        const etag: string | undefined = item.eTag ?? item["@odata.etag"];
        context.log(`RespondToExpenseReportAction: report ${requestId}, etag=${etag ?? "MISSING"}`);
        if (!etag) {
            // Without a concurrency token we can't guard the PATCH below against the
            // app/email double-response race (see RespondToExpenseReport.ts's matching
            // comment) — fail loudly rather than silently proceeding unprotected.
            throw new Error(`Could not read a concurrency token (ETag) for report ${requestId} from Graph.`);
        }

        // Status must be checked in addition to CurrentUser: once a report is
        // DECLINED, CurrentUser is repointed to the requester (so they can edit
        // and resubmit it — see ResubmitExpenseReport.ts) — a completely
        // different meaning than "awaiting this person's approval". Without this
        // check, a requester who is also somewhere in their own approval chain
        // (e.g. testing) could still "approve" a stale Approve/Decline email for
        // a request that was already declined, because CurrentUser now happens
        // to match them too.
        if (fields.Status !== "PENDING") {
            return {
                status: 200,
                headers: { "CARD-UPDATE-IN-BODY": "true" },
                jsonBody: buildActionRecordedCard(
                    `This expense report is no longer awaiting approval (current status: ${fields.Status}).`,
                    "This report was already answered."
                ),
            };
        }
        if (fields.CurrentUser?.toLowerCase() !== actingUserPrincipal.toLowerCase()) {
            return {
                status: 200,
                headers: { "CARD-UPDATE-IN-BODY": "true" },
                jsonBody: buildActionRecordedCard(
                    "This expense report is not currently awaiting your action.",
                    "This report was already answered."
                ),
            };
        }

        const actingUserName = resolveName(actingUserPrincipal);
        const requester = fields.EmployeeCode;
        const requesterName = resolveName(requester);
        const remainingChain = (fields.ApprovalFlow ?? "")
            .split(";")
            .map((s) => s.trim())
            .filter(Boolean);

        const subjectDetail = { label: "Subject", value: fields.RequestSubject };
        const totalDetail = { label: "Total", value: `${fields.TotalAmount} ${fields.Currency}` };

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
            notifyMessage = `Your expense report "${fields.RequestSubject}" was declined.`;
            notifyRecipients = [requester];
            notifyDetails = [subjectDetail, totalDetail, { label: "Declined by", value: actingUserName }];
        } else if (remainingChain.length > 0) {
            const [nextApprover, ...rest] = remainingChain;
            newStatus = "PENDING";
            newCurrentUser = nextApprover;
            newRemainingChain = rest;
            notifyMessage = `${requesterName} needs your approval for an expense report.`;
            notifyRecipients = [nextApprover];
            notifyDetails = [subjectDetail, totalDetail, { label: "Requested by", value: requesterName }];
            requesterUpdateMessage = `Your expense report "${fields.RequestSubject}" was approved by ${actingUserName} and is now awaiting approval from ${resolveName(nextApprover)}.`;
            requesterUpdateDetails = [
                subjectDetail,
                totalDetail,
                { label: "Approved by", value: actingUserName },
                { label: "Next approver", value: resolveName(nextApprover) },
            ];
        } else {
            newStatus = "APPROVED";
            newCurrentUser = "";
            newRemainingChain = [];
            notifyMessage = `Your expense report "${fields.RequestSubject}" has been fully approved.`;
            notifyRecipients = [requester];
            notifyDetails = [subjectDetail, totalDetail, { label: "Approved by", value: actingUserName }];
        }

        if (comment) {
            notifyDetails = [...notifyDetails, { label: "Comment", value: comment }];
            if (requesterUpdateDetails) {
                requesterUpdateDetails = [...requesterUpdateDetails, { label: "Comment", value: comment }];
            }
        }

        notifyRecipients = Array.from(new Map(notifyRecipients.map((r) => [r.toLowerCase(), r])).values());

        try {
            const patchRequest = graphClient.api(`/sites/${siteId}/lists/${expenseReportsListId}/items/${requestId}/fields`);
            // If-Match makes this update conditional on the item being unchanged
            // since our GET above — guards against a race where this email click
            // and an app response (RespondToExpenseReport.ts) are both acted on
            // for the same pending step at nearly the same time. Without this,
            // both reads would see the same "still pending for me" state and both
            // writes would succeed, silently applying contradicting decisions.
            patchRequest.header("If-Match", etag);
            await patchRequest.patch({
                Status: newStatus,
                CurrentUser: newCurrentUser,
                ApprovalFlow: newRemainingChain.join(";"),
                Comments: comment ?? fields.Comments ?? "",
            });
            context.log(`RespondToExpenseReportAction: patch for report ${requestId} succeeded with If-Match=${etag}`);
        } catch (patchError) {
            context.log(
                `RespondToExpenseReportAction: patch for report ${requestId} failed, statusCode=${(patchError as { statusCode?: number })?.statusCode}, message=${(patchError as Error)?.message}`
            );
            if (isEtagConflictError(patchError)) {
                return {
                    status: 200,
                    headers: { "CARD-UPDATE-IN-BODY": "true" },
                    jsonBody: buildActionRecordedCard(
                        "Someone already responded to this expense report (from the app or a previous email click) — this click wasn't applied.",
                        "This report was already answered."
                    ),
                };
            }
            throw patchError;
        }

        try {
            await logAuditEntry(graphClient, siteId, auditLogListId, {
                actor: actingUserPrincipal,
                action: decision === "declined" ? "Declined" : newStatus === "APPROVED" ? "ApprovedFinal" : "Approved",
                employeeCode: requester,
                requestId,
                details: comment ?? "",
            });
        } catch (auditError) {
            context.warn("logAuditEntry failed (non-blocking)", auditError);
        }

        // Only the mid-chain "next approver" email is an approval-needed notification —
        // that's the one that gets Approve/Decline Actionable Message buttons.
        const nextApproverCardScriptTag =
            newStatus === "PENDING"
                ? buildApprovalCardScriptTag({
                      requesterName,
                      requestSubject: fields.RequestSubject,
                      totalAmount: Number(fields.TotalAmount) || 0,
                      currency: fields.Currency,
                      requestId,
                      tabBaseUrl,
                      teamsDeepLink,
                      comment: fields.Comments || undefined,
                  })
                : null;

        const subject = `Expense Report - ${requesterName}`;
        await Promise.all([
            ...notifyRecipients.map((recipient) =>
                newStatus === "PENDING" && recipient === newCurrentUser && nextApproverCardScriptTag
                    ? sendApprovalActionEmail(tenantId, recipient, subject, notifyMessage, teamsDeepLink, nextApproverCardScriptTag, notifyDetails).catch(
                          (emailError: unknown) => context.warn(`sendApprovalActionEmail to ${recipient} failed (non-blocking)`, emailError)
                      )
                    : sendExpenseEmailNotification(tenantId, recipient, subject, notifyMessage, teamsDeepLink, notifyDetails).catch((emailError) =>
                          context.warn(`sendExpenseEmailNotification to ${recipient} failed (non-blocking)`, emailError)
                      )
            ),
            ...(requesterUpdateMessage
                ? [
                      sendExpenseEmailNotification(tenantId, requester, subject, requesterUpdateMessage, teamsDeepLink, requesterUpdateDetails).catch(
                          (emailError) => context.warn("sendExpenseEmailNotification to requester failed (non-blocking)", emailError)
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
                ? "You declined this expense report."
                : newStatus === "APPROVED"
                  ? "You approved this expense report. It has been fully approved."
                  : `You approved this expense report. It's now awaiting approval from ${resolveName(newCurrentUser)}.`;
        return {
            status: 200,
            headers: { "CARD-UPDATE-IN-BODY": "true" },
            jsonBody: buildActionRecordedCard(recordedMessage),
        };
    } catch (error) {
        context.error("RespondToExpenseReportAction failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_RespondToExpenseReportAction", {
    route: "expense/RespondToExpenseReportAction",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: RespondToExpenseReportAction,
});
