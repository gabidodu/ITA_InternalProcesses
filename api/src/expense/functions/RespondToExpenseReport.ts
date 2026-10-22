import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows } from "../lib/setup";
import { sendExpenseEmailNotification, sendApprovalActionEmail } from "../lib/emailNotification";
import { logAuditEntry } from "../../shared/lib/auditLog";
import { buildDisplayNameMap, getEmployeeDirectory } from "../../shared/lib/employees";
import { getTeamsDeepLink } from "../../shared/lib/teamsLink";
import { buildApprovalCardScriptTag } from "../lib/actionableMessage";

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

type RespondToExpenseReportBody = {
    id: string;
    decision: Decision;
    comment?: string;
};

function parseAndValidateBody(raw: unknown): RespondToExpenseReportBody {
    const body = raw as Partial<RespondToExpenseReportBody> | null;
    if (!body?.id || typeof body.id !== "string") {
        throw new Error("id is required.");
    }
    if (body.decision !== "approved" && body.decision !== "declined") {
        throw new Error('decision must be "approved" or "declined".');
    }
    return { id: body.id, decision: body.decision, comment: body.comment?.trim() || undefined };
}

/**
 * Approves or declines an expense report as the current approver. Advances
 * the remaining approval chain (stored in ApprovalFlow) to the next
 * approver, or resolves the request once the chain is exhausted or on a
 * decline. Final approval sets Status=APPROVED with CurrentUser cleared —
 * that's the terminal state, approval is the last step of the workflow. A
 * decline sets Status=DECLINED with CurrentUser back on the requester, who
 * can then edit and resubmit it (see ResubmitExpenseReport.ts) — that
 * re-enters this same chain from the top. Every step is logged to
 * DE_AuditLog.
 */
export async function RespondToExpenseReport(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const tabBaseUrl = process.env.TAB_BASE_URL;
        const teamsAppId = process.env.TEAMS_APP_ID;
        if (!tabBaseUrl || !teamsAppId) {
            throw new Error("TAB_BASE_URL / TEAMS_APP_ID are not configured.");
        }
        const teamsDeepLink = getTeamsDeepLink(teamsAppId, tabBaseUrl, "expense");

        const body = parseAndValidateBody(await request.json());

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected", "GroupMember.Read.All", "User.ReadBasic.All"]);
        const { siteId, expenseReportsListId, setupListId, auditLogListId } = await getExpenseConfig(tenantId);

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

        const item = await graphClient.api(`/sites/${siteId}/lists/${expenseReportsListId}/items/${body.id}`).expand("fields").get();
        const fields = item.fields as Record<string, string>;
        const etag: string | undefined = item.eTag ?? item["@odata.etag"];
        context.log(`RespondToExpenseReport: report ${body.id}, etag=${etag ?? "MISSING"}`);
        if (!etag) {
            // Without a concurrency token we can't guard the PATCH below against the
            // app/email double-response race (see RespondToExpenseReportAction.ts's
            // matching comment) — fail loudly rather than silently proceeding unprotected.
            throw new Error(`Could not read a concurrency token (ETag) for report ${body.id} from Graph.`);
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
                status: 400,
                jsonBody: { error: `This expense report is no longer awaiting approval (current status: ${fields.Status}).` },
            };
        }
        if (fields.CurrentUser?.toLowerCase() !== currentUserPrincipal?.toLowerCase()) {
            return {
                status: 403,
                jsonBody: { error: "This expense report is not currently awaiting your action." },
            };
        }

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

        if (body.decision === "declined") {
            newStatus = "DECLINED";
            newCurrentUser = requester;
            newRemainingChain = [];
            notifyMessage = `Your expense report "${fields.RequestSubject}" was declined.`;
            notifyRecipients = [requester];
            notifyDetails = [subjectDetail, totalDetail, { label: "Declined by", value: currentUserName }];
        } else if (remainingChain.length > 0) {
            const [nextApprover, ...rest] = remainingChain;
            newStatus = "PENDING";
            newCurrentUser = nextApprover;
            newRemainingChain = rest;
            notifyMessage = `${requesterName} needs your approval for an expense report.`;
            notifyRecipients = [nextApprover];
            notifyDetails = [subjectDetail, totalDetail, { label: "Requested by", value: requesterName }];
            requesterUpdateMessage = `Your expense report "${fields.RequestSubject}" was approved by ${currentUserName} and is now awaiting approval from ${resolveName(nextApprover)}.`;
            requesterUpdateDetails = [
                subjectDetail,
                totalDetail,
                { label: "Approved by", value: currentUserName },
                { label: "Next approver", value: resolveName(nextApprover) },
            ];
        } else {
            newStatus = "APPROVED";
            newCurrentUser = "";
            newRemainingChain = [];
            notifyMessage = `Your expense report "${fields.RequestSubject}" has been fully approved.`;
            notifyRecipients = [requester];
            notifyDetails = [subjectDetail, totalDetail, { label: "Approved by", value: currentUserName }];
        }

        if (body.comment) {
            notifyDetails = [...notifyDetails, { label: "Comment", value: body.comment }];
            if (requesterUpdateDetails) {
                requesterUpdateDetails = [...requesterUpdateDetails, { label: "Comment", value: body.comment }];
            }
        }

        // Dedupe case-insensitively (e.g. the requester also being the next
        // approver in a testing scenario) so nobody gets the same email twice.
        notifyRecipients = Array.from(new Map(notifyRecipients.map((r) => [r.toLowerCase(), r])).values());

        try {
            const patchRequest = graphClient.api(`/sites/${siteId}/lists/${expenseReportsListId}/items/${body.id}/fields`);
            // If-Match makes this update conditional on the item being unchanged
            // since our GET above — guards against a race where the app and the
            // Outlook email Approve/Decline button (RespondToExpenseReportAction.ts)
            // are both acted on for the same pending step at nearly the same time.
            // Without this, both reads would see the same "still pending for me"
            // state and both writes would succeed, silently applying contradicting
            // decisions.
            patchRequest.header("If-Match", etag);
            await patchRequest.patch({
                Status: newStatus,
                CurrentUser: newCurrentUser,
                ApprovalFlow: newRemainingChain.join(";"),
                Comments: body.comment ?? fields.Comments ?? "",
            });
            context.log(`RespondToExpenseReport: patch for report ${body.id} succeeded with If-Match=${etag}`);
        } catch (patchError) {
            context.log(
                `RespondToExpenseReport: patch for report ${body.id} failed, statusCode=${(patchError as { statusCode?: number })?.statusCode}, message=${(patchError as Error)?.message}`
            );
            if (isEtagConflictError(patchError)) {
                return {
                    status: 409,
                    jsonBody: {
                        error: "This expense report was already responded to — from another device or the email notification — just now.",
                    },
                };
            }
            throw patchError;
        }

        try {
            await logAuditEntry(graphClient, siteId, auditLogListId, {
                actor: currentUserPrincipal,
                action: body.decision === "declined" ? "Declined" : newStatus === "APPROVED" ? "ApprovedFinal" : "Approved",
                employeeCode: requester,
                requestId: body.id,
                // The history view (GetExpenseReportHistory) shows only Timestamp/
                // Action/Actor plus this as an optional comment — no narrative here.
                details: body.comment ?? "",
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
                      requestId: body.id,
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

        return { jsonBody: { id: body.id, status: newStatus, currentUser: newCurrentUser } };
    } catch (error) {
        context.error("RespondToExpenseReport failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_RespondToExpenseReport", {
    route: "expense/RespondToExpenseReport",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: RespondToExpenseReport,
});
