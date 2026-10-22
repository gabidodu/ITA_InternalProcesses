import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { getSetupRows, canAccessSettings } from "../lib/setup";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";
import { sendExpenseEmailNotification } from "../lib/emailNotification";
import { logAuditEntry, wasApproverOn } from "../../shared/lib/auditLog";
import { buildDisplayNameMap, getEmployeeDirectory } from "../../shared/lib/employees";
import { getTeamsDeepLink } from "../../shared/lib/teamsLink";

type CancelExpenseReportBody = {
    id: string;
};

function parseAndValidateBody(raw: unknown): CancelExpenseReportBody {
    const body = raw as Partial<CancelExpenseReportBody> | null;
    if (!body?.id || typeof body.id !== "string") {
        throw new Error("id is required.");
    }
    return { id: body.id };
}

/**
 * Cancels an expense report. Allowed for: the original requester, any
 * HR/Admin, or a past approver on this specific request (checked against
 * DE_AuditLog "Approved"/"ApprovedFinal" entries — see lib/auditLog.ts's
 * wasApproverOn). Only while still PENDING — approval is the final step of
 * the workflow, so an already-APPROVED report (like a DECLINED/CANCELLED
 * one) is done and can't be cancelled. Note: the frontend deliberately never
 * renders Cancel alongside Approve/Decline on the same row — this is what
 * prevents the leave-app's known double-click race (see
 * [[project-leave-teams-app]]), kept as a UI convention here rather than a
 * server-side lock.
 */
export async function CancelExpenseReport(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
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

        const item = await graphClient.api(`/sites/${siteId}/lists/${expenseReportsListId}/items/${body.id}`).expand("fields").get();
        const fields = item.fields as Record<string, string>;

        if (fields.Status !== "PENDING") {
            return { status: 400, jsonBody: { error: "This expense report is already resolved and cannot be cancelled." } };
        }

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const nameByMail = buildDisplayNameMap(employees);
        const resolveName = (mail: string | undefined) => (mail && nameByMail.get(mail.toLowerCase())) || mail || "";

        const isRequester = fields.EmployeeCode?.toLowerCase() === currentUserPrincipal?.toLowerCase();
        const isSettingsAdmin = canAccessSettings(setupRows, currentUserPrincipal);
        const wasParticipant =
            !isRequester && !isSettingsAdmin && (await wasApproverOn(graphClient, siteId, auditLogListId, body.id, currentUserPrincipal));

        if (!isRequester && !isSettingsAdmin && !wasParticipant) {
            return {
                status: 403,
                jsonBody: { error: "Only the requester, HR/Admin, or a past approver can cancel this expense report." },
            };
        }

        const requester = fields.EmployeeCode;
        const requesterName = resolveName(requester);

        await graphClient.api(`/sites/${siteId}/lists/${expenseReportsListId}/items/${body.id}/fields`).patch({
            Status: "CANCELLED",
            CurrentUser: requester,
            ApprovalFlow: "",
        });

        try {
            await logAuditEntry(graphClient, siteId, auditLogListId, {
                actor: currentUserPrincipal,
                action: "Cancelled",
                employeeCode: requester,
                requestId: body.id,
                // The history view (GetExpenseReportHistory) shows only Timestamp/
                // Action/Actor plus this as an optional comment — no comment concept
                // exists for a cancellation, so nothing to put here.
                details: "",
            });
        } catch (auditError) {
            context.warn("logAuditEntry failed (non-blocking)", auditError);
        }

        const notifyRecipients = new Set<string>([requester]);
        if (fields.CurrentUser && fields.CurrentUser.toLowerCase() !== requester?.toLowerCase()) {
            notifyRecipients.add(fields.CurrentUser);
        }
        const notifyMessage = `${requesterName}'s expense report "${fields.RequestSubject}" was cancelled.`;
        const subject = `Expense Report - ${requesterName}`;

        await Promise.all(
            Array.from(notifyRecipients).map((recipient) =>
                sendExpenseEmailNotification(tenantId, recipient, subject, notifyMessage, teamsDeepLink).catch((emailError) =>
                    context.warn(`sendExpenseEmailNotification to ${recipient} failed (non-blocking)`, emailError)
                )
            )
        );

        return { jsonBody: { id: body.id, status: "CANCELLED" } };
    } catch (error) {
        context.error("CancelExpenseReport failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_CancelExpenseReport", {
    route: "expense/CancelExpenseReport",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: CancelExpenseReport,
});
