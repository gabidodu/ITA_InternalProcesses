import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows, isHrAdmin } from "../lib/setup";
import { buildDisplayNameMap, getEmployeeDirectory, isDirectManagerOf } from "../../shared/lib/employees";
import { getAllListItems } from "../../shared/lib/graphPaging";

type HistoryEntry = {
    timestamp: string;
    actor: string;
    actorName: string;
    action: string;
    details: string;
};

/**
 * Returns the full audit trail (who did what, when) for a single leave
 * request — who approved/declined each step, comments, cancellations.
 * Viewable by the requester, HR, the current approver, anyone who has ever
 * acted as an approver on this specific request (per LR_AuditLog), or the
 * requester's own direct manager (e.g. viewing the "Details" of one of
 * their reports' requests from the Reports tab).
 */
export async function GetRequestHistory(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const requestId = request.query.get("id");
        if (!requestId) {
            throw new Error("id query parameter is required.");
        }

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, [
            "Sites.Selected",
            "GroupMember.Read.All",
            "User.ReadBasic.All",
        ]);
        const { siteId, leaveRequestsListId, auditLogListId, setupListId } = await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;

        const item = await graphClient
            .api(`/sites/${siteId}/lists/${leaveRequestsListId}/items/${requestId}`)
            .expand("fields")
            .get();
        const fields = item.fields as Record<string, string>;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const isRequester = fields.EmployeeCode?.toLowerCase() === currentUserPrincipal?.toLowerCase();
        const isCurrentApprover = fields.CurrentUser?.toLowerCase() === currentUserPrincipal?.toLowerCase();
        const isHr = isHrAdmin(setupRows, currentUserPrincipal);

        const auditItems = await getAllListItems(graphClient, siteId, auditLogListId);

        const entriesForRequest = auditItems.filter(
            (auditItem: { fields: Record<string, string> }) => auditItem.fields.RequestId === requestId
        );

        const wasParticipant = entriesForRequest.some(
            (auditItem: { fields: Record<string, string> }) =>
                auditItem.fields.Actor?.toLowerCase() === currentUserPrincipal?.toLowerCase()
        );

        const isAuthorized =
            isRequester ||
            isHr ||
            isCurrentApprover ||
            wasParticipant ||
            (await isDirectManagerOf(graphClient, currentUserPrincipal, fields.EmployeeCode));
        if (!isAuthorized) {
            return { status: 403, jsonBody: { error: "You don't have access to this request's history." } };
        }

        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const nameByMail = buildDisplayNameMap(employees);
        const resolveName = (mail: string | undefined) => (mail && nameByMail.get(mail.toLowerCase())) || mail || "";

        const history: HistoryEntry[] = entriesForRequest
            .map((auditItem: { fields: Record<string, string> }) => ({
                timestamp: auditItem.fields.Timestamp,
                actor: auditItem.fields.Actor,
                actorName: resolveName(auditItem.fields.Actor),
                action: auditItem.fields.Action,
                details: auditItem.fields.Details,
            }))
            .sort((a: HistoryEntry, b: HistoryEntry) => new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime());

        return { jsonBody: { history } };
    } catch (error) {
        context.error("GetRequestHistory failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_GetRequestHistory", {
    route: "leave/GetRequestHistory",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetRequestHistory,
});
