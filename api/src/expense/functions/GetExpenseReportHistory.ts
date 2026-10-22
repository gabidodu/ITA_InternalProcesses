import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows } from "../lib/setup";
import { buildDisplayNameMap, getEmployeeDirectory } from "../../shared/lib/employees";

type HistoryEntry = {
    timestamp: string;
    actor: string;
    actorName: string;
    action: string;
    details: string;
};

/**
 * Returns the full DE_AuditLog trail for one expense report (?id=), oldest
 * first — every creation/approval/decline/resubmit/cancellation logged by
 * CreateExpenseReport/RespondToExpenseReport/ResubmitExpenseReport/
 * CancelExpenseReport. Not further access-restricted beyond requiring a
 * valid Teams SSO token: a report's history is only ever reachable through
 * a report card the caller could already see (their own report, one
 * they're approving, or the Admin/Finance-only company-wide view).
 */
export async function GetExpenseReportHistory(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const id = request.query.get("id");
        if (!id) {
            throw new Error("id is required.");
        }

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected", "GroupMember.Read.All", "User.ReadBasic.All"]);
        const { siteId, auditLogListId, setupListId } = await getExpenseConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const nameByMail = buildDisplayNameMap(employees);
        const resolveName = (mail: string | undefined) => (mail && nameByMail.get(mail.toLowerCase())) || mail || "";

        const response = await graphClient.api(`/sites/${siteId}/lists/${auditLogListId}/items`).expand("fields").get();

        const history: HistoryEntry[] = response.value
            .filter((item: { fields: Record<string, string> }) => item.fields.RequestId === id)
            .map((item: { fields: Record<string, string> }) => ({
                timestamp: item.fields.Timestamp,
                actor: item.fields.Actor,
                actorName: resolveName(item.fields.Actor),
                action: item.fields.Action,
                details: item.fields.Details ?? "",
            }))
            .sort((a: HistoryEntry, b: HistoryEntry) => new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime());

        return { jsonBody: { history } };
    } catch (error) {
        context.error("GetExpenseReportHistory failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_GetExpenseReportHistory", {
    route: "expense/GetExpenseReportHistory",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetExpenseReportHistory,
});
