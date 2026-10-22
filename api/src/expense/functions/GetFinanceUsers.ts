import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";
import { getFinanceEmails, getSetupRows } from "../lib/setup";

/**
 * Returns the emails currently granted the Finance role (DE_Setup
 * Category=Finance) — who get full visibility into every expense report
 * (see GetAllExpenseReports.ts). Read is unrestricted (same precedent as
 * GetApprovalFlow); only SetFinanceUsers enforces the Admin-only write.
 */
export async function GetFinanceUsers(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected"]);
        const { siteId, setupListId } = await getExpenseConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);

        return { jsonBody: { emails: getFinanceEmails(setupRows) } };
    } catch (error) {
        context.error("GetFinanceUsers failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_GetFinanceUsers", {
    route: "expense/GetFinanceUsers",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetFinanceUsers,
});
