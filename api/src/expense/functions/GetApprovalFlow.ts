import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";
import { getApprovalChainFor, getSetupRows, hasApprovalChainOverride, parseApprovalChain } from "../lib/setup";

/**
 * Returns an approval chain: the global standard one (Setting/approvalFlow)
 * if no ?employeeCode is given, or that specific employee's override chain
 * if they have one — falling back to the standard chain as a starting point
 * (isOverride: false) so HR sees what they'd inherit before customizing it.
 */
export async function GetApprovalFlow(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const employeeCode = request.query.get("employeeCode") ?? undefined;

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected"]);
        const { siteId, setupListId } = await getExpenseConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const employeesWithOverrides = setupRows
            .filter((r) => r.category === "ApprovalFlow")
            .map((r) => r.settingKey?.toLowerCase())
            .filter((key): key is string => Boolean(key));

        if (!employeeCode) {
            return { jsonBody: { chain: parseApprovalChain(setupRows), isOverride: true, employeesWithOverrides } };
        }

        return {
            jsonBody: {
                chain: getApprovalChainFor(setupRows, employeeCode),
                isOverride: hasApprovalChainOverride(setupRows, employeeCode),
                employeesWithOverrides,
            },
        };
    } catch (error) {
        context.error("GetApprovalFlow failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_GetApprovalFlow", {
    route: "expense/GetApprovalFlow",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetApprovalFlow,
});
