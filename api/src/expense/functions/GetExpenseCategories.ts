import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";

type ExpenseCategory = {
    id: number;
    name: string;
    nameRo: string;
    code: string;
};

/**
 * Returns the expense-category catalog (DE_ExpenseCategory list) so the
 * client can populate the "category" dropdown on each expense line.
 */
export async function GetExpenseCategories(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected"]);
        const { siteId, expenseCategoryListId: categoryListId } = await getExpenseConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const itemsResponse = await graphClient.api(`/sites/${siteId}/lists/${categoryListId}/items`).expand("fields").get();

        const categories: ExpenseCategory[] = itemsResponse.value
            .map((item: { fields: Record<string, string> }) => ({
                id: Number(item.fields.IDCategory),
                name: item.fields.CategoryName,
                nameRo: item.fields.CategoryNameRO,
                code: item.fields.CategoryCode,
            }))
            .sort((a: ExpenseCategory, b: ExpenseCategory) => a.id - b.id);

        return { jsonBody: { categories } };
    } catch (error) {
        context.error("GetExpenseCategories failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_GetExpenseCategories", {
    route: "expense/GetExpenseCategories",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetExpenseCategories,
});
