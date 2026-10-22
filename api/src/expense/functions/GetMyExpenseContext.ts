import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";
import { canAccessSettings, getSetupRows, isFinanceUser } from "../lib/setup";
import { buildDisplayNameMap, getEmployeeDirectory } from "../../shared/lib/employees";
import { byRequestDateDesc, loadAllExpenseReports } from "../lib/expenseData";

/**
 * Returns the signed-in user's expense-report context: their own reports
 * ("myReports", every status) and reports awaiting their approval
 * ("pendingApprovals" — CurrentUser === me && Status === PENDING). Approval
 * is the final step of the workflow — there's no payment layer. Each report
 * is enriched with its lines (DE_ExpenseLines) and attachments
 * (DE_Attachments) via lib/expenseData.ts, shared with GetAllExpenseReports
 * (the company-wide, Admin/Finance-only overview of every report).
 */
export async function GetMyExpenseContext(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected", "GroupMember.Read.All", "User.ReadBasic.All"]);
        const { siteId, expenseReportsListId, expenseLinesListId, expenseCategoryListId: categoryListId, expenseAttachmentsListId: attachmentsListId, setupListId } = await getExpenseConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("displayName,mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const nameByMail = buildDisplayNameMap(employees);
        const resolveName = (mail: string | undefined) => (mail && nameByMail.get(mail.toLowerCase())) || mail || "";

        const isHrAdmin = canAccessSettings(setupRows, currentUserPrincipal);
        const isFinance = isFinanceUser(setupRows, currentUserPrincipal);

        const allReports = await loadAllExpenseReports(
            graphClient,
            { siteId, expenseReportsListId, expenseLinesListId, categoryListId, attachmentsListId },
            resolveName
        );

        const myReports = allReports
            .filter((r) => r.employeeCode?.toLowerCase() === currentUserPrincipal?.toLowerCase())
            .sort(byRequestDateDesc);

        const pendingApprovals = allReports
            .filter((r) => r.currentUser?.toLowerCase() === currentUserPrincipal?.toLowerCase() && r.status === "PENDING")
            .sort(byRequestDateDesc);

        return {
            jsonBody: {
                displayName: me.displayName,
                mail: currentUserPrincipal,
                myReports,
                pendingApprovals,
                isHrAdmin,
                isFinance,
            },
        };
    } catch (error) {
        context.error("GetMyExpenseContext failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_GetMyExpenseContext", {
    route: "expense/GetMyExpenseContext",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetMyExpenseContext,
});
