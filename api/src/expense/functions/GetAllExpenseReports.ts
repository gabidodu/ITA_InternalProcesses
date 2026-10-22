import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";
import { canAccessSettings, getSetupRows, isFinanceUser } from "../lib/setup";
import { buildDisplayNameMap, getEmployeeDirectory } from "../../shared/lib/employees";
import { byRequestDateDesc, loadAllExpenseReports } from "../lib/expenseData";

/**
 * Company-wide, read-only overview of every expense report regardless of
 * status (PENDING/APPROVED/DECLINED/CANCELLED), optionally narrowed by year
 * (RequestDate), employeeCode, and/or month (1-12). Restricted to Admin or
 * Finance users, enforced server-side — this exposes every employee's
 * expense history, not just the caller's own. Finance's own visibility into
 * every report (including each line's receipt/invoice) goes through this
 * endpoint, same as Admin's. Results are sorted newest-first and paginated
 * (`offset`/`limit`, default page size 50) — the response's `total` and
 * `hasMore` describe the full filtered set, `reports` is just the current page.
 */
export async function GetAllExpenseReports(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const yearParam = request.query.get("year");
        const year = yearParam ? Number(yearParam) : new Date().getFullYear();
        if (!Number.isInteger(year)) {
            throw new Error("year must be a whole number.");
        }
        const employeeCode = request.query.get("employeeCode")?.trim() || undefined;
        const monthParam = request.query.get("month");
        const month = monthParam ? Number(monthParam) : undefined;
        if (month !== undefined && (!Number.isInteger(month) || month < 1 || month > 12)) {
            throw new Error("month must be a whole number from 1 to 12.");
        }
        const offsetParam = request.query.get("offset");
        const offset = offsetParam ? Number(offsetParam) : 0;
        if (!Number.isInteger(offset) || offset < 0) {
            throw new Error("offset must be a non-negative whole number.");
        }
        const limitParam = request.query.get("limit");
        const limit = limitParam ? Number(limitParam) : 50;
        if (!Number.isInteger(limit) || limit < 1 || limit > 500) {
            throw new Error("limit must be a whole number from 1 to 500.");
        }

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected", "GroupMember.Read.All", "User.ReadBasic.All"]);
        const { siteId, expenseReportsListId, expenseLinesListId, expenseCategoryListId: categoryListId, expenseAttachmentsListId: attachmentsListId, setupListId } = await getExpenseConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        if (!canAccessSettings(setupRows, currentUserPrincipal) && !isFinanceUser(setupRows, currentUserPrincipal)) {
            return { status: 403, jsonBody: { error: "Only Admin or Finance users can view the company-wide report." } };
        }

        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const nameByMail = buildDisplayNameMap(employees);
        const resolveName = (mail: string | undefined) => (mail && nameByMail.get(mail.toLowerCase())) || mail || "";

        const allReports = await loadAllExpenseReports(
            graphClient,
            { siteId, expenseReportsListId, expenseLinesListId, categoryListId, attachmentsListId },
            resolveName
        );

        const matchingReports = allReports
            .filter((r) => new Date(r.requestDate).getFullYear() === year)
            .filter((r) => !employeeCode || r.employeeCode?.toLowerCase() === employeeCode.toLowerCase())
            .filter((r) => month === undefined || new Date(r.requestDate).getMonth() + 1 === month)
            .sort(byRequestDateDesc);

        const total = matchingReports.length;
        const totalAmount = matchingReports.reduce((sum, r) => sum + r.totalAmount, 0);
        const reports = matchingReports.slice(offset, offset + limit);

        return { jsonBody: { year, reports, total, totalAmount, hasMore: offset + limit < total } };
    } catch (error) {
        context.error("GetAllExpenseReports failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_GetAllExpenseReports", {
    route: "expense/GetAllExpenseReports",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetAllExpenseReports,
});
