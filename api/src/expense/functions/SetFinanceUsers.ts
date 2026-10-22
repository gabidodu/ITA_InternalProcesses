import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";
import { canAccessSettings, titleForSetupRow, SetupRow } from "../lib/setup";
import { logAuditEntry } from "../../shared/lib/auditLog";
import { buildDisplayNameMap, getEmployeeDirectory } from "../../shared/lib/employees";

type SetFinanceUsersBody = {
    emails: string[];
};

function parseAndValidateBody(raw: unknown): SetFinanceUsersBody {
    const body = raw as Partial<SetFinanceUsersBody> | null;
    if (!body || !Array.isArray(body.emails) || body.emails.some((e) => typeof e !== "string" || !e.trim())) {
        throw new Error("emails must be an array of user emails.");
    }
    return { emails: Array.from(new Set(body.emails.map((e) => e.trim().toLowerCase()))) };
}

/**
 * Replaces the full set of Finance-role users (DE_Setup Category=Finance) —
 * who get full visibility into every expense report (see
 * GetAllExpenseReports.ts), a role check rather than a chain. Restricted to
 * HR/Admin, enforced server-side (canAccessSettings), same as
 * SetApprovalFlow. Unlike the approval chain, an empty Finance list is
 * allowed — it just means no one has that broader visibility until an Admin
 * adds someone back, not that anyone is blocked from doing their own work.
 */
export async function SetFinanceUsers(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const body = parseAndValidateBody(await request.json());

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected", "GroupMember.Read.All", "User.ReadBasic.All"]);
        const { siteId, setupListId, auditLogListId } = await getExpenseConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("displayName,mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;
        const currentUserName: string = me.displayName ?? currentUserPrincipal;

        const itemsResponse = await graphClient.api(`/sites/${siteId}/lists/${setupListId}/items`).expand("fields").get();

        const setupRows: SetupRow[] = itemsResponse.value.map((item: { fields: Record<string, string> }) => ({
            settingKey: item.fields.SettingKey,
            settingValue: item.fields.SettingValue,
            category: item.fields.Category,
        }));

        if (!canAccessSettings(setupRows, currentUserPrincipal)) {
            return { status: 403, jsonBody: { error: "Only HR/Admin can edit Finance access." } };
        }

        const financeItems = itemsResponse.value.filter((item: { fields: Record<string, string> }) => item.fields.Category === "Finance");
        const existingEmails = new Set(financeItems.map((item: { fields: Record<string, string> }) => item.fields.SettingKey?.toLowerCase()));
        const newEmails = new Set(body.emails);

        const toRemove = financeItems.filter((item: { fields: Record<string, string> }) => !newEmails.has(item.fields.SettingKey?.toLowerCase()));
        const toAdd = body.emails.filter((email) => !existingEmails.has(email));

        for (const item of toRemove) {
            await graphClient.api(`/sites/${siteId}/lists/${setupListId}/items/${item.id}`).delete();
        }
        for (const email of toAdd) {
            await graphClient.api(`/sites/${siteId}/lists/${setupListId}/items`).post({
                fields: { Title: titleForSetupRow("Finance", email), Category: "Finance", SettingKey: email, SettingValue: "Finance" },
            });
        }

        try {
            const employees = await getEmployeeDirectory(graphClient, setupRows);
            const nameByMail = buildDisplayNameMap(employees);
            const resolveName = (mail: string) => nameByMail.get(mail.toLowerCase()) ?? mail;
            await logAuditEntry(graphClient, siteId, auditLogListId, {
                actor: currentUserPrincipal,
                action: "FinanceAccessChanged",
                details: `${currentUserName} set Finance access to: ${body.emails.map(resolveName).join(", ") || "(none)"}`,
            });
        } catch (auditError) {
            context.warn("logAuditEntry failed (non-blocking)", auditError);
        }

        return { jsonBody: { emails: body.emails } };
    } catch (error) {
        context.error("SetFinanceUsers failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_SetFinanceUsers", {
    route: "expense/SetFinanceUsers",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: SetFinanceUsers,
});
