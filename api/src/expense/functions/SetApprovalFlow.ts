import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";
import { getApprovalChainFor, canAccessSettings, labelForChainEntry, titleForSetupRow, SetupRow } from "../lib/setup";
import { logAuditEntry } from "../../shared/lib/auditLog";
import { buildDisplayNameMap, getEmployeeDirectory } from "../../shared/lib/employees";

type SetApprovalFlowBody = {
    chain: string[];
    employeeCode?: string;
};

function parseAndValidateBody(raw: unknown): SetApprovalFlowBody {
    const body = raw as Partial<SetApprovalFlowBody> | null;
    if (!body || !Array.isArray(body.chain) || body.chain.some((e) => typeof e !== "string" || !e.trim())) {
        throw new Error("chain must be an array of approver emails.");
    }
    if (!body.employeeCode && body.chain.length === 0) {
        throw new Error("The default chain cannot be emptied — it would block everyone from submitting expense reports.");
    }
    return { chain: body.chain.map((e) => e.trim()), employeeCode: body.employeeCode?.trim() || undefined };
}

/**
 * Updates an approval chain in DE_Setup: the global standard one
 * (Setting/approvalFlow) if no employeeCode is given, or a specific
 * employee's override (ApprovalFlow/<employeeCode>) otherwise. Sending an
 * empty chain for a specific employee deletes their override, reverting them
 * to the standard chain. Restricted to HR admins, enforced server-side.
 */
export async function SetApprovalFlow(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
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
            return { status: 403, jsonBody: { error: "Only HR/Admin can edit approval chains." } };
        }

        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const nameByMail = buildDisplayNameMap(employees);
        const resolveName = (mail: string | undefined) => (mail && nameByMail.get(mail.toLowerCase())) || mail || "";
        const chainLabel = (chain: string[]) => chain.map((entry) => labelForChainEntry(entry, nameByMail)).join(" -> ");

        const targetCategory = body.employeeCode ? "ApprovalFlow" : "Setting";
        const targetKey = body.employeeCode ?? "approvalFlow";
        const targetItem = itemsResponse.value.find(
            (item: { fields: Record<string, string> }) =>
                item.fields.Category === targetCategory && item.fields.SettingKey?.toLowerCase() === targetKey.toLowerCase()
        );

        let responseChain = body.chain;
        let responseIsOverride = true;

        if (body.employeeCode && body.chain.length === 0) {
            if (targetItem) {
                await graphClient.api(`/sites/${siteId}/lists/${setupListId}/items/${targetItem.id}`).delete();
            }
            const remainingRows = setupRows.filter(
                (r) => !(r.category === targetCategory && r.settingKey?.toLowerCase() === targetKey.toLowerCase())
            );
            responseChain = getApprovalChainFor(remainingRows, body.employeeCode);
            responseIsOverride = false;
        } else {
            const newValue = body.chain.join(";");
            const title = titleForSetupRow(targetCategory, targetKey);
            if (targetItem) {
                await graphClient.api(`/sites/${siteId}/lists/${setupListId}/items/${targetItem.id}/fields`).patch({ Title: title, SettingValue: newValue });
            } else {
                await graphClient.api(`/sites/${siteId}/lists/${setupListId}/items`).post({
                    fields: { Title: title, Category: targetCategory, SettingKey: targetKey, SettingValue: newValue },
                });
            }
        }

        try {
            const targetLabel = body.employeeCode ? resolveName(body.employeeCode) : "the standard chain";
            await logAuditEntry(graphClient, siteId, auditLogListId, {
                actor: currentUserPrincipal,
                action: "ApprovalChainChanged",
                employeeCode: body.employeeCode ?? "(default)",
                details: body.employeeCode
                    ? body.chain.length === 0
                        ? `${currentUserName} removed ${targetLabel}'s override, reverted to standard chain.`
                        : `${currentUserName} set ${targetLabel}'s override to: ${chainLabel(body.chain)}`
                    : `${currentUserName} set the standard chain to: ${chainLabel(body.chain)}`,
            });
        } catch (auditError) {
            context.warn("logAuditEntry failed (non-blocking)", auditError);
        }

        return { jsonBody: { chain: responseChain, isOverride: responseIsOverride } };
    } catch (error) {
        context.error("SetApprovalFlow failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_SetApprovalFlow", {
    route: "expense/SetApprovalFlow",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: SetApprovalFlow,
});
