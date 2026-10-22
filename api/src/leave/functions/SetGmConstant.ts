import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { isHrAdmin, SetupRow } from "../lib/setup";
import { logAuditEntry } from "../../shared/lib/auditLog";
import { getAllListItems } from "../../shared/lib/graphPaging";

type SetGmConstantBody = {
    gmEmail: string;
};

function parseAndValidateBody(raw: unknown): SetGmConstantBody {
    const body = raw as Partial<SetGmConstantBody> | null;
    if (!body?.gmEmail || typeof body.gmEmail !== "string") {
        throw new Error("gmEmail is required.");
    }
    return { gmEmail: body.gmEmail.trim() };
}

/** Sets the GM constant in LR_Setup (Setting/GM), used to resolve the {GM} chain token. Restricted to HR admins. */
export async function SetGmConstant(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const body = parseAndValidateBody(await request.json());

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected"]);
        const { siteId, setupListId, auditLogListId } = await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;

        const setupItems = await getAllListItems(graphClient, siteId, setupListId);

        const setupRows: SetupRow[] = setupItems.map((item: { fields: Record<string, string> }) => ({
            settingKey: item.fields.SettingKey,
            settingValue: item.fields.SettingValue,
            category: item.fields.Category,
        }));

        if (!isHrAdmin(setupRows, currentUserPrincipal)) {
            return { status: 403, jsonBody: { error: "Only HR admins can set the GM constant." } };
        }

        const existingItem = setupItems.find(
            (item: { fields: Record<string, string> }) => item.fields.Category === "Setting" && item.fields.SettingKey === "GM"
        );

        if (existingItem) {
            await graphClient
                .api(`/sites/${siteId}/lists/${setupListId}/items/${existingItem.id}/fields`)
                .patch({ SettingValue: body.gmEmail });
        } else {
            await graphClient.api(`/sites/${siteId}/lists/${setupListId}/items`).post({
                fields: { Category: "Setting", SettingKey: "GM", SettingValue: body.gmEmail },
            });
        }

        try {
            await logAuditEntry(graphClient, siteId, auditLogListId, {
                actor: currentUserPrincipal,
                action: "GmConstantChanged",
                details: `GM constant set to: ${body.gmEmail}`,
            });
        } catch (auditError) {
            context.warn("logAuditEntry failed (non-blocking)", auditError);
        }

        return { jsonBody: { gmEmail: body.gmEmail } };
    } catch (error) {
        context.error("SetGmConstant failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_SetGmConstant", {
    route: "leave/SetGmConstant",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: SetGmConstant,
});
