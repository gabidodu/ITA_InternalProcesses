import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getGmEmail, getSetupRows } from "../lib/setup";

/** Returns the GM constant configured in LR_Setup (Setting/GM), used to resolve the {GM} chain token. */
export async function GetGmConstant(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected"]);
        const { siteId, setupListId } = await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        return { jsonBody: { gmEmail: getGmEmail(setupRows) ?? null } };
    } catch (error) {
        context.error("GetGmConstant failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_GetGmConstant", {
    route: "leave/GetGmConstant",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetGmConstant,
});
