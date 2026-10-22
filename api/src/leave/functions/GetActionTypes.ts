import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getAllListItems } from "../../shared/lib/graphPaging";

type ActionType = {
    id: number;
    name: string;
    nameRo: string;
    code: string;
};

/**
 * Returns the leave/day-off type catalog (LR_ActionType list) so the client
 * can populate the "type" dropdown when creating a new leave request.
 */
export async function GetActionTypes(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected"]);
        const { siteId, actionTypeListId } = await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const items = await getAllListItems(graphClient, siteId, actionTypeListId);

        const actionTypes: ActionType[] = items
            .map((item: { fields: Record<string, string> }) => ({
                id: Number(item.fields.IDActionType),
                name: item.fields.ActionTypeName,
                nameRo: item.fields.ActionTypeNameRO,
                code: item.fields.ActionTypeCode,
            }))
            .sort((a: ActionType, b: ActionType) => a.id - b.id);

        return { jsonBody: { actionTypes } };
    } catch (error) {
        context.error("GetActionTypes failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_GetActionTypes", {
    route: "leave/GetActionTypes",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetActionTypes,
});
