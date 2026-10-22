import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows } from "../lib/setup";
import { getEmployeeDirectory } from "../../shared/lib/employees";

/**
 * Returns the members of the "all employees" Entra ID group (its id is
 * configured in LR_Setup, category AzureADGroupID / allEmployeesGroupID), so
 * the client can populate people-pickers for stand-in selection and for the
 * HR approval-chain editor.
 */
export async function GetPickableEmployees(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, [
            "Sites.Selected",
            "GroupMember.Read.All",
            "User.ReadBasic.All",
        ]);
        const { siteId, setupListId } = await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const employees = (await getEmployeeDirectory(graphClient, setupRows)).sort((a, b) =>
            a.displayName.localeCompare(b.displayName)
        );

        return { jsonBody: { employees } };
    } catch (error) {
        context.error("GetPickableEmployees failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_GetPickableEmployees", {
    route: "leave/GetPickableEmployees",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetPickableEmployees,
});
