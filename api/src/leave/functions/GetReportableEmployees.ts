import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows, canAccessSettings } from "../lib/setup";
import { getEmployeeDirectory, getMyDirectReports } from "../../shared/lib/employees";

/**
 * Returns the employees the signed-in user is allowed to view a leave report
 * for: the full AD-group directory for HR/Admin/GM (canAccessSettings), or
 * just their own direct reports (per Entra ID) for a plain manager. Powers
 * the Reports tab's employee picker.
 */
export async function GetReportableEmployees(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, [
            "Sites.Selected",
            "GroupMember.Read.All",
            "User.ReadBasic.All",
            "User.Read.All",
        ]);
        const { siteId, setupListId } = await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);

        const employees = canAccessSettings(setupRows, currentUserPrincipal)
            ? (await getEmployeeDirectory(graphClient, setupRows)).sort((a, b) => a.displayName.localeCompare(b.displayName))
            : (await getMyDirectReports(graphClient)).sort((a, b) => a.displayName.localeCompare(b.displayName));

        return { jsonBody: { employees } };
    } catch (error) {
        context.error("GetReportableEmployees failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_GetReportableEmployees", {
    route: "leave/GetReportableEmployees",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetReportableEmployees,
});
