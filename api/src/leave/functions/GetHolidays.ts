import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getRelevantYears } from "../../shared/lib/dates";
import { getAllListItems } from "../../shared/lib/graphPaging";

type Holiday = {
    id: string;
    name: string;
    date: string;
    year: number;
    type: string;
};

/**
 * Returns the standard company holiday calendar (LR_HolidaySetup), newest
 * first, limited to last year / current year / next year — the full history
 * (2018+) is still in SharePoint, just not relevant to show/manage here.
 */
export async function GetHolidays(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected"]);
        const { siteId, holidaySetupListId } = await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const items = await getAllListItems(graphClient, siteId, holidaySetupListId);

        const relevantYears = getRelevantYears();

        const holidays: Holiday[] = items
            .map((item: { id: string; fields: Record<string, string> }) => ({
                id: item.id,
                name: item.fields.HolidayName,
                date: item.fields.HolidayDate,
                year: Number(item.fields.HolidayYear),
                type: item.fields.HolidayType,
            }))
            .filter((h: Holiday) => relevantYears.has(h.year))
            .sort((a: Holiday, b: Holiday) => new Date(b.date).getTime() - new Date(a.date).getTime());

        return { jsonBody: { holidays } };
    } catch (error) {
        context.error("GetHolidays failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_GetHolidays", {
    route: "leave/GetHolidays",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetHolidays,
});
