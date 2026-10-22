import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows, canAccessSettings } from "../lib/setup";

type AddHolidayBody = {
    name: string;
    date: string;
};

function parseAndValidateBody(raw: unknown): AddHolidayBody {
    const body = raw as Partial<AddHolidayBody> | null;
    if (!body?.name || typeof body.name !== "string") {
        throw new Error("name is required.");
    }
    if (!body.date || Number.isNaN(new Date(body.date).getTime())) {
        throw new Error("date is required and must be valid.");
    }
    return { name: body.name, date: body.date };
}

/** Adds a new standard holiday to LR_HolidaySetup. Restricted to HR admins. */
export async function AddHoliday(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const body = parseAndValidateBody(await request.json());

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected"]);
        const { siteId, holidaySetupListId, setupListId } = await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        if (!canAccessSettings(setupRows, currentUserPrincipal)) {
            return { status: 403, jsonBody: { error: "Only HR, Admin, or GM can manage the holiday calendar." } };
        }

        const date = new Date(body.date);
        const created = await graphClient.api(`/sites/${siteId}/lists/${holidaySetupListId}/items`).post({
            fields: {
                HolidayName: body.name,
                HolidayDate: date.toISOString(),
                HolidayType: "Legal",
                HolidayYear: date.getFullYear(),
            },
        });

        return {
            status: 201,
            jsonBody: { id: created.id, name: body.name, date: date.toISOString(), year: date.getFullYear(), type: "Legal" },
        };
    } catch (error) {
        context.error("AddHoliday failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_AddHoliday", {
    route: "leave/AddHoliday",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: AddHoliday,
});
