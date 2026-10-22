import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows, canAccessSettings } from "../lib/setup";
import { getAllListItems } from "../../shared/lib/graphPaging";

type DeleteHolidayBody = {
    id: string;
};

function parseAndValidateBody(raw: unknown): DeleteHolidayBody {
    const body = raw as Partial<DeleteHolidayBody> | null;
    if (!body?.id || typeof body.id !== "string") {
        throw new Error("id is required.");
    }
    return { id: body.id };
}

/**
 * Deletes a holiday from LR_HolidaySetup — but refuses if that date falls
 * within any existing leave request's start/end range, since it may have
 * already been factored into that request's taken-days count. Restricted to
 * HR admins.
 */
export async function DeleteHoliday(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const body = parseAndValidateBody(await request.json());

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected"]);
        const { siteId, holidaySetupListId, leaveRequestsListId, setupListId } = await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        if (!canAccessSettings(setupRows, currentUserPrincipal)) {
            return { status: 403, jsonBody: { error: "Only HR, Admin, or GM can manage the holiday calendar." } };
        }

        const holidayItem = await graphClient
            .api(`/sites/${siteId}/lists/${holidaySetupListId}/items/${body.id}`)
            .expand("fields")
            .get();
        const holidayDate = new Date(holidayItem.fields.HolidayDate).getTime();

        const leaveRequestItems = await getAllListItems(graphClient, siteId, leaveRequestsListId);

        const isInUse = leaveRequestItems.some((item: { fields: Record<string, string> }) => {
            if (!item.fields.StartDate || !item.fields.EndDate) {
                return false;
            }
            const start = new Date(item.fields.StartDate).getTime();
            const end = new Date(item.fields.EndDate).getTime();
            return holidayDate >= start && holidayDate <= end;
        });

        if (isInUse) {
            return {
                status: 409,
                jsonBody: {
                    error: "You can't delete this holiday: the date has already been used in an existing leave request.",
                },
            };
        }

        await graphClient.api(`/sites/${siteId}/lists/${holidaySetupListId}/items/${body.id}`).delete();

        return { jsonBody: { id: body.id, deleted: true } };
    } catch (error) {
        context.error("DeleteHoliday failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_DeleteHoliday", {
    route: "leave/DeleteHoliday",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: DeleteHoliday,
});
