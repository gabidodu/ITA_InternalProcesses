import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";

/**
 * Streams back the raw file content of one DE_Attachments item (?id=), so
 * receipts/invoices are viewable/downloadable from the Teams tab regardless
 * of whether the viewer happens to have their own direct SharePoint site
 * permissions on AVI_Site — this app otherwise never asks an end user to
 * browse SharePoint directly, only to call Graph via the OBO-exchanged
 * token, and a plain `webUrl` link would silently break that for anyone
 * without separate SharePoint site access. The client fetches this with the
 * X-Teams-Sso-Token header (same pattern as every other call) and turns the
 * response into a blob URL — a plain `<a href>` can't carry that header.
 */
export async function GetExpenseAttachment(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const id = request.query.get("id");
        if (!id) {
            throw new Error("id is required.");
        }

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected"]);
        const { siteId, expenseAttachmentsListId: attachmentsListId } = await getExpenseConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const drive = await graphClient.api(`/sites/${siteId}/lists/${attachmentsListId}/drive`).get();

        const item: { name: string; file?: { mimeType?: string } } = await graphClient
            .api(`/drives/${drive.id}/items/${id}`)
            .select("name,file")
            .get();

        // The Graph JS client's .get() assumes a JSON response, which corrupts
        // binary content — fetch the raw bytes directly with the same token instead.
        const contentResponse = await fetch(`https://graph.microsoft.com/v1.0/drives/${drive.id}/items/${id}/content`, {
            headers: { Authorization: `Bearer ${graphToken}` },
        });
        if (!contentResponse.ok) {
            throw new Error(`Graph returned ${contentResponse.status} fetching the file content.`);
        }
        const buffer = Buffer.from(await contentResponse.arrayBuffer());

        return {
            status: 200,
            body: buffer,
            headers: {
                "Content-Type": item.file?.mimeType || "application/octet-stream",
                "Content-Disposition": `inline; filename="${(item.name || "attachment").replace(/"/g, "")}"`,
            },
        };
    } catch (error) {
        context.error("GetExpenseAttachment failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_GetExpenseAttachment", {
    route: "expense/GetExpenseAttachment",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetExpenseAttachment,
});
