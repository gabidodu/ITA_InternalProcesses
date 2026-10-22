import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";

// Graph's "simple upload" (a single PUT with the whole file body) is capped at
// 4 MiB — anything larger needs a resumable upload session instead, which
// this endpoint doesn't implement yet (no real need has come up: receipt/
// invoice photos and scans comfortably fit under this limit).
const MAX_UPLOAD_BYTES = 4 * 1024 * 1024;

type UploadExpenseAttachmentBody = {
    expenseReportId: number;
    expenseLineId?: number;
    documentType: "Receipt" | "Invoice" | "Other";
    fileName: string;
    contentType: string;
    fileBase64: string;
};

function parseAndValidateBody(raw: unknown): UploadExpenseAttachmentBody {
    const body = raw as Partial<UploadExpenseAttachmentBody> | null;
    if (!body || typeof body.expenseReportId !== "number") {
        throw new Error("expenseReportId is required and must be a number.");
    }
    if (body.documentType !== "Receipt" && body.documentType !== "Invoice" && body.documentType !== "Other") {
        throw new Error('documentType must be "Receipt", "Invoice", or "Other".');
    }
    if (!body.fileName || typeof body.fileName !== "string") {
        throw new Error("fileName is required.");
    }
    if (!body.fileBase64 || typeof body.fileBase64 !== "string") {
        throw new Error("fileBase64 is required.");
    }
    if (typeof body.expenseLineId !== "number" && body.expenseLineId !== undefined) {
        throw new Error("expenseLineId must be a number if provided.");
    }
    return {
        expenseReportId: body.expenseReportId,
        expenseLineId: body.expenseLineId,
        documentType: body.documentType,
        fileName: body.fileName,
        contentType: body.contentType || "application/octet-stream",
        fileBase64: body.fileBase64,
    };
}

// Strips characters SharePoint/OneDrive file names can't contain, and prefixes
// with the report id + a timestamp so two people uploading "receipt.jpg" the
// same second never collide.
function buildSafeFileName(expenseReportId: number, originalName: string): string {
    const sanitized = originalName.replace(/[\\/:*?"<>|]/g, "_");
    return `${expenseReportId}_${Date.now()}_${sanitized}`;
}

/**
 * Uploads a receipt/invoice file into the DE_Attachments document library and
 * sets its metadata columns (ExpenseReportId, ExpenseLineId, DocumentType,
 * UploadedBy, UploadedDate) so it can be found again by report/line. This is
 * net-new — no prior function in the sibling LeaveRequestForTeams app does a
 * SharePoint file/drive upload (it only ever emails a generated PDF).
 *
 * The client sends the file as base64 JSON rather than multipart/form-data —
 * simpler to validate and forward through Azure Static Web Apps' managed API
 * proxy, at the cost of ~33% size overhead (acceptable given the 4 MiB cap).
 */
export async function UploadExpenseAttachment(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const body = parseAndValidateBody(await request.json());

        const fileBuffer = Buffer.from(body.fileBase64, "base64");
        if (fileBuffer.byteLength === 0) {
            throw new Error("The uploaded file is empty.");
        }
        if (fileBuffer.byteLength > MAX_UPLOAD_BYTES) {
            throw new Error(
                `File is too large (${(fileBuffer.byteLength / (1024 * 1024)).toFixed(1)} MB). ` +
                    "Please compress or rescan at a lower resolution — the limit is 4 MB per file."
            );
        }

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected"]);
        const { siteId, expenseAttachmentsListId: attachmentsListId } = await getExpenseConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;

        // DE_Attachments is a document library, so its content lives in its own
        // drive, not the site's default "Documents" drive — resolve it from the
        // list rather than assuming /sites/{id}/drive.
        const drive = await graphClient.api(`/sites/${siteId}/lists/${attachmentsListId}/drive`).get();
        const safeFileName = buildSafeFileName(body.expenseReportId, body.fileName);

        const uploaded = await graphClient
            .api(`/drives/${drive.id}/root:/${encodeURIComponent(safeFileName)}:/content`)
            .header("Content-Type", body.contentType)
            .put(fileBuffer);

        // Omit ExpenseLineId entirely when absent — Graph's field PATCH rejects
        // an explicit null for a Number column, so "not attached to a specific
        // line" is represented by leaving the column unset, not null.
        const metadataFields: Record<string, string | number> = {
            ExpenseReportId: body.expenseReportId,
            DocumentType: body.documentType,
            UploadedBy: currentUserPrincipal,
            UploadedDate: new Date().toISOString(),
        };
        if (body.expenseLineId !== undefined) {
            metadataFields.ExpenseLineId = body.expenseLineId;
        }

        await graphClient.api(`/drives/${drive.id}/items/${uploaded.id}/listItem/fields`).patch(metadataFields);

        return {
            status: 201,
            jsonBody: {
                id: uploaded.id,
                name: uploaded.name,
                webUrl: uploaded.webUrl,
                size: uploaded.size,
                documentType: body.documentType,
                expenseLineId: body.expenseLineId ?? null,
            },
        };
    } catch (error) {
        context.error("UploadExpenseAttachment failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_UploadExpenseAttachment", {
    route: "expense/UploadExpenseAttachment",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: UploadExpenseAttachment,
});
