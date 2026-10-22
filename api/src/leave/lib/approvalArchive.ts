import { Client } from "@microsoft/microsoft-graph-client";

const INVALID_PATH_CHARS = /["*:<>?/\\|]/g;

function sanitizePathSegment(segment: string): string {
    const cleaned = segment.replace(INVALID_PATH_CHARS, "_").trim().replace(/[. ]+$/, "");
    return cleaned || "Unknown";
}

/**
 * Archives the approved-leave-request PDF into LR_Approvals, under a
 * per-employee subfolder — a permanent, browsable copy alongside the email
 * attachment RespondToLeaveRequest(Action).ts already sends, so an approved
 * request survives even if that email gets deleted. The subfolder is created
 * on first use: Graph's simple-upload PUT isn't documented to reliably
 * auto-create missing parent folders via path addressing, so it's created
 * explicitly first, treating a 409 ("already exists") as success rather than
 * checking existence beforehand — cheaper than a separate lookup call, and
 * still correct under concurrent first-time uploads for the same employee.
 * Best-effort: callers wrap this in try/catch, same as the PDF email send it
 * runs alongside — a SharePoint hiccup here must never block the approval
 * itself.
 */
export type ApprovedLeaveRequestArchiveEntry = {
    employeeFullName: string;
    employeeCode: string;
    requestId: string;
    fileName: string;
    pdfBuffer: Buffer;
};

export async function saveApprovedLeaveRequestPdf(
    graphClient: Client,
    siteId: string,
    approvalsLibraryId: string,
    entry: ApprovedLeaveRequestArchiveEntry
): Promise<void> {
    const drive = await graphClient.api(`/sites/${siteId}/lists/${approvalsLibraryId}/drive`).get();
    const folderName = sanitizePathSegment(entry.employeeFullName);

    try {
        await graphClient.api(`/drives/${drive.id}/root/children`).post({
            name: folderName,
            folder: {},
            "@microsoft.graph.conflictBehavior": "fail",
        });
    } catch (error) {
        if ((error as { statusCode?: number })?.statusCode !== 409) {
            throw error;
        }
    }

    const safeFileName = sanitizePathSegment(entry.fileName);
    const uploaded = await graphClient
        .api(`/drives/${drive.id}/root:/${encodeURIComponent(folderName)}/${encodeURIComponent(safeFileName)}:/content`)
        .header("Content-Type", "application/pdf")
        .put(entry.pdfBuffer);

    await graphClient.api(`/drives/${drive.id}/items/${uploaded.id}/listItem/fields`).patch({
        EmployeeCode: entry.employeeCode,
        RequestId: entry.requestId,
        ApprovedDate: new Date().toISOString(),
    });
}
