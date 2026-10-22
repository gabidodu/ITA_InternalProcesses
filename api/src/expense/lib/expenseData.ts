import { Client } from "@microsoft/microsoft-graph-client";

/**
 * Graph paginates /items responses (default page size well under most lists'
 * real row counts) via @odata.nextLink — without following it, any list past
 * that threshold gets silently truncated to its first page. Since items come
 * back in ascending id/creation order, it's the NEWEST rows (the ones anyone
 * would actually care about) that go missing, with no error anywhere.
 * `buildFirstPage` issues the first request (with whatever .expand()/.select()
 * it needs); every subsequent page is fetched from the server-provided
 * nextLink as-is, which already carries those same query params.
 */
async function fetchAllPages(graphClient: Client, buildFirstPage: () => { get(): Promise<{ value: unknown[] }> }): Promise<unknown[]> {
    let response = (await buildFirstPage().get()) as { value: unknown[]; "@odata.nextLink"?: string };
    let items = response.value;
    while (response["@odata.nextLink"]) {
        response = (await graphClient.api(response["@odata.nextLink"]!).get()) as { value: unknown[]; "@odata.nextLink"?: string };
        items = items.concat(response.value);
    }
    return items;
}

export type ExpenseLine = {
    id: string;
    lineNumber: number;
    expenseDate: string;
    categoryId: number;
    categoryName: string;
    vendor: string;
    description: string;
    amount: number;
    currency: string;
    exchangeRate: number;
};

export type ExpenseAttachment = {
    id: string;
    name: string;
    webUrl: string;
    documentType: string;
    expenseLineId: number | null;
    uploadedBy: string;
    uploadedDate: string;
};

export type ExpenseReport = {
    id: string;
    subject: string;
    requestDate: string;
    totalAmount: number;
    currency: string;
    status: string;
    currentUser: string;
    currentUserName: string;
    employeeCode: string;
    employeeName: string;
    comments: string;
    lines: ExpenseLine[];
    attachments: ExpenseAttachment[];
};

export type ExpenseListIds = {
    siteId: string;
    expenseReportsListId: string;
    expenseLinesListId: string;
    categoryListId: string;
    attachmentsListId: string;
};

/**
 * Loads and enriches EVERY item in DE_ExpenseReports with its DE_ExpenseLines
 * and DE_Attachments (category names resolved from DE_ExpenseCategory).
 * Shared by GetMyExpenseContext (which then filters this into myReports/
 * pendingApprovals) and GetAllExpenseReports (the company-wide, Admin/
 * Finance-only overview, every status) — the per-item enrichment logic is
 * identical between the two, only the filtering differs.
 */
export async function loadAllExpenseReports(
    graphClient: Client,
    ids: ExpenseListIds,
    resolveName: (mail: string | undefined) => string
): Promise<ExpenseReport[]> {
    const { siteId, expenseReportsListId, expenseLinesListId, categoryListId, attachmentsListId } = ids;

    const categoriesItems = await fetchAllPages(graphClient, () => graphClient.api(`/sites/${siteId}/lists/${categoryListId}/items`).expand("fields"));
    const categoryNameById = new Map<number, string>();
    (categoriesItems as { fields: Record<string, string> }[]).forEach((item) => {
        categoryNameById.set(Number(item.fields.IDCategory), item.fields.CategoryName || item.fields.CategoryNameRO);
    });

    const linesItems = await fetchAllPages(graphClient, () => graphClient.api(`/sites/${siteId}/lists/${expenseLinesListId}/items`).expand("fields"));
    const linesByReportId = new Map<number, ExpenseLine[]>();
    (linesItems as { id: string; fields: Record<string, string> }[]).forEach((item) => {
        const f = item.fields;
        const reportId = Number(f.ExpenseReportId);
        const line: ExpenseLine = {
            id: item.id,
            lineNumber: Number(f.LineNumber),
            expenseDate: f.ExpenseDate,
            categoryId: Number(f.CategoryId),
            categoryName: categoryNameById.get(Number(f.CategoryId)) ?? "",
            vendor: f.Vendor ?? "",
            description: f.Description ?? "",
            amount: Number(f.Amount) || 0,
            currency: f.Currency ?? "",
            exchangeRate: Number(f.ExchangeRate) || 1,
        };
        const existing = linesByReportId.get(reportId) ?? [];
        existing.push(line);
        linesByReportId.set(reportId, existing);
    });
    linesByReportId.forEach((lines) => lines.sort((a, b) => a.lineNumber - b.lineNumber));

    const attachmentsItems = await fetchAllPages(graphClient, () =>
        graphClient.api(`/sites/${siteId}/lists/${attachmentsListId}/items`).expand("fields,driveItem")
    );
    const attachmentsByReportId = new Map<number, ExpenseAttachment[]>();
    (attachmentsItems as { id: string; fields: Record<string, string>; driveItem?: { id?: string; webUrl?: string; name?: string } }[]).forEach(
        (item) => {
        const f = item.fields;
        const reportId = Number(f.ExpenseReportId);
        const attachment: ExpenseAttachment = {
            // The SharePoint LIST item id (item.id) and the underlying drive
            // item id are different id spaces for the same file — GetExpenseAttachment
            // looks it up via /drives/{driveId}/items/{id}, which needs the drive item
            // id (item.driveItem.id), not the list item id. Passing the wrong one is
            // exactly what produced "The resource could not be found" on every "View".
            id: item.driveItem?.id ?? item.id,
            name: item.driveItem?.name ?? "",
            webUrl: item.driveItem?.webUrl ?? "",
            documentType: f.DocumentType ?? "Other",
            expenseLineId: f.ExpenseLineId ? Number(f.ExpenseLineId) : null,
            uploadedBy: f.UploadedBy ?? "",
            uploadedDate: f.UploadedDate ?? "",
        };
        const existing = attachmentsByReportId.get(reportId) ?? [];
        existing.push(attachment);
        attachmentsByReportId.set(reportId, existing);
    });

    const reportsItems = await fetchAllPages(graphClient, () => graphClient.api(`/sites/${siteId}/lists/${expenseReportsListId}/items`).expand("fields"));

    return (reportsItems as { id: string; fields: Record<string, string> }[]).map(
        (item): ExpenseReport => ({
            id: item.id,
            subject: item.fields.RequestSubject,
            requestDate: item.fields.RequestDate,
            totalAmount: Number(item.fields.TotalAmount) || 0,
            currency: item.fields.Currency ?? "",
            status: item.fields.Status,
            currentUser: item.fields.CurrentUser ?? "",
            currentUserName: resolveName(item.fields.CurrentUser),
            employeeCode: item.fields.EmployeeCode,
            employeeName: resolveName(item.fields.EmployeeCode),
            comments: item.fields.Comments ?? "",
            lines: linesByReportId.get(Number(item.id)) ?? [],
            attachments: attachmentsByReportId.get(Number(item.id)) ?? [],
        })
    );
}

export const byRequestDateDesc = (a: ExpenseReport, b: ExpenseReport) => new Date(b.requestDate).getTime() - new Date(a.requestDate).getTime();
