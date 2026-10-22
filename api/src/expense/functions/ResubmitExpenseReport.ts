import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";
import { getApprovalChainFor, getSetupRows, resolveChainTokens } from "../lib/setup";
import { sendApprovalActionEmail } from "../lib/emailNotification";
import { getTeamsDeepLink } from "../../shared/lib/teamsLink";
import { logAuditEntry } from "../../shared/lib/auditLog";
import { buildApprovalCardScriptTag } from "../lib/actionableMessage";

const SUPPORTED_CURRENCIES = ["RON", "EUR", "USD"];
const HOME_CURRENCY = "RON";

type ExpenseLineInput = {
    // Present when this line already existed on the report (its DE_ExpenseLines
    // item id) and the requester kept it (edited or not) — it's then PATCHed in
    // place so its attachment (linked via ExpenseLineId) stays put. Absent for a
    // line added during this edit, which is created fresh, same as CreateExpenseReport.
    id?: string;
    expenseDate: string;
    categoryId: number;
    vendor?: string;
    description: string;
    amount: number;
    currency: string;
    exchangeRate: number;
};

type ResubmitExpenseReportBody = {
    id: string;
    subject: string;
    lines: ExpenseLineInput[];
    comments?: string;
};

function parseAndValidateBody(raw: unknown): ResubmitExpenseReportBody {
    const body = raw as Partial<ResubmitExpenseReportBody> | null;
    if (!body?.id || typeof body.id !== "string") {
        throw new Error("id is required.");
    }
    if (!body.subject || typeof body.subject !== "string") {
        throw new Error("subject is required.");
    }
    if (!Array.isArray(body.lines) || body.lines.length === 0) {
        throw new Error("At least one expense line is required.");
    }

    const lines = body.lines.map((line, index) => {
        if (!line || typeof line.categoryId !== "number") {
            throw new Error(`Line ${index + 1}: categoryId is required and must be a number.`);
        }
        if (line.id !== undefined && typeof line.id !== "string") {
            throw new Error(`Line ${index + 1}: id must be a string if provided.`);
        }
        if (!line.expenseDate || Number.isNaN(new Date(line.expenseDate).getTime())) {
            throw new Error(`Line ${index + 1}: expenseDate must be a valid date.`);
        }
        if (typeof line.amount !== "number" || !(line.amount > 0)) {
            throw new Error(`Line ${index + 1}: amount must be a positive number.`);
        }
        if (!line.currency || !SUPPORTED_CURRENCIES.includes(line.currency)) {
            throw new Error(`Line ${index + 1}: currency must be one of ${SUPPORTED_CURRENCIES.join(", ")}.`);
        }
        const isHomeCurrency = line.currency === HOME_CURRENCY;
        if (!isHomeCurrency && !(typeof line.exchangeRate === "number" && line.exchangeRate > 0)) {
            throw new Error(`Line ${index + 1}: exchangeRate is required and must be positive for a non-${HOME_CURRENCY} currency.`);
        }
        return {
            id: line.id,
            expenseDate: line.expenseDate,
            categoryId: line.categoryId,
            vendor: line.vendor?.trim() || "",
            description: line.description?.trim() || "",
            amount: line.amount,
            currency: line.currency,
            exchangeRate: isHomeCurrency ? 1 : line.exchangeRate,
        };
    });

    return {
        id: body.id,
        subject: body.subject,
        lines,
        comments: body.comments?.trim() || undefined,
    };
}

/**
 * Lets the original requester edit and resubmit a DECLINED expense report —
 * the counterpart to RespondToExpenseReport's decline branch. Re-enters the
 * approval chain from the top (recomputed fresh via getApprovalChainFor,
 * since it may have changed since the original submission), same as
 * CreateExpenseReport.
 *
 * Line reconciliation: a submitted line carrying an existing `id` is PATCHed
 * in place, so its attachment (linked via DE_Attachments.ExpenseLineId) stays
 * associated with it untouched. A submitted line with no `id` is a new line,
 * created fresh (its attachment, if any, is uploaded by the client afterward
 * via UploadExpenseAttachment, same two-step flow as CreateExpenseReport). An
 * existing line the requester removed (its id no longer present in the
 * submitted set) is deleted, along with any attachment that was linked to it
 * — otherwise that attachment would be orphaned (pointing at a line that no
 * longer exists).
 */
export async function ResubmitExpenseReport(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const tabBaseUrl = process.env.TAB_BASE_URL;
        const teamsAppId = process.env.TEAMS_APP_ID;
        if (!tabBaseUrl || !teamsAppId) {
            throw new Error("TAB_BASE_URL / TEAMS_APP_ID are not configured.");
        }
        const teamsDeepLink = getTeamsDeepLink(teamsAppId, tabBaseUrl, "expense");

        const body = parseAndValidateBody(await request.json());

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, ["Sites.Selected", "User.Read"]);
        const { siteId, expenseReportsListId, expenseLinesListId, expenseAttachmentsListId: attachmentsListId, setupListId, auditLogListId } = await getExpenseConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("displayName,mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;
        const currentUserName: string = me.displayName ?? currentUserPrincipal;

        const item = await graphClient.api(`/sites/${siteId}/lists/${expenseReportsListId}/items/${body.id}`).expand("fields").get();
        const fields = item.fields as Record<string, string>;

        if (fields.EmployeeCode?.toLowerCase() !== currentUserPrincipal?.toLowerCase()) {
            return { status: 403, jsonBody: { error: "Only the original requester can edit and resubmit this expense report." } };
        }
        if (fields.Status !== "DECLINED") {
            return { status: 400, jsonBody: { error: `This report cannot be resubmitted (current status: ${fields.Status}).` } };
        }

        const reportIdNum = Number(body.id);

        const existingLinesResponse = await graphClient.api(`/sites/${siteId}/lists/${expenseLinesListId}/items`).expand("fields").get();
        const existingLineIds = new Set<string>(
            existingLinesResponse.value
                .filter((li: { fields: Record<string, string> }) => Number(li.fields.ExpenseReportId) === reportIdNum)
                .map((li: { id: string }) => li.id)
        );

        for (const line of body.lines) {
            if (line.id && !existingLineIds.has(line.id)) {
                throw new Error("A line references an id that doesn't belong to this report.");
            }
        }

        const keptLineIds = new Set(body.lines.map((l) => l.id).filter((id): id is string => Boolean(id)));
        const removedLineIds = [...existingLineIds].filter((id) => !keptLineIds.has(id));

        if (removedLineIds.length > 0) {
            const attachmentsResponse = await graphClient
                .api(`/sites/${siteId}/lists/${attachmentsListId}/items`)
                .expand("fields,driveItem")
                .get();
            const orphanedAttachments = attachmentsResponse.value.filter(
                (a: { fields: Record<string, string> }) =>
                    Number(a.fields.ExpenseReportId) === reportIdNum && removedLineIds.includes(String(Number(a.fields.ExpenseLineId)))
            );
            if (orphanedAttachments.length > 0) {
                const drive = await graphClient.api(`/sites/${siteId}/lists/${attachmentsListId}/drive`).get();
                await Promise.all(
                    orphanedAttachments.map((a: { driveItem?: { id?: string }; id: string }) =>
                        graphClient.api(`/drives/${drive.id}/items/${a.driveItem?.id ?? a.id}`).delete()
                    )
                );
            }
            await Promise.all(removedLineIds.map((lineId) => graphClient.api(`/sites/${siteId}/lists/${expenseLinesListId}/items/${lineId}`).delete()));
        }

        const finalLines: { id: string; lineNumber: number }[] = [];
        for (let i = 0; i < body.lines.length; i++) {
            const line = body.lines[i];
            const lineFields = {
                ExpenseReportId: reportIdNum,
                LineNumber: i + 1,
                ExpenseDate: new Date(line.expenseDate).toISOString(),
                CategoryId: line.categoryId,
                Vendor: line.vendor,
                Description: line.description,
                Amount: line.amount,
                Currency: line.currency,
                ExchangeRate: line.exchangeRate,
            };
            if (line.id) {
                await graphClient.api(`/sites/${siteId}/lists/${expenseLinesListId}/items/${line.id}/fields`).patch(lineFields);
                finalLines.push({ id: line.id, lineNumber: i + 1 });
            } else {
                const createdLine = await graphClient.api(`/sites/${siteId}/lists/${expenseLinesListId}/items`).post({ fields: lineFields });
                finalLines.push({ id: createdLine.id, lineNumber: i + 1 });
            }
        }

        const totalAmount = body.lines.reduce((sum, line) => sum + line.amount * line.exchangeRate, 0);

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const employeeChain = getApprovalChainFor(setupRows, currentUserPrincipal);
        const fullChain = await resolveChainTokens(graphClient, setupRows, employeeChain);
        if (fullChain.length === 0) {
            throw new Error("No approval chain is configured. Please contact HR.");
        }
        const [firstApprover, ...remainingChain] = fullChain;

        await graphClient.api(`/sites/${siteId}/lists/${expenseReportsListId}/items/${body.id}/fields`).patch({
            RequestSubject: body.subject,
            Comments: body.comments ?? "",
            TotalAmount: totalAmount,
            Status: "PENDING",
            CurrentUser: firstApprover,
            ApprovalFlow: remainingChain.join(";"),
        });

        try {
            await logAuditEntry(graphClient, siteId, auditLogListId, {
                actor: currentUserPrincipal,
                action: "Resubmitted",
                employeeCode: currentUserPrincipal,
                requestId: body.id,
                details: "",
            });
        } catch (auditError) {
            context.warn("logAuditEntry failed (non-blocking)", auditError);
        }

        try {
            const cardScriptTag = buildApprovalCardScriptTag({
                requesterName: currentUserName,
                requestSubject: body.subject,
                totalAmount,
                currency: HOME_CURRENCY,
                requestId: body.id,
                tabBaseUrl,
                teamsDeepLink,
            });
            await sendApprovalActionEmail(tenantId, 
                firstApprover,
                `Expense Report - ${currentUserName}`,
                `${currentUserName} needs your approval for a resubmitted expense report ("${body.subject}", ${totalAmount.toFixed(2)} ${HOME_CURRENCY}).`,
                teamsDeepLink,
                cardScriptTag,
                [
                    { label: "Subject", value: body.subject },
                    { label: "Total", value: `${totalAmount.toFixed(2)} ${HOME_CURRENCY}` },
                    { label: "Requested by", value: currentUserName },
                ]
            );
        } catch (emailError) {
            context.warn("sendApprovalActionEmail failed (non-blocking)", emailError);
        }

        return {
            jsonBody: {
                id: body.id,
                subject: body.subject,
                totalAmount,
                currency: HOME_CURRENCY,
                status: "PENDING",
                currentUser: firstApprover,
                lines: finalLines,
            },
        };
    } catch (error) {
        context.error("ResubmitExpenseReport failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_ResubmitExpenseReport", {
    route: "expense/ResubmitExpenseReport",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: ResubmitExpenseReport,
});
