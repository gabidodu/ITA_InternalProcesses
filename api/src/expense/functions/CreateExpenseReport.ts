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
    expenseDate: string;
    categoryId: number;
    vendor?: string;
    description: string;
    amount: number;
    currency: string;
    exchangeRate: number;
};

type CreateExpenseReportBody = {
    subject: string;
    lines: ExpenseLineInput[];
    comments?: string;
};

function parseAndValidateBody(raw: unknown): CreateExpenseReportBody {
    const body = raw as Partial<CreateExpenseReportBody> | null;
    if (!body?.subject || typeof body.subject !== "string") {
        throw new Error("subject is required.");
    }
    if (!Array.isArray(body.lines) || body.lines.length === 0) {
        throw new Error("At least one expense line is required.");
    }

    const lines = body.lines.map((line, index) => {
        if (!line || typeof line.categoryId !== "number") {
            throw new Error(`Line ${index + 1}: categoryId is required and must be a number.`);
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
        // The report's own total is always kept in HOME_CURRENCY — a non-home
        // currency line needs its own exchange rate to convert into it. A
        // home-currency line's rate is fixed at 1 regardless of what's sent.
        const isHomeCurrency = line.currency === HOME_CURRENCY;
        if (!isHomeCurrency && !(typeof line.exchangeRate === "number" && line.exchangeRate > 0)) {
            throw new Error(`Line ${index + 1}: exchangeRate is required and must be positive for a non-${HOME_CURRENCY} currency.`);
        }
        return {
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
        subject: body.subject,
        lines,
        comments: body.comments?.trim() || undefined,
    };
}

/**
 * Creates a new expense report (DE_ExpenseReports header + one DE_ExpenseLines
 * row per line item). The approval chain mirrors the sibling
 * LeaveRequestForTeams app: the standard chain HR configured in DE_Setup
 * (Setting/approvalFlow), or that employee's own override if one exists.
 * CurrentUser is set to the first approver; the rest is kept
 * (semicolon-joined) in ApprovalFlow and consumed step by step by
 * RespondToExpenseReport. Once the chain is fully exhausted, the report
 * moves to Status=APPROVED — the terminal, final state (approval is the
 * last step of the workflow; there's no payment layer). If instead it's
 * declined, the requester can edit and resubmit it — see
 * ResubmitExpenseReport.ts, which re-enters this same chain from the top.
 *
 * Lines may mix currencies (RON/EUR/USD) — each line keeps its own original
 * amount/currency, but the report header's TotalAmount/Currency is always
 * HOME_CURRENCY, computed from each line's own exchangeRate (1 for a
 * HOME_CURRENCY line). There's no live FX-rate lookup in this app; the rate
 * is whatever the employee enters (e.g. off the receipt, or the day's bank
 * rate), recorded per line for Finance's own reference.
 */
export async function CreateExpenseReport(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
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
        const { siteId, expenseReportsListId, expenseLinesListId, setupListId, auditLogListId } = await getExpenseConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("displayName,mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;
        const currentUserName: string = me.displayName ?? currentUserPrincipal;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const employeeChain = getApprovalChainFor(setupRows, currentUserPrincipal);
        const fullChain = await resolveChainTokens(graphClient, setupRows, employeeChain);

        if (fullChain.length === 0) {
            throw new Error("No approval chain is configured. Please contact HR.");
        }

        const [firstApprover, ...remainingChain] = fullChain;

        const totalAmount = body.lines.reduce((sum, line) => sum + line.amount * line.exchangeRate, 0);

        const created = await graphClient
            .api(`/sites/${siteId}/lists/${expenseReportsListId}/items`)
            .post({
                fields: {
                    EmployeeCode: currentUserPrincipal,
                    RequestSubject: body.subject,
                    RequestDate: new Date().toISOString(),
                    TotalAmount: totalAmount,
                    Currency: HOME_CURRENCY,
                    Status: "PENDING",
                    CurrentUser: firstApprover,
                    ApprovalFlow: remainingChain.join(";"),
                    Comments: body.comments ?? "",
                },
            });

        const createdLines = [];
        for (let i = 0; i < body.lines.length; i++) {
            const line = body.lines[i];
            const createdLine = await graphClient
                .api(`/sites/${siteId}/lists/${expenseLinesListId}/items`)
                .post({
                    fields: {
                        ExpenseReportId: Number(created.id),
                        LineNumber: i + 1,
                        ExpenseDate: new Date(line.expenseDate).toISOString(),
                        CategoryId: line.categoryId,
                        Vendor: line.vendor,
                        Description: line.description,
                        Amount: line.amount,
                        Currency: line.currency,
                        ExchangeRate: line.exchangeRate,
                    },
                });
            createdLines.push({ id: createdLine.id, lineNumber: i + 1, ...line });
        }

        try {
            await logAuditEntry(graphClient, siteId, auditLogListId, {
                actor: currentUserPrincipal,
                action: "Created",
                employeeCode: currentUserPrincipal,
                requestId: created.id,
                details: body.comments ?? "",
            });
        } catch (auditError) {
            context.warn("logAuditEntry failed (non-blocking)", auditError);
        }

        const notifyMessage = `${currentUserName} needs your approval for an expense report ("${body.subject}", ${totalAmount.toFixed(2)} ${HOME_CURRENCY}).`;
        try {
            const cardScriptTag = buildApprovalCardScriptTag({
                requesterName: currentUserName,
                requestSubject: body.subject,
                totalAmount,
                currency: HOME_CURRENCY,
                requestId: created.id,
                tabBaseUrl,
                teamsDeepLink,
                comment: body.comments,
            });
            await sendApprovalActionEmail(tenantId, 
                firstApprover,
                `Expense Report - ${currentUserName}`,
                notifyMessage,
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
            status: 201,
            jsonBody: {
                id: created.id,
                subject: body.subject,
                totalAmount,
                currency: HOME_CURRENCY,
                status: "PENDING",
                currentUser: firstApprover,
                lines: createdLines,
            },
        };
    } catch (error) {
        context.error("CreateExpenseReport failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Expense_CreateExpenseReport", {
    route: "expense/CreateExpenseReport",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: CreateExpenseReport,
});
