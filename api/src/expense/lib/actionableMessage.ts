import { getOriginatorId } from "../../shared/lib/originator";
/**
 * The Approve/Decline Adaptive Card for one pending expense-report approval
 * step, embedded via Outlook Actionable Messages' plain
 * `<script type="application/adaptivecard+json">` method (verified via
 * originator registration + SPF/DKIM sender alignment, not a signed card) —
 * this is the sibling LeaveRequestForTeams app's own final, working
 * approach; an earlier signed RSA-JWS variant was tried there and dropped
 * (never rendered), so this app goes straight to what's proven to work.
 * `hideOriginalBody: true` tells Outlook to hide the email's normal HTML
 * body when it DOES render the card, so clients that don't support
 * Actionable Messages still fall back to a readable plain email (see
 * emailNotification.ts's sendApprovalActionEmail, which builds that
 * fallback body) instead of showing nothing. `version: "1.0"` — every
 * official Microsoft example card uses 1.0, the safest bet for consistent
 * rendering across Outlook clients including mobile.
 */
function buildExpenseApprovalCard(params: {
    requesterName: string;
    requestSubject: string;
    totalAmount: number;
    currency: string;
    actionBaseUrl: string;
    requestId: string;
    teamsDeepLink: string;
    comment?: string;
}): object {
    const { requesterName, requestSubject, totalAmount, currency, actionBaseUrl, requestId, teamsDeepLink, comment } = params;
    return {
        type: "AdaptiveCard",
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        version: "1.0",
        originator: getOriginatorId("expense"),
        hideOriginalBody: true,
        body: [
            {
                type: "Container",
                style: "emphasis",
                bleed: true,
                items: [{ type: "TextBlock", text: "EXPENSE REPORTS", weight: "Bolder", size: "Small", color: "Accent", spacing: "None" }],
            },
            {
                type: "TextBlock",
                text: "Expense report awaiting your approval",
                weight: "Bolder",
                size: "Medium",
                wrap: true,
                spacing: "Medium",
            },
            { type: "TextBlock", text: `${requesterName} — ${requestSubject}`, wrap: true, spacing: "Small" },
            {
                type: "FactSet",
                facts: [{ title: "Total", value: `${totalAmount.toFixed(2)} ${currency}` }],
            },
            ...(comment
                ? [{ type: "TextBlock", text: `"${comment}"`, wrap: true, isSubtle: true, spacing: "Small" }]
                : []),
            {
                type: "Input.Text",
                id: "commentInput",
                placeholder: "Add a comment (optional)",
                isMultiline: true,
                spacing: "Medium",
            },
        ],
        actions: [
            {
                type: "Action.Http",
                title: "Approve",
                style: "positive",
                method: "POST",
                url: `${actionBaseUrl}?id=${encodeURIComponent(requestId)}&decision=approved`,
                // Azure Static Web Apps' managed-API proxy overwrites the standard
                // Authorization header with its own internal SWA<->Function token before a
                // request reaches our code — setting Authorization to "" here tells Outlook
                // to send its real bearer token via Action-Authorization instead, which the
                // proxy leaves untouched.
                headers: [{ name: "Authorization", value: "" }],
                // The raw comment text becomes the entire POST body (see
                // RespondToExpenseReportAction.ts, which reads it verbatim, not as JSON) —
                // deliberately not wrapped in hand-built JSON, since Outlook's Input.Text
                // substitution isn't escaped and a comment containing a quote would produce
                // an invalid payload otherwise.
                body: "{{commentInput.value}}",
            },
            {
                type: "Action.Http",
                title: "Decline",
                style: "destructive",
                method: "POST",
                url: `${actionBaseUrl}?id=${encodeURIComponent(requestId)}&decision=declined`,
                headers: [{ name: "Authorization", value: "" }],
                body: "{{commentInput.value}}",
            },
            { type: "Action.OpenUrl", title: "Open in Teams", url: teamsDeepLink },
        ],
    };
}

/**
 * Builds just the `<script type="application/adaptivecard+json">` tag for an
 * approval-needed notification — meant to sit in the `<head>` of an email
 * whose `<body>` carries this app's normal styled fallback content (see
 * emailNotification.ts's sendApprovalActionEmail, which combines the two).
 */
export function buildApprovalCardScriptTag(params: {
    requesterName: string;
    requestSubject: string;
    totalAmount: number;
    currency: string;
    requestId: string;
    tabBaseUrl: string;
    teamsDeepLink: string;
    comment?: string;
}): string {
    const actionBaseUrl = `${params.tabBaseUrl}/api/expense/RespondToExpenseReportAction`;
    const card = buildExpenseApprovalCard({ ...params, actionBaseUrl });
    return `<script type="application/adaptivecard+json">${JSON.stringify(card)}</script>`;
}

/**
 * A "refresh card" — no actions, just a confirmation message — returned as
 * the HTTP response body from RespondToExpenseReportAction (alongside the
 * CARD-UPDATE-IN-BODY: true response header) so Outlook replaces the
 * Approve/Decline card in the email with this one, permanently removing the
 * buttons so the same click can't be submitted twice.
 */
export function buildActionRecordedCard(message: string, title: string = "Your answer has been recorded."): object {
    return {
        type: "AdaptiveCard",
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        version: "1.0",
        // Without this, Outlook falls back to showing the original email's
        // plain-HTML body (built for clients that can't render Actionable
        // Messages) underneath this refresh card — the original card's own
        // hideOriginalBody:true only applies to that first render, not to a
        // later CARD-UPDATE-IN-BODY replacement. The same goes for `originator`
        // (the leave module's recorded card carries it for the same reason).
        originator: getOriginatorId("expense"),
        hideOriginalBody: true,
        body: [
            {
                type: "Container",
                style: "emphasis",
                bleed: true,
                items: [{ type: "TextBlock", text: "EXPENSE REPORTS", weight: "Bolder", size: "Small", color: "Accent", spacing: "None" }],
            },
            { type: "TextBlock", text: title, weight: "Bolder", size: "Medium", wrap: true, spacing: "Medium" },
            { type: "TextBlock", text: message, wrap: true, spacing: "Small" },
        ],
    };
}
