import { getOriginatorId } from "../../shared/lib/originator";
/**
 * The Approve/Decline Adaptive Card for one pending leave-request approval
 * step, embedded via Outlook Actionable Messages' plain
 * `<script type="application/adaptivecard+json">` method (verified via
 * SPF/DKIM sender alignment, not a signed card) — the only embedding method
 * confirmed to actually render in this tenant's Outlook (a separate signed
 * RSA-JWS approach was tried and dropped — never rendered). `hideOriginalBody:
 * true` tells Outlook to hide the email's normal HTML body when it DOES
 * render the card, so clients that don't support Actionable Messages still
 * fall back to a readable plain email (see emailNotification.ts's
 * sendApprovalActionEmail, which builds that fallback body) instead of
 * showing nothing. `style: "positive"/"destructive"` on the actions gives
 * Outlook's renderer a hint to color Approve/Decline green/red, mirroring
 * this app's own good/critical status palette. `version: "1.0"` — every
 * official Microsoft example card uses 1.0 (not a newer schema version); the
 * card uses no feature that needs anything newer, and 1.0 is the safest bet
 * for consistent rendering across Outlook clients, including mobile.
 */
function buildLeaveApprovalCard(params: {
    requesterName: string;
    requestSubject: string;
    startDateLabel: string;
    endDateLabel: string;
    takenDays: number;
    actionBaseUrl: string;
    requestId: string;
    teamsDeepLink: string;
    originatorOverride?: string;
}): object {
    const { requesterName, requestSubject, startDateLabel, endDateLabel, takenDays, actionBaseUrl, requestId, teamsDeepLink } = params;
    return {
        type: "AdaptiveCard",
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        version: "1.0",
        originator: params.originatorOverride || getOriginatorId("leave"),
        hideOriginalBody: true,
        body: [
            {
                type: "Container",
                style: "emphasis",
                bleed: true,
                items: [{ type: "TextBlock", text: "LEAVE REQUESTS", weight: "Bolder", size: "Small", color: "Accent", spacing: "None" }],
            },
            {
                type: "TextBlock",
                text: "Leave request awaiting your approval",
                weight: "Bolder",
                size: "Medium",
                wrap: true,
                spacing: "Medium",
            },
            { type: "TextBlock", text: `${requesterName} — ${requestSubject}`, wrap: true, spacing: "Small" },
            {
                type: "FactSet",
                facts: [
                    { title: "Period", value: `${startDateLabel} – ${endDateLabel}` },
                    { title: "Days", value: String(takenDays) },
                ],
            },
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
                // Authorization header with its own internal SWA↔Function token before a
                // request reaches our code — the exact same gotcha hit with Teams SSO
                // early in this project. Setting Authorization to "" here tells Outlook
                // to send its real bearer token via Action-Authorization instead, which
                // the proxy leaves untouched.
                headers: [{ name: "Authorization", value: "" }],
                // Input value substitution — the raw comment text becomes the entire POST
                // body (see RespondToLeaveRequestAction.ts, which reads it verbatim, not as
                // JSON). Deliberately NOT wrapped in a hand-built JSON string: Outlook's
                // substitution is a plain string replace, not escaped, so a comment
                // containing a quote or backslash would produce invalid JSON — sending it
                // as the whole body sidesteps that entirely, since id/decision already
                // travel via the URL, a separate channel.
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
    startDateLabel: string;
    endDateLabel: string;
    takenDays: number;
    requestId: string;
    tabBaseUrl: string;
    teamsDeepLink: string;
    originatorOverride?: string;
}): string {
    const actionBaseUrl = `${params.tabBaseUrl}/api/leave/RespondToLeaveRequestAction`;
    const card = buildLeaveApprovalCard({ ...params, actionBaseUrl });
    return `<script type="application/adaptivecard+json">${JSON.stringify(card)}</script>`;
}

/**
 * A "refresh card" — no actions, just a confirmation message — returned as
 * the HTTP response body from RespondToLeaveRequestAction (alongside the
 * CARD-UPDATE-IN-BODY: true response header) so Outlook replaces the
 * Approve/Decline card in the email with this one, permanently removing the
 * buttons so the same click can't be submitted twice. Needs its own
 * `originator`/`hideOriginalBody: true` — Outlook evaluates each card render
 * independently, so without these (missing here originally) the original
 * card's hideOriginalBody didn't carry over, and the plain HTML fallback body
 * (meant only for clients that can't render Actionable Messages) stayed
 * visible underneath the "recorded" confirmation card instead of being hidden.
 * `title` is overridable for the "nothing was applied" outcomes (not awaiting
 * this user, or a concurrent response won the race), where "recorded" would
 * be misleading.
 */
export function buildActionRecordedCard(message: string, title: string = "Your answer has been recorded."): object {
    return {
        type: "AdaptiveCard",
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        version: "1.0",
        originator: getOriginatorId("leave"),
        hideOriginalBody: true,
        body: [
            {
                type: "Container",
                style: "emphasis",
                bleed: true,
                items: [{ type: "TextBlock", text: "LEAVE REQUESTS", weight: "Bolder", size: "Small", color: "Accent", spacing: "None" }],
            },
            { type: "TextBlock", text: title, weight: "Bolder", size: "Medium", wrap: true, spacing: "Medium" },
            { type: "TextBlock", text: message, wrap: true, spacing: "Small" },
        ],
    };
}
