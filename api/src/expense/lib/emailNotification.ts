import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { getAppOnlyGraphToken } from "../../shared/auth/appOnly";
import { getExpenseConfig } from "../../shared/lib/tenantConfig";

/**
 * All outbound notification email sends from the client tenant's shared
 * mailbox (`mailSenderAddress` in TenantConfigs, shared with the leave
 * module), not the acting user's own — consistent branding, and app-only auth
 * (Application Mail.Send) since sending "as" an arbitrary shared mailbox
 * isn't possible with a delegated, per-user token. The app-only token is
 * minted for that same tenant, so this works for any provisioned client.
 */
async function getMailSender(tenantId: string): Promise<{ graphClient: Client; senderAddress: string }> {
    const [token, config] = await Promise.all([getAppOnlyGraphToken(tenantId), getExpenseConfig(tenantId)]);
    return {
        graphClient: Client.init({ authProvider: (done) => done(null, token) }),
        senderAddress: config.mailSenderAddress,
    };
}

function escapeHtml(text: string): string {
    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

export type EmailDetail = { label: string; value: string };

function buildDetailsHtml(details?: EmailDetail[]): string {
    if (!details?.length) return "";
    const rows = details
        .map(
            (d) => `
          <tr>
            <td style="padding:5px 14px 5px 0;font-size:11px;font-weight:700;letter-spacing:0.05em;color:#6b7280;text-transform:uppercase;white-space:nowrap;vertical-align:top;">${escapeHtml(d.label)}</td>
            <td style="padding:5px 0;font-size:14px;color:#1c2333;font-weight:600;">${escapeHtml(d.value)}</td>
          </tr>`
        )
        .join("");
    return `
        <table role="presentation" cellpadding="0" cellspacing="0" style="margin:0 0 24px 0;background-color:#f8f9fb;border-radius:8px;padding:4px 14px;">
          ${rows}
        </table>`;
}

// A small, table-based transactional-email template (inline styles only, no
// flexbox/grid) so it renders consistently in Outlook desktop's limited HTML
// engine as well as modern web/mobile mail clients. `details` is an optional
// list of facts (total, who acted, ...) rendered as a compact label/value
// block between the message and the button.
function buildEmailHtml(messageText: string, tabUrl: string, details?: EmailDetail[]): string {
    const safeMessage = escapeHtml(messageText);
    return `
<div style="font-family:'Segoe UI',Helvetica,Arial,sans-serif;background-color:#f2f3f6;padding:32px 16px;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:480px;margin:0 auto;background-color:#ffffff;border-radius:10px;overflow:hidden;border:1px solid #e2e4ea;">
    <tr>
      <td style="background-color:#2b3a67;padding:18px 24px;">
        <span style="color:#ffffff;font-size:13px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;">Expense Reports</span>
      </td>
    </tr>
    <tr>
      <td style="padding:28px 24px 8px 24px;">
        <p style="margin:0 0 20px 0;font-size:15px;line-height:1.6;color:#1c2333;">${safeMessage}</p>
        ${buildDetailsHtml(details)}
        <table role="presentation" cellpadding="0" cellspacing="0">
          <tr>
            <td style="border-radius:6px;background-color:#2b3a67;">
              <a href="${tabUrl}" style="display:inline-block;padding:11px 22px;font-size:14px;font-weight:600;color:#ffffff;text-decoration:none;border-radius:6px;">Open in Teams</a>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td style="padding:18px 24px;background-color:#f8f9fb;border-top:1px solid #e2e4ea;">
        <p style="margin:0;font-size:12px;line-height:1.5;color:#6b7280;">
          This is an automated notification from the Expense Reports app. Please don't reply to this email.
        </p>
      </td>
    </tr>
  </table>
</div>`.trim();
}

/**
 * Sends a plain email from the tenant's shared mailbox (app-only
 * `Mail.Send`) to a specific recipient, with a link back into the Teams tab.
 * Best-effort: callers should wrap this in try/catch so a mail failure never
 * blocks the actual expense-report action.
 */
export async function sendExpenseEmailNotification(
    tenantId: string,
    to: string,
    subject: string,
    messageText: string,
    tabUrl: string,
    details?: EmailDetail[]
): Promise<void> {
    const { graphClient, senderAddress } = await getMailSender(tenantId);
    await graphClient.api(`/users/${senderAddress}/sendMail`).post({
        message: {
            subject,
            body: {
                contentType: "HTML",
                content: buildEmailHtml(messageText, tabUrl, details),
            },
            toRecipients: [{ emailAddress: { address: to } }],
        },
        saveToSentItems: false,
    });
}

/**
 * Sends the "needs your approval" email using Outlook Actionable Messages'
 * plain `<script type="application/adaptivecard+json">` embedding (see
 * lib/actionableMessage.ts's buildApprovalCardScriptTag) — the script tag
 * sits in `<head>`, and the `<body>` carries this app's normal styled
 * template (same `buildEmailHtml` every other notification uses). The card
 * itself sets `hideOriginalBody: true`, so Outlook hides this body when it
 * DOES render the card; a recipient whose client doesn't support Actionable
 * Messages just sees the normal styled email instead of a blank one.
 */
export async function sendApprovalActionEmail(
    tenantId: string,
    to: string,
    subject: string,
    messageText: string,
    tabUrl: string,
    cardScriptTag: string,
    details?: EmailDetail[]
): Promise<void> {
    const { graphClient, senderAddress } = await getMailSender(tenantId);
    const html = `<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">${cardScriptTag}</head><body>${buildEmailHtml(messageText, tabUrl, details)}</body></html>`;
    await graphClient.api(`/users/${senderAddress}/sendMail`).post({
        message: {
            subject,
            body: {
                contentType: "HTML",
                content: html,
            },
            toRecipients: [{ emailAddress: { address: to } }],
        },
        saveToSentItems: false,
    });
}
