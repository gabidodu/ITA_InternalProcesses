import { Client } from "@microsoft/microsoft-graph-client";

/**
 * Sends a Teams Activity Feed notification (the bell icon) to a specific user,
 * deep-linking into this app's tab. Requires the "leaveRequestUpdate" activity
 * type declared in appPackage/manifest.json's `activities.activityTypes`, and
 * the recipient must have the app installed.
 *
 * Best-effort: failures here (e.g. TeamsActivity.Send not yet consented) must
 * never block the leave-request action itself, so callers should swallow
 * errors from this function rather than let them fail the request.
 */
export async function sendLeaveActivityNotification(
    graphClient: Client,
    recipientUpn: string,
    messageText: string,
    webUrl: string
): Promise<void> {
    await graphClient.api(`/users/${recipientUpn}/teamwork/sendActivityNotification`).post({
        topic: {
            source: "text",
            value: "Cerere de concediu",
            webUrl,
        },
        activityType: "leaveRequestUpdate",
        previewText: {
            content: messageText,
        },
        templateParameters: [
            {
                name: "message",
                value: messageText,
            },
        ],
    });
}
