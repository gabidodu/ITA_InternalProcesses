import { ConfidentialClientApplication } from "@azure/msal-node";

const clientsByTenant = new Map<string, ConfidentialClientApplication>();

function getConfidentialClient(tenantId: string): ConfidentialClientApplication {
    const cached = clientsByTenant.get(tenantId);
    if (cached) {
        return cached;
    }

    const clientId = process.env.AAD_APP_CLIENT_ID;
    const clientSecret = process.env.AAD_APP_CLIENT_SECRET;
    if (!clientId || !clientSecret) {
        throw new Error("AAD_APP_CLIENT_ID / AAD_APP_CLIENT_SECRET are not configured.");
    }

    const client = new ConfidentialClientApplication({
        auth: {
            clientId,
            clientSecret,
            authority: `https://login.microsoftonline.com/${tenantId}`,
        },
    });
    clientsByTenant.set(tenantId, client);
    return client;
}

/**
 * Acquires an app-only Microsoft Graph token (client-credentials flow, no
 * signed-in user) for a specific client tenant — used for sending email from
 * the shared mailbox (Application-type `Mail.Send`, see
 * lib/emailNotification.ts) and for the Outlook Actionable Messages entry
 * point (RespondToLeaveRequestAction.ts), which has no delegated/OBO context
 * to piggyback on. Same app registration + certificate/secret credential as
 * everywhere else in this app, just requesting a token scoped to the target
 * tenant's own consent — this is what makes app-only calls work across
 * multiple client tenants rather than one hardcoded tenant.
 */
export async function getAppOnlyGraphToken(tenantId: string): Promise<string> {
    const client = getConfidentialClient(tenantId);

    const result = await client.acquireTokenByClientCredential({
        scopes: ["https://graph.microsoft.com/.default"],
    });

    if (!result?.accessToken) {
        throw new Error("App-only token acquisition returned no access token.");
    }

    return result.accessToken;
}
