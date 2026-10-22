import { ConfidentialClientApplication } from "@azure/msal-node";

const TENANT_AUTHORITY_TEMPLATE = "https://login.microsoftonline.com/{tenantId}";

let cachedApp: ConfidentialClientApplication | undefined;

function getConfidentialClient(): ConfidentialClientApplication {
    if (cachedApp) {
        return cachedApp;
    }

    const clientId = process.env.AAD_APP_CLIENT_ID;
    const clientSecret = process.env.AAD_APP_CLIENT_SECRET;
    if (!clientId || !clientSecret) {
        throw new Error("AAD_APP_CLIENT_ID / AAD_APP_CLIENT_SECRET are not configured.");
    }

    cachedApp = new ConfidentialClientApplication({
        auth: {
            clientId,
            clientSecret,
            // "organizations" accepts sign-in from any Entra ID tenant (multi-tenant SaaS),
            // the actual tenant is resolved per-request from the caller's token (tid claim).
            authority: TENANT_AUTHORITY_TEMPLATE.replace("{tenantId}", "organizations"),
        },
    });

    return cachedApp;
}

export type OboResult = {
    accessToken: string;
    /** The signed-in user's own Entra tenant id — resolves which client's SharePoint config to use (see lib/tenantConfig.ts). */
    tenantId: string;
};

/**
 * Reads `tid` directly off the incoming Teams SSO token (unverified decode —
 * safe here because the token is already about to be cryptographically
 * validated by the OBO exchange itself; if that exchange succeeds, the token
 * was genuine). Used instead of trusting MSAL's `AuthenticationResult` to
 * carry a tenant id: in practice `result.tenantId`/`result.account?.tenantId`
 * both came back empty for this OBO flow (confirmed via a real failed request
 * — "On-Behalf-Of token acquisition did not return a tenant id"), so the
 * input token is the reliable source, not the SDK's output.
 */
function decodeTenantIdFromToken(token: string): string {
    try {
        const payload = token.split(".")[1];
        const decoded = JSON.parse(Buffer.from(payload, "base64").toString("utf8"));
        if (typeof decoded.tid === "string" && decoded.tid) {
            return decoded.tid;
        }
    } catch {
        // fall through to the error below
    }
    throw new Error("Could not determine the tenant id from the SSO token.");
}

/**
 * Exchanges the Teams SSO token (sent by the tab as a bearer token, audience =
 * this app's own Application ID URI) for a Microsoft Graph access token scoped
 * to the signed-in user, via the On-Behalf-Of flow. Also returns the caller's
 * tenant id, decoded from the SSO token itself (see decodeTenantIdFromToken).
 */
export async function getGraphTokenOnBehalfOf(userAssertionToken: string, scopes: string[]): Promise<OboResult> {
    const client = getConfidentialClient();
    const tenantId = decodeTenantIdFromToken(userAssertionToken);

    const result = await client.acquireTokenOnBehalfOf({
        oboAssertion: userAssertionToken,
        scopes,
    });

    if (!result?.accessToken) {
        throw new Error("On-Behalf-Of token acquisition returned no access token.");
    }

    return { accessToken: result.accessToken, tenantId };
}

/**
 * Reads the Teams SSO token from the custom "X-Teams-Sso-Token" header.
 * Not read from the standard "Authorization" header because Azure Static Web
 * Apps' managed API proxy overwrites that header for its own edge-to-function
 * authentication before the request reaches this function.
 */
export function extractSsoToken(ssoTokenHeader: string | null): string {
    if (!ssoTokenHeader) {
        throw new Error("Missing X-Teams-Sso-Token header.");
    }
    return ssoTokenHeader;
}
