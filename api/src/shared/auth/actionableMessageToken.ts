import jwt, { GetPublicKeyOrSecret, JwtPayload } from "jsonwebtoken";
import jwksClient, { JwksClient } from "jwks-rsa";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { getAppOnlyGraphToken } from "./appOnly";

// Microsoft's fixed "Actions" service app id — every Outlook Actionable
// Messages Action.Http request is issued a token with this as its client
// (azp/appid). Not tenant-specific, so it's a constant, not an app setting.
const MICROSOFT_ACTIONS_APP_ID = "48af08dc-f6d2-435f-b2a7-069abd99c086";

let cachedJwksClient: JwksClient | undefined;

// The "common" endpoint validates a v2.0 token issued by ANY Entra ID tenant
// (not just one hardcoded tenant) — required now that Approve/Decline emails
// can come from any client tenant that's provisioned this app, not just
// IT ASSIST's own. Which specific tenant issued the token is still checked
// afterwards, against the token's own `tid` claim plus the `azp`/`appid`
// check below — this endpoint only proves the signature is genuinely
// Microsoft's, not which tenant to trust.
function getJwksClient(): JwksClient {
    if (!cachedJwksClient) {
        cachedJwksClient = jwksClient({
            jwksUri: "https://login.microsoftonline.com/common/discovery/v2.0/keys",
            cache: true,
            rateLimit: true,
        });
    }
    return cachedJwksClient;
}

// Best-effort, unverified decode — used only to enrich error messages with the
// actual claims a rejected token carried (aud/iss/azp/appid), since none of
// that is visible otherwise (no Application Insights wired up for this app,
// and Outlook never surfaces our response body to the end user).
function decodeUnverifiedClaims(token: string): Record<string, unknown> | null {
    try {
        const payload = token.split(".")[1];
        return JSON.parse(Buffer.from(payload, "base64").toString("utf8"));
    } catch {
        return null;
    }
}

/**
 * Validates the Microsoft Entra ID bearer token Outlook's Actionable Messages
 * "Actions" service attaches to an Action.Http request when a user clicks
 * Approve/Decline in an email — the email-native counterpart to Teams SSO for
 * this app's other entry point. Confirms the token is genuinely issued by
 * this tenant, scoped to ACTIONABLE_MESSAGE_AUDIENCE (this app's dedicated
 * Actionable Messages API scope), and requested by Microsoft's own Actions
 * service rather than any other caller with a token. Returns the acting
 * user's email AND their tenant id (this endpoint has no Teams SSO context to
 * get one from any other way) so callers can apply the same "must equal
 * CurrentUser" authorization rule as the in-Teams approve/decline flow, and
 * resolve which client's SharePoint config to use.
 *
 * The token's exact issuer format isn't pinned down in Microsoft's own sample
 * (docs show `iss` without the usual `/v2.0` suffix despite `ver: "2.0"`), so
 * issuer is checked leniently (must reference the token's own `tid`) rather
 * than via jsonwebtoken's strict `issuer` option — signature, audience and
 * algorithm are still strictly enforced, which is the security-critical part.
 */
export async function validateActionableMessageToken(
    authorizationHeader: string | null
): Promise<{ actingUser: string; tenantId: string }> {
    if (!authorizationHeader?.toLowerCase().startsWith("bearer ")) {
        throw new Error("Missing or malformed Authorization header.");
    }
    const token = authorizationHeader.slice(7).trim();

    try {
        const audienceSetting = process.env.ACTIONABLE_MESSAGE_AUDIENCE;
        const clientId = process.env.AAD_APP_CLIENT_ID;
        if (!audienceSetting) {
            throw new Error("ACTIONABLE_MESSAGE_AUDIENCE is not configured.");
        }
        // ACTIONABLE_MESSAGE_AUDIENCE may hold several comma-separated AppIdUris —
        // each Actionable Message provider registration (one per originator, e.g.
        // the org-approved one and the older test-scope one) mints tokens whose
        // `aud` is that registration's own AppIdUri, and both must keep validating
        // as long as cards signed with either originator can still be in inboxes.
        const audiences = audienceSetting.split(",").map((a) => a.trim()).filter(Boolean);
        // Entra v2 tokens for this app can legitimately carry either the full
        // AppIdUri (one of `audiences`) or the bare client id as `aud` —
        // observed in practice to be the bare client id, despite the AppIdUri being
        // what's actually registered as the scope's resource. Accept either.
        const acceptedAudiences = (clientId ? [...audiences, clientId] : audiences) as [string, ...string[]];

        const client = getJwksClient();
        const getKey: GetPublicKeyOrSecret = (header, callback) => {
            client.getSigningKey(header.kid, (err, key) => {
                if (err || !key) {
                    callback(err ?? new Error("Signing key not found."));
                    return;
                }
                callback(null, key.getPublicKey());
            });
        };

        const decoded = await new Promise<JwtPayload>((resolve, reject) => {
            jwt.verify(token, getKey, { algorithms: ["RS256"], audience: acceptedAudiences }, (err, payload) => {
                if (err || !payload || typeof payload === "string") {
                    reject(err ?? new Error("Token verification returned no payload."));
                    return;
                }
                resolve(payload);
            });
        });

        const tenantId = decoded.tid as string | undefined;
        if (!tenantId) {
            throw new Error("Token does not include a tenant id (tid).");
        }
        if (typeof decoded.iss !== "string" || !decoded.iss.includes(tenantId)) {
            throw new Error("Token issuer does not match its own tenant id.");
        }

        const callerAppId = (decoded.azp as string | undefined) ?? (decoded.appid as string | undefined);
        if (callerAppId !== MICROSOFT_ACTIONS_APP_ID) {
            throw new Error("Token was not issued to the Microsoft Actionable Messages Actions service.");
        }

        const tokenIdentity = (decoded.preferred_username as string | undefined) ?? (decoded.upn as string | undefined);
        if (!tokenIdentity) {
            throw new Error("Token does not include the acting user's identity.");
        }

        // The token's preferred_username/upn is the user's sign-in name, which can
        // differ from their mail address (e.g. Mircea Mitu signs in as mache@itassist.eu
        // but his mail — what CreateLeaveRequest/RespondToLeaveRequest store on
        // CurrentUser — is mircea.mitu@itassist.eu). Resolve the real mail via an
        // app-only lookup so this matches the Teams-app path's own `me.mail ??
        // me.userPrincipalName`, instead of comparing CurrentUser against a UPN it
        // was never written with — falls back to the raw token identity if the
        // lookup fails, rather than blocking the approval outright.
        let actingUser = tokenIdentity;
        const oid = decoded.oid as string | undefined;
        if (oid) {
            try {
                const appGraphToken = await getAppOnlyGraphToken(tenantId);
                const appGraphClient = Client.init({ authProvider: (done) => done(null, appGraphToken) });
                const resolvedUser = await appGraphClient.api(`/users/${oid}`).select("mail,userPrincipalName").get();
                actingUser = resolvedUser.mail ?? resolvedUser.userPrincipalName ?? tokenIdentity;
            } catch (lookupError) {
                // Non-blocking — the CurrentUser comparison just falls back to the
                // token's own identity, same as before this fix existed.
            }
        }

        return { actingUser, tenantId };
    } catch (error) {
        const claims = decodeUnverifiedClaims(token);
        const baseMessage = error instanceof Error ? error.message : "Unknown error";
        throw new Error(`${baseMessage} | unverified claims: ${JSON.stringify(claims)}`);
    }
}
