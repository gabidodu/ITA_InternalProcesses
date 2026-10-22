"use client";

import { useEffect } from "react";
import { app } from "@microsoft/teams-js";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, INCREMENTAL_CONSENT_SCOPES } from "@/lib/msalConfig";

/**
 * Opened by Teams as a popup (via authentication.authenticate()) when the API's
 * silent On-Behalf-Of call fails because the signed-in user hasn't yet consented
 * to a downstream Microsoft Graph scope (e.g. Sites.Selected). Redirects to the
 * Microsoft identity platform to collect that consent, then hands off to
 * /auth-end, which reports success back to the main tab.
 */
export default function AuthStart() {
    useEffect(() => {
        async function run() {
            await app.initialize();
            const msalInstance = new PublicClientApplication(msalConfig);
            await msalInstance.initialize();
            await msalInstance.loginRedirect({
                scopes: INCREMENTAL_CONSENT_SCOPES,
                redirectUri: `${window.location.origin}/auth-end`,
                prompt: "consent",
            });
        }
        void run();
    }, []);

    return <p style={{ padding: 16 }}>Requesting access…</p>;
}
