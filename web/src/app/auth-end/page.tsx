"use client";

import { useEffect, useState } from "react";
import { app, authentication } from "@microsoft/teams-js";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig } from "@/lib/msalConfig";

/**
 * Completes the interactive consent redirect started in /auth-start, then
 * reports success (or failure) back to the tab that opened this popup via
 * Teams' authentication.notifySuccess/notifyFailure.
 */
export default function AuthEnd() {
    const [errorMessage, setErrorMessage] = useState<string | null>(null);

    useEffect(() => {
        async function run() {
            try {
                await app.initialize();
                const msalInstance = new PublicClientApplication(msalConfig);
                await msalInstance.initialize();
                const result = await msalInstance.handleRedirectPromise();

                if (!result) {
                    throw new Error("No redirect result — consent may have been cancelled.");
                }

                authentication.notifySuccess("consent-granted");
            } catch (error) {
                const message = error instanceof Error ? error.message : "Unknown error";
                setErrorMessage(message);
                authentication.notifyFailure(message);
            }
        }
        void run();
    }, []);

    return <p style={{ padding: 16 }}>{errorMessage ?? "Finishing up…"}</p>;
}
