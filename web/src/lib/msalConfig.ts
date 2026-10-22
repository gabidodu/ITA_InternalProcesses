import { Configuration } from "@azure/msal-browser";

export const AAD_APP_CLIENT_ID = "359906f4-0bbf-4810-97f4-0d1b28aa5f1a";

export const msalConfig: Configuration = {
    auth: {
        clientId: AAD_APP_CLIENT_ID,
        authority: "https://login.microsoftonline.com/organizations",
    },
    cache: {
        cacheLocation: "sessionStorage",
    },
};

// The downstream Microsoft Graph permission our API needs via On-Behalf-Of.
// Requesting it here interactively is what records the user's consent grant,
// so the API's silent OBO call can succeed afterwards.
export const INCREMENTAL_CONSENT_SCOPES = ["https://graph.microsoft.com/Sites.Selected"];
