/**
 * Builds a Teams deep link so a plain link (e.g. in an email) opens this
 * app's tab INSIDE Microsoft Teams, rather than as a raw website. This
 * matters because Teams SSO (`@microsoft/teams-js`) only initializes when
 * the page is loaded inside Teams' own iframe/parent window — opening the
 * bare site URL directly in a browser tab fails with
 * "Authentication failed: Initialization Failed. No Parent window found."
 * The tab's `entityId` is always "home" (the only `staticTabs[].entityId`
 * in appPackage/manifest.json); `module` is passed through as the deep
 * link's `subEntityId`, which the home screen reads on load (teams-js
 * `page.subPageId`) to jump straight into that module instead of the
 * request-type picker.
 */
export function getTeamsDeepLink(teamsAppId: string, tabBaseUrl: string, module?: "leave" | "expense"): string {
    const webUrl = module ? `${tabBaseUrl}/${module}` : tabBaseUrl;
    let link = `https://teams.microsoft.com/l/entity/${teamsAppId}/home?webUrl=${encodeURIComponent(webUrl)}`;
    if (module) {
        link += `&context=${encodeURIComponent(JSON.stringify({ subEntityId: module }))}`;
    }
    return link;
}
