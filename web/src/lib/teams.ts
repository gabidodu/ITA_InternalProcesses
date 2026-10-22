import { app, authentication } from "@microsoft/teams-js";

/**
 * Local preview mode (`NEXT_PUBLIC_PREVIEW=1`, see .env.preview): the UI runs
 * in a plain browser tab outside Teams, with no SSO and canned API responses
 * from lib/preview.ts — for looking at layout and navigation, never for
 * testing real data. Baked in at build time, so a normal build can never
 * end up in preview mode by accident.
 */
export const isPreview = process.env.NEXT_PUBLIC_PREVIEW === "1";

let initPromise: Promise<void> | undefined;

/**
 * Initializes teams-js once per page load. The home screen and both module
 * pages each call this on mount; client-side navigation between them keeps
 * the same JS context, so later calls reuse the first initialization instead
 * of re-running it.
 */
export function initTeams(): Promise<void> {
  if (isPreview) return Promise.resolve();
  initPromise ??= app.initialize();
  return initPromise;
}

/** Teams SSO token for the API's X-Teams-Sso-Token header. */
export async function getSsoToken(): Promise<string> {
  if (isPreview) return "preview";
  await initTeams();
  return authentication.getAuthToken();
}

/**
 * The deep link's `subEntityId` (set by the API's getTeamsDeepLink for links
 * in notification emails) — which module to open directly instead of the
 * home screen. Answered only once per page load: Teams keeps reporting the
 * same subPageId for the tab's whole lifetime, so without this, the "← All
 * requests" link would bounce straight back into the deep-linked module.
 */
let requestedModuleConsumed = false;

export async function getRequestedModule(): Promise<string | undefined> {
  if (isPreview || requestedModuleConsumed) return undefined;
  requestedModuleConsumed = true;
  await initTeams();
  const context = await app.getContext();
  return context.page.subPageId || undefined;
}
