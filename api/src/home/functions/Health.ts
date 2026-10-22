import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";

/**
 * Liveness/readiness target only — no SharePoint/Graph calls, so it answers
 * as soon as the Functions host has finished booting. The web container's
 * nginx proxies /healthz here (see web/nginx.conf) so Container Apps can
 * probe the "web" container's own port and still find out whether the
 * sibling "api" container is actually reachable yet.
 */
export async function Health(_request: HttpRequest, _context: InvocationContext): Promise<HttpResponseInit> {
    return { jsonBody: { status: "ok" } };
}

app.http("Health", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: Health,
});
