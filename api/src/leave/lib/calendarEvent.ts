import { Client } from "@microsoft/microsoft-graph-client";
import { getAppOnlyGraphToken } from "../../shared/auth/appOnly";
import { getTenantConfig } from "../../shared/lib/tenantConfig";

const DEFAULT_CALENDAR_NAME = "Leave Requests";

function getCalendarName(): string {
    return process.env.CALENDAR_NAME || DEFAULT_CALENDAR_NAME;
}

async function getCalendarGraphClient(tenantId: string): Promise<{ graphClient: Client; mailbox: string }> {
    const [token, tenant] = await Promise.all([getAppOnlyGraphToken(tenantId), getTenantConfig(tenantId)]);
    return { graphClient: Client.init({ authProvider: (done) => done(null, token) }), mailbox: tenant.mailSenderAddress };
}

// Keyed by tenant — each client's shared mailbox has its own "Leave Requests"
// calendar, so the id cached for one tenant must never be reused for another.
const cachedCalendarIdByTenant = new Map<string, string>();

/**
 * Finds the shared "Leave Requests" calendar in the mailbox by name (case-
 * insensitive), creating it if it doesn't exist yet — no manual setup step
 * needed. Result is cached in memory for the lifetime of this Function App
 * instance, since the calendar itself never changes once created.
 */
async function getOrCreateCalendarId(tenantId: string, graphClient: Client, mailbox: string): Promise<string> {
    const cached = cachedCalendarIdByTenant.get(tenantId);
    if (cached) {
        return cached;
    }

    const calendarName = getCalendarName();

    const calendarsResponse = await graphClient.api(`/users/${mailbox}/calendars`).get();
    const existing = calendarsResponse.value.find(
        (cal: { id: string; name: string }) => cal.name.toLowerCase() === calendarName.toLowerCase()
    );
    if (existing) {
        cachedCalendarIdByTenant.set(tenantId, existing.id);
        return existing.id;
    }

    const created = await graphClient.api(`/users/${mailbox}/calendars`).post({ name: calendarName });
    cachedCalendarIdByTenant.set(tenantId, created.id);
    return created.id;
}

/** Graph all-day events use an exclusive end date — a leave from day 1 to day 1 needs end = day 2. */
function toAllDayDate(date: Date): string {
    return date.toISOString().slice(0, 10);
}

/**
 * Creates an all-day event spanning the leave period in the shared "Leave
 * Requests" calendar, subject "{requesterName} (pending)". Returns the new
 * event's id, to be stored on the LR_LeaveRequests item (CalendarEventId) so
 * it can be renamed/deleted later. Best-effort: callers should wrap this in
 * try/catch so a calendar failure never blocks request creation.
 */
export async function createPendingLeaveCalendarEvent(
    tenantId: string,
    requesterName: string,
    startDate: Date,
    endDate: Date
): Promise<string> {
    const { graphClient, mailbox } = await getCalendarGraphClient(tenantId);
    const calendarId = await getOrCreateCalendarId(tenantId, graphClient, mailbox);

    const exclusiveEnd = new Date(endDate);
    exclusiveEnd.setDate(exclusiveEnd.getDate() + 1);

    const created = await graphClient.api(`/users/${mailbox}/calendars/${calendarId}/events`).post({
        subject: `${requesterName} (pending)`,
        isAllDay: true,
        start: { dateTime: toAllDayDate(startDate), timeZone: "UTC" },
        end: { dateTime: toAllDayDate(exclusiveEnd), timeZone: "UTC" },
        showAs: "free",
    });

    return created.id;
}

/**
 * Renames a leave calendar event to drop the "(pending)" suffix once the
 * request is finally approved. Best-effort — see above.
 */
export async function markLeaveCalendarEventApproved(tenantId: string, eventId: string, requesterName: string): Promise<void> {
    const { graphClient, mailbox } = await getCalendarGraphClient(tenantId);
    await graphClient.api(`/users/${mailbox}/events/${eventId}`).patch({ subject: requesterName });
}

/**
 * Deletes a leave calendar event (on decline or cancellation). Tolerates the
 * event already being gone (404) rather than treating it as a failure.
 */
export async function deleteLeaveCalendarEvent(tenantId: string, eventId: string): Promise<void> {
    const { graphClient, mailbox } = await getCalendarGraphClient(tenantId);
    try {
        await graphClient.api(`/users/${mailbox}/events/${eventId}`).delete();
    } catch (error) {
        const statusCode = (error as { statusCode?: number })?.statusCode;
        if (statusCode !== 404) {
            throw error;
        }
    }
}
