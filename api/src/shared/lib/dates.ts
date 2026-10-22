// Deliberately uses local getters (.getDate()/.getMonth()), not UTC — some
// stored dates (anything the legacy pre-migration app wrote) are local
// Romania midnight saved as UTC without conversion (e.g. "...T21:00:00Z" for
// what's really the next day locally). Reading them back with local getters
// undoes that, IF the process's own local timezone is also Romania's — which
// it isn't by default on Container Apps (Linux containers default to UTC).
// The container's `api` service therefore needs the env var
// TZ=Europe/Bucharest set (not tracked in this repo — no IaC defines this
// container app, it's provisioned imperatively via `az`) for this file to
// format dates correctly; without it, every date this function touches comes
// out one day early for any request the legacy app originally created.
/** Formats a date as dd/mm/yyyy for user-facing messages. */
export function formatDate(date: Date | string): string {
    const d = typeof date === "string" ? new Date(date) : date;
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yyyy = d.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
}

/** Formats a date as dd/mm/yyyy HH:mm for signature-style timestamps. */
export function formatDateTime(date: Date | string): string {
    const d = typeof date === "string" ? new Date(date) : date;
    const hh = String(d.getHours()).padStart(2, "0");
    const min = String(d.getMinutes()).padStart(2, "0");
    return `${formatDate(d)} ${hh}:${min}`;
}

/** Last year / current year / next year — the shared "recent enough to show" window used across lists (holidays, leave requests). */
export function getRelevantYears(): Set<number> {
    const currentYear = new Date().getFullYear();
    return new Set([currentYear - 1, currentYear, currentYear + 1]);
}
