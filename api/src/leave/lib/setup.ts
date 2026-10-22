export { SetupRow, getSetupRows, getSetupValue, isHrAdmin } from "../../shared/lib/setupBase";
import { Client } from "@microsoft/microsoft-graph-client";
import { SetupRow, getSetupValue, isHrAdmin } from "../../shared/lib/setupBase";

export function getHrAdminEmails(rows: SetupRow[]): string[] {
    return rows.filter((r) => r.category === "Admin").map((r) => r.settingKey);
}

/**
 * Whether this user may access the Settings tab (employee/holiday/chain/
 * allocation management): the existing Admin role (Category=Admin), a new HR
 * role (Category=HR, same shape), or the single GM constant (Setting/GM).
 * All three are LR_Setup rows edited directly in SharePoint — there is no
 * in-app UI to grant or revoke any of these roles, by design.
 */
export function canAccessSettings(rows: SetupRow[], userPrincipal: string | undefined): boolean {
    const principal = userPrincipal?.toLowerCase();
    if (!principal) return false;
    if (isHrAdmin(rows, userPrincipal ?? "")) return true;
    const isHrRole = rows.some((r) => r.category === "HR" && r.settingKey?.toLowerCase() === principal);
    const gmEmail = getGmEmail(rows)?.toLowerCase();
    return isHrRole || gmEmail === principal;
}

// Reserved chain-entry tokens, resolved at request-creation time rather than
// stored as a literal email — {manager} is always the requester's own direct
// manager (looked up dynamically per request), {GM} is the person HR marks in
// LR_Setup with Category=GM (SettingKey holds their email, same row shape as
// Admin/HR rows). Mirrors the legacy app's "manager;GM" chain convention, but
// resolved explicitly here instead of implicitly.
export const MANAGER_TOKEN = "{manager}";
export const GM_TOKEN = "{GM}";

export function getGmEmail(rows: SetupRow[]): string | undefined {
    return rows.find((r) => r.category === "GM")?.settingKey;
}

/** The default number of Holiday days granted per employee per year (Setting/defaultAnnualAllocDays), falling back to 21 if HR hasn't set it. */
export function getDefaultAnnualAllocDays(rows: SetupRow[]): number {
    const raw = getSetupValue(rows, "Setting", "defaultAnnualAllocDays");
    const parsed = Number(raw);
    return Number.isFinite(parsed) && parsed > 0 ? parsed : 21;
}

/**
 * Keeps only the last occurrence of each entry (case-insensitive), dropping
 * earlier repeats while preserving the order of what remains. Used so a
 * person who appears more than once in a chain (e.g. the same email listed
 * literally and also reached via {manager}/{GM}) only has to approve once,
 * at the latest point they'd otherwise have appeared.
 */
export function dedupeChainKeepLastOccurrence(chain: string[]): string[] {
    return chain.filter((entry, index) => {
        const key = entry.toLowerCase();
        return !chain.some((other, otherIndex) => otherIndex > index && other.toLowerCase() === key);
    });
}

/**
 * Resolves {manager}/{GM} tokens in a chain into real emails — {manager} via
 * a live `/me/manager` lookup (the signed-in requester's own manager, so no
 * extra permission beyond the already-consented User.Read), {GM} via the
 * LR_Setup constant. Unresolvable entries (e.g. no manager set in Entra ID)
 * are dropped rather than failing the whole request. The resolved list is
 * then deduped, keeping only each person's last occurrence.
 */
export async function resolveChainTokens(
    graphClient: Client,
    setupRows: SetupRow[],
    chain: string[]
): Promise<string[]> {
    const resolved: string[] = [];
    for (const entry of chain) {
        if (entry === MANAGER_TOKEN) {
            try {
                const manager = await graphClient.api("/me/manager").select("mail,userPrincipalName").get();
                const managerEmail = manager.mail ?? manager.userPrincipalName;
                if (managerEmail) {
                    resolved.push(managerEmail);
                }
            } catch {
                // No manager set for this user in Entra ID — skip rather than fail the request.
            }
        } else if (entry === GM_TOKEN) {
            const gmEmail = getGmEmail(setupRows);
            if (gmEmail) {
                resolved.push(gmEmail);
            }
        } else {
            resolved.push(entry);
        }
    }
    return dedupeChainKeepLastOccurrence(resolved);
}

/** Renders a chain entry (real email or {manager}/{GM} token) as a friendly name for prose/audit text. */
export function labelForChainEntry(entry: string, nameByMail: Map<string, string>): string {
    if (entry === MANAGER_TOKEN) return "Direct Manager";
    if (entry === GM_TOKEN) return "GM";
    return nameByMail.get(entry.toLowerCase()) ?? entry;
}

function parseChainValue(raw: string | undefined): string[] {
    return (raw ?? "")
        .split(";")
        .map((s) => s.trim())
        .filter(Boolean);
}

/** The global standard chain (Setting/approvalFlow) — used when an employee has no specific override. */
export function parseApprovalChain(rows: SetupRow[]): string[] {
    return parseChainValue(getSetupValue(rows, "Setting", "approvalFlow"));
}

/** Whether this specific employee has their own override chain (Category=ApprovalFlow). */
export function hasApprovalChainOverride(rows: SetupRow[], employeeCode: string): boolean {
    return rows.some(
        (r) => r.category === "ApprovalFlow" && r.settingKey?.toLowerCase() === employeeCode?.toLowerCase()
    );
}

/** This employee's override chain if set, otherwise falls back to the global standard chain. */
export function getApprovalChainFor(rows: SetupRow[], employeeCode: string): string[] {
    const overrideRow = rows.find(
        (r) => r.category === "ApprovalFlow" && r.settingKey?.toLowerCase() === employeeCode?.toLowerCase()
    );
    return overrideRow ? parseChainValue(overrideRow.settingValue) : parseApprovalChain(rows);
}
