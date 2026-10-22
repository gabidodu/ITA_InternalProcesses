export { SetupRow, getSetupRows, getSetupValue, isHrAdmin } from "../../shared/lib/setupBase";
import { Client } from "@microsoft/microsoft-graph-client";
import { SetupRow, getSetupValue, isHrAdmin } from "../../shared/lib/setupBase";

/**
 * Who gets full, read-only visibility into every expense report regardless
 * of status — every line and its receipt/invoice — via
 * GetAllExpenseReports.ts. A role check, not a CurrentUser chain match.
 */
export function isFinanceUser(rows: SetupRow[], userPrincipal: string): boolean {
    return rows.some((r) => r.category === "Finance" && r.settingKey?.toLowerCase() === userPrincipal?.toLowerCase());
}

export function getFinanceEmails(rows: SetupRow[]): string[] {
    return rows.filter((r) => r.category === "Finance").map((r) => r.settingKey);
}

/** Whether this user may access the HR/Admin settings (approval-chain editor, category management). Admin-category rows only — there is no in-app UI to grant or revoke this role, by design; it's edited directly in SharePoint. */
export function canAccessSettings(rows: SetupRow[], userPrincipal: string | undefined): boolean {
    return Boolean(userPrincipal) && isHrAdmin(rows, userPrincipal ?? "");
}

/** Human-readable Title for a DE_Setup row, so the list stays legible when browsed directly in SharePoint instead of showing the default blank/"Item" title. */
export function titleForSetupRow(category: string, settingKey: string | undefined): string {
    switch (category) {
        case "Admin":
            return "Admin Member";
        case "Finance":
            return "Finance Member";
        case "ApprovalFlow":
            return "Custom Approval Flow";
        case "AzureADGroupID":
            return "Employee Group ID";
        case "Setting":
            if (settingKey?.toLowerCase() === "approvalflow") return "Standard Approval Flow";
            if (settingKey?.toLowerCase() === "gm") return "General Manager";
            return "Setting";
        default:
            return category;
    }
}

// Reserved chain-entry tokens, resolved at request-creation time rather than
// stored as a literal email — {manager} is always the requester's own direct
// manager (looked up dynamically per request), {GM} is the constant HR sets
// once in DE_Setup (Setting/GM). Mirrors the sibling leave app's convention.
export const MANAGER_TOKEN = "{manager}";
export const GM_TOKEN = "{GM}";

export function getGmEmail(rows: SetupRow[]): string | undefined {
    return getSetupValue(rows, "Setting", "GM");
}

/**
 * Resolves {manager}/{GM} tokens in a chain into real emails — {manager} via
 * a live `/me/manager` lookup (the signed-in requester's own manager, so no
 * extra permission beyond the already-consented User.Read), {GM} via the
 * DE_Setup constant. Unresolvable entries (e.g. no manager set in Entra ID)
 * are dropped rather than failing the whole request.
 */
export async function resolveChainTokens(graphClient: Client, setupRows: SetupRow[], chain: string[]): Promise<string[]> {
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
    return resolved;
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
    return rows.some((r) => r.category === "ApprovalFlow" && r.settingKey?.toLowerCase() === employeeCode?.toLowerCase());
}

/** This employee's override chain if set, otherwise falls back to the global standard chain. */
export function getApprovalChainFor(rows: SetupRow[], employeeCode: string): string[] {
    const overrideRow = rows.find((r) => r.category === "ApprovalFlow" && r.settingKey?.toLowerCase() === employeeCode?.toLowerCase());
    return overrideRow ? parseChainValue(overrideRow.settingValue) : parseApprovalChain(rows);
}
