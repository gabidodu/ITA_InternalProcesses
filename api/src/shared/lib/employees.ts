import { Client } from "@microsoft/microsoft-graph-client";
import { getSetupValue, SetupRow } from "./setupBase";

export type Employee = {
    mail: string;
    displayName: string;
};

/**
 * Returns the members of the "all employees" Entra ID group (id configured in
 * the module's Setup list, AzureADGroupID/allEmployeesGroupID) — the one directory source
 * this app can read without extra permissions, reused both for people-pickers
 * and for resolving an email to a display name.
 */
export async function getEmployeeDirectory(graphClient: Client, setupRows: SetupRow[]): Promise<Employee[]> {
    const groupId = getSetupValue(setupRows, "AzureADGroupID", "allEmployeesGroupID");
    if (!groupId) {
        return [];
    }

    // Graph's untyped /members (directoryObject) collection silently returns
    // null for $select-ed properties beyond a small default set — casting to
    // the concrete /microsoft.graph.user subtype is required to actually get
    // displayName/mail/userPrincipalName back. Confirmed in production: same
    // call without the cast returned rawMemberCount > 0 but every field null.
    const membersResponse = await graphClient
        .api(`/groups/${groupId}/members/microsoft.graph.user`)
        .select("displayName,mail,userPrincipalName")
        .get();

    return membersResponse.value
        .map((u: { displayName: string; mail: string | null; userPrincipalName: string }) => ({
            displayName: u.displayName,
            mail: u.mail ?? u.userPrincipalName,
        }))
        .filter((e: Employee) => e.mail);
}

export function buildDisplayNameMap(employees: Employee[]): Map<string, string> {
    const map = new Map<string, string>();
    employees.forEach((e) => map.set(e.mail.toLowerCase(), e.displayName));
    return map;
}

/**
 * Whether `viewerPrincipal` is `employeeCode`'s direct manager, per Entra ID
 * (live lookup, not cached) — used to let a manager view their own directs'
 * leave reports without needing the broader HR/Admin/GM access. Returns
 * false (not throws) if the employee has no manager set, or the lookup
 * otherwise fails — absence of a manager relationship is not an error here.
 */
export async function isDirectManagerOf(graphClient: Client, viewerPrincipal: string, employeeCode: string): Promise<boolean> {
    try {
        const manager = await graphClient
            .api(`/users/${encodeURIComponent(employeeCode)}/manager`)
            .select("mail,userPrincipalName")
            .get();
        const managerEmail: string | undefined = manager.mail ?? manager.userPrincipalName;
        return managerEmail?.toLowerCase() === viewerPrincipal?.toLowerCase();
    } catch {
        return false;
    }
}

/** The signed-in user's own direct reports, per Entra ID — used to gate/scope the Reports tab for plain managers. */
export async function getMyDirectReports(graphClient: Client): Promise<Employee[]> {
    const response = await graphClient.api("/me/directReports").select("displayName,mail,userPrincipalName").get();
    return response.value
        .map((u: { displayName: string; mail: string | null; userPrincipalName: string }) => ({
            displayName: u.displayName,
            mail: u.mail ?? u.userPrincipalName,
        }))
        .filter((e: Employee) => e.mail);
}
