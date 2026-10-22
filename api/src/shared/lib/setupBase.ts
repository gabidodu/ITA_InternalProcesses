import { Client } from "@microsoft/microsoft-graph-client";
import { getAllListItems } from "../../shared/lib/graphPaging";

export type SetupRow = {
    settingKey: string;
    settingValue: string;
    category: string;
};

/**
 * Reads a module's Setup key/value list (LR_Setup / DE_Setup) (SettingKey/SettingValue/Category), the
 * app's generic configuration store — same shape in every module: the legacy
 * SetupList convention (allEmployeesGroupID, Admin role rows, approvalFlow, ...)
 * plus whatever role categories a module adds (HR/GM for leave, Finance for expense).
 */
export async function getSetupRows(graphClient: Client, siteId: string, setupListId: string): Promise<SetupRow[]> {
    const items = await getAllListItems(graphClient, siteId, setupListId);

    return items.map((item: { fields: Record<string, string> }) => ({
        settingKey: item.fields.SettingKey,
        settingValue: item.fields.SettingValue,
        category: item.fields.Category,
    }));
}

export function getSetupValue(rows: SetupRow[], category: string, settingKey: string): string | undefined {
    return rows.find((r) => r.category === category && r.settingKey === settingKey)?.settingValue;
}

export function isHrAdmin(rows: SetupRow[], userPrincipal: string): boolean {
    return rows.some(
        (r) => r.category === "Admin" && r.settingKey?.toLowerCase() === userPrincipal?.toLowerCase()
    );
}
