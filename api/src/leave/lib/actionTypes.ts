import { Client } from "@microsoft/microsoft-graph-client";
import { getAllListItems } from "../../shared/lib/graphPaging";

/** Looks up an LR_ActionType row's IDActionType by its short code (e.g. "CO" for Holiday, "CB" for Bonus Days). */
export async function getActionTypeIdByCode(
    graphClient: Client,
    siteId: string,
    actionTypeListId: string,
    code: string
): Promise<number | undefined> {
    const items = await getAllListItems(graphClient, siteId, actionTypeListId);

    const row = items.find((item: { fields: Record<string, string> }) => item.fields.ActionTypeCode === code);
    return row ? Number(row.fields.IDActionType) : undefined;
}

/** Looks up an LR_ActionType row's short code (e.g. "CO", "CB") by its IDActionType. */
export async function getActionTypeCodeById(
    graphClient: Client,
    siteId: string,
    actionTypeListId: string,
    id: number
): Promise<string | undefined> {
    const items = await getAllListItems(graphClient, siteId, actionTypeListId);

    const row = items.find((item: { fields: Record<string, string> }) => Number(item.fields.IDActionType) === id);
    return row?.fields.ActionTypeCode;
}
