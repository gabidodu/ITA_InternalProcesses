import { Client } from "@microsoft/microsoft-graph-client";

/**
 * Fetches every item of a SharePoint list via Graph, following @odata.nextLink
 * until exhausted. Graph pages list-items responses (~200/page) regardless of
 * list size — every single-shot `.items().expand("fields").get()` call in this
 * codebase silently returned only the first page. Harmless on the small dev
 * lists used throughout development, but real production lists (1000+ items,
 * e.g. LR_LeaveRequests) return only their oldest page without this.
 */
export async function getAllListItems<T = { id: string; fields: Record<string, string> }>(
    graphClient: Client,
    siteId: string,
    listId: string
): Promise<T[]> {
    const items: T[] = [];
    let response = await graphClient.api(`/sites/${siteId}/lists/${listId}/items`).expand("fields").get();
    items.push(...response.value);
    while (response["@odata.nextLink"]) {
        response = await graphClient.api(response["@odata.nextLink"]).get();
        items.push(...response.value);
    }
    return items;
}
