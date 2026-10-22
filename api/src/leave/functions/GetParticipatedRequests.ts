import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getSetupRows } from "../lib/setup";
import { buildDisplayNameMap, getEmployeeDirectory } from "../../shared/lib/employees";
import { getRelevantYears } from "../../shared/lib/dates";
import { getAllListItems } from "../../shared/lib/graphPaging";

type ParticipatedRequest = {
    id: string;
    subject: string;
    startDate: string;
    endDate: string;
    takenDays: number;
    status: string;
    currentUser: string;
    currentUserName: string;
    employeeCode: string;
    employeeName: string;
};

/**
 * Returns leave requests the signed-in user has ever acted on as an approver
 * (approved or declined a step, even if the request has since moved on or
 * been resolved) — sourced from LR_AuditLog since the request item itself
 * doesn't keep approval history. Lets a manager see everything they've
 * participated in, not just what's currently awaiting them.
 */
export async function GetParticipatedRequests(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, [
            "Sites.Selected",
            "GroupMember.Read.All",
            "User.ReadBasic.All",
        ]);
        const { siteId, leaveRequestsListId, auditLogListId, setupListId } = await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;

        const auditItems = await getAllListItems(graphClient, siteId, auditLogListId);

        const participatedIds = new Set<string>(
            auditItems
                .filter(
                    (item: { fields: Record<string, string> }) =>
                        item.fields.Actor?.toLowerCase() === currentUserPrincipal?.toLowerCase() &&
                        ["Approved", "ApprovedFinal", "Declined"].includes(item.fields.Action) &&
                        item.fields.RequestId
                )
                .map((item: { fields: Record<string, string> }) => item.fields.RequestId)
        );

        if (participatedIds.size === 0) {
            return { jsonBody: { requests: [] } };
        }

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const employees = await getEmployeeDirectory(graphClient, setupRows);
        const nameByMail = buildDisplayNameMap(employees);
        const resolveName = (mail: string | undefined) => (mail && nameByMail.get(mail.toLowerCase())) || mail || "";

        const leaveRequestItems = await getAllListItems(graphClient, siteId, leaveRequestsListId);

        const relevantYears = getRelevantYears();

        const requests: ParticipatedRequest[] = leaveRequestItems
            .filter(
                (item: { id: string; fields: Record<string, string> }) =>
                    participatedIds.has(item.id) && relevantYears.has(Number(item.fields.Year))
            )
            .map((item: { id: string; fields: Record<string, string> }) => ({
                id: item.id,
                subject: item.fields.RequestSubject,
                startDate: item.fields.StartDate,
                endDate: item.fields.EndDate,
                takenDays: Number(item.fields.TakenDays),
                status: item.fields.Status,
                currentUser: item.fields.CurrentUser,
                currentUserName: resolveName(item.fields.CurrentUser),
                employeeCode: item.fields.EmployeeCode,
                employeeName: resolveName(item.fields.EmployeeCode),
            }))
            .sort((a: ParticipatedRequest, b: ParticipatedRequest) => new Date(b.startDate).getTime() - new Date(a.startDate).getTime());

        return { jsonBody: { requests } };
    } catch (error) {
        context.error("GetParticipatedRequests failed", error);
        return {
            status: 401,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_GetParticipatedRequests", {
    route: "leave/GetParticipatedRequests",
    methods: ["GET"],
    authLevel: "anonymous",
    handler: GetParticipatedRequests,
});
