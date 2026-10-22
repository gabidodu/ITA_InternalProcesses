import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { extractSsoToken, getGraphTokenOnBehalfOf } from "../../shared/auth/obo";
import { getTenantConfig } from "../../shared/lib/tenantConfig";
import { getApprovalChainFor, getSetupRows, resolveChainTokens } from "../lib/setup";
import { sendLeaveActivityNotification } from "../lib/activityNotification";
import { sendApprovalActionEmail } from "../lib/emailNotification";
import { formatDate } from "../../shared/lib/dates";
import { adjustTakenDays } from "../lib/leaveStatus";
import { getActionTypeCodeById } from "../lib/actionTypes";
import { getTeamsDeepLink } from "../../shared/lib/teamsLink";
import { createPendingLeaveCalendarEvent } from "../lib/calendarEvent";
import { buildApprovalCardScriptTag } from "../lib/actionableMessage";
import { getAllListItems } from "../../shared/lib/graphPaging";
import { isBlockingStatus } from "../lib/requestStatus";

// Only these leave types have a settable maximum (LR_LeaveStatus.AllocDays) —
// Medical/Unpaid/etc. are deliberately uncapped, per an earlier explicit
// decision, so their balance is never checked before creating a request.
const CAPPED_ACTION_TYPE_CODES = ["CO", "CB"];

type CreateLeaveRequestBody = {
    requestType: number;
    requestSubject: string;
    startDate: string;
    endDate: string;
    comments?: string;
    standInEmployee?: string;
};

function rangesOverlap(startA: Date, endA: Date, startB: Date, endB: Date): boolean {
    return startA <= endB && startB <= endA;
}

function countBusinessDays(start: Date, end: Date): number {
    let count = 0;
    const day = new Date(start);
    while (day <= end) {
        const weekday = day.getDay();
        if (weekday !== 0 && weekday !== 6) {
            count++;
        }
        day.setDate(day.getDate() + 1);
    }
    return count;
}

function parseAndValidateBody(raw: unknown): CreateLeaveRequestBody {
    const body = raw as Partial<CreateLeaveRequestBody> | null;
    if (!body || typeof body.requestType !== "number") {
        throw new Error("requestType is required and must be a number.");
    }
    if (!body.requestSubject || typeof body.requestSubject !== "string") {
        throw new Error("requestSubject is required.");
    }
    if (!body.startDate || !body.endDate) {
        throw new Error("startDate and endDate are required.");
    }

    const startDate = new Date(body.startDate);
    const endDate = new Date(body.endDate);
    if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime())) {
        throw new Error("startDate / endDate must be valid dates.");
    }
    if (startDate > endDate) {
        throw new Error("startDate must not be after endDate.");
    }

    return {
        requestType: body.requestType,
        requestSubject: body.requestSubject,
        startDate: body.startDate,
        endDate: body.endDate,
        comments: body.comments,
        standInEmployee: body.standInEmployee?.trim() || undefined,
    };
}

/**
 * Creates a new leave request in the client's SharePoint site (LR_LeaveRequests
 * list). The approval chain is: optional stand-in first, then the standard
 * chain HR configured in LR_Setup (Setting/approvalFlow). CurrentUser is set
 * to the first approver, the rest is kept (semicolon-joined) in ApprovalFLow
 * and consumed step by step by RespondToLeaveRequest. For Holiday/Bonus (the
 * two capped leave types), rejects up front if not enough days remain, then
 * reserves the days on LR_LeaveStatus immediately — not at approval time —
 * so multiple pending requests correctly compound against the same balance.
 */
export async function CreateLeaveRequest(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const tabBaseUrl = process.env.TAB_BASE_URL;
        const teamsAppId = process.env.TEAMS_APP_ID;
        if (!tabBaseUrl || !teamsAppId) {
            throw new Error("TAB_BASE_URL / TEAMS_APP_ID app settings are not fully configured.");
        }
        const teamsDeepLink = getTeamsDeepLink(teamsAppId, tabBaseUrl, "leave");

        const body = parseAndValidateBody(await request.json());

        const ssoToken = extractSsoToken(request.headers.get("x-teams-sso-token"));
        const { accessToken: graphToken, tenantId } = await getGraphTokenOnBehalfOf(ssoToken, [
            "Sites.Selected",
            "TeamsActivity.Send",
            "User.Read",
        ]);
        const { siteId, leaveRequestsListId, leaveStatusListId, actionTypeListId, setupListId } =
            await getTenantConfig(tenantId);

        const graphClient = Client.init({
            authProvider: (done) => done(null, graphToken),
        });

        const me = await graphClient.api("/me").select("displayName,mail,userPrincipalName").get();
        const currentUserPrincipal: string = me.mail ?? me.userPrincipalName;
        const currentUserName: string = me.displayName ?? currentUserPrincipal;

        const setupRows = await getSetupRows(graphClient, siteId, setupListId);
        const employeeChain = getApprovalChainFor(setupRows, currentUserPrincipal);
        const rawChain = body.standInEmployee ? [body.standInEmployee, ...employeeChain] : employeeChain;
        const fullChain = await resolveChainTokens(graphClient, setupRows, rawChain);

        if (fullChain.length === 0) {
            throw new Error("No approval chain is configured. Please contact HR.");
        }

        const [firstApprover, ...remainingChain] = fullChain;

        const startDate = new Date(body.startDate);
        const endDate = new Date(body.endDate);
        const takenDays = countBusinessDays(startDate, endDate);

        const existingRequestItems = await getAllListItems(graphClient, siteId, leaveRequestsListId);

        const conflict = existingRequestItems.find((item: { fields: Record<string, string> }) => {
            const f = item.fields;
            if (f.EmployeeCode?.toLowerCase() !== currentUserPrincipal?.toLowerCase()) {
                return false;
            }
            if (!isBlockingStatus(f.Status) || !f.StartDate || !f.EndDate) {
                return false;
            }
            return rangesOverlap(startDate, endDate, new Date(f.StartDate), new Date(f.EndDate));
        });

        if (conflict) {
            const cf = conflict.fields;
            throw new Error(
                `You already have a request that overlaps these dates: "${cf.RequestSubject}" ` +
                    `(${formatDate(cf.StartDate)} – ${formatDate(cf.EndDate)}, ${cf.Status}).`
            );
        }

        // Holiday/Bonus have a settable maximum — reject up front rather than let the
        // requester go into debt. Medical/Unpaid/etc. are deliberately uncapped (per an
        // earlier decision) and skip this check entirely.
        const actionTypeCode = await getActionTypeCodeById(graphClient, siteId, actionTypeListId, body.requestType);
        if (actionTypeCode && CAPPED_ACTION_TYPE_CODES.includes(actionTypeCode)) {
            const statusItems = await getAllListItems(graphClient, siteId, leaveStatusListId);
            const balanceRow = statusItems.find(
                (item: { fields: Record<string, string> }) =>
                    item.fields.EmployeeCode?.toLowerCase() === currentUserPrincipal?.toLowerCase() &&
                    Number(item.fields.RequestType) === body.requestType
            );
            const remaining = balanceRow
                ? (Number(balanceRow.fields.AllocDays) || 0) - (Number(balanceRow.fields.TakenDays) || 0)
                : 0;
            if (takenDays > remaining) {
                throw new Error(`You only have ${remaining} day(s) remaining for this leave type, but requested ${takenDays}.`);
            }
        }

        let calendarEventId = "";
        try {
            calendarEventId = await createPendingLeaveCalendarEvent(tenantId, currentUserName, startDate, endDate);
        } catch (calendarError) {
            context.warn("createPendingLeaveCalendarEvent failed (non-blocking)", calendarError);
        }

        const created = await graphClient
            .api(`/sites/${siteId}/lists/${leaveRequestsListId}/items`)
            .post({
                fields: {
                    RequestDate: new Date().toISOString(),
                    RequestType: body.requestType,
                    RequestSubject: body.requestSubject,
                    StartDate: startDate.toISOString(),
                    EndDate: endDate.toISOString(),
                    TakenDays: takenDays,
                    Status: "PENDING",
                    Comments: body.comments ?? "",
                    EmployeeCode: currentUserPrincipal,
                    CurrentUser: firstApprover,
                    StandInEmployee: body.standInEmployee ?? "",
                    ApprovalFLow: remainingChain.join(";"),
                    Year: startDate.getFullYear(),
                    CalendarEventId: calendarEventId,
                },
            });

        // Days are reserved the moment the request is created (not at approval) — this is
        // what makes the balance check above meaningful across multiple pending requests.
        // Released back on decline (RespondToLeaveRequest) or cancellation (CancelLeaveRequest).
        await adjustTakenDays(graphClient, tenantId, siteId, leaveStatusListId, currentUserPrincipal, body.requestType, takenDays);

        const notifyMessage = `${currentUserName} needs your approval for a leave request.`;

        try {
            await sendLeaveActivityNotification(graphClient, firstApprover, notifyMessage, tabBaseUrl);
        } catch (notificationError) {
            context.warn("sendLeaveActivityNotification failed (non-blocking)", notificationError);
        }

        let emailSent = false;
        try {
            const cardScriptTag = buildApprovalCardScriptTag({
                requesterName: currentUserName,
                requestSubject: body.requestSubject,
                startDateLabel: formatDate(startDate),
                endDateLabel: formatDate(endDate),
                takenDays,
                requestId: created.id,
                tabBaseUrl,
                teamsDeepLink,
            });
            const notifyDetails = [
                { label: "Period", value: `${formatDate(startDate)} – ${formatDate(endDate)}` },
                { label: "Days", value: String(takenDays) },
                { label: "Requested by", value: currentUserName },
            ];
            await sendApprovalActionEmail(
                tenantId,
                firstApprover,
                `Leave Request - ${currentUserName}`,
                notifyMessage,
                teamsDeepLink,
                cardScriptTag,
                notifyDetails
            );
            emailSent = true;
        } catch (emailError) {
            context.warn("sendApprovalActionEmail failed (non-blocking)", emailError);
        }

        return {
            status: 201,
            jsonBody: {
                id: created.id,
                subject: body.requestSubject,
                startDate: startDate.toISOString(),
                endDate: endDate.toISOString(),
                takenDays,
                status: "PENDING",
                currentUser: firstApprover,
                emailSent,
            },
        };
    } catch (error) {
        context.error("CreateLeaveRequest failed", error);
        return {
            status: 400,
            jsonBody: { error: error instanceof Error ? error.message : "Unknown error" },
        };
    }
}

app.http("Leave_CreateLeaveRequest", {
    route: "leave/CreateLeaveRequest",
    methods: ["POST"],
    authLevel: "anonymous",
    handler: CreateLeaveRequest,
});
