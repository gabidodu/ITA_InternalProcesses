/**
 * Whether a leave request's reserved days still count against the employee's
 * balance. Days are reserved (LR_LeaveStatus.TakenDays incremented) the
 * instant a request is created, not at final approval, and released back
 * only on decline or cancellation — so PENDING and every mid-chain status
 * still hold the reservation, exactly like APPROVED. Used both to block
 * overlapping requests (CreateLeaveRequest) and to compute "days taken"
 * anywhere it's reported (GetEmployeeLeaveReport, GetMyLeaveContext) —
 * filtering on Status === "APPROVED" alone undercounts real balance usage.
 */
export function isBlockingStatus(status: string | undefined): boolean {
    const s = status?.toUpperCase() ?? "";
    return !s.startsWith("DECLINED") && !s.startsWith("CANCELLED");
}
