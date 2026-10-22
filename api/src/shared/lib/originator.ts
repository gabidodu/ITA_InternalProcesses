/**
 * The Actionable Messages originator id to stamp on a module's cards. Normally
 * one originator (ACTIONABLE_MESSAGE_ORIGINATOR_ID) covers both modules; a
 * module can still get its own via LEAVE_ACTIONABLE_MESSAGE_ORIGINATOR_ID /
 * EXPENSE_ACTIONABLE_MESSAGE_ORIGINATOR_ID if the two are ever split again.
 */
export function getOriginatorId(module: "leave" | "expense"): string | undefined {
    const perModule = module === "leave" ? process.env.LEAVE_ACTIONABLE_MESSAGE_ORIGINATOR_ID : process.env.EXPENSE_ACTIONABLE_MESSAGE_ORIGINATOR_ID;
    return perModule || process.env.ACTIONABLE_MESSAGE_ORIGINATOR_ID;
}
