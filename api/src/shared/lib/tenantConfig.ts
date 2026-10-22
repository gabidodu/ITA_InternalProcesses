import { TableClient } from "@azure/data-tables";

/**
 * Per-client SharePoint wiring — resolved at request time from the caller's
 * own tenant id (`tid`) rather than from fixed app settings, so one shared
 * deployment of this app can serve multiple client tenants, each keeping
 * their data in their own SharePoint site (never centralized). One row per
 * tenant in the `TenantConfigs` table (PartitionKey = RowKey = tenantId),
 * written once per client by `scripts/leave/Register-ClientTenant.ps1` during
 * onboarding.
 *
 * One row carries the wiring for every module (leave, expense, ...). A
 * module is enabled for a tenant exactly when all of its required columns
 * are filled in — see getTenantModules, which drives which tiles the home
 * screen shows. A tenant that only bought leave simply has no expense
 * columns, and never sees the expense module.
 */
export type TenantConfig = {
    tenantId: string;
    siteId: string;
    leaveRequestsListId: string;
    leaveStatusListId: string;
    leaveSetupListId: string;
    actionTypeListId: string;
    setupListId: string;
    auditLogListId: string;
    holidaySetupListId: string;
    /** Shared mailbox this tenant's notification emails/calendar events send from — must be a real mailbox inside THIS tenant, since an app-only token minted for this tenant can't send as a mailbox in a different one. Shared by every module. */
    mailSenderAddress: string;
    /** LR_Approvals document library (per-employee subfolders of approved-request PDFs — see leave/lib/approvalArchive.ts). Optional and NOT in LEAVE_REQUIRED_FIELDS: tenants registered before this field existed keep working, the archive step is just skipped for them until re-provisioned (New-LeaveApprovalsLibrary.ps1 + a Register-ClientTenant.ps1 re-run). */
    approvalsLibraryId?: string;
    /** LR_LeaveRequests(Archive) — pre-existing historical requests some tenants (e.g. IT ASSIST's own, migrated from the legacy site) split into a separate list before this app existed. Optional: only set for tenants that actually have one. Read-only, and only merged into GetEmployeeLeaveReport's history/yearly totals — never into live balance checks (CreateLeaveRequest overlap, GetMyLeaveContext), since those only ever cared about currently-open requests. */
    leaveRequestsArchiveListId?: string;
};

/** The expense module's DE_* lists. Lives on the same site as the leave lists unless `expenseSiteId` says otherwise. */
export type ExpenseConfig = {
    tenantId: string;
    siteId: string;
    expenseReportsListId: string;
    expenseLinesListId: string;
    expenseCategoryListId: string;
    expenseAttachmentsListId: string;
    setupListId: string;
    auditLogListId: string;
    mailSenderAddress: string;
};

export type ModuleKey = "leave" | "expense";

const TABLE_NAME = "TenantConfigs";
const CACHE_TTL_MS = 5 * 60 * 1000;

let cachedClient: TableClient | undefined;
const cache = new Map<string, { value: Record<string, unknown>; expiresAt: number }>();

function getTableClient(): TableClient {
    if (cachedClient) {
        return cachedClient;
    }
    const connectionString = process.env.TENANT_CONFIG_STORAGE_CONNECTION_STRING;
    if (!connectionString) {
        throw new Error("TENANT_CONFIG_STORAGE_CONNECTION_STRING is not configured.");
    }
    cachedClient = TableClient.fromConnectionString(connectionString, TABLE_NAME);
    return cachedClient;
}

const LEAVE_REQUIRED_FIELDS = [
    "siteId",
    "leaveRequestsListId",
    "leaveStatusListId",
    "leaveSetupListId",
    "actionTypeListId",
    "setupListId",
    "auditLogListId",
    "holidaySetupListId",
    "mailSenderAddress",
] as const;

// expenseSiteId is optional (falls back to siteId), so it isn't listed here.
const EXPENSE_REQUIRED_FIELDS = [
    "expenseReportsListId",
    "expenseLinesListId",
    "expenseCategoryListId",
    "expenseAttachmentsListId",
    "expenseSetupListId",
    "expenseAuditLogListId",
    "mailSenderAddress",
] as const;

/**
 * Raw TenantConfigs row for a tenant, cached for CACHE_TTL_MS. Throws a clear
 * error (surfaced to the caller as a 4xx, same pattern as every other
 * validation error in this app) if the tenant hasn't been provisioned yet —
 * never falls back to another tenant's data.
 */
async function getTenantEntity(tenantId: string): Promise<Record<string, unknown>> {
    const cached = cache.get(tenantId);
    if (cached && cached.expiresAt > Date.now()) {
        return cached.value;
    }

    const client = getTableClient();
    let entity: Record<string, unknown>;
    try {
        entity = await client.getEntity(tenantId, tenantId);
    } catch (error) {
        if ((error as { statusCode?: number })?.statusCode === 404) {
            throw new Error(
                `No tenant configuration found for tenant "${tenantId}". This client hasn't been provisioned yet — contact IT ASSIST.`
            );
        }
        throw error;
    }

    cache.set(tenantId, { value: entity, expiresAt: Date.now() + CACHE_TTL_MS });
    return entity;
}

function missingFields(entity: Record<string, unknown>, fields: readonly string[]): string[] {
    return fields.filter((field) => !entity[field]);
}

/** Leave module wiring (the LR_* lists). Throws if the tenant isn't provisioned for leave. */
export async function getTenantConfig(tenantId: string): Promise<TenantConfig> {
    const entity = await getTenantEntity(tenantId);

    const missing = missingFields(entity, LEAVE_REQUIRED_FIELDS);
    if (missing.length > 0) {
        throw new Error(`Leave configuration for tenant "${tenantId}" is incomplete — missing: ${missing.join(", ")}.`);
    }

    return {
        tenantId,
        siteId: entity.siteId as string,
        leaveRequestsListId: entity.leaveRequestsListId as string,
        leaveStatusListId: entity.leaveStatusListId as string,
        leaveSetupListId: entity.leaveSetupListId as string,
        actionTypeListId: entity.actionTypeListId as string,
        setupListId: entity.setupListId as string,
        auditLogListId: entity.auditLogListId as string,
        holidaySetupListId: entity.holidaySetupListId as string,
        mailSenderAddress: entity.mailSenderAddress as string,
        approvalsLibraryId: (entity.approvalsLibraryId as string) || undefined,
        leaveRequestsArchiveListId: (entity.leaveRequestsArchiveListId as string) || undefined,
    };
}

/** Expense module wiring (the DE_* lists). Throws if the tenant isn't provisioned for expense. */
export async function getExpenseConfig(tenantId: string): Promise<ExpenseConfig> {
    const entity = await getTenantEntity(tenantId);

    const siteId = (entity.expenseSiteId as string) || (entity.siteId as string);
    const missing = missingFields(entity, EXPENSE_REQUIRED_FIELDS);
    if (!siteId) missing.unshift("expenseSiteId/siteId");
    if (missing.length > 0) {
        throw new Error(`Expense configuration for tenant "${tenantId}" is incomplete — missing: ${missing.join(", ")}.`);
    }

    return {
        tenantId,
        siteId,
        expenseReportsListId: entity.expenseReportsListId as string,
        expenseLinesListId: entity.expenseLinesListId as string,
        expenseCategoryListId: entity.expenseCategoryListId as string,
        expenseAttachmentsListId: entity.expenseAttachmentsListId as string,
        setupListId: entity.expenseSetupListId as string,
        auditLogListId: entity.expenseAuditLogListId as string,
        mailSenderAddress: entity.mailSenderAddress as string,
    };
}

/** Which modules this tenant is provisioned for — drives the home screen's tiles. Never throws for a half-configured module, it's just left out. */
export async function getTenantModules(tenantId: string): Promise<ModuleKey[]> {
    const entity = await getTenantEntity(tenantId);
    const modules: ModuleKey[] = [];
    if (missingFields(entity, LEAVE_REQUIRED_FIELDS).length === 0) modules.push("leave");
    if ((entity.expenseSiteId || entity.siteId) && missingFields(entity, EXPENSE_REQUIRED_FIELDS).length === 0) modules.push("expense");
    return modules;
}
