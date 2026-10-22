# Internal Processes for Teams

One Teams app — leave requests and expense reports — in one Container App.
Merged from `LeaveRequestForTeams` and `ExpenseReportForTeams` (their git
history stays in those repos).

## Layout

```
api/src/
  shared/auth     obo, appOnly, actionableMessageToken (multi-tenant, from leave)
  shared/lib      tenantConfig (one TenantConfigs row per tenant, all modules), teamsLink, graphPaging,
                  auditLog, employees, dates, setupBase (Setup list reader + Admin role) — one copy for both modules
  home/functions  Health, GetHomeSummary (tiles + combined "awaiting your approval")
  leave/          LR_* functions + libs     → /api/leave/*
  expense/        DE_* functions + libs     → /api/expense/*
web/src/app/
  page.tsx               home: request-type picker + combined approvals inbox
  (modules)/layout.tsx   "← All requests" bar
  (modules)/leave        the former leave app page
  (modules)/expense      the former expense app page
scripts/leave, scripts/expense   SharePoint provisioning
appPackage/              Teams manifest (keeps the leave app id 661a64f0…)
```

Still duplicated per module: the role/approval-chain half of `setup.ts`, and
`emailNotification` / `actionableMessage` (they diverged too much to merge
blindly — leave has the stale-click handling, expense doesn't yet).

## Modules per tenant

`TenantConfigs` (storage account `stleaverequestforteams`) has one row per
client tenant. A module is enabled when all its columns are filled:

| Module  | Columns |
|---------|---------|
| leave   | siteId, leaveRequestsListId, leaveStatusListId, leaveSetupListId, actionTypeListId, setupListId, auditLogListId, holidaySetupListId, mailSenderAddress |
| expense | expenseSiteId (optional, defaults to siteId), expenseReportsListId, expenseLinesListId, expenseCategoryListId, expenseAttachmentsListId, expenseSetupListId, expenseAuditLogListId, mailSenderAddress |

`scripts/leave/Register-ClientTenant.ps1` writes the leave columns;
`scripts/expense/Enable-ExpenseModule.ps1` merges in the expense ones.

## Running locally

- **UI only, no Teams:** `cd web && npm run preview` → http://localhost:3100.
  Sample data from `web/src/lib/preview.ts`, no SSO, nothing is written.
- **API:** `cd api && npm run build && func start` with `local.settings.json`
  (copy `local.settings.example.json`). Every real call needs a Teams SSO
  token, so end-to-end testing needs the tab running inside Teams.
- **Container:** `docker compose up --build` builds the production image.

## Before this can go live

1. Create Container App `ca-internal-processes` in `rg-internal-processes`
   (single container, port 8080, image from the root `Dockerfile`).
2. Entra app `359906f4…`: add identifierUri
   `api://ca-internal-processes.politesky-08438219.westeurope.azurecontainerapps.io/359906f4-0bbf-4810-97f4-0d1b28aa5f1a`
   and the `/auth-end` redirect URI on that host.
3. Actionable Messages: neither existing originator accepts new domains, so
   one new originator was registered for the new host (details below); set
   `ACTIONABLE_MESSAGE_ORIGINATOR_ID` = `e1414a49-1a3e-4bba-94b4-e27134577476` and
   `ACTIONABLE_MESSAGE_AUDIENCE` = `api://auth-am-e1414a49-1a3e-4bba-94b4-e27134577476/359906f4-0bbf-4810-97f4-0d1b28aa5f1a`.
4. Other env: `TAB_BASE_URL` = the new host, `TEAMS_APP_ID` =
   `661a64f0-7ed1-413c-a709-4ea98ef49fd7`, `AAD_APP_*`,
   `TENANT_CONFIG_STORAGE_CONNECTION_STRING`, `TZ=Europe/Bucharest`
   (`shared/lib/dates.ts` formats with local getters).
5. Run `Enable-ExpenseModule.ps1` for the IT ASSIST tenant.
6. Upload `appPackage` (version 1.1.0) to replace the Leave Request app;
   retire the Expense Reports app (`bf70ae37…`).
7. Keep the two old Container Apps running a few weeks: emails already sent
   point at their hosts and their originators.

### The new Actionable Messages originator

Provider `e1414a49-1a3e-4bba-94b4-e27134577476`, API ID URI `api://auth-am-e1414a49-1a3e-4bba-94b4-e27134577476/359906f4-0bbf-4810-97f4-0d1b28aa5f1a`.

| Field | Value |
|---|---|
| Sender (from) | the shared mailbox the emails are sent from (`mailSenderAddress`, e.g. `ita_biz@itassist.eu`) |
| Target URLs | `https://ca-internal-processes.politesky-08438219.westeurope.azurecontainerapps.io/api/leave/RespondToLeaveRequestAction` and `…/api/expense/RespondToExpenseReportAction` |
| Scope | Organization (as the existing ones) |

Registering it also creates the Entra resource `api://auth-am-<id>`; the app
`359906f4` must be allowed to receive tokens for it, same as for the old ones.

## Scaling

`infra/scale.json` is the saved policy for all three Container Apps:
`minReplicas = 0`, max 3, an HTTP rule, and a cron rule keeping one replica
running Monday–Friday 08:00–20:00 (Europe/Bucharest); outside that window the
app scales to zero and the first request starts it cold. `scripts/Set-ContainerAppScale.ps1`
checks each app against it (`-Apply` fixes any drift). Weekends are not covered —
change `1-5` in the cron `start`/`end` to `*` for every day.
