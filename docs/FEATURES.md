# Detailed Features

This page expands the GitHub front page into a feature catalog for managers, admins, AP operators, and future maintainers.

## Roles

| Role | Main Use |
| --- | --- |
| Employee | Basic dashboard visibility. |
| Front Desk | Dashboard, checkout/return workflow, and search. |
| AP Operator | Front Desk plus people/assets view, search, and reports. |
| Manager | People/assets editing, Manager page, reports, system tools, alerts, logs, and backups. |
| Admin | Full access including admin tools and protected system actions. |

## Pages

### Dashboard

- Live count cards grouped around operations, risk, and system status.
- Click-through details for out assets, keys, radios, tablets, temp badges, late returns, alerts, errors, backup status, available assets, and issue assets.
- Export visible dashboard detail popups to Excel.
- Context menu actions for opening related asset/user records and audit logs.

### Front Desk

- Operator sign-in context is shown clearly.
- Guided flow keeps checkout steps visible.
- Employee scan and asset scan support keyboard/scanner input.
- Due time, condition, and notes are captured at checkout.
- Return flow handles normal returns and wrong-user returns.
- Manager/Admin blocks protect restricted checkout conditions.

### Users

- Search, filter, view, and edit people.
- Import users from CSV template.
- Assign database-backed groups/roles.
- Add or review AP alerts on people.
- View person audit history and checkout history.

### Assets

- Asset search and filtering by text, type, status, sort field, and sort direction.
- Add/Edit Asset form rebuilds its fields by selected asset type.
- Secondary actions are grouped under More/context menus to keep the toolbar clean.
- Asset list uses status coloring and row striping.
- Export selected asset, all assets, or one asset type to Excel.
- Retire assets with history, delete unused assets where allowed.

### Manager

- Manager counts: Open, Late, Issues, Available, Active Alerts, Errors Today.
- Backup status band: date, time, status, refresh countdown, and manual refresh.
- Open Assets and Issue Assets tables.
- Manager notification review area with filters and note/resolve workflows.
- AP alert review and export tools.
- Shortcuts for reports, closeout, settings, log review, audit/error review, backup history, and export history.

### Reports

- Current Out.
- Overdue.
- Asset Issues.
- Employee History.
- Asset History.
- Alert History.
- Group History.
- End of Shift.
- Detailed Audit.
- Detailed Error.
- Save TXT, HTML, and Excel outputs.

### Logs

- Combined log viewer for audit, errors, notifications, and AP alerts.
- Filter by type, user/source, asset/target, action, date range, and text.
- Export selected rows or filtered views to Excel.
- Archive old logs before clearing.
- Open log folder from System/Logs workflows.

### System

- Data location display and shared database setup.
- Backup/restore.
- Import templates.
- App settings.
- Scanner diagnostics.
- Role test.
- Health check.
- Self-test info.
- Export system report.
- Admin tools.

## Asset Type Detail Map

| Asset Type | Common Fields | Type-Specific Fields |
| --- | --- | --- |
| Key | Type, barcode, name, status, location, notes | Controlled key number, key set number, ring serial, number of keys, ring location, ring use, key/access list |
| Radio | Type, barcode, name, status, location, notes | Serial/device tag, radio number, factory serial, radio serial, assigned area |
| Temp Badge | Type, barcode, name, status, location, notes | Badge number or temp badge details |
| Scanner | Type, barcode, name, status, location, notes | Serial/device tag and scanner-specific details |
| Tablet | Type, barcode, name, status, location, notes | Serial/device tag, IMEI/license data, accessories, custom accessory |
| Item | Type, barcode, name, status, location, notes | General tracked item details |
| Other | Type, barcode, name, status, location, notes | General custom asset details |

## Safety And Accountability

- Automatic daily backup is enabled by default.
- Manual backup is available from Manager and System.
- Restore backs up the current database before replacing it.
- Manager notifications are generated for sensitive actions.
- Permission changes require a reason.
- Group deletion requires a reason and replacement group selection when users are assigned.
- Wrong-user returns require Manager/Admin access and a reason.
- Blocked/restricted checkout attempts are logged and notify managers.
- Log cleanup requires Admin access, confirmation, and a reason.

## Export Features

- Formatted Excel asset workbook by all assets or asset type.
- Dashboard detail Excel export.
- Report Excel export.
- Manager notification Excel export.
- AP alert Excel export.
- Group user Excel export.
- Selected log row Excel export.
- Filtered log Excel export.
- Raw CSV bundle export for people, assets, activity, audit, errors, manager notifications, AP alerts, groups, and settings.

