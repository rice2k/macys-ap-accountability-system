# Detailed Features

This page expands the GitHub front page into a feature catalog for managers, admins, AP operators, and future maintainers.

## Roles

| Role | Main Use |
| --- | --- |
| Employee | Basic dashboard visibility. |
| Front Desk | Dashboard, checkout/return workflow, and search. |
| AP Operator | Front Desk plus people/assets view, search, and reports. |
| Manager | People/assets editing, Manager page, reports, system tools, alerts, logs, and backups. |
| Admin | Full access including admin tools, protected system actions, and Scan-Only Front Desk setting. |

## Dashboard

- Live count cards grouped around operations, risk, and system status.
- Cards for all out, keys out, radios out, tablets out, temp badges out, late returns, active alerts, errors today, backup status, available assets, and issue assets.
- Click-through detail popups.
- Excel export from detail popups.
- Context menus for copy, refresh, related records, and related logs.
- Compact health badges for database, backup, auto-refresh, late returns, and errors.

## Front Desk

![Front Desk Scan-Only](screenshots/front-desk-scan-only.png)

- Scanner/mouse/keyboard friendly checkout and return.
- Admin-controlled Scan-Only Front Desk Mode.
- Header shows Scan-Only Front Desk when no operator login is required.
- One scan box handles employee badges, asset barcodes, controlled key numbers, serials, and device tags.
- Checkout flow: scan employee badge, scan item, pick due time/condition/notes, check out.
- Scan-only checkout flow: employee badge plus item scan, no operator login needed.
- Return flow: scan return-by badge and checked-out item.
- Any scanned return-by badge can return the item.
- Different-badge returns are recorded in notes/audit without blocking.
- Selected Employee, Selected Asset, and Return By cards.
- Large next-step banner.
- Checkout Mode / Return Mode / Scan-Only Mode label.
- Quick Scan buttons: New Checkout, Return Item, Next Item Same Employee, Quick Guide.
- Due presets and More custom due time.
- Function keys: F2 checkout, F3 return, F4 next item same employee, F8 check out, F9 return, Esc start over.

## Users

- Search, filter, view, and edit people.
- Import users from CSV template.
- Assign database-backed groups/roles.
- Add or review AP alerts on people.
- View person audit history and checkout history.
- Deactivate/reactivate support.
- Context menus for profile, AP alerts, audit history, and row copying.

## Assets

![Asset entry details](screenshots/asset-entry-details.png)

- Asset types: Key, Radio, Temp Badge, Scanner, Tablet, Item, Other.
- Search by barcode, name, key number, serial, device tag, location, holder, status, and notes.
- Filter by type and status.
- Sort by type, name, barcode, status, location, key number, radio serial, or holder.
- Add, edit, duplicate, retire, export selected, export all, and export by type.
- Status tools for Available, Repair, Missing, and Retired.
- Asset profile with history, print/export, AP alert actions, and audit links.
- Type-specific asset entry prevents irrelevant fields from showing or being saved.

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

## Manager

![Manager priorities](screenshots/manager-priorities.png)

- Today's Priorities cards for Backup, Late Returns, Alerts, and Open Assets.
- Priority colors for normal, warning, critical, and informational states.
- Only Show Problems toggle.
- Backup status band: date, time, status, refresh countdown, and manual refresh.
- Open Assets and Issue Assets tables.
- Manager notification review area with filters and note/resolve workflows.
- AP alert review and export tools.
- Shortcuts for reports, closeout, settings, log review, audit/error review, backup history, and export history.

## Reports

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

## Logs

- Combined log viewer for audit, errors, notifications, and AP alerts.
- Filter by type, user/source, asset/target, action, date range, and text.
- Export selected rows or filtered views to Excel.
- Archive old logs before clearing.
- Open log folder from System/Logs workflows.
- Rich audit/error fields include date, 12-hour time, 24-hour time, operator role, active page, status, and computer name.

## System

- Data location display and shared database setup.
- Backup/restore.
- Import templates.
- App settings.
- Admin-only Scan-Only Front Desk setting.
- Scanner diagnostics.
- Role test.
- Health check.
- Self-test info.
- Export system report.
- Admin tools.
- Training Mode demo reset for sample users/assets.

## Groups And Permissions

- Default groups: Employee, Front Desk, AP Operator, Manager, Admin.
- Default groups are protected.
- Permissions are database-backed.
- Managers/Admins can add groups, save rights, rename non-protected groups, delete non-protected groups with a reason, view assigned users, and assign users to groups.
- Rights are grouped by Access, People, Assets, and Admin.
- Group changes create audit entries and manager notifications.

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

## Safety And Accountability

- Automatic daily backup is enabled by default.
- Manual backup is available from Manager and System.
- Restore backs up the current database before replacing it.
- Manager notifications are generated for sensitive actions.
- Permission changes require a reason.
- Group deletion requires a reason and replacement group selection when users are assigned.
- Blocked/restricted checkout attempts are logged and notify managers.
- Different-badge returns are accepted and recorded.
- Log cleanup requires Admin access, confirmation, and a reason.
