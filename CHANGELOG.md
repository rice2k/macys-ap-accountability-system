# Changelog

## v5.2.7 Flow + Visual Polish - July 13, 2026

### Added

- Admin-only Front Desk Scan-Only Mode.
- Scan-only checkout flow: scan employee badge, scan item, complete checkout.
- Scan-only return flow: scan return-by badge, scan checked-out item, complete return.
- Return acceptance by any scanned return-by badge, even when different from the original checkout holder.
- Notes/audit labeling for Scan-Only Front Desk checkout and return.
- Front Desk large next-step banner.
- Front Desk Checkout Mode / Return Mode / Scan-Only Mode label.
- Quick Scan controls for New Checkout, Return Item, Next Item Same Employee, and Quick Guide.
- Function-key shortcuts for scanner stations: F2, F3, F4, F8, F9, Esc.
- Manager Today's Priorities strip for Backup, Late Returns, Alerts, and Open Assets.
- Manager Only Show Problems toggle.
- Printable Quick Scan guide export.
- Training Mode demo reset for sample users/assets.
- Public GitHub visuals and documentation images.
- `app/` source package folder in the public repo.

### Changed

- Return flow no longer blocks when the return-by badge differs from the original checkout holder.
- Different-badge returns are recorded in notes/audit without Manager/Admin approval.
- Front Desk buttons enable only when the current action is ready.
- Front Desk step tracker adapts between checkout and return.
- Manager backup display now emphasizes last backup date/time and refresh countdown.
- Asset entry keeps fields specific to the selected asset type.
- README, How It Works, Data Info, Screenshots, Features, Versions, and Public Upload docs now match the latest workflow.

### Validation

- Self-test result: PASS.
- Syntax check: PASS.
- ZIP rebuilt: `Macys_AP_v5_2_7_Flow_Visual_Polish_20260713.zip`.

## v5.2.7 Asset Entry / Manager / Export Upgrade

### Added

- Manager backup status band with last backup date, last backup time, backup status, refresh countdown, and manual refresh.
- Type-specific asset detail fields.
- Key ring detail tracking.
- Tablet IMEI/license/accessory tracking.
- Manager notifications.
- AP Alerts for people and assets.
- Database-backed groups and protected default roles.
- Dashboard clickable cards and detail popups.
- Excel exports for dashboard details, assets, reports, notifications, alerts, groups, and logs.
- Combined logs page with filtering.
- Backup/export history shortcuts.
- Display density setting.
- Context menus across major tables.

### Changed

- Dashboard focuses on live operational counts and compact health badges.
- System owns data location, backup/restore, imports/templates, settings, diagnostics, and admin support tools.
- Asset entry hides or disables fields that do not apply to the selected asset type.
- CSV asset import uses the same validation as manual entry.
- Exports use saved folders and safer filenames.
- Auto-refresh honors saved refresh timer when configured.
- Left navigation is grouped and simplified.
- Reports are grouped by Daily, History, Export, and Output.
- Tables use consistent row striping, selected-row styling, and scrollbars.

### Safety

- Restricted checkout conditions require Manager/Admin access.
- Critical/blocking AP alerts require Manager/Admin access.
- Group rights changes require a reason.
- Group deletion requires a reason before moving assigned users.
- Old log cleanup archives records before deleting them.
- Restore backs up the active database first.
