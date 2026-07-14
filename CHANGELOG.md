# Changelog

## v5.2.7 - Latest Verified Build

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

### Security And Safety

- Restricted checkout conditions require Manager/Admin access.
- Critical/blocking AP alerts require Manager/Admin access.
- Wrong-user returns require Manager/Admin access and reason entry.
- Group rights changes require a reason.
- Group deletion requires a reason before moving assigned users.
- Old log cleanup archives records before deleting them.
- Restore backs up the active database first.

### Validation

- Self-test result: PASS.
- Validated database creation, default admin, seeded groups/assets, asset mapping, asset details, checkout, double-checkout guard, return, status updates, Excel write, counts, dashboard type counts, manager notifications, AP alerts, audit/error records, rich log fields, and cleanup.

