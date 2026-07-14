# Security Notes

This app handles AP operations, employee/user records, asset history, alerts, logs, reports, and backup data. Because this repository is public, commit documentation, approved source files, sanitized visuals, and sample/template files only.

## Keep Out Of GitHub

- Live SQLite databases.
- Backup ZIP files containing live data.
- Exported Excel workbooks with live data.
- Exported reports with live data.
- Audit/error/system logs from live operation.
- Real employee records.
- Real badge numbers.
- AP alert notes.
- Incident notes.
- Store-specific paths.
- Network paths.
- Screenshots with live data.

## Public Repository Safety

- Use sample data in documentation and images.
- Review screenshots before upload.
- Do not include production database files in issues.
- Do not attach backup/export/log files to issues.
- Keep release ZIPs public only when they contain approved source/template files and no live data.

## App Controls

The app includes role-based access, database-backed groups, protected Admin/Manager actions, audit logs, manager notifications, backup/restore, and log archival before cleanup.

Admin can enable Scan-Only Front Desk Mode. When enabled, checkout/return can run without operator login, but the app still records the scanned employee/return-by badge and labels the action as Scan-Only Front Desk in notes/audit.

## Reporting Security Issues

Open a GitHub issue using the bug template, or contact the app owner/administrator directly. Do not attach live data. Include steps to reproduce, expected behavior, actual behavior, and whether live data could be exposed.
