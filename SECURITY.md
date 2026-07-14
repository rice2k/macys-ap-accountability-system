# Security Notes

This app handles AP operations, employee/user records, asset history, alerts, logs, reports, and backup data. If this repository is public, commit documentation and sanitized samples only.

## Sensitive Data

Keep these out of GitHub:

- Live SQLite databases.
- Backup ZIP files.
- Exported Excel workbooks.
- Exported reports.
- Audit/error/system logs.
- Real employee records.
- Badge numbers.
- AP alert notes.
- Incident notes.
- Store/network paths.

## Public Repository Safety

- Use a public repository only for approved documentation and approved source files.
- Keep live operational data outside GitHub.
- Review screenshots before upload.
- Use sample data in documentation.
- Do not include production database files in issues.

## App Controls

The app includes role-based access, database-backed groups, manager approval for high-risk actions, audit logs, manager notifications, backup/restore, and log archival before cleanup. These controls depend on operators signing in with the correct role and on the database being protected at the file/folder level.

## Reporting Security Issues

Open a GitHub issue using the bug template, or contact the app owner/administrator directly. Do not attach live data. Include steps to reproduce, expected behavior, actual behavior, and whether live data could be exposed.
