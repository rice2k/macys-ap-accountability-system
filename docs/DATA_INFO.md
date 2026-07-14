# Data Info

![Data and backup](screenshots/data-and-backup.png)

## Database

The app uses SQLite. It can run from a local database on one workstation or a shared database path for multiple workstations.

Recommended shared path style:

```text
\\SERVER\Share\Macys_AP_Data\macys_ap_data.db
```

## Main Tables

| Table | Purpose |
| --- | --- |
| `people` | Employees, operators, badges, roles, status, department, shift, notes. |
| `assets` | Keys, radios, tablets, scanners, temp badges, items, status, holder, asset details. |
| `activity` | Checkout and return records, due times, conditions, notes, operators. |
| `audit` | User actions and system events. |
| `errors` | App errors and diagnostics. |
| `manager_notifications` | Events managers may need to review. |
| `ap_alerts` | Person or asset AP alerts. |
| `groups` | Database-backed roles and permission rights. |
| `settings` | Backup folders, export folders, refresh timing, display density, scan-only mode, data path settings. |

## Backup Contents

Backups include:

- Active SQLite database copy.
- CSV exports for people, assets, activity, audit, errors, manager notifications, AP alerts, groups, and settings.
- Backup metadata/info file.

Automatic daily backup is on by default. Manual backup is available from Manager and System.

## Export Folders

Settings can save default folders for:

- Backups.
- Excel exports.
- Report exports.
- Audit logs.
- Error logs.
- System logs.

## Public Repo Safety

Safe for this public repo:

- Source files approved for public release.
- Import templates with sample-only structure.
- README/docs/changelog.
- Sanitized visuals with sample data.
- Self-test result files that do not contain live data.

Do not upload:

- `.db`, `.sqlite`, or `.sqlite3` files.
- Backup ZIPs containing live data.
- Logs, reports, exports, or Excel files with live data.
- Real employee data.
- Real badge numbers.
- AP alert details from live operations.
- Incident notes.
- Network paths or store-sensitive paths.
- Screenshots showing live names, badges, AP alerts, or database paths.

## Scan-Only Data Notes

When Scan-Only Front Desk Mode is on, checkout/return can happen without operator login. The app records the actor as Scan-Only Front Desk and adds scan-only notes. The checkout employee and return-by badge are still saved from scans.

Different-badge returns are allowed. The return is accepted and the notes/audit record show both the original checkout holder and the return-by badge.
