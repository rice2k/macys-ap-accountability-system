# Version History

The current app version is v5.2.7. The package history below is based on the available build files and release notes from the Macy's AP workspace.

## Current Build

| Field | Value |
| --- | --- |
| Current version | v5.2.7 |
| Current package | Macys_AP_Beta_v5_2_7_Asset_Entry_Improved.zip |
| Current source | `macys_ap_v5_2_7.py` |
| Export helper | `macys_ap_export.py` |
| Latest validation | Self-test PASS |

## v5.2.7 Build Timeline

| Build / Package | Main Focus |
| --- | --- |
| Clean Improved | Cleaned package, startup files, README, validation files. |
| Network Improved Final | Shared/custom database support and faster refresh for shared data. |
| Polished Network Final | Additional shared data polish and operational packaging. |
| Assets Manager Final | Asset page and manager page improvements. |
| Button Layout Final | Toolbar/action grouping and button layout cleanup. |
| Consistency Final | Naming, page consistency, and final layout checks. |
| Final Checked | Validated checked package with self-test output. |
| Mapped Final | Asset type field mapping so irrelevant key/serial fields are cleared. |
| Manager Asset Excel Upgrade | Manager backup status band, asset detail database field, and Excel asset export. |
| Modular Export Upgrade | Excel workbook generation moved into `macys_ap_export.py`. |
| Text Document Upgrade | Manager notifications, dashboard detail popups, saved folders, and clearer checkout warnings. |
| Complete Audit Upgrade | Expanded audit/error records and reporting support. |
| Full Recommendations | Broader recommendations and system/report improvements. |
| Text Followup | Dashboard All Out and Backup cards, filtered log export, and manager notes. |
| Text File Complete Update | AP Alerts table, alert review workflows, search/report integration, richer exports. |
| DOCX Upgrade Complete | Right-click/context menus, selected-row export, expanded log fields, sign-in/page logging context. |
| DOCX Upgrade Continuation | Direct Manager shortcuts, default asset export type, refresh timer seconds, related log actions. |
| Asset Entry Improved | Type-specific asset forms, final UI polish, backup/export/log safety improvements, display density, and final self-test PASS. |

## Available Package Inventory

| Package | Timestamp |
| --- | --- |
| `Macys_AP_Beta_v5_2_7_Clean_Improved.zip` | 2026-06-17 02:14 AM |
| `Macys_AP_v5_2_7_Network_Improved_Final.zip` | 2026-06-17 02:57 AM |
| `Macys_AP_v5_2_7_Polished_Network_Final.zip` | 2026-06-17 03:19 AM |
| `Macys_AP_v5_2_7_Assets_Manager_Final.zip` | 2026-06-17 03:31 AM |
| `Macys_AP_v5_2_7_Button_Layout_Final.zip` | 2026-06-17 03:36 AM |
| `Macys_AP_v5_2_7_Consistency_Final.zip` | 2026-06-17 03:42 AM |
| `Macys_AP_v5_2_7_Final_Checked.zip` | 2026-06-17 03:47 AM |
| `Macys_AP_v5_2_7_Mapped_Final_20260617_1312.zip` | 2026-06-17 01:12 PM |
| `Macys_AP_v5_2_7_Manager_Asset_Excel_Upgrade_20260617_1325.zip` | 2026-06-17 01:35 PM |
| `Macys_AP_v5_2_7_Modular_Export_Upgrade_20260617_1400.zip` | 2026-06-17 01:58 PM |
| `Macys_AP_v5_2_7_Text_Document_Upgrade_20260617_1420.zip` | 2026-06-17 02:07 PM |
| `Macys_AP_v5_2_7_Complete_Audit_Upgrade_20260617_1435.zip` | 2026-06-17 02:41 PM |
| `Macys_AP_v5_2_7_Full_Recommendations_20260617_1455.zip` | 2026-06-17 03:00 PM |
| `Macys_AP_v5_2_7_Text_Followup_20260617_1745.zip` | 2026-06-17 05:52 PM |
| `Macys_AP_Beta_v5_2_7_Text_File_Complete_Update_20260617.zip` | 2026-06-17 07:39 PM |
| `Macys_AP_Beta_v5_2_7_Text_File_Complete_Update_20260617_CLEAN.zip` | 2026-06-17 07:40 PM |
| `Macys_AP_Beta_v5_2_7_DOCX_Upgrade_Complete_20260617.zip` | 2026-06-17 08:01 PM |
| `Macys_AP_Beta_v5_2_7_DOCX_Upgrade_Continuation_20260617.zip` | 2026-06-17 08:24 PM |
| `Macys_AP_Beta_v5_2_7_Asset_Entry_Improved.zip` | 2026-06-18 12:53 AM |

## Major v5.2.7 Updates

### Asset Entry

- Guided Add/Edit Asset flow.
- Auto-Fill Name.
- Save + New.
- Duplicate checks.
- Type-specific guidance.
- Type-specific detail fields.
- Key ring detail tracking.
- Tablet IMEI/license/accessory tracking.
- CSV import validation aligned with manual entry.

### Manager Page

- Last backup date.
- Last backup time.
- Backup status.
- Refresh countdown.
- Manual refresh.
- Open and issue asset tables.
- Manager notifications.
- AP alert review.
- Backup/export history shortcuts.
- Log, audit, error, settings, and admin shortcuts.

### Dashboard

- Larger grouped cards.
- All Out and Backup status cards.
- Click-through detail popups.
- Excel export from detail popups.
- Compact system health badges.
- Reduced duplicated search/backup controls.

### Logs And Alerts

- Manager notifications table.
- AP Alerts table.
- Alert save/review/resolve workflows.
- Checkout alert checks.
- Critical/blocking alert manager override.
- Wrong-user return manager approval.
- Rich audit/error fields.
- Combined log viewer with filters.
- Archive before old log deletion.

### Exports

- Formatted Excel asset exporter.
- Export All and Export Type.
- Report Excel save.
- Dashboard detail Excel export.
- Manager notification export.
- AP alert export.
- Group user export.
- Filtered log export.
- Selected log row export.
- CSV bundle now includes manager notifications, AP alerts, groups, and settings.

### Permissions

- Database-backed groups.
- Protected default roles.
- Group rights editing with reason.
- Group deletion with reason and replacement group handling.
- Assign user to group.
- Permission checks read from database groups.

### UI Polish

- Wider grouped left navigation.
- Cleaner header and sign-in area.
- Image-based Macy star/logo.
- Hover feedback and stronger active-page highlight.
- Cleaner title/subtitle bars with quick refresh.
- Improved table headers, row striping, selected rows, and scrollbars.
- Display density setting.
- More menus for lower-priority actions.
