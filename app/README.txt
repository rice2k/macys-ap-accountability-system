Macy's AP Accountability System
Version v5.2.7

Quick Start
1. Double-click Run_Macys_AP_v5_2_7.bat.
2. Sign in with an operator badge or employee ID.
3. Open Front Desk.
4. Scan the employee.
5. Scan the asset.
6. Click Check Out Item or Return Item.

Scan-Only Front Desk Option
- Admin can turn on Front Desk Scan-Only Mode in System > Settings.
- When it is on, Front Desk checkout/return can run without operator login.
- Checkout flow: scan the employee badge, scan the item, then click Check Out Item or press F8.
- Return flow: scan the return-by badge and the checked-out item, then click Return Item or press F9.
- Scan-only checkout and return records are marked in notes/audit as Scan-Only Front Desk.

Default Admin
- Employee ID: F984717
- Badge: 88984717

Shared Network Data
- Put the database in one shared folder.
- On each computer, open System > Change Data File Location.
- Point every computer to the same .db file.
- Prefer a UNC path such as \\SERVER\Share\Macys_AP_Data\macys_ap_data.db.
- The app shows Local Data or Shared Data in the header.
- Shared/custom data refreshes every 5 minutes by default.

Backups
- Automatic daily backup is on by default.
- Change this in System > Settings.
- Manual backups are available from Manager or System.
- Backups include the database and CSV exports of people, assets, activity, audit, errors, Manager notifications, AP alerts, groups, and settings.
- Backup status shows the latest backup time from the active database settings.

Main Pages
- Dashboard: live counts, recent activity, backup status.
- Front Desk: checkout and return workflow.
- People / Users: employee and role management.
- Assets: asset list, add/edit, import, export, retire.
- Manager: open assets, issue assets, backup status, refresh countdown, and quick reports.
- Manager notifications: filter, review, resolve, add notes, and export blocked checkouts, overrides, different-badge returns, failed logins, export events, and settings changes.
- AP Alerts: Manager/Admin users can add AP alerts to a person or asset, review/resolve them, export them, and see them during checkout.
- Group Management: Admin/Manager users can add/save custom groups, edit rights, assign users, view assigned users, rename groups, and delete non-protected groups with a reason.
- Left menu now focuses on Dashboard, Front Desk, Users, Assets, Search, Manager, Reports, Logs, and System.
- Groups / Permissions, Settings, and Admin Tools are managed from the System hub instead of separate left-menu pages.
- Search: people, assets, activity, audit, errors, and AP alerts.
- Reports: current out, overdue, issue lists, alert history, employee/asset history, and exports.
- System: data location, backups, imports, settings, diagnostics.

Right-Click Menus
- Dashboard cards: open details, refresh details, and copy count.
- Dashboard detail popups: view row details, export visible list, refresh, copy selected row, and close.
- Dashboard detail popups can also open related asset/user records and related audit logs.
- Asset list: view, edit, duplicate, retire, checkout/return, export selected, mark repair/missing, mark missing accessory, add/review alerts, view audit log, and copy row.
- People / Users: open/edit profile, deactivate/reactivate, add/view AP alerts, view audit log, and copy row.
- Front Desk current-out table: return selected asset, open asset/user record, view checkout, add notes, refresh, and copy row.
- Manager notifications: mark reviewed/resolved, add note, view full notification, open related record, export selected/all, and copy row.
- Manager notifications can also open related audit/error logs for the selected event.
- Logs: view full entry, export selected row, export visible logs, copy details, open log folder, and clear old logs with permission.
- Settings path fields: browse, open, reset one path, and validate write access.

Asset Page Improvements
- Search assets by barcode, name, key number, serial, location, holder, status, or notes.
- Filter by asset type and status.
- Select an asset to see a clear summary.
- Quickly add or edit assets; secondary actions such as open profile, duplicate, import, export selected, and retire are grouped under More/context menus.
- Asset filters and actions are grouped by Find/Type/Status/Sort, Manage, Status, Excel, and More.
- Asset rows use subtle status colors for easier scanning.
- Asset type mapping keeps fields clean:
  Keys use controlled key number only.
  Radios, scanners, and tablets use serial number / device tag only.
  Temp badges, items, and other assets use common fields only.
- Add/Edit Asset now changes its detail section by asset type.
- Key rings can track key set number, ring serial, key count, ring use, and each key on the ring.
- Tablets can track IMEI/license data and accessories.
- Export All saves every asset record.
- Export Type saves only the selected asset type.
- Export Excel saves a readable .xlsx workbook organized by asset type with an Export Info sheet.
- Asset Export All and Export Type now both use the formatted Excel exporter.
- Dashboard summary boxes open detail popups when clicked.
- Dashboard detail popups include an Excel export button.
- Dashboard now includes All Out and Backup status cards.
- Reports can be saved as TXT, HTML, or Excel.
- Asset page supports duplicate asset, export selected asset, export all/type, and profile export/print.
- Front Desk checks saved AP alerts plus AP alert/restriction notes before checkout.
- Critical or blocking AP alerts require Manager/Admin access before checkout can continue.
- Returns can be completed by any person as long as their badge and the item are scanned; if the return badge differs from the checkout holder, the app records it in notes/audit without blocking the return.

Layout Improvements
- Page buttons are grouped by purpose where possible.
- Banner now uses the bundled star image icon instead of a text asterisk.
- Banner star rendering uses the bundled icon artwork with transparent background cleanup so it displays cleanly on the dark header.
- The star/logo is aligned inside the main header brand area.
- Left navigation is wider, grouped into sections, and uses generated pictogram icons beside each page label.
- Active navigation has a stronger red highlight and inactive navigation has hover feedback.
- The operator/sign-in area is styled as a header bar with colored operator and role badges.
- Sign In and Sign Out now swap visibility based on whether an operator is signed in.
- Training mode now uses a compact on/off button style plus a clear mode badge.
- Top-level screens use a cleaner title bar with title, subtitle, and a quick Refresh action.
- Cards/panels use softer borders and consistent padding.
- Dashboard cards are now larger, color-coded by category, grouped into cleaner rows, and visibly clickable.
- Dashboard cards include compact icon tags and less text density.
- Dashboard search and Backup Now were removed from the main dashboard so Search and backup tools live in their proper sections.
- Dashboard system status now uses compact badges instead of a text-heavy status dump.
- Users, Assets, Search, Logs, and Manager toolbars now move lower-priority actions into More menus.
- System is now the main maintenance hub for data/storage, backup/restore, templates/imports, app/security, diagnostics, and advanced tools.
- Group Management rights are grouped by Access, People, Assets, and Admin.
- Tree tables now use shared styling, controlled column resizing behavior, horizontal/vertical scrollbars, and easier selected-row highlighting.
- Log and search tables now color important row types such as audit, error, manager, AP alert, person, and asset rows.
- Front Desk now shows a visible Operator > Employee > Asset > Due Time > Complete step flow.
- Front Desk includes a custom More due-time option for exact HH:MM or date/time entries.
- Selected Employee, Selected Asset, and Return By sections use profile-card style layouts with avatar placeholders.
- Front Desk due time, condition, and notes fields are aligned in a cleaner form row.
- Main action buttons use clearer visual styles for checkout, return, scan, and reset.
- Tables use improved header padding, selected-row highlighting, and alternating row shading.
- Empty tables now show a clear non-action placeholder row.
- Sorting tables reapplies the shared row striping and status coloring.
- Long selected row details and System detail output now open in scrollable readable panes.
- Reports are grouped into Daily, History, Export, and Output sections.
- Front Desk controls separate actions from due-time presets.
- Manager shortcuts are grouped into Daily, Reports, and Setup sections.
- Manager keeps only the daily workflow actions visible and moves review/history tools into More.
- Manager backup status shows last backup date, last backup time, backup status, countdown, and a manual refresh button.
- Manager refresh countdown now displays minutes and seconds.
- Settings shows the current main data file and includes saved folders for backups, Excel exports, reports, and log output.
- Settings includes default asset export type and refresh timer seconds, with 300 seconds / 5 minutes as the default.
- Settings includes display density choices that adjust button/table spacing.
- Settings includes an Admin-only Front Desk Scan-Only Mode switch for badge/item checkout and return without operator login.
- Default groups are protected so Admin/Manager access cannot be accidentally removed.
- Group rights changes require a reason, and group deletion does not move assigned users until the delete reason is confirmed.
- Manager page includes backup/export history shortcuts.
- Manager notifications can be reviewed, resolved, annotated with notes, and exported.
- Manager notification exports and secondary refresh/open-log actions are grouped under More.
- Logs page can filter by type, user/source, asset/target, action, date range, and search text.
- Logs page can export the current filtered log view to Excel and can archive old log records before clearing them with Admin access and a reason.
- Excel workbooks include active page/filter/search context on the Export Info sheet where available.
- Export filenames are cleaned so asset barcodes, group names, and filters with Windows-invalid characters do not break saves.
- Audit/error log rows include date, 12-hour time, 24-hour time, operator role, active page, status, and computer name fields.

Final QA
- Self-test checks database creation, admin seed, asset seed, asset field mapping, detail save/reload, Excel export write, add/reload, checkout, double-checkout guard, return, edit/reload, status update, dashboard counts, AP alert save/resolve, rich audit/error fields, and cleanup.
- Restore now uses the current active data file, including a shared/custom database path.

Included Files
- macys_ap_v5_2_7.py: main app.
- macys_ap_export.py: Excel workbook export helper module.
- Run_Macys_AP_v5_2_7.bat: launcher.
- Run_Self_Test_No_GUI.bat: validation test.
- Build_EXE.bat: optional EXE builder.
- people_import_template.csv: people import template.
- assets_import_template.csv: asset import template.
- macys_star_icon.ico / .png: app icons.

Flow + Visual Polish Update
- Front Desk now shows a large plain-language next-step banner so users know exactly what to scan or click next.
- Front Desk now has Quick Scan buttons for New Checkout, Return Item, and Next Item Same Employee.
- Check Out Item and Return Item are enabled only when that action is ready.
- Checkout Mode and Return Mode are shown clearly above the scan box.
- Manager now has a Today's Priorities strip for Backup, Late Returns, Alerts, and Open Assets.
- Manager priority cards change color when action is needed.
- Manager now includes an Only Show Problems toggle.
- Scanner stations now have keyboard shortcuts: F2 New Checkout, F3 Return Mode, F4 Next Item Same Employee, F8 Check Out, F9 Return, Esc Start Over.
- Front Desk now gives visual scan feedback for success, return mode, and blocked actions.
- Front Desk can export a printable Quick Scan guide.
- Front Desk can run in Admin-enabled Scan-Only Mode: checkout uses employee badge plus item, and return uses return-by badge plus item.
- System includes Reset Demo Data for Training Mode sample users/assets only.
- VISUAL_FLOW_IMPROVEMENTS_30.txt includes 30 additional visual/workflow improvement suggestions.
