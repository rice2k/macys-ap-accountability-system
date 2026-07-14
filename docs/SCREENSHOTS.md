# Screenshots And Capture Plan

The repository includes placeholder SVGs so the GitHub page looks organized before real screenshots are added. Replace each placeholder in `docs/screenshots/` with a real `.png` using the same base filename, then update the image links in `README.md` if needed.

## Recommended Screenshot Files

| File | Screen | Capture Notes |
| --- | --- | --- |
| `dashboard.png` | Dashboard | Capture live cards, backup card, alert/error cards, and system health badges. |
| `dashboard-detail.png` | Dashboard detail popup | Open a dashboard card and show the filtered table plus Excel export option. |
| `front-desk.png` | Front Desk | Show the checkout flow, selected employee card, selected asset card, due time, condition, and notes. |
| `assets.png` | Assets | Show search/filter/sort, status tools, Excel export type dropdown, and asset list. |
| `asset-entry-key.png` | Add/Edit Asset - Key | Show key-specific fields such as key set number, ring serial, key count, ring use, and key/access list. |
| `asset-entry-tablet.png` | Add/Edit Asset - Tablet | Show tablet-specific serial/device, IMEI/license, and accessories fields. |
| `asset-profile.png` | Asset Profile | Show detail text, holder/status, history, AP alert, print/export options. |
| `manager.png` | Manager | Show counts, backup status, refresh countdown, open assets, issue assets, and shortcuts. |
| `manager-notifications.png` | Manager Notifications | Show filters, note/resolve actions, export, and related-log options. |
| `ap-alerts.png` | AP Alerts | Show active/resolved alert review workflow. |
| `reports.png` | Reports | Show Daily, History, Export, and Output groups with report text. |
| `logs.png` | Logs | Show filters, search, selected row export, and archive/clear options. |
| `system.png` | System | Show data/storage, backup/restore, imports/templates, app/security, diagnostics, and advanced tools. |
| `settings.png` | Settings | Show folder paths, default export type, refresh timer, display density, and current data file. |
| `groups.png` | Groups / Permissions | Show rights grouped by Access, People, Assets, and Admin. |

## Screenshot Guidelines

- Use sample data only.
- Do not show real employee information, real badge numbers, live store records, or real incident notes.
- Hide or blur the full database path if it contains a username, network share, or store-sensitive folder name.
- Capture at the app's normal desktop size so buttons and tables are readable.
- Keep image names lowercase with hyphens.
- Prefer PNG for GitHub display.
- Add a short caption under each screenshot in `README.md`.

## Suggested GitHub Gallery Layout

```markdown
| Dashboard | Front Desk |
| --- | --- |
| ![Dashboard](docs/screenshots/dashboard.png) | ![Front Desk](docs/screenshots/front-desk.png) |

| Assets | Manager |
| --- | --- |
| ![Assets](docs/screenshots/assets.png) | ![Manager](docs/screenshots/manager.png) |
```

