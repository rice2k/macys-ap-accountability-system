# Public GitHub Upload

This package is ready for a public GitHub repository or public GitHub Pages site. Because the app supports AP operations and can contain sensitive employee, asset, alert, log, and backup data, only sanitized documentation and approved source files should be uploaded.

## Recommended Public Repository Setup

1. Create a new GitHub repository.
2. Set visibility to Public.
3. Upload this documentation package.
4. Keep `README.md` at the repository root.
5. Confirm `.gitignore` is present before adding app files.
6. Add app source files only if they are approved for public release.
7. Do not upload live `.db`, backup ZIPs, exports, reports, logs, employee data, badge numbers, AP alert details, or incident notes.

## GitHub Pages Setup

1. Go to the repository on GitHub.
2. Open Settings.
3. Open Pages.
4. Under Build and deployment, choose Deploy from a branch.
5. Choose the main branch and root folder.
6. Save.
7. Wait a few minutes, then open the GitHub Pages URL.

Official GitHub references:

- Repository visibility: https://docs.github.com/repositories/managing-your-repositorys-settings-and-features/managing-repository-settings/setting-repository-visibility
- GitHub Pages quickstart: https://docs.github.com/pages/quickstart

## What To Upload

Safe for public upload:

- `README.md`
- `CHANGELOG.md`
- `SECURITY.md`
- `docs/`
- `.github/ISSUE_TEMPLATE/`
- `_config.yml`
- Sanitized screenshots
- App source files only if approved
- Import templates only if they contain sample data
- Icons and static assets

Do not upload:

- `*.db`
- `Backups/`
- `Logs/`
- `Exports/`
- `Reports/`
- Real employee data
- Real badge numbers
- Real incident notes
- Real AP alert notes
- Real checkout/return history
- Network paths or store-specific paths
- Backup ZIP files
- Generated Excel exports with live data

## Suggested Public Repo Description

```text
Documentation and release notes for the Macy's AP Accountability System.
```

## Suggested Topics

```text
asset-tracking, checkout-return, tkinter, sqlite, documentation, windows-desktop
```

## First Public Issue To Create

Title:

```text
Add sanitized production screenshots to README gallery
```

Body:

```text
Replace placeholder screenshot SVGs with sanitized PNG screenshots:
- Dashboard
- Front Desk
- Assets
- Add/Edit Asset - Key
- Add/Edit Asset - Tablet
- Manager
- Logs
- System

Make sure no live employee, badge, incident, AP alert, database path, or network-path details are visible.
```

