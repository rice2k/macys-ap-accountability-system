# How It Works

## Daily Checkout Flow

```mermaid
flowchart TD
    A["Operator signs in"] --> B["Front Desk"]
    B --> C["Scan employee badge or ID"]
    C --> D["Employee profile loaded"]
    D --> E["Scan asset barcode, key number, serial, or device tag"]
    E --> F["Asset profile loaded"]
    F --> G["Select due time, condition, and notes"]
    G --> H{"Restriction or AP alert?"}
    H -->|No| I["Checkout saved"]
    H -->|Yes| J["Manager/Admin approval required"]
    J --> I
    I --> K["Asset marked Checked Out"]
    K --> L["Activity, audit, and manager counts update"]
```

## Return Flow

```mermaid
flowchart TD
    A["Scan checked-out asset"] --> B["Open checkout record found"]
    B --> C["Scan return-by person when needed"]
    C --> D{"Returned by assigned holder?"}
    D -->|Yes| E["Return accepted"]
    D -->|No| F["Manager/Admin review and reason required"]
    F --> E
    E --> G["Activity marked Returned"]
    G --> H["Asset marked Available"]
    H --> I["Audit record written"]
```

## Backup Flow

```mermaid
flowchart TD
    A["Auto backup enabled"] --> B{"Backup already done today?"}
    B -->|Yes| C["Status shows OK today"]
    B -->|No| D["Create backup ZIP"]
    D --> E["Copy database"]
    E --> F["Export key tables as CSV"]
    F --> G["Write backup info"]
    G --> H["Save last backup timestamp"]
    H --> I["Manager backup band updates"]
```

## Data Model Summary

The app uses a SQLite database. The current build includes these major data areas:

- `people` for users, employees, operators, roles, status, shifts, and notes.
- `assets` for controlled assets, status, holder, type-specific details, and notes.
- `activity` for checkouts, due times, returns, conditions, and operators.
- `audit` for user actions and system events.
- `errors` for application errors and diagnostics.
- `manager_notifications` for events managers should review.
- `ap_alerts` for person or asset AP alerts.
- `groups` for database-backed roles and permission sets.
- `settings` for folders, refresh timing, backup settings, display density, and saved paths.

## Access Control

Permission checks happen before protected actions. Manager/Admin access is required for high-risk work such as:

- Restricted checkout overrides.
- Wrong-user returns.
- Missing/repair/retired asset checkout.
- Critical or blocking AP alerts.
- Group rights changes.
- Group deletion or rename.
- Log cleanup.
- Admin tools.

## Shared Data Mode

The app can use a local database or a shared database path. In shared/custom mode, the header shows Shared Data and the app refreshes more often. Every workstation should point to the same `.db` file through System > Change Data File Location.

Recommended shared path format:

```text
\\SERVER\Share\Macys_AP_Data\macys_ap_data.db
```

## File Output

The app can save output to configured folders:

- Backups.
- Excel exports.
- Report exports.
- Audit logs.
- Error logs.
- System logs.

Generated output should not be committed to GitHub unless it is sanitized sample data.

