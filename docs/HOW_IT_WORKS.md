# How It Works

## Checkout Flow

```mermaid
flowchart TD
    A["Open Front Desk"] --> B{"Scan-Only Mode on?"}
    B -->|No| C["Operator signs in"]
    B -->|Yes| D["No operator login needed"]
    C --> E["Scan employee badge or ID"]
    D --> E
    E --> F["Employee profile loaded"]
    F --> G["Scan asset barcode, key number, serial, or device tag"]
    G --> H["Asset profile loaded"]
    H --> I["Pick due time, condition, and notes"]
    I --> J{"Restricted checkout or AP alert?"}
    J -->|No| K["Checkout saved"]
    J -->|Yes| L["Manager/Admin approval required"]
    L --> K
    K --> M["Asset marked Checked Out"]
    M --> N["Activity, audit, dashboard, manager counts update"]
```

## Return Flow

```mermaid
flowchart TD
    A["Open Front Desk"] --> B["Scan return-by badge"]
    B --> C["Return-by person recorded"]
    C --> D["Scan checked-out item"]
    D --> E["Open checkout record found"]
    E --> F["Return accepted"]
    F --> G{"Return-by badge matches checkout holder?"}
    G -->|Yes| H["Record normal return"]
    G -->|No| I["Record different-badge return in notes/audit"]
    H --> J["Activity marked Returned"]
    I --> J
    J --> K["Asset marked Available"]
    K --> L["Manager/dashboard/report counts refresh"]
```

Return rule: it does not matter who returns the item. The app only needs the return-by badge and the checked-out item scan. If the return-by badge is different from the original checkout holder, the return still completes and the difference is recorded.

## Scan-Only Front Desk Mode

Scan-Only Mode is an Admin-controlled setting in System > Settings.

When off:

- Front Desk requires a signed-in operator with Front Desk access.
- Checkout and return actions are attributed to the signed-in operator.

When on:

- Front Desk checkout/return can run without operator login.
- The header displays Scan-Only Front Desk.
- Checkout uses employee badge plus item scan.
- Return uses return-by badge plus checked-out item scan.
- Audit/notes label the action as Scan-Only Front Desk.

## Manager Oversight Flow

```mermaid
flowchart LR
    A["Checkout/return activity"] --> B["Dashboard cards"]
    A --> C["Manager page"]
    A --> D["Reports"]
    A --> E["Audit logs"]
    C --> F["Today's Priorities"]
    C --> G["Open assets"]
    C --> H["Issue assets"]
    C --> I["Manager notifications"]
    C --> J["Backup status"]
```

Managers can focus on what needs attention: late returns, active alerts, open assets, issue assets, backup status, failed logins, blocked checkout attempts, export events, settings changes, and log/audit history.

## Backup Flow

```mermaid
flowchart TD
    A["Auto backup enabled"] --> B{"Backup already done today?"}
    B -->|Yes| C["Status shows OK today"]
    B -->|No| D["Create backup ZIP"]
    D --> E["Copy active SQLite database"]
    E --> F["Export key tables as CSV"]
    F --> G["Write backup info"]
    G --> H["Save last backup timestamp"]
    H --> I["Manager backup band updates"]
```

## Data Flow Summary

The live app stores data in SQLite, writes audit and error records, and can operate from either a local database or a shared/network database. Exports and backups are saved to folders configured in Settings.

More detail is in [Data Info](DATA_INFO.md).
