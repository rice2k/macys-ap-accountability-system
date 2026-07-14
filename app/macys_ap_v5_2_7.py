
import csv
import datetime as dt
import json
import math
import os
import shutil
import sqlite3
import sys
import time
import traceback
import uuid
import zipfile
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

from macys_ap_export import write_xlsx_workbook

try:
    from PIL import Image, ImageTk, ImageDraw
except Exception:
    Image = None
    ImageTk = None
    ImageDraw = None

APP_TITLE = "Macy's AP Accountability System"
APP_VERSION = "v5.2.7"
DB_FILE = "macys_ap_v527.db"
CONFIG_FILE = "macys_ap_config.json"

ROLES = ["Employee", "Front Desk", "AP Operator", "Manager", "Admin"]
ROLE_RANK = {"Guest": 0, "Employee": 1, "Front Desk": 2, "AP Operator": 3, "Manager": 4, "Admin": 5}
ROLE_RIGHTS = {
    "Guest": {"dashboard"},
    "Employee": {"dashboard"},
    "Front Desk": {"dashboard", "frontdesk", "search"},
    "AP Operator": {"dashboard", "frontdesk", "people_view", "assets_view", "search", "reports"},
    "Manager": {"dashboard", "frontdesk", "people_view", "people_edit", "assets_view", "assets_edit", "search", "reports", "manager", "system"},
    "Admin": {"dashboard", "frontdesk", "people_view", "people_edit", "assets_view", "assets_edit", "search", "reports", "manager", "system", "admin"},
}
ALL_RIGHTS = sorted({right for rights in ROLE_RIGHTS.values() for right in rights})
ASSET_TYPES = ["Key", "Radio", "Temp Badge", "Scanner", "Tablet", "Item", "Other"]
ASSET_STATUS = ["Available", "Checked Out", "Missing", "Repair", "Retired"]
PEOPLE_STATUS = ["Active", "Manual Review", "Inactive", "Suspended", "Deleted"]
ASSET_FIELDS = ["asset_type", "barcode", "name", "status", "location", "controlled_key_number", "serial_number", "notes"]
ASSET_DB_FIELDS = ASSET_FIELDS + ["asset_details"]
ASSET_COMMON_FIELDS = ("asset_type", "barcode", "name", "status", "location", "notes")
TABLET_ACCESSORIES = ["Tablet cover", "Tablet strap", "iPad keyboard", "Charging cable", "Charging block", "Docking station", "Stylus", "Other accessory"]
ASSET_DETAIL_SCHEMAS = {
    "Key": {
        "title": "Key Ring Details",
        "fields": [
            {"key": "key_set_number", "label": "Key set number", "required": True, "unique": True},
            {"key": "key_ring_serial_number", "label": "Key ring serial number", "required": True, "unique": True},
            {"key": "number_of_keys", "label": "Number of keys on the ring", "required": True, "type": "int", "min": 1, "default": "1"},
            {"key": "key_ring_location", "label": "Key ring location", "required": True},
            {"key": "key_ring_use", "label": "What this key ring is used for", "required": True},
        ],
        "list_fields": ("key_set_number", "key_ring_serial_number", "number_of_keys", "key_ring_location", "key_ring_use"),
    },
    "Radio": {
        "title": "Radio Details",
        "fields": [
            {"key": "radio_number", "label": "Radio number", "required": True, "unique": True},
            {"key": "factory_serial_number", "label": "Factory serial number", "required": True, "unique": True},
            {"key": "radio_serial_number", "label": "Radio serial number", "required": True, "unique": True},
            {"key": "radio_location", "label": "Radio location", "required": True},
            {"key": "assigned_area", "label": "Assigned area", "required": False},
        ],
        "list_fields": ("radio_number", "factory_serial_number", "radio_serial_number", "radio_location", "assigned_area"),
    },
    "Tablet": {
        "title": "Tablet Details",
        "fields": [
            {"key": "tablet_number", "label": "Tablet number", "required": True, "unique": True},
            {"key": "tablet_factory_serial_number", "label": "Tablet factory serial number", "required": True, "unique": True},
            {"key": "imei_number", "label": "IMEI number", "required": False, "unique": True},
            {"key": "windows_license_key", "label": "Windows key or license key", "required": False},
            {"key": "tablet_location", "label": "Tablet location", "required": True},
            {"key": "assigned_area", "label": "Assigned area", "required": False},
            {"key": "other_accessory", "label": "Other accessory", "required": False},
        ],
        "accessories": TABLET_ACCESSORIES,
        "list_fields": ("tablet_number", "tablet_factory_serial_number", "imei_number", "tablet_location", "assigned_area"),
    },
    "Temp Badge": {
        "title": "Temp Badge Details",
        "fields": [
            {"key": "badge_number", "label": "Badge number", "required": True, "unique": True},
            {"key": "badge_location", "label": "Badge location", "required": True},
            {"key": "assigned_area", "label": "Assigned area", "required": False},
        ],
        "list_fields": ("badge_number", "badge_location", "assigned_area"),
    },
    "Scanner": {
        "title": "Scanner Details",
        "fields": [
            {"key": "scanner_number", "label": "Scanner number", "required": False, "unique": True},
            {"key": "scanner_serial_number", "label": "Scanner serial number", "required": True, "unique": True},
            {"key": "scanner_location", "label": "Scanner location", "required": True},
            {"key": "assigned_area", "label": "Assigned area", "required": False},
        ],
        "list_fields": ("scanner_number", "scanner_serial_number", "scanner_location", "assigned_area"),
    },
    "Item": {
        "title": "Item Details",
        "fields": [
            {"key": "item_number", "label": "Item number", "required": False, "unique": True},
            {"key": "item_location", "label": "Item location", "required": False},
            {"key": "assigned_area", "label": "Assigned area", "required": False},
            {"key": "item_use", "label": "What this item is used for", "required": False},
        ],
        "list_fields": ("item_number", "item_location", "assigned_area", "item_use"),
    },
    "Other": {
        "title": "Other Asset Details",
        "fields": [
            {"key": "asset_number", "label": "Asset number", "required": False, "unique": True},
            {"key": "asset_location", "label": "Asset location", "required": False},
            {"key": "assigned_area", "label": "Assigned area", "required": False},
            {"key": "asset_use", "label": "What this asset is used for", "required": False},
        ],
        "list_fields": ("asset_number", "asset_location", "assigned_area", "asset_use"),
    },
}
ASSET_TYPE_FIELD_MAP = {
    "Key": {
        "fields": ASSET_COMMON_FIELDS + ("controlled_key_number",),
        "recommended": ("controlled_key_number",),
        "field_notes": "Keys use the controlled key number field. Serial number stays blank.",
    },
    "Radio": {
        "fields": ASSET_COMMON_FIELDS + ("serial_number",),
        "recommended": ("serial_number",),
        "field_notes": "Radios use the serial number field. Controlled key number stays blank.",
    },
    "Temp Badge": {
        "fields": ASSET_COMMON_FIELDS,
        "recommended": (),
        "field_notes": "Temporary badges use barcode, name, status, location, and notes only.",
    },
    "Scanner": {
        "fields": ASSET_COMMON_FIELDS + ("serial_number",),
        "recommended": ("serial_number",),
        "field_notes": "Scanners use the serial number field. Controlled key number stays blank.",
    },
    "Tablet": {
        "fields": ASSET_COMMON_FIELDS + ("serial_number",),
        "recommended": ("serial_number",),
        "field_notes": "Tablets use the serial number field. Controlled key number stays blank.",
    },
    "Item": {
        "fields": ASSET_COMMON_FIELDS,
        "recommended": (),
        "field_notes": "General items use the common asset fields only.",
    },
    "Other": {
        "fields": ASSET_COMMON_FIELDS,
        "recommended": (),
        "field_notes": "Other assets use the common asset fields only.",
    },
}
ASSET_TYPE_GUIDANCE = {
    "Key": {
        "name_prefix": "Key Ring",
        "recommended": ("controlled_key_number",),
        "notes": "Keys should include a controlled key number when available.",
    },
    "Radio": {
        "name_prefix": "Radio",
        "recommended": ("serial_number",),
        "notes": "Radios should include the serial number printed on the unit.",
    },
    "Temp Badge": {
        "name_prefix": "Temp Badge",
        "recommended": (),
        "notes": "Temporary badges only need a barcode and clear display name.",
    },
    "Scanner": {
        "name_prefix": "Scanner",
        "recommended": ("serial_number",),
        "notes": "Scanners should include the serial number when available.",
    },
    "Tablet": {
        "name_prefix": "Tablet",
        "recommended": ("serial_number",),
        "notes": "Tablets should include the serial number or device tag when available.",
    },
    "Item": {
        "name_prefix": "Item",
        "recommended": (),
        "notes": "Use Item for scanners, tablets, pouches, and other tracked equipment.",
    },
    "Other": {
        "name_prefix": "Asset",
        "recommended": (),
        "notes": "Use Other only when the asset does not fit a standard category.",
    },
}

def asset_type_rule(asset_type):
    return ASSET_TYPE_FIELD_MAP.get(asset_type, ASSET_TYPE_FIELD_MAP["Item"])

def asset_detail_schema(asset_type):
    return ASSET_DETAIL_SCHEMAS.get(asset_type, ASSET_DETAIL_SCHEMAS["Item"])

def apply_asset_type_field_map(d):
    allowed = set(asset_type_rule(d.get("asset_type"))["fields"])
    for type_field in ("controlled_key_number", "serial_number"):
        if type_field not in allowed:
            d[type_field] = ""
    return d

def parse_asset_details(value):
    if isinstance(value, dict):
        return dict(value)
    if not value:
        return {}
    try:
        parsed = json.loads(str(value))
        return parsed if isinstance(parsed, dict) else {}
    except Exception:
        return {}

def normalize_asset_details(asset_type, source):
    details = parse_asset_details(source.get("asset_details", ""))
    for spec in asset_detail_schema(asset_type)["fields"]:
        key = spec["key"]
        value = clean(source.get(key, details.get(key, spec.get("default", ""))))
        if spec.get("type") == "int":
            value = "".join(ch for ch in value if ch.isdigit())
        details[key] = value
    if asset_type == "Key":
        count_text = details.get("number_of_keys") or "1"
        count = max(1, min(50, int(count_text) if count_text.isdigit() else 1))
        details["number_of_keys"] = str(count)
        existing_keys = details.get("keys", [])
        if not isinstance(existing_keys, list):
            existing_keys = []
        key_rows = []
        for idx in range(count):
            previous = existing_keys[idx] if idx < len(existing_keys) and isinstance(existing_keys[idx], dict) else {}
            serial = clean(source.get(f"key_{idx+1}_serial", previous.get("serial", ""))).upper()
            access = clean(source.get(f"key_{idx+1}_access", previous.get("access", "")))
            key_rows.append({"serial": serial, "access": access})
        details["keys"] = key_rows
    if asset_type == "Tablet":
        accessories = []
        selected = source.get("accessories", details.get("accessories", []))
        if isinstance(selected, str):
            selected = [x.strip() for x in selected.split(";") if x.strip()]
        for item in selected if isinstance(selected, list) else []:
            if item in TABLET_ACCESSORIES and item not in accessories:
                accessories.append(item)
        details["accessories"] = accessories
        if "Other accessory" not in accessories:
            details["other_accessory"] = ""
    return details

def asset_details_to_json(details):
    return json.dumps(details or {}, sort_keys=True)

def asset_detail_text(asset_type, details):
    details = parse_asset_details(details)
    schema = asset_detail_schema(asset_type)
    parts = []
    for spec in schema["fields"]:
        value = details.get(spec["key"])
        if value:
            parts.append(f"{spec['label']}: {value}")
    if asset_type == "Tablet" and details.get("accessories"):
        parts.append("Accessories: " + ", ".join(details["accessories"]))
        if details.get("other_accessory"):
            parts.append("Other accessory: " + details["other_accessory"])
    if asset_type == "Key" and details.get("keys"):
        for idx, item in enumerate(details["keys"], start=1):
            serial = item.get("serial", "")
            access = item.get("access", "")
            if serial or access:
                parts.append(f"Key {idx}: {serial} - {access}".strip(" -"))
    return " | ".join(parts)

def now_iso():
    return dt.datetime.now().isoformat(timespec="seconds")

def pretty(s):
    if not s:
        return ""
    try:
        return dt.datetime.fromisoformat(str(s)).strftime("%m/%d/%Y %I:%M %p")
    except Exception:
        return str(s)

def clean(s):
    return str(s or "").strip()

def safe_filename_part(s, fallback="file"):
    text = clean(s) or fallback
    for ch in '<>:"/\\|?*':
        text = text.replace(ch, "-")
    text = " ".join(text.split()).strip(" .")
    return (text[:80] or fallback)

def stamp():
    return dt.datetime.now().strftime("%Y%m%d_%H%M%S")

def plus_hours(h):
    return (dt.datetime.now() + dt.timedelta(hours=h)).isoformat(timespec="seconds")

def due_today(hour=18, minute=30):
    target = dt.datetime.now().replace(hour=hour, minute=minute, second=0, microsecond=0)
    if target < dt.datetime.now():
        target += dt.timedelta(days=1)
    return target.isoformat(timespec="seconds")


class ToolTip:
    """Small hover description helper for buttons, entries, dropdowns, and tables."""
    def __init__(self, widget, text, delay=450):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tip = None
        self.job = None
        widget.bind("<Enter>", self.schedule, add="+")
        widget.bind("<Leave>", self.hide, add="+")
        widget.bind("<ButtonPress>", self.hide, add="+")

    def schedule(self, event=None):
        self.cancel()
        self.job = self.widget.after(self.delay, self.show)

    def cancel(self):
        if self.job:
            try:
                self.widget.after_cancel(self.job)
            except Exception:
                pass
            self.job = None

    def show(self):
        if self.tip or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 20
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
            self.tip = tk.Toplevel(self.widget)
            self.tip.wm_overrideredirect(True)
            self.tip.wm_geometry(f"+{x}+{y}")
            frame = tk.Frame(self.tip, bg="#111111", bd=1, relief="solid")
            frame.pack()
            tk.Label(frame, text=self.text, justify="left", bg="#111111", fg="white",
                     font=("Segoe UI", 9), padx=9, pady=6, wraplength=360).pack()
        except Exception:
            self.tip = None

    def hide(self, event=None):
        self.cancel()
        if self.tip:
            try:
                self.tip.destroy()
            except Exception:
                pass
            self.tip = None


class DB:
    def __init__(self, path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(self.path, timeout=30, isolation_level=None)
        self.conn.row_factory = sqlite3.Row
        self.configure_connection()
        self.setup()

    def configure_connection(self):
        # These settings make shared-drive use more tolerant of brief write locks.
        self.conn.execute("PRAGMA busy_timeout=30000")
        self.conn.execute("PRAGMA foreign_keys=ON")
        self.conn.execute("PRAGMA synchronous=NORMAL")
        self.conn.execute("PRAGMA journal_mode=DELETE")

    def execute_with_retry(self, sql, params=(), write=False):
        last_error = None
        for attempt in range(6):
            try:
                cur = self.conn.execute(sql, params)
                return cur
            except sqlite3.OperationalError as e:
                last_error = e
                msg = str(e).lower()
                if "locked" not in msg and "busy" not in msg:
                    raise
                time.sleep(0.25 * (attempt + 1))
        raise last_error

    def begin_immediate(self):
        self.execute_with_retry("BEGIN IMMEDIATE")

    def commit(self):
        self.conn.commit()

    def rollback(self):
        try:
            self.conn.rollback()
        except Exception:
            pass

    def run(self, sql, params=()):
        cur = self.execute_with_retry(sql, params, write=True)
        return cur

    def one(self, sql, params=()):
        return self.execute_with_retry(sql, params).fetchone()

    def all(self, sql, params=()):
        return self.execute_with_retry(sql, params).fetchall()

    def setup(self):
        self.run("""CREATE TABLE IF NOT EXISTS people(
            id INTEGER PRIMARY KEY,
            employee_id TEXT UNIQUE,
            badge TEXT UNIQUE,
            first_name TEXT,
            last_name TEXT,
            department TEXT,
            status TEXT DEFAULT 'Active',
            shift TEXT,
            role TEXT DEFAULT 'Employee',
            notes TEXT,
            created_at TEXT,
            updated_at TEXT
        )""")
        self.run("""CREATE TABLE IF NOT EXISTS assets(
            id INTEGER PRIMARY KEY,
            asset_type TEXT,
            barcode TEXT UNIQUE,
            name TEXT,
            status TEXT DEFAULT 'Available',
            location TEXT,
            controlled_key_number TEXT,
            serial_number TEXT,
            asset_details TEXT,
            notes TEXT,
            current_holder_id TEXT,
            current_holder_name TEXT,
            created_at TEXT,
            updated_at TEXT
        )""")
        self.run("""CREATE TABLE IF NOT EXISTS activity(
            id INTEGER PRIMARY KEY,
            log_id TEXT UNIQUE,
            asset_barcode TEXT,
            asset_type TEXT,
            asset_name TEXT,
            employee_id TEXT,
            employee_badge TEXT,
            employee_name TEXT,
            status TEXT,
            checked_out_at TEXT,
            due_back_at TEXT,
            returned_at TEXT,
            checkout_operator TEXT,
            return_operator TEXT,
            condition_out TEXT,
            condition_in TEXT,
            checkout_notes TEXT,
            return_notes TEXT
        )""")
        self.run("""CREATE TABLE IF NOT EXISTS audit(
            id INTEGER PRIMARY KEY,
            timestamp TEXT,
            date TEXT,
            time_12 TEXT,
            time_24 TEXT,
            action TEXT,
            actor TEXT,
            actor_role TEXT,
            page TEXT,
            target TEXT,
            old_value TEXT,
            new_value TEXT,
            reason TEXT,
            status TEXT,
            details TEXT,
            computer_name TEXT
        )""")
        self.run("""CREATE TABLE IF NOT EXISTS errors(
            id INTEGER PRIMARY KEY,
            timestamp TEXT,
            date TEXT,
            time_12 TEXT,
            time_24 TEXT,
            source TEXT,
            message TEXT,
            details TEXT,
            actor TEXT,
            actor_role TEXT,
            page TEXT,
            status TEXT,
            computer_name TEXT
        )""")
        self.run("""CREATE TABLE IF NOT EXISTS settings(
            key TEXT PRIMARY KEY,
            value TEXT
        )""")
        self.run("""CREATE TABLE IF NOT EXISTS manager_notifications(
            id INTEGER PRIMARY KEY,
            timestamp TEXT,
            time_12 TEXT,
            time_24 TEXT,
            event_type TEXT,
            severity TEXT,
            user_involved TEXT,
            asset_involved TEXT,
            action_taken TEXT,
            handled_by TEXT,
            status TEXT DEFAULT 'New',
            notes TEXT,
            reason TEXT,
            computer_name TEXT
        )""")
        self.run("""CREATE TABLE IF NOT EXISTS ap_alerts(
            id INTEGER PRIMARY KEY,
            timestamp TEXT,
            time_12 TEXT,
            time_24 TEXT,
            target_type TEXT,
            target_key TEXT,
            target_label TEXT,
            alert_type TEXT,
            severity TEXT,
            note TEXT,
            required_action TEXT,
            created_by TEXT,
            status TEXT DEFAULT 'Active',
            reviewed_by TEXT,
            reviewed_at TEXT,
            resolved_by TEXT,
            resolved_at TEXT,
            reason TEXT,
            computer_name TEXT
        )""")
        self.run("""CREATE TABLE IF NOT EXISTS groups(
            id INTEGER PRIMARY KEY,
            name TEXT UNIQUE,
            rights TEXT,
            protected TEXT DEFAULT 'No',
            notes TEXT,
            created_at TEXT,
            updated_at TEXT
        )""")
        self.ensure_column("assets", "asset_details", "TEXT")
        for column, definition in {
            "date": "TEXT",
            "time_12": "TEXT",
            "time_24": "TEXT",
            "actor_role": "TEXT",
            "page": "TEXT",
            "old_value": "TEXT",
            "new_value": "TEXT",
            "reason": "TEXT",
            "status": "TEXT",
            "computer_name": "TEXT",
        }.items():
            self.ensure_column("audit", column, definition)
        for column, definition in {
            "date": "TEXT",
            "time_12": "TEXT",
            "time_24": "TEXT",
            "actor": "TEXT",
            "actor_role": "TEXT",
            "page": "TEXT",
            "status": "TEXT",
            "computer_name": "TEXT",
        }.items():
            self.ensure_column("errors", column, definition)
        self.seed_defaults()

    def ensure_column(self, table, column, definition):
        cols = [r["name"] for r in self.all(f"PRAGMA table_info({table})")]
        if column not in cols:
            self.run(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")

    def seed_defaults(self):
        if not self.one("SELECT value FROM settings WHERE key='backup_folder'"):
            self.run("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", ("backup_folder", str(Path.cwd() / "Backups")))
        for key, folder in {
            "excel_export_folder": "Exports",
            "report_export_folder": "Reports",
            "error_log_folder": "Logs",
            "audit_log_folder": "Logs",
            "system_log_folder": "Logs",
        }.items():
            if not self.one("SELECT value FROM settings WHERE key=?", (key,)):
                self.run("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (key, str(Path.cwd() / folder)))
        if not self.one("SELECT value FROM settings WHERE key='last_operator_scan'"):
            self.run("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", ("last_operator_scan", "F984717"))
        for key, value in {
            "default_export_type": "All",
            "refresh_seconds": "300",
            "current_operator_name": "Unknown",
            "current_operator_role": "Guest",
            "current_page": "Startup",
            "display_density": "Comfortable",
            "frontdesk_scan_only_mode": "No",
        }.items():
            if not self.one("SELECT value FROM settings WHERE key=?", (key,)):
                self.run("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (key, value))
        # v5.1.9 role migration: regular people default to Employee.
        self.run("UPDATE people SET role='Employee' WHERE role IS NULL OR role='' OR role='User'")
        for group_name in ["Employee", "Front Desk", "AP Operator", "Manager", "Admin"]:
            rights = sorted(ROLE_RIGHTS.get(group_name, set()))
            self.run("INSERT OR IGNORE INTO groups(name,rights,protected,notes,created_at,updated_at) VALUES(?,?,?,?,?,?)",
                     (group_name, ",".join(rights), "Yes", "Default protected group.", now_iso(), now_iso()))
        # v5.2.2 Admin format migration: badge starts with 88 and employee ID starts with F.
        existing_admin = self.one("SELECT * FROM people WHERE employee_id IN ('F984717','88984717') OR badge IN ('F984717','88984717') LIMIT 1")
        if existing_admin:
            self.run("UPDATE people SET employee_id='F984717', badge='88984717', role='Admin', status='Active', updated_at=? WHERE id=?", (now_iso(), existing_admin["id"]))

        if not self.find_person("F984717"):
            self.add_person({
                "employee_id": "F984717",
                "badge": "88984717",
                "first_name": "Christopher",
                "last_name": "Schumacher",
                "department": "Asset Protection",
                "status": "Active",
                "shift": "Default",
                "role": "Admin",
                "notes": "Default Admin account."
            }, "System")
        else:
            p = self.find_person("F984717")
            self.run("UPDATE people SET role='Admin', status='Active', updated_at=? WHERE id=?", (now_iso(), p["id"]))
        # Simple starter assets for testing
        for asset in [
            {"asset_type":"Item","barcode":"ITEM-001","name":"Scanner 001","status":"Available","location":"Front Desk","controlled_key_number":"","serial_number":"","asset_details":"{}","notes":"Sample asset"},
            {"asset_type":"Key","barcode":"KEY-001","name":"Key Ring 001","status":"Available","location":"AP","controlled_key_number":"001","serial_number":"","asset_details":asset_details_to_json({"key_set_number":"001","key_ring_serial_number":"KEYRING001","number_of_keys":"1","key_ring_location":"AP","key_ring_use":"Sample controlled key ring","keys":[{"serial":"001","access":"Sample access"}]}),"notes":"Sample key"},
            {"asset_type":"Radio","barcode":"RADIO-001","name":"Radio 001","status":"Available","location":"AP","controlled_key_number":"","serial_number":"RAD001","asset_details":asset_details_to_json({"radio_number":"RADIO-001","factory_serial_number":"RAD001","radio_serial_number":"RAD001","radio_location":"AP","assigned_area":"AP"}),"notes":"Sample radio"},
        ]:
            if not self.find_asset(asset["barcode"]):
                self.add_asset(asset, "System")

    def setting(self, key, default=""):
        r = self.one("SELECT value FROM settings WHERE key=?", (key,))
        return r["value"] if r else default

    def set_setting(self, key, value):
        self.run("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (key, str(value)))

    def person_name(self, p):
        return f"{p['first_name']} {p['last_name']}".strip()

    def group_names(self):
        rows = self.all("SELECT name FROM groups ORDER BY CASE protected WHEN 'Yes' THEN 0 ELSE 1 END, name")
        names = [r["name"] for r in rows]
        return names or list(ROLES)

    def group_rights(self, name):
        row = self.one("SELECT rights FROM groups WHERE name=?", (name,))
        if row:
            return {x for x in str(row["rights"] or "").split(",") if x}
        return ROLE_RIGHTS.get(name, set())

    def find_person(self, scan):
        s = clean(scan)
        return self.one("SELECT * FROM people WHERE lower(employee_id)=lower(?) OR lower(badge)=lower(?)", (s, s))

    def find_asset(self, scan):
        s = clean(scan)
        return self.one("SELECT * FROM assets WHERE lower(barcode)=lower(?) OR lower(controlled_key_number)=lower(?) OR lower(serial_number)=lower(?)", (s, s, s))

    def add_person(self, d, actor):
        self.run("""INSERT INTO people(employee_id,badge,first_name,last_name,department,status,shift,role,notes,created_at,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
            (d.get("employee_id",""), d.get("badge",""), d.get("first_name",""), d.get("last_name",""), d.get("department",""),
             d.get("status","Active"), d.get("shift",""), d.get("role","Employee"), d.get("notes",""), now_iso(), now_iso()))
        self.audit("ADD PERSON", actor, d.get("employee_id") or d.get("badge"), json.dumps(d))

    def update_person(self, row_id, d, actor):
        before = self.one("SELECT * FROM people WHERE id=?", (row_id,))
        self.run("""UPDATE people SET employee_id=?,badge=?,first_name=?,last_name=?,department=?,status=?,shift=?,role=?,notes=?,updated_at=? WHERE id=?""",
            (d.get("employee_id",""), d.get("badge",""), d.get("first_name",""), d.get("last_name",""), d.get("department",""),
             d.get("status","Active"), d.get("shift",""), d.get("role","Employee"), d.get("notes",""), now_iso(), row_id))
        self.audit("UPDATE PERSON", actor, d.get("employee_id") or d.get("badge"), f"Before role={before['role'] if before else ''}; After role={d.get('role')}")

    def add_asset(self, d, actor):
        self.run("""INSERT INTO assets(asset_type,barcode,name,status,location,controlled_key_number,serial_number,asset_details,notes,current_holder_id,current_holder_name,created_at,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (d.get("asset_type","Item"), d.get("barcode",""), d.get("name",""), d.get("status","Available"), d.get("location",""),
             d.get("controlled_key_number",""), d.get("serial_number",""), d.get("asset_details","{}"), d.get("notes",""), "", "", now_iso(), now_iso()))
        self.audit("ADD ASSET", actor, d.get("barcode"), json.dumps(d))

    def update_asset(self, row_id, d, actor):
        self.run("""UPDATE assets SET asset_type=?,barcode=?,name=?,status=?,location=?,controlled_key_number=?,serial_number=?,asset_details=?,notes=?,updated_at=? WHERE id=?""",
            (d.get("asset_type","Item"), d.get("barcode",""), d.get("name",""), d.get("status","Available"), d.get("location",""),
             d.get("controlled_key_number",""), d.get("serial_number",""), d.get("asset_details","{}"), d.get("notes",""), now_iso(), row_id))
        self.audit("UPDATE ASSET", actor, d.get("barcode"), json.dumps(d))

    def open_log(self, barcode):
        return self.one("SELECT * FROM activity WHERE lower(asset_barcode)=lower(?) AND status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') ORDER BY id DESC LIMIT 1", (barcode,))

    def checkout(self, person, asset, due, operator, condition="Good", notes=""):
        log_id = "LOG-" + dt.datetime.now().strftime("%Y%m%d%H%M%S") + "-" + uuid.uuid4().hex[:5].upper()
        try:
            self.begin_immediate()
            if self.open_log(asset["barcode"]):
                self.rollback()
                return None
            self.run("""INSERT INTO activity(log_id,asset_barcode,asset_type,asset_name,employee_id,employee_badge,employee_name,status,checked_out_at,due_back_at,checkout_operator,condition_out,checkout_notes)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (log_id, asset["barcode"], asset["asset_type"], asset["name"], person["employee_id"], person["badge"], self.person_name(person),
                 "OUT", now_iso(), due, operator, condition, notes))
            self.run("UPDATE assets SET status='Checked Out', current_holder_id=?, current_holder_name=?, updated_at=? WHERE id=?",
                     (person["employee_id"], self.person_name(person), now_iso(), asset["id"]))
            self.audit("CHECKOUT", operator, asset["barcode"], f"{self.person_name(person)} due {pretty(due)}")
            self.commit()
            return log_id
        except Exception:
            self.rollback()
            raise

    def return_asset(self, log, operator, condition="Good", notes=""):
        try:
            self.begin_immediate()
            self.run("""UPDATE activity SET status='RETURNED', returned_at=?, return_operator=?, condition_in=?, return_notes=? WHERE id=?""",
                     (now_iso(), operator, condition, notes, log["id"]))
            self.run("UPDATE assets SET status='Available', current_holder_id='', current_holder_name='', updated_at=? WHERE lower(barcode)=lower(?)",
                     (now_iso(), log["asset_barcode"]))
            self.audit("RETURN", operator, log["asset_barcode"], notes)
            self.commit()
        except Exception:
            self.rollback()
            raise

    def current_log_context(self):
        try:
            role = self.setting("current_operator_role", "Unknown")
            page = self.setting("current_page", "Unknown")
        except Exception:
            role = "Unknown"
            page = "Unknown"
        return role, page, os.environ.get("COMPUTERNAME", "")

    def audit(self, action, actor, target, details="", old_value="", new_value="", reason="", status="Success"):
        # v5.1.7: keep audit entries detailed and consistent.
        # Details can be plain text or JSON; reports show the full value.
        if details is None:
            details = ""
        now = dt.datetime.now()
        role, page, computer = self.current_log_context()
        self.run("""INSERT INTO audit(timestamp,date,time_12,time_24,action,actor,actor_role,page,target,old_value,new_value,reason,status,details,computer_name)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (now.isoformat(timespec="seconds"), now.strftime("%Y-%m-%d"), now.strftime("%I:%M:%S %p"), now.strftime("%H:%M:%S"),
             action, actor or "Unknown", role, page, target or "", str(old_value or ""), str(new_value or ""), str(reason or ""),
             status or "Success", str(details), computer))

    def error(self, source, message, details=""):
        now = dt.datetime.now()
        role, page, computer = self.current_log_context()
        actor = self.setting("current_operator_name", "Unknown")
        self.run("""INSERT INTO errors(timestamp,date,time_12,time_24,source,message,details,actor,actor_role,page,status,computer_name)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
            (now.isoformat(timespec="seconds"), now.strftime("%Y-%m-%d"), now.strftime("%I:%M:%S %p"), now.strftime("%H:%M:%S"),
             source, message, details, actor, role, page, "Error", computer))

    def notify_manager(self, event_type, severity, user_involved="", asset_involved="", action_taken="", handled_by="", notes="", reason="", status="New"):
        now = dt.datetime.now()
        self.run("""INSERT INTO manager_notifications(timestamp,time_12,time_24,event_type,severity,user_involved,asset_involved,action_taken,handled_by,status,notes,reason,computer_name)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (now.isoformat(timespec="seconds"), now.strftime("%I:%M:%S %p"), now.strftime("%H:%M:%S"), event_type, severity,
             user_involved, asset_involved, action_taken, handled_by or "Unknown", status, notes, reason, os.environ.get("COMPUTERNAME", "")))

    def add_ap_alert(self, target_type, target_key, target_label, alert_type, severity, note, required_action, actor):
        now = dt.datetime.now()
        self.run("""INSERT INTO ap_alerts(timestamp,time_12,time_24,target_type,target_key,target_label,alert_type,severity,note,required_action,created_by,status,computer_name)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (now.isoformat(timespec="seconds"), now.strftime("%I:%M:%S %p"), now.strftime("%H:%M:%S"), target_type, target_key,
             target_label, alert_type, severity, note, required_action, actor or "Unknown", "Active", os.environ.get("COMPUTERNAME", "")))
        self.audit("AP ALERT ADDED", actor, f"{target_type}:{target_key}", f"{severity} | {alert_type} | {note} | Required action: {required_action}")
        self.notify_manager("Manager alert created", severity, target_label if target_type == "Person" else "", target_label if target_type == "Asset" else "",
                            "AP alert added", actor, note, required_action)

    def active_ap_alerts(self, target_type=None, target_key=None):
        sql = "SELECT * FROM ap_alerts WHERE status='Active'"
        params = []
        if target_type:
            sql += " AND target_type=?"
            params.append(target_type)
        if target_key:
            sql += " AND lower(target_key)=lower(?)"
            params.append(target_key)
        sql += " ORDER BY CASE severity WHEN 'Critical' THEN 0 WHEN 'Warning' THEN 1 ELSE 2 END, timestamp DESC"
        return self.all(sql, tuple(params))

    def update_ap_alert_status(self, alert_id, status, actor, reason=""):
        now = now_iso()
        if status == "Reviewed":
            self.run("UPDATE ap_alerts SET status=?, reviewed_by=?, reviewed_at=?, reason=? WHERE id=?", (status, actor, now, reason, alert_id))
        elif status == "Resolved":
            self.run("UPDATE ap_alerts SET status=?, resolved_by=?, resolved_at=?, reason=? WHERE id=?", (status, actor, now, reason, alert_id))
        else:
            self.run("UPDATE ap_alerts SET status=?, reason=? WHERE id=?", (status, reason, alert_id))
        self.audit("AP ALERT UPDATED", actor, alert_id, f"Status={status}; Reason={reason}")
        self.notify_manager("Manager alert cleared or marked reviewed", "Info", actor, "", f"AP alert set to {status}", actor, f"Alert ID {alert_id}", reason, status="Reviewed")

    def counts(self):
        out = {}
        out["keys_out"] = self.one("SELECT COUNT(*) c FROM activity WHERE status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') AND asset_type='Key'")["c"]
        out["radios_out"] = self.one("SELECT COUNT(*) c FROM activity WHERE status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') AND asset_type='Radio'")["c"]
        out["items_out"] = self.one("SELECT COUNT(*) c FROM activity WHERE status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') AND asset_type NOT IN ('Key','Radio')")["c"]
        out["tablets_out"] = self.one("SELECT COUNT(*) c FROM activity WHERE status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') AND asset_type='Tablet'")["c"]
        out["temp_badges_out"] = self.one("SELECT COUNT(*) c FROM activity WHERE status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') AND asset_type='Temp Badge'")["c"]
        out["late_returns"] = self.one("SELECT COUNT(*) c FROM activity WHERE status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') AND due_back_at < ?", (now_iso(),))["c"]
        out["people"] = self.one("SELECT COUNT(*) c FROM people WHERE status!='Deleted'")["c"]
        out["employee_users"] = self.one("SELECT COUNT(*) c FROM people WHERE role='Employee' AND status!='Deleted'")["c"]
        out["front_desk_users"] = self.one("SELECT COUNT(*) c FROM people WHERE role='Front Desk' AND status!='Deleted'")["c"]
        out["ap_operators"] = self.one("SELECT COUNT(*) c FROM people WHERE role='AP Operator' AND status!='Deleted'")["c"]
        out["managers"] = self.one("SELECT COUNT(*) c FROM people WHERE role='Manager' AND status!='Deleted'")["c"]
        out["admins"] = self.one("SELECT COUNT(*) c FROM people WHERE role='Admin' AND status!='Deleted'")["c"]
        out["inactive_people"] = self.one("SELECT COUNT(*) c FROM people WHERE status NOT IN ('Active','Manual Review')")["c"]
        out["assets"] = self.one("SELECT COUNT(*) c FROM assets WHERE status!='Retired'")["c"]
        out["keys_total"] = self.one("SELECT COUNT(*) c FROM assets WHERE asset_type='Key' AND status!='Retired'")["c"]
        out["radios_total"] = self.one("SELECT COUNT(*) c FROM assets WHERE asset_type='Radio' AND status!='Retired'")["c"]
        out["items_total"] = self.one("SELECT COUNT(*) c FROM assets WHERE asset_type NOT IN ('Key','Radio') AND status!='Retired'")["c"]
        out["available_assets"] = self.one("SELECT COUNT(*) c FROM assets WHERE status='Available'")["c"]
        out["inactive_assets"] = self.one("SELECT COUNT(*) c FROM assets WHERE status IN ('Retired','Repair','Missing')")["c"]
        out["out"] = out["keys_out"] + out["radios_out"] + out["items_out"]
        out["errors"] = self.one("SELECT COUNT(*) c FROM errors WHERE timestamp LIKE ?", (dt.datetime.now().strftime("%Y-%m-%d")+"%",))["c"]
        manager_new = self.one("SELECT COUNT(*) c FROM manager_notifications WHERE status='New'")["c"]
        ap_active = self.one("SELECT COUNT(*) c FROM ap_alerts WHERE status='Active'")["c"]
        out["manager_alerts"] = manager_new + ap_active
        return out

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.app_dir = Path(__file__).resolve().parent
        self.db = DB(self.resolve_db_path())
        self.title(f"{APP_TITLE} - {APP_VERSION}")
        self.geometry("1550x920")
        self.minsize(1240, 760)
        self.colors = {
            "bg": "#eef1f5",
            "panel": "#ffffff",
            "panel_alt": "#f8fafc",
            "black": "#111111",
            "nav": "#171717",
            "nav_button": "#202020",
            "nav_hover": "#2d2d2d",
            "red": "#e21a2c",
            "green": "#198754",
            "purple": "#6f42c1",
            "line": "#d8dee7",
            "muted": "#5f6b7a",
        }
        self.operator = tk.StringVar(value="No operator signed in")
        self.operator_display = tk.StringVar(value="No operator signed in")
        self.operator_role = tk.StringVar(value="Guest")
        self.operator_scan = tk.StringVar()
        self.status = tk.StringVar(value="Ready.")
        self.data_status = tk.StringVar(value="Data: checking...")
        self.training_mode = tk.BooleanVar(value=False)
        self.scan_mode = tk.StringVar(value="Mode: Check Out / Scan Employee")
        self.selected_person = None
        self.selected_asset = None
        self.returning_person = None
        self.open_windows = {}
        self.pages = {}
        self.nav = {}
        self.current_page = "Dashboard"
        self.current_asset_filter = "All"
        self.asset_sort_by = tk.StringVar(value="Type")
        self.asset_sort_dir = tk.StringVar(value="A-Z")
        self.asset_search_var = tk.StringVar()
        self.asset_status_filter = tk.StringVar(value="All Status")
        self.asset_quick_status_var = tk.StringVar(value="Repair")
        self.asset_selected_var = tk.StringVar(value="Select an asset to view details and actions.")
        default_export_type = self.db.setting("default_export_type", "All")
        self.asset_export_type_var = tk.StringVar(value=default_export_type if default_export_type in (["All"] + ASSET_TYPES) else "All")
        self.manager_backup_date_var = tk.StringVar(value="Not backed up")
        self.manager_backup_time_var = tk.StringVar(value="Not backed up")
        self.manager_backup_status_var = tk.StringVar(value="Checking backup status...")
        self.manager_refresh_countdown_var = tk.StringVar(value="Refresh in --")
        self.manager_health_var = tk.StringVar(value="System health: checking...")
        self.manager_notification_status_var = tk.StringVar(value="All")
        self.manager_notification_severity_var = tk.StringVar(value="All")
        self.manager_notification_search_var = tk.StringVar()
        self.manager_show_problems_only = tk.BooleanVar(value=False)
        self.frontdesk_next_step_var = tk.StringVar(value="Sign in, then scan an employee badge to begin.")
        self.frontdesk_mode_var = tk.StringVar(value="Checkout Mode")
        self.frontdesk_preferred_mode = "checkout"
        self.last_frontdesk_scan_kind = ""
        self.manager_focus_vars = {
            "backup": tk.StringVar(value="Checking backup status"),
            "late": tk.StringVar(value="Checking late returns"),
            "alerts": tk.StringVar(value="Checking alerts"),
            "open": tk.StringVar(value="Checking open assets"),
        }
        self.next_auto_refresh_at = time.time() + (self.refresh_interval_ms() / 1000)
        self.report_callback_exception = self.tk_error
        self.style()
        self.apply_icon()
        self.after(250, self.apply_icon)
        self.shell()
        self.build_pages()
        self.v51_apply_enhancements()
        self.restore_operator()
        self.show("Dashboard")
        self.update_data_status()
        self.auto_daily_backup()
        self.tick_refresh_countdown()
        self.after(self.refresh_interval_ms(), self.auto_refresh)


    def tip(self, widget, text):
        try:
            ToolTip(widget, text)
        except Exception:
            pass
        return widget

    def bind_enter(self, widget, command):
        def _run(event=None):
            command()
            return "break"
        widget.bind("<Return>", _run, add="+")
        widget.bind("<KP_Enter>", _run, add="+")
        return widget

    def refresh_scan_mode(self):
        try:
            scan_only = self.scan_only_frontdesk_enabled()
            if self.selected_asset and self.db.open_log(self.selected_asset["barcode"]):
                if scan_only and self.returning_person:
                    self.scan_mode.set("Mode: SCAN-ONLY RETURN READY - click Return Item")
                elif scan_only:
                    self.scan_mode.set("Mode: SCAN-ONLY RETURN - scan the return-by badge, then click Return Item")
                else:
                    self.scan_mode.set("Mode: RETURN - scan the person returning the item, then click Return Item")
            elif self.selected_person and not self.selected_asset:
                self.scan_mode.set("Mode: SCAN-ONLY CHECKOUT - scan the item for this employee" if scan_only else "Mode: CHECK OUT - scan an asset for the selected employee")
            elif self.selected_person and self.selected_asset:
                self.scan_mode.set("Mode: SCAN-ONLY CHECKOUT READY - click Check Out Item" if scan_only else "Mode: CHECK OUT READY - click Check Out Item")
            elif scan_only and getattr(self, "frontdesk_preferred_mode", "checkout") == "return":
                self.scan_mode.set("Mode: SCAN-ONLY RETURN - scan return-by badge and checked-out item")
            elif scan_only:
                self.scan_mode.set("Mode: SCAN-ONLY CHECKOUT - scan employee badge, then item")
            else:
                self.scan_mode.set("Mode: CHECK OUT - scan employee badge first")
        except Exception:
            self.scan_mode.set("Mode: Ready")
        self.update_frontdesk_workflow()

    def update_frontdesk_workflow(self):
        if not hasattr(self, "workflow_var"):
            return
        due_ok = False
        try:
            due_ok = bool(clean(self.due_var.get())) and bool(dt.datetime.fromisoformat(self.due_var.get()))
        except Exception:
            due_ok = False
        scan_only = self.scan_only_frontdesk_enabled()
        mode_ready = scan_only or self.role() != "Guest"
        selected_open_log = None
        if self.selected_asset:
            selected_open_log = self.db.open_log(self.selected_asset["barcode"])
        return_flow = bool(selected_open_log)
        person_done = bool(self.returning_person) if return_flow else bool(self.selected_person)
        due_done = True if return_flow else due_ok
        complete_done = bool(
            self.selected_asset and (
                (return_flow and (self.returning_person or not scan_only))
                or (not return_flow and self.selected_person and due_ok)
            )
        )
        steps = [
            ("Mode", mode_ready),
            ("Return By" if return_flow else "Employee", person_done),
            ("Asset", bool(self.selected_asset)),
            ("Return Ready" if return_flow else "Due Time", due_done),
            ("Complete", complete_done),
        ]
        self.workflow_var.set("  >  ".join([f"[OK] {label}" if done else f"[  ] {label}" for label, done in steps]))
        if hasattr(self, "workflow_step_widgets"):
            first_open = next((i for i, (_, done) in enumerate(steps) if not done), len(steps) - 1)
            for i, ((label, done), widget) in enumerate(zip(steps, self.workflow_step_widgets)):
                if done:
                    bg, fg, text = self.colors["green"], "white", f"[OK] {label}"
                elif i == first_open:
                    bg, fg, text = "#0d6efd", "white", f"[..] {label}"
                else:
                    bg, fg, text = "#e9edf3", "#4b5563", f"[  ] {label}"
                widget.configure(text=text, bg=bg, fg=fg)
        warning = ""
        if self.selected_asset:
            status = clean(self.selected_asset["status"])
            if selected_open_log:
                warning = f"Checked out to {selected_open_log['employee_name']} until {pretty(selected_open_log['due_back_at'])}. Scan return-by badge, then Return Item."
            elif status in ("Missing", "Repair", "Retired"):
                warning = f"Warning: this asset status is {status}. It should not be checked out unless a manager approves."
            elif status != "Available":
                warning = f"Warning: this asset status is {status}."
        self.asset_warning_var.set(warning)
        if hasattr(self, "frontdesk_next_step_var"):
            self.frontdesk_next_step_var.set(self.frontdesk_next_step_text(due_ok, selected_open_log))
        if hasattr(self, "frontdesk_next_step_label"):
            if selected_open_log:
                bg, fg = "#fff3cd", "#664d03"
            elif self.selected_person and self.selected_asset and due_ok:
                bg, fg = "#d1e7dd", "#0f5132"
            elif self.role() == "Guest" and not scan_only:
                bg, fg = "#f8d7da", "#842029"
            else:
                bg, fg = "#e7f1ff", "#084298"
            self.frontdesk_next_step_label.configure(bg=bg, fg=fg)
        if hasattr(self, "frontdesk_mode_label"):
            return_ready = bool(selected_open_log) or (getattr(self, "frontdesk_preferred_mode", "checkout") == "return" and not self.selected_person and not self.selected_asset)
            mode = "Return Mode" if return_ready else "Checkout Mode"
            if scan_only:
                mode = "Scan-Only " + mode
            self.frontdesk_mode_var.set(mode)
            self.frontdesk_mode_label.configure(bg="#fff3cd" if return_ready else "#d1e7dd", fg="#664d03" if return_ready else "#0f5132")
        if hasattr(self, "checkout_button"):
            checkout_ready = bool(self.selected_person and self.selected_asset and due_ok and not selected_open_log)
            self.checkout_button.configure(state=("normal" if checkout_ready else "disabled"))
        if hasattr(self, "return_button"):
            return_ready = bool(selected_open_log and (self.returning_person or not scan_only))
            self.return_button.configure(state=("normal" if return_ready else "disabled"))
        if hasattr(self, "same_employee_button"):
            self.same_employee_button.configure(state=("normal" if self.selected_person else "disabled"))

    def frontdesk_next_step_text(self, due_ok=False, selected_open_log=None):
        scan_only = self.scan_only_frontdesk_enabled()
        if self.role() == "Guest" and not scan_only:
            return "Sign in at the top right before checking out or returning assets."
        if getattr(self, "frontdesk_preferred_mode", "checkout") == "return" and not self.selected_person and not self.selected_asset:
            return "Scan-only return: scan the return-by badge and checked-out item." if scan_only else "Return mode: scan the checked-out item first."
        if self.selected_asset and selected_open_log:
            if self.returning_person:
                return "Ready to return: click Return Item."
            return "Scan-only return: scan the return-by badge, then click Return Item." if scan_only else "Return mode: scan the person returning it, then click Return Item."
        if self.selected_person and self.selected_asset:
            if not due_ok:
                return "Pick a valid due time, then click Check Out Item."
            return "Ready to check out: click Check Out Item."
        if self.selected_person:
            return "Employee selected. Scan the asset barcode, key number, serial, or tag."
        if self.selected_asset:
            return "Asset selected. Scan the employee badge to check it out."
        if scan_only:
            return "Scan-only mode: for checkout scan employee badge then item. For return scan return-by badge and checked-out item."
        return "Scan an employee badge first. For a return, scan the item being returned."

    def set_frontdesk_quick_mode(self, mode):
        scan_only = self.scan_only_frontdesk_enabled()
        if mode == "checkout":
            self.frontdesk_preferred_mode = "checkout"
            self.selected_asset = None
            self.returning_person = None
            self.sel_asset_var.set("No asset selected.")
            self.return_person_var.set("Return by: scan badge required for scan-only returns." if scan_only else "Return by: signed-in operator unless another badge is scanned.")
            self.fd_message.set("Scan-only checkout: scan employee badge, then item." if scan_only else "Checkout mode: scan employee, then scan asset.")
        elif mode == "return":
            self.frontdesk_preferred_mode = "return"
            self.selected_person = None
            self.selected_asset = None
            self.returning_person = None
            self.sel_person_var.set("No employee selected.")
            self.sel_asset_var.set("No asset selected.")
            self.return_person_var.set("Return by: scan the person returning the item.")
            self.fd_message.set("Scan-only return: scan return-by badge and checked-out item." if scan_only else "Return mode: scan the checked-out item first.")
        elif mode == "same_employee":
            self.frontdesk_preferred_mode = "checkout"
            if not self.selected_person:
                self.fd_message.set("Scan or select an employee before using Next Item.")
            else:
                self.selected_asset = None
                self.returning_person = None
                self.sel_asset_var.set("No asset selected. Scan the next item for this employee.")
                self.return_person_var.set("Return by: scan badge required for scan-only returns." if scan_only else "Return by: signed-in operator unless another badge is scanned.")
                self.fd_message.set("Same employee: scan the next asset.")
        self.frontdesk_mode_var.set("Return Mode" if mode == "return" else "Checkout Mode")
        self.refresh_scan_mode()
        try:
            self.scan_entry.focus_set()
        except Exception:
            pass

    def frontdesk_feedback(self, kind, message=None):
        if message:
            try:
                self.fd_message.set(message)
                self.status.set(message)
            except Exception:
                pass
        colors = {
            "success": ("#d1e7dd", "#0f5132"),
            "return": ("#fff3cd", "#664d03"),
            "blocked": ("#f8d7da", "#842029"),
            "info": ("#e7f1ff", "#084298"),
        }
        bg, fg = colors.get(kind, colors["info"])
        try:
            self.bell()
        except Exception:
            pass
        label = getattr(self, "frontdesk_next_step_label", None)
        if label:
            try:
                original_bg = label.cget("bg")
                original_fg = label.cget("fg")
                label.configure(bg=bg, fg=fg)
                self.after(700, lambda: label.configure(bg=original_bg, fg=original_fg))
            except Exception:
                pass

    def bind_frontdesk_shortcuts(self):
        if getattr(self, "_frontdesk_shortcuts_bound", False):
            return
        self._frontdesk_shortcuts_bound = True
        shortcuts = {
            "<F2>": lambda e: self.frontdesk_shortcut(lambda: self.set_frontdesk_quick_mode("checkout")),
            "<F3>": lambda e: self.frontdesk_shortcut(lambda: self.set_frontdesk_quick_mode("return")),
            "<F4>": lambda e: self.frontdesk_shortcut(lambda: self.set_frontdesk_quick_mode("same_employee")),
            "<F8>": lambda e: self.frontdesk_shortcut(self.checkout),
            "<F9>": lambda e: self.frontdesk_shortcut(self.return_scanned),
            "<Escape>": lambda e: self.frontdesk_shortcut(self.clear_frontdesk),
        }
        for key, handler in shortcuts.items():
            self.bind_all(key, handler, add="+")

    def frontdesk_shortcut(self, command):
        if getattr(self, "current_page", "") != "Front Desk":
            return None
        command()
        return "break"

    def default_due_iso(self):
        value = self.db.setting("default_due_time", "18:30")
        try:
            hour, minute = [int(x) for x in value.split(":", 1)]
            return due_today(hour, minute)
        except Exception:
            return due_today(18, 30)

    def set_custom_due_time(self):
        value = simpledialog.askstring(
            "Custom Due Time",
            "Enter due time as HH:MM for today, or YYYY-MM-DD HH:MM for a specific date.",
            parent=self,
        )
        value = clean(value)
        if not value:
            return
        try:
            if "-" not in value and ":" in value:
                hour, minute = [int(x) for x in value.split(":", 1)]
                if not (0 <= hour <= 23 and 0 <= minute <= 59):
                    raise ValueError
                self.due_var.set(due_today(hour, minute))
            else:
                normalized = value.replace(" ", "T", 1)
                due_dt = dt.datetime.fromisoformat(normalized)
                self.due_var.set(due_dt.isoformat(timespec="seconds"))
            self.status.set("Custom due time set.")
        except Exception:
            messagebox.showwarning("Invalid Due Time", "Use HH:MM, like 18:30, or YYYY-MM-DD HH:MM.")

    def unknown_person_prompt(self, scan_value):
        if not messagebox.askyesno("Person Not Found", f"'{scan_value}' was not found as an employee/person.\\n\\nDo you want to add this person now?"):
            self.fd_message.set(f"Unknown scan: {scan_value}")
            return
        self.popup_person_prefill(scan_value)

    def popup_person_prefill(self, scan_value):
        if not self.require("people_edit"):
            messagebox.showwarning("Access Denied", "You do not have access to add people. Ask an Admin or Manager.")
            return
        self.popup_person_prefill_value = scan_value
        self.popup_person(None)



    def config_path(self):
        return self.app_dir / CONFIG_FILE

    def load_config(self):
        path = self.config_path()
        if path.exists():
            try:
                return json.loads(path.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def save_config(self, cfg):
        self.config_path().write_text(json.dumps(cfg, indent=2), encoding="utf-8")

    def resolve_db_path(self):
        cfg = self.load_config()
        custom = clean(cfg.get("data_file", ""))
        if custom:
            p = Path(custom)
            try:
                p.parent.mkdir(parents=True, exist_ok=True)
                return p
            except Exception:
                pass
        return self.app_dir / DB_FILE

    def current_data_file(self):
        try:
            return Path(self.db.path).resolve()
        except Exception:
            return Path(self.app_dir / DB_FILE).resolve()

    def is_unc_path(self, path):
        return str(path).startswith("\\\\")

    def is_shared_data_file(self):
        cfg = self.load_config()
        custom = clean(cfg.get("data_file", ""))
        return bool(custom) or self.is_unc_path(self.current_data_file())

    def refresh_interval_ms(self):
        configured = clean(self.db.setting("refresh_seconds", ""))
        if configured.isdigit():
            return max(5, min(600, int(configured))) * 1000
        return 300000

    def update_data_status(self):
        if not hasattr(self, "data_status"):
            return
        path = self.current_data_file()
        mode = "Shared Data" if self.is_shared_data_file() else "Local Data"
        refresh = int(self.refresh_interval_ms() / 1000)
        self.data_status.set(f"{mode} | Auto-refresh {refresh}s | Last {dt.datetime.now().strftime('%I:%M:%S %p')} | {path.name}")

    def folder_setting(self, key, default_name):
        folder = Path(self.db.setting(key, str(self.app_dir / default_name)))
        try:
            folder.mkdir(parents=True, exist_ok=True)
        except Exception:
            return self.app_dir / default_name
        return folder

    def show_data_location(self):
        path = self.current_data_file()
        cfg = self.load_config()
        lines = [
            "DATABASE / DATA FILE LOCATION",
            "=" * 72,
            f"Current data file:",
            str(path),
            "",
            f"Config file:",
            str(self.config_path()),
            "",
            f"Custom data file set:",
            "Yes" if clean(cfg.get("data_file", "")) else "No",
            "",
            f"Network/shared mode:",
            f"{'Yes - shared data file' if self.is_shared_data_file() else 'No - local data file'}; auto-refresh every {int(self.refresh_interval_ms() / 1000)} seconds",
            "",
            "Notes:",
            "- This .db file stores people, assets, checkouts, returns, audit logs, errors, and settings.",
            "- For multiple computers, use one shared database on a UNC path like \\\\SERVER\\Share\\macys_ap_data.db.",
            "- Use Change Data File Location to move the database to another folder or network share.",
            "- After changing the data file location, restart the app.",
        ]
        self.system_text.delete("1.0", "end")
        self.system_text.insert("1.0", "\n".join(lines))
        self.db.audit("VIEW DATA LOCATION", self.actor(), "System", str(path))

    def open_data_folder(self):
        path = self.current_data_file()
        folder = path.parent
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)
            else:
                messagebox.showinfo("Data Folder", str(folder))
        except Exception:
            messagebox.showinfo("Data Folder", str(folder))
        self.db.audit("OPEN DATA FOLDER", self.actor(), "System", str(folder))

    def change_data_location(self):
        if not self.require("system"):
            return
        current = self.current_data_file()
        message = (
            "Choose the new .db data file location.\n\n"
            "The current database will be copied there and the app will use it after restart.\n\n"
            "For network use, choose a shared UNC path like \\\\SERVER\\Share\\macys_ap_data.db."
        )
        messagebox.showinfo("Change Data File Location", message)
        target = filedialog.asksaveasfilename(
            title="Choose new database/data file location",
            defaultextension=".db",
            initialfile="macys_ap_data.db",
            filetypes=[("SQLite database", "*.db"), ("All files", "*.*")]
        )
        if not target:
            return
        target = Path(target)
        if target.resolve() == current.resolve():
            messagebox.showinfo("No Change", "That is already the current data file.")
            return
        if not messagebox.askyesno(
            "Confirm Data File Move",
            f"Copy current database from:\n{current}\n\nTo:\n{target}\n\nContinue?"
        ):
            return
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            self.db.conn.commit()
            shutil.copy2(current, target)
            cfg = self.load_config()
            cfg["data_file"] = str(target)
            self.save_config(cfg)
            self.db.audit("CHANGE DATA LOCATION", self.actor(), "System", f"New data file set for next startup: {target}")
            messagebox.showinfo(
                "Data Location Updated",
                f"Database copied to:\n{target}\n\nRestart the app to use the new location."
            )
            self.show_data_location()
        except Exception:
            self.log_exception("Change Data Location")



    def style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        density = clean(self.db.setting("display_density", "Comfortable"))
        if density == "Compact":
            button_padding = (8, 5)
            tree_height = 26
            tree_font = ("Segoe UI", 9)
        elif density == "Spacious":
            button_padding = (12, 9)
            tree_height = 34
            tree_font = ("Segoe UI", 10)
        else:
            button_padding = (10, 7)
            tree_height = 30
            tree_font = ("Segoe UI", 10)
        s.configure("TButton", padding=button_padding, font=("Segoe UI", 10, "bold"), relief="flat")
        s.map("TButton", background=[("active", "#eef2f7")])
        s.configure("Red.TButton", background=self.colors["red"], foreground="white", relief="flat")
        s.configure("Green.TButton", background=self.colors["green"], foreground="white", relief="flat")
        s.configure("Amber.TButton", background="#ffc107", foreground="#111111", relief="flat")
        s.configure("Secondary.TButton", background="#e9edf3", foreground="#17202a", relief="flat")
        s.configure("Panel.TLabel", background=self.colors["panel"], font=("Segoe UI", 10))
        s.configure("Title.TLabel", background=self.colors["panel"], font=("Segoe UI", 18, "bold"))
        s.configure("Sub.TLabel", background=self.colors["panel"], foreground="#555")
        s.configure("Treeview", rowheight=tree_height, font=tree_font, fieldbackground="white", borderwidth=0)
        s.map("Treeview", background=[("selected", "#cfe2ff")], foreground=[("selected", "#111111")])
        s.configure("Treeview.Heading", background="#111", foreground="white", font=("Segoe UI", 10, "bold"), padding=(8, 7), anchor="w")

    def apply_icon(self, win=None):
        """Apply Macy star icon to main window and popups. The EXE also uses this icon."""
        win = win or self
        ico = self.app_dir / "macys_star_icon.ico"
        png = self.app_dir / "macys_star_icon.png"
        try:
            if ico.exists():
                try:
                    win.iconbitmap(default=str(ico))
                except TypeError:
                    win.iconbitmap(str(ico))
                except Exception:
                    win.iconbitmap(str(ico))
        except Exception:
            pass
        try:
            if png.exists():
                if not hasattr(self, "_icon_photo"):
                    self._icon_photo = tk.PhotoImage(file=str(png))
                win.iconphoto(True, self._icon_photo)
        except Exception:
            pass

    def header_star(self, parent):
        size = 54
        png = self.app_dir / "macys_star_icon.png"
        if Image is not None and ImageTk is not None and png.exists():
            try:
                img = Image.open(png).convert("RGBA")
                transparent_pixels = []
                for r, g, b, a in img.getdata():
                    if r > 245 and g > 245 and b > 245:
                        transparent_pixels.append((r, g, b, 0))
                    else:
                        transparent_pixels.append((r, g, b, a))
                img.putdata(transparent_pixels)
                resample = getattr(getattr(Image, "Resampling", Image), "LANCZOS", getattr(Image, "BICUBIC", 3))
                img = img.resize((size, size), resample)
                self._header_star_photo = ImageTk.PhotoImage(img)
                label = tk.Label(parent, image=self._header_star_photo, bg=self.colors["black"], width=size + 8, height=size + 8, bd=0, highlightthickness=0)
                self.tip(label, "Macy's AP")
                return label
            except Exception:
                pass
        canvas = tk.Canvas(parent, width=size, height=size, bg=self.colors["black"], highlightthickness=0, bd=0)
        cx = cy = size / 2
        outer = 23
        inner = 10
        points = []
        for i in range(10):
            radius = outer if i % 2 == 0 else inner
            angle = -math.pi / 2 + i * math.pi / 5
            points.extend((cx + radius * math.cos(angle), cy + radius * math.sin(angle)))
        canvas.create_polygon(points, fill=self.colors["red"], outline=self.colors["red"], width=1, joinstyle="round")
        self.tip(canvas, "Macy's AP")
        return canvas


    def shell(self):
        top = tk.Frame(self, bg=self.colors["black"], height=104)
        top.pack(fill="x")
        top.pack_propagate(False)

        brand = tk.Frame(top, bg=self.colors["black"])
        brand.pack(side="left", fill="both", expand=True, padx=(20, 8))
        logo_slot = tk.Frame(brand, bg=self.colors["black"], width=72)
        logo_slot.pack(side="left", fill="y")
        logo_slot.pack_propagate(False)
        logo = self.header_star(logo_slot)
        logo.place(relx=0.5, rely=0.5, anchor="center")
        title_area = tk.Frame(brand, bg=self.colors["black"])
        title_area.pack(side="left", fill="both", expand=True, padx=(8, 0))
        tk.Label(title_area, text=APP_TITLE, bg=self.colors["black"], fg="white", font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(18,2))
        tk.Label(title_area, text=f"Version {APP_VERSION}", bg=self.colors["black"], fg="#f2f2f2", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(title_area, textvariable=self.data_status, bg=self.colors["black"], fg="#d7f7df", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(3,0))

        op = tk.Frame(top, bg="#1f1f1f", width=690, highlightbackground="#404040", highlightthickness=1)
        op.pack(side="right", fill="y", padx=(8, 14), pady=10)
        op.pack_propagate(False)
        tk.Label(op, text="Operator", bg="#1f1f1f", fg="#cfd4dc", font=("Segoe UI", 9, "bold")).grid(row=0,column=0,sticky="w",padx=10,pady=(8,2))
        tk.Label(op, text="Badge / ID", bg="#1f1f1f", fg="#cfd4dc", font=("Segoe UI", 9, "bold")).grid(row=0,column=1,sticky="w",padx=8,pady=(8,2))
        self.operator_badge = tk.Label(op, textvariable=self.operator_display, bg="#6c757d", fg="white", font=("Segoe UI", 10, "bold"), anchor="w", padx=10, pady=4)
        self.operator_badge.grid(row=1,column=0,sticky="ew",padx=10)
        self.role_label = tk.Label(op, text="", bg="#6c757d", fg="white", font=("Segoe UI",9,"bold"), anchor="w", padx=10, pady=3)
        self.role_label.grid(row=2,column=0,sticky="ew",padx=10,pady=(3,8))
        self.training_label = tk.Label(op, text="Live Mode", bg="#198754", fg="white", font=("Segoe UI",9,"bold"), padx=10, pady=3)
        self.training_label.grid(row=2,column=1,sticky="ew",padx=8,pady=(3,8))
        self.operator_entry = ttk.Entry(op, textvariable=self.operator_scan, width=18)
        self.operator_entry.grid(row=1,column=1,padx=8,sticky="ew")
        self.bind_enter(self.operator_entry, self.operator_login)
        self.tip(self.operator_entry, "Scan or type the operator badge/ID, then press Enter or click Sign In.")
        self.sign_in_button = self.tip(ttk.Button(op, text="Sign In", command=self.operator_login), "Signs in the operator using the badge/ID in the box.")
        self.sign_in_button.grid(row=1,column=2,rowspan=2,padx=(6,3),pady=(0,8),sticky="ns")
        self.sign_out_button = self.tip(ttk.Button(op, text="Sign Out", command=self.operator_logout), "Signs out the current operator and returns the app to Guest mode.")
        self.sign_out_button.grid(row=1,column=2,rowspan=2,padx=(6,3),pady=(0,8),sticky="ns")
        self.training_button = self.tip(ttk.Button(op, text="Training Off", style="Secondary.TButton", command=self.toggle_training_mode), "Turn Training/Demo Mode on or off. Training actions are marked in logs.")
        self.training_button.grid(row=1,column=3,rowspan=2,padx=(3,10),pady=(0,8),sticky="ns")
        op.grid_columnconfigure(0, weight=1)
        op.grid_columnconfigure(1, weight=1)

        main = tk.Frame(self, bg=self.colors["bg"])
        main.pack(fill="both", expand=True)
        nav = tk.Frame(main, bg=self.colors["nav"], width=258)
        nav.pack(side="left", fill="y")
        nav.pack_propagate(False)
        tk.Label(nav, text="MENU", bg=self.colors["nav"], fg="#cccccc", font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=18, pady=(18,6))
        content = tk.Frame(main, bg=self.colors["bg"])
        content.pack(side="left", fill="both", expand=True, padx=14, pady=14)
        self.content = content
        nav_items = [
            ("WORKFLOW", None),
            ("Dashboard", "DB"),
            ("Front Desk", "FD"),
            ("Users", "US"),
            ("Assets", "AS"),
            ("Search", "SE"),
            ("OVERSIGHT", None),
            ("Manager", "MG"),
            ("Reports", "RP"),
            ("Logs", "LG"),
            ("SETUP", None),
            ("System", "SY"),
        ]
        for name, icon in nav_items:
            if icon is None:
                tk.Label(nav, text=name, bg=self.colors["nav"], fg="#8f99a8", font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=18, pady=(16,4))
                continue
            image = self.nav_icon(name)
            b = tk.Button(nav, text=name, image=image, compound="left", anchor="w", bg=self.colors["nav_button"], fg="white", activebackground=self.colors["nav_hover"], activeforeground="white", bd=0, relief="flat", font=("Segoe UI",10,"bold"), padx=16, pady=9, cursor="hand2", command=lambda n=name:self.show(n))
            b.pack(fill="x", padx=8, pady=2)
            b.image = image
            self.tip(b, f"Open the {name} page.")
            b.bind("<Enter>", lambda e, btn=b, n=name: self.nav_hover(btn, n, True), add="+")
            b.bind("<Leave>", lambda e, btn=b, n=name: self.nav_hover(btn, n, False), add="+")
            self.nav[name] = b
        status = tk.Label(self, textvariable=self.status, anchor="w", bg="#fff3cd", fg="#111", font=("Segoe UI", 9, "bold"), padx=10, pady=6)
        status.pack(fill="x")

    def nav_icon(self, name):
        if not hasattr(self, "_nav_icon_cache"):
            self._nav_icon_cache = {}
        if name in self._nav_icon_cache:
            return self._nav_icon_cache[name]
        size = 22
        if Image is None or ImageTk is None or ImageDraw is None:
            img = tk.PhotoImage(width=size, height=size)
            self._nav_icon_cache[name] = img
            return img
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        white = (255, 255, 255, 235)
        muted = (255, 255, 255, 150)
        if name == "Dashboard":
            for x, y in [(4, 12), (9, 8), (14, 5)]:
                d.rounded_rectangle([x, y, x + 3, 18], radius=1, fill=white)
        elif name == "Front Desk":
            d.rounded_rectangle([3, 5, 19, 15], radius=2, outline=white, width=2)
            d.line([6, 18, 16, 18], fill=white, width=2)
            d.line([8, 9, 16, 9], fill=muted, width=2)
        elif name == "Users":
            d.ellipse([7, 4, 15, 12], outline=white, width=2)
            d.arc([4, 11, 18, 24], 205, 335, fill=white, width=2)
        elif name == "Assets":
            d.polygon([(11, 3), (18, 7), (18, 15), (11, 19), (4, 15), (4, 7)], outline=white)
            d.line([4, 7, 11, 11, 18, 7], fill=muted, width=1)
            d.line([11, 11, 11, 19], fill=muted, width=1)
        elif name == "Search":
            d.ellipse([4, 4, 14, 14], outline=white, width=2)
            d.line([13, 13, 18, 18], fill=white, width=2)
        elif name == "Manager":
            d.polygon([(11, 3), (18, 6), (17, 13), (11, 20), (5, 13), (4, 6)], outline=white)
            d.line([8, 11, 10, 14, 15, 8], fill=white, width=2)
        elif name == "Reports":
            d.rounded_rectangle([5, 3, 17, 19], radius=2, outline=white, width=2)
            for y in (8, 12, 16):
                d.line([8, y, 15, y], fill=muted, width=1)
        elif name == "Logs":
            d.rounded_rectangle([4, 4, 18, 18], radius=2, outline=white, width=2)
            for y in (8, 11, 14):
                d.line([7, y, 15, y], fill=muted, width=1)
        else:
            d.ellipse([5, 5, 17, 17], outline=white, width=2)
            d.ellipse([9, 9, 13, 13], fill=white)
            for x1, y1, x2, y2 in [(11, 2, 11, 5), (11, 17, 11, 20), (2, 11, 5, 11), (17, 11, 20, 11)]:
                d.line([x1, y1, x2, y2], fill=white, width=2)
        photo = ImageTk.PhotoImage(img)
        self._nav_icon_cache[name] = photo
        return photo

    def nav_hover(self, button, name, inside):
        if getattr(self, "active_nav_name", "") == name:
            return
        button.configure(bg=self.colors["nav_hover"] if inside else self.colors["nav_button"])

    def set_nav_button_state(self, name, active):
        button = self.nav.get(name)
        if not button:
            return
        if active:
            button.configure(bg=self.colors["red"], fg="white", activebackground=self.colors["red"], font=("Segoe UI", 10, "bold"))
        else:
            button.configure(bg=self.colors["nav_button"], fg="white", activebackground=self.colors["nav_hover"], font=("Segoe UI", 10, "bold"))

    def role_badge_colors(self, role):
        colors = {
            "Guest": ("#6c757d", "white"),
            "Employee": ("#3d6f8e", "white"),
            "Front Desk": ("#0d6efd", "white"),
            "AP Operator": ("#6f42c1", "white"),
            "Manager": ("#b36b00", "white"),
            "Admin": (self.colors["red"], "white"),
        }
        return colors.get(role or "Guest", ("#6c757d", "white"))

    def update_operator_badges(self):
        role = self.role()
        role_bg, role_fg = self.role_badge_colors(role)
        raw_operator = self.operator.get() or "No operator signed in"
        signed_in = role != "Guest" and raw_operator != "No operator signed in"
        scan_only = self.scan_only_frontdesk_enabled()
        if hasattr(self, "operator_display"):
            self.operator_display.set("Scan-Only Front Desk" if scan_only and not signed_in else raw_operator)
        if scan_only and not signed_in:
            role_bg, role_fg = "#0d6efd", "white"
        operator_bg = self.colors["green"] if signed_in else "#6c757d"
        if scan_only and not signed_in:
            operator_bg = "#0d6efd"
        if hasattr(self, "operator_badge"):
            self.operator_badge.configure(bg=operator_bg, fg="white")
        if hasattr(self, "role_label"):
            role_text = "Mode: Scan-Only Front Desk" if scan_only and not signed_in else f"Role: {role}"
            self.role_label.configure(text=role_text, bg=role_bg, fg=role_fg)
        if hasattr(self, "sign_in_button") and hasattr(self, "sign_out_button"):
            if signed_in:
                self.sign_in_button.grid_remove()
                self.sign_out_button.grid()
            else:
                self.sign_out_button.grid_remove()
                self.sign_in_button.grid()
        self.update_training_badge()

    def update_training_badge(self):
        training = self.is_training()
        try:
            self.training_label.configure(
                text="TRAINING MODE ON" if training else "Live Mode",
                bg="#ffc107" if training else self.colors["green"],
                fg="#111111" if training else "white",
            )
        except Exception:
            pass
        try:
            self.training_button.configure(
                text="Training On" if training else "Training Off",
                style="Amber.TButton" if training else "Secondary.TButton",
            )
        except Exception:
            pass

    def page(self, name):
        f = tk.Frame(self.content, bg=self.colors["bg"])
        self.pages[name] = f
        return f

    def page_title_actions(self, parent, title):
        actions = {
            "Dashboard": [("Refresh", self.refresh_dashboard)],
            "Front Desk": [("Start Over", self.clear_frontdesk)],
            "Manager": [("Refresh Now", self.manual_manager_refresh)],
            "System": [("Data Location", self.show_data_location)],
        }
        return actions.get(title, [])

    def chip_button(self, parent, text, command, active=False):
        bg = "#0d6efd" if active else "#e9edf3"
        fg = "white" if active else "#17202a"
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground="#0b5ed7" if active else "#dde5ef",
            activeforeground="white" if active else "#17202a",
            relief="flat",
            bd=0,
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=5,
            cursor="hand2",
        )
        return btn

    def set_chip_active(self, button, active):
        try:
            button.configure(
                bg="#0d6efd" if active else "#e9edf3",
                fg="white" if active else "#17202a",
                activebackground="#0b5ed7" if active else "#dde5ef",
                activeforeground="white" if active else "#17202a",
            )
        except Exception:
            pass

    def more_button(self, parent, actions, text="More"):
        btn = tk.Menubutton(parent, text=text, bg="#e9edf3", fg="#17202a", activebackground="#dde5ef", relief="flat", font=("Segoe UI", 10, "bold"), padx=10, pady=6, cursor="hand2")
        menu = tk.Menu(btn, tearoff=0)
        for item in actions:
            if item is None:
                menu.add_separator()
                continue
            label, command = item[0], item[1]
            enabled = item[2] if len(item) > 2 else True
            menu.add_command(label=label, command=command, state=("normal" if enabled else "disabled"))
        btn.configure(menu=menu)
        return btn

    def status_badge(self, parent, label, var, color="#0d6efd"):
        badge = tk.Frame(parent, bg="#f8fafc", highlightbackground=self.colors["line"], highlightthickness=1)
        badge.pack(side="left", fill="x", expand=True, padx=(0,8), pady=2)
        tk.Label(badge, text=label, bg="#f8fafc", fg=self.colors["muted"], font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=10, pady=(7,0))
        tk.Label(badge, textvariable=var, bg="#f8fafc", fg=color, font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=10, pady=(0,7))
        return badge

    def count_accent(self, key):
        if key in ("late_returns", "errors", "inactive_people"):
            return self.colors["red"]
        if key in ("inactive_assets", "manager_alerts"):
            return "#b36b00"
        if key in ("available_assets",):
            return self.colors["green"]
        if key in ("keys_total", "keys_out"):
            return "#6f42c1"
        if key in ("radios_total", "radios_out"):
            return "#0d6efd"
        if key in ("tablets_out", "temp_badges_out"):
            return "#3d6f8e"
        return "#17202a"

    def panel(self, parent, title, sub=""):
        p = tk.Frame(parent, bg=self.colors["panel"], highlightbackground=self.colors["line"], highlightthickness=1, bd=0)
        head = tk.Frame(p, bg=self.colors["panel"])
        head.pack(fill="x", padx=16, pady=(14,6))
        title_col = tk.Frame(head, bg=self.colors["panel"])
        title_col.pack(side="left", fill="x", expand=True)
        tk.Label(title_col, text=title, bg=self.colors["panel"], fg="#17202a", font=("Segoe UI", 18, "bold")).pack(anchor="w")
        if sub:
            tk.Label(title_col, text=sub, bg=self.colors["panel"], fg=self.colors["muted"], font=("Segoe UI", 10), anchor="w", justify="left", wraplength=980).pack(anchor="w", fill="x", pady=(2,0))
        actions = self.page_title_actions(parent, title)
        if actions:
            action_bar = tk.Frame(head, bg=self.colors["panel"])
            action_bar.pack(side="right", padx=(12,0))
            for label, cmd in actions:
                self.tip(ttk.Button(action_bar, text=label, command=cmd), f"{label} {title}.").pack(side="left", padx=3)
        tk.Frame(p, bg=self.colors["line"], height=1).pack(fill="x", padx=16, pady=(0,0))
        body = tk.Frame(p, bg=self.colors["panel"])
        body.pack(fill="both", expand=True, padx=16, pady=(12,16))
        p.body = body
        return p

    def tree(self, parent, cols, height=14):
        wrap = tk.Frame(parent, bg=self.colors["panel"])
        wrap.pack(fill="both", expand=True)
        t = ttk.Treeview(wrap, columns=cols, displaycolumns=cols, show="headings", height=height)
        y = ttk.Scrollbar(wrap, orient="vertical", command=t.yview)
        x = ttk.Scrollbar(wrap, orient="horizontal", command=t.xview)
        t.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        for c in cols:
            t.heading(c, text=c, command=lambda cc=c, tt=t: self.sort_tree(tt, cc))
            t.column(c, width=max(120, len(c)*12), anchor="w")
        t.grid(row=0,column=0,sticky="nsew")
        y.grid(row=0,column=1,sticky="ns")
        x.grid(row=1,column=0,sticky="ew")
        wrap.grid_rowconfigure(0, weight=1)
        wrap.grid_columnconfigure(0, weight=1)
        def stop_uncontrolled_drag(event):
            if t.identify_region(event.x, event.y) == "separator":
                return "break"
            return None
        t.bind("<ButtonPress-1>", stop_uncontrolled_drag, add="+")
        t.bind("<B1-Motion>", stop_uncontrolled_drag, add="+")
        return t

    def fill(self, tree, rows):
        tree.delete(*tree.get_children())
        cols = list(tree["columns"])
        status_idx = cols.index("Status") if "Status" in cols else None
        type_idx = cols.index("Type") if "Type" in cols else None
        for tag, color in {
            "row_even": "#ffffff",
            "row_odd": "#f6f8fb",
            "audit": "#eef6ff",
            "audit_odd": "#e2f0ff",
            "error": "#ffecec",
            "error_odd": "#ffe0e0",
            "manager": "#fff8e6",
            "manager_odd": "#fff2cf",
            "ap_alert": "#fff1f1",
            "ap_alert_odd": "#ffe2e2",
            "person": "#f1f7ff",
            "person_odd": "#e7f1ff",
            "asset": "#f6f0ff",
            "asset_odd": "#efe4ff",
            "available": "#f4fff7",
            "available_odd": "#ecfff2",
            "checked_out": "#fff8e6",
            "checked_out_odd": "#fff2cf",
            "missing": "#ffecec",
            "missing_odd": "#ffe0e0",
            "repair": "#fff3cd",
            "repair_odd": "#ffe9a8",
            "retired": "#eeeeee",
            "retired_odd": "#e3e6ea",
            "returned": "#f4fff7",
            "returned_odd": "#ecfff2",
            "overdue": "#ffecec",
            "overdue_odd": "#ffe0e0",
        }.items():
            try:
                tree.tag_configure(tag, background=color)
            except Exception:
                pass
        try:
            tree.tag_configure("empty_state", background="#ffffff", foreground=self.colors["muted"])
        except Exception:
            pass
        if not rows:
            values = ["No records found."] + [""] * max(0, len(cols) - 1)
            tree.insert("", "end", values=values, tags=("row_even", "empty_state"))
            return
        for idx, r in enumerate(rows):
            parity = "odd" if idx % 2 else "even"
            tags = (f"row_{parity}",)
            if status_idx is not None and status_idx < len(r):
                status = str(r[status_idx] or "").lower().replace(" ", "_")
                status_tag = f"{status}_{parity}" if parity == "odd" else status
                tags = (status_tag,)
            elif type_idx is not None and type_idx < len(r):
                row_type = str(r[type_idx] or "").lower().replace(" ", "_").replace("/", "_")
                type_tag = f"{row_type}_{parity}" if parity == "odd" else row_type
                tags = (type_tag,)
            tree.insert("", "end", values=r, tags=tags)

    def sort_tree(self, tree, col):
        # v5.1.6 fix: keep sort keys one type so Python never compares float to str.
        cols = list(tree["columns"])
        idx = cols.index(col)
        children = tree.get_children()
        if not children or all("empty_state" in tree.item(item, "tags") for item in children):
            return
        reverse = not getattr(tree, "_sort_reverse", {}).get(col, False)
        data = []
        for item in children:
            vals = list(tree.item(item, "values"))
            val = vals[idx] if idx < len(vals) else ""
            raw = str(val).strip()
            try:
                key = (0, float(raw.replace(",", "")))
            except Exception:
                key = (1, raw.lower())
            data.append((key, vals))
        data.sort(reverse=reverse, key=lambda x: x[0])
        self.fill(tree, [vals for _, vals in data])
        tree._sort_reverse = {col: reverse}
        for c in cols:
            label = c + (" v" if c == col and reverse else " ^" if c == col else "")
            tree.heading(c, text=label, command=lambda cc=c, tt=tree: self.sort_tree(tt, cc))
        self.status.set(f"Sorted by {col}.")

    def build_pages(self):
        self.dashboard_page()
        self.frontdesk_page()
        self.people_page()
        self.assets_page()
        self.search_page()
        self.manager_page()
        self.reports_page()
        self.logs_page()
        self.system_page()
        self.patch_assets_delete_button()
        self.apply_default_tooltips()



    def mark_required(self, widget, is_bad):
        try:
            widget.configure(background="#fff3cd" if is_bad else "white")
        except Exception:
            pass

    def selected_tree_values(self, tree):
        sel = tree.selection()
        if not sel:
            return None
        if "empty_state" in tree.item(sel[0], "tags"):
            return None
        return tree.item(sel[0], "values")

    def copy_tree_row(self, tree):
        values = self.selected_tree_values(tree)
        if not values:
            messagebox.showinfo("Copy Row", "Select a row first.")
            return
        text = "\t".join(str(v) for v in values)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status.set("Selected row copied.")
        self.db.audit("COPY TABLE ROW", self.actor(), self.current_page, text)

    def show_context_menu(self, event, tree, actions):
        row_id = tree.identify_row(event.y)
        if row_id:
            tree.selection_set(row_id)
            try:
                tree.focus(row_id)
            except Exception:
                pass
        menu = tk.Menu(self, tearoff=0)
        added = False
        for item in actions:
            if item is None:
                if added:
                    menu.add_separator()
                continue
            label, cmd = item[0], item[1]
            enabled = item[2] if len(item) > 2 else True
            menu.add_command(label=label, command=cmd, state=("normal" if enabled else "disabled"))
            added = True
        if added:
            menu.tk_popup(event.x_root, event.y_root)

    def bind_context_menu(self, tree, builder):
        def popup(event):
            row_id = tree.identify_row(event.y)
            if row_id:
                tree.selection_set(row_id)
                try:
                    tree.focus(row_id)
                except Exception:
                    pass
            actions = builder(tree)
            self.show_context_menu(event, tree, actions)
            return "break"
        tree.bind("<Button-3>", popup, add="+")
        tree.bind("<Button-2>", popup, add="+")
        return tree

    def dashboard_card_context(self, event, key, title):
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Open Details", command=lambda: self.dashboard_detail_popup(key, title))
        menu.add_command(label="Refresh Dashboard", command=self.refresh_dashboard)
        menu.add_command(label="Copy Count", command=lambda: self.copy_dashboard_count(key, title))
        menu.tk_popup(event.x_root, event.y_root)
        return "break"

    def copy_dashboard_count(self, key, title):
        value = self.metric_vars[key].get() if hasattr(self, "metric_vars") and key in self.metric_vars else ""
        text = f"{title}: {value}"
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status.set(f"Copied {text}.")

    def selected_person_from_tree(self):
        if not hasattr(self, "people_tree"):
            return None
        values = self.selected_tree_values(self.people_tree)
        if not values:
            return None
        return self.db.find_person(values[0])

    def set_selected_person_status(self, status):
        if not self.require("people_edit"):
            return
        person = self.selected_person_from_tree()
        if not person:
            messagebox.showinfo("Select User", "Select a user first.")
            return
        if person["status"] == status:
            messagebox.showinfo("User Status", f"User is already {status}.")
            return
        reason = simpledialog.askstring("Reason Required", f"Reason for setting {self.db.person_name(person)} to {status}:", parent=self)
        if not clean(reason):
            messagebox.showwarning("Reason Required", "A reason is required.")
            return
        d = {k: person[k] for k in ["employee_id","badge","first_name","last_name","department","status","shift","role","notes"]}
        old_status = d["status"]
        d["status"] = status
        self.db.update_person(person["id"], d, self.actor())
        self.db.audit("USER STATUS UPDATED", self.actor(), person["employee_id"], f"{old_status} -> {status}; Reason: {clean(reason)}", old_status, status, clean(reason))
        self.db.notify_manager("User status changed", "Info", self.db.person_name(person), "", f"{old_status} -> {status}", self.actor(), clean(reason), clean(reason), status="Reviewed")
        self.refresh_people()

    def selected_current_log(self, tree=None):
        tree = tree or getattr(self, "current_tree", None)
        if not tree:
            return None
        values = self.selected_tree_values(tree)
        if not values:
            return None
        asset_text = str(values[2] if tree is getattr(self, "current_tree", None) else values[3] if len(values) > 3 else "")
        barcode = clean(asset_text.split("/", 1)[0])
        return self.db.open_log(barcode) if barcode else None

    def open_current_tree_asset(self):
        log = self.selected_current_log()
        if not log:
            messagebox.showinfo("Open Asset", "Select an open checkout row first.")
            return
        asset = self.db.find_asset(log["asset_barcode"])
        if asset:
            self.asset_profile(asset)

    def open_current_tree_user(self):
        log = self.selected_current_log()
        if not log:
            messagebox.showinfo("Open User", "Select an open checkout row first.")
            return
        person = self.db.find_person(log["employee_id"])
        if person:
            self.person_profile(person)

    def open_current_tree_logs(self):
        log = self.selected_current_log()
        if not log:
            messagebox.showinfo("View Log History", "Select an open checkout row first.")
            return
        self.audit_target_popup(f"Checkout Logs - {log['asset_barcode']}", log["asset_barcode"])

    def return_selected_current_asset(self):
        log = self.selected_current_log()
        if not log:
            messagebox.showinfo("Return Asset", "Select an open checkout row first.")
            return
        asset = self.db.find_asset(log["asset_barcode"])
        if asset:
            self.selected_asset = asset
            self.sel_asset_var.set(f"{asset['asset_type']} | {asset['barcode']} | {asset['name']} | {asset['status']}")
            self.return_scanned()

    def add_notes_to_selected_current_log(self):
        log = self.selected_current_log()
        if not log:
            messagebox.showinfo("Add Return Notes", "Select an open checkout row first.")
            return
        note = simpledialog.askstring("Add Checkout / Return Note", "Add note for this open checkout:", parent=self)
        if not clean(note):
            return
        existing = log["checkout_notes"] or ""
        combined = (existing + " | " if existing else "") + f"{pretty(now_iso())} {self.actor()}: {clean(note)}"
        self.db.run("UPDATE activity SET checkout_notes=? WHERE id=?", (combined, log["id"]))
        self.db.audit("ADD CHECKOUT NOTE", self.actor(), log["asset_barcode"], clean(note))
        self.refresh_current()

    def show_full_tree_row(self, tree, title="Row Details"):
        values = self.selected_tree_values(tree)
        if not values:
            messagebox.showinfo(title, "Select a row first.")
            return
        labels = list(tree["columns"])
        text = "\n".join(f"{label}: {value}" for label, value in zip(labels, values))
        win = tk.Toplevel(self)
        win.title(title)
        win.geometry("820x520")
        self.apply_icon(win)
        pan = self.panel(win, title, "Full selected row details.")
        pan.pack(fill="both", expand=True, padx=14, pady=14)
        detail = tk.Text(pan.body, wrap="word", font=("Segoe UI", 10), bg="white", fg="#111", padx=12, pady=10)
        scroll = ttk.Scrollbar(pan.body, orient="vertical", command=detail.yview)
        detail.configure(yscrollcommand=scroll.set)
        detail.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        pan.body.grid_rowconfigure(0, weight=1)
        pan.body.grid_columnconfigure(0, weight=1)
        detail.insert("1.0", text)
        detail.configure(state="disabled")
        buttons = tk.Frame(pan.body, bg=self.colors["panel"])
        buttons.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10,0))
        ttk.Button(buttons, text="Copy", command=lambda: (self.clipboard_clear(), self.clipboard_append(text), self.status.set("Details copied."))).pack(side="left", padx=4)
        ttk.Button(buttons, text="Close", command=win.destroy).pack(side="left", padx=4)

    def open_related_from_detail_tree(self, tree):
        values = self.selected_tree_values(tree)
        if not values:
            messagebox.showinfo("Related Record", "Select a row first.")
            return
        text = " ".join(str(v) for v in values)
        asset_token = clean(str(values[3]).split("/", 1)[0]) if len(values) > 3 else ""
        asset = self.db.find_asset(asset_token)
        if asset:
            self.asset_profile(asset)
            return
        log = self.db.one("SELECT * FROM activity WHERE asset_barcode LIKE ? OR asset_name LIKE ? OR employee_name LIKE ? ORDER BY id DESC LIMIT 1", (f"%{asset_token}%", f"%{asset_token}%", f"%{asset_token}%"))
        if log:
            asset = self.db.find_asset(log["asset_barcode"])
            person = self.db.find_person(log["employee_id"])
            if asset:
                self.asset_profile(asset)
            elif person:
                self.person_profile(person)
            return
        for token in text.replace("/", " ").replace("|", " ").split():
            person = self.db.find_person(token)
            if person:
                self.person_profile(person)
                return
        messagebox.showinfo("Related Record", "No matching asset or user was found for the selected row.")

    def open_logs_for_detail_tree(self, tree):
        values = self.selected_tree_values(tree)
        if not values:
            messagebox.showinfo("Related Logs", "Select a row first.")
            return
        target = clean(str(values[3]).split("/", 1)[0]) if len(values) > 3 else clean(str(values[0]))
        self.audit_target_popup(f"Related Logs - {target}", target)

    def add_alert_for_selected_asset(self):
        asset = self.selected_asset_from_tree()
        if not asset:
            messagebox.showinfo("Select Asset", "Select an asset first.")
            return
        self.add_ap_alert_popup("Asset", asset["barcode"], f"{asset['barcode']} | {asset['name']} | {asset['asset_type']}")

    def add_asset_issue_alert(self, alert_type, required_action="Review before checkout"):
        asset = self.selected_asset_from_tree()
        if not asset:
            messagebox.showinfo("Select Asset", "Select an asset first.")
            return
        note = simpledialog.askstring(alert_type, "Describe the issue:", parent=self)
        if not clean(note):
            return
        self.db.add_ap_alert("Asset", asset["barcode"], f"{asset['barcode']} | {asset['name']} | {asset['asset_type']}", alert_type, "Warning", clean(note), required_action, self.actor())
        self.refresh_assets()
        self.status.set(f"{alert_type} alert added for {asset['barcode']}.")

    def view_alerts_for_selected_asset(self):
        asset = self.selected_asset_from_tree()
        if not asset:
            messagebox.showinfo("Select Asset", "Select an asset first.")
            return
        self.ap_alerts_popup("Asset", asset["barcode"], f"{asset['barcode']} | {asset['name']} | {asset['asset_type']}")

    def add_alert_for_selected_person(self):
        person = self.selected_person_from_tree()
        if not person:
            messagebox.showinfo("Select User", "Select a user first.")
            return
        self.add_ap_alert_popup("Person", person["employee_id"], f"{self.db.person_name(person)} | {person['employee_id']} | {person['badge']}")

    def view_alerts_for_selected_person(self):
        person = self.selected_person_from_tree()
        if not person:
            messagebox.showinfo("Select User", "Select a user first.")
            return
        self.ap_alerts_popup("Person", person["employee_id"], f"{self.db.person_name(person)} | {person['employee_id']} | {person['badge']}")

    def current_out_for_selected_person(self):
        person = self.selected_person_from_tree()
        if not person:
            messagebox.showinfo("Select User", "Select a user first.")
            return
        win = tk.Toplevel(self)
        win.title(f"Current Out - {self.db.person_name(person)}")
        win.geometry("980x560")
        self.apply_icon(win)
        pan = self.panel(win, f"Current Out - {self.db.person_name(person)}", "Assets currently checked out to this user.")
        pan.pack(fill="both", expand=True, padx=14, pady=14)
        tree = self.tree(pan.body, ("Status","Type","Asset","Due","Operator","Notes"), 12)
        rows = self.db.all("SELECT * FROM activity WHERE employee_id=? AND status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') ORDER BY due_back_at", (person["employee_id"],))
        self.fill(tree, [(r["status"], r["asset_type"], f"{r['asset_barcode']} / {r['asset_name']}", pretty(r["due_back_at"]), r["checkout_operator"], r["checkout_notes"]) for r in rows])
        self.bind_context_menu(tree, lambda t: [
            ("Open Details", lambda: self.show_full_tree_row(t, "Current Out"), bool(self.selected_tree_values(t))),
            ("Copy Row", lambda: self.copy_tree_row(t), bool(self.selected_tree_values(t))),
            ("Close", win.destroy),
        ])
        ttk.Button(pan.body, text="Close", command=win.destroy).pack(anchor="w", pady=8)

    def assign_selected_person_group(self):
        if not self.require("people_edit"):
            return
        person = self.selected_person_from_tree()
        if not person:
            messagebox.showinfo("Select User", "Select a user first.")
            return
        groups = self.db.group_names()
        new_group = simpledialog.askstring("Assign Group", "Enter group name:\n" + ", ".join(groups), initialvalue=person["role"] or "Employee", parent=self)
        new_group = clean(new_group)
        if new_group not in groups:
            messagebox.showwarning("Invalid Group", "Choose an existing group.")
            return
        old_group = person["role"] or "Employee"
        if old_group == new_group:
            return
        if not self.can_assign_role(old_group, new_group, person):
            return
        reason = simpledialog.askstring("Reason Required", f"Why assign {self.db.person_name(person)} to {new_group}?", parent=self)
        if not clean(reason):
            messagebox.showwarning("Reason Required", "A reason is required.")
            return
        d = {k: person[k] for k in ["employee_id","badge","first_name","last_name","department","status","shift","role","notes"]}
        d["role"] = new_group
        self.db.update_person(person["id"], d, self.actor())
        detail = f"Old group: {old_group}; New group: {new_group}; Reason: {clean(reason)}"
        self.db.audit("USER GROUP ASSIGNED", self.actor(), person["employee_id"], detail, old_group, new_group, clean(reason))
        self.db.notify_manager("Permission changed", "Info", self.db.person_name(person), "", "User group assigned", self.actor(), detail, clean(reason), status="Reviewed")
        self.refresh_people()

    def audit_for_selected_asset(self):
        asset = self.selected_asset_from_tree()
        if asset:
            self.audit_target_popup(f"Audit Log - {asset['barcode']}", asset["barcode"])

    def audit_for_selected_person(self):
        person = self.selected_person_from_tree()
        if person:
            self.audit_target_popup(f"Audit Log - {person['employee_id']}", person["employee_id"])

    def audit_target_popup(self, title, target):
        win = tk.Toplevel(self)
        win.title(title)
        win.geometry("1000x560")
        self.apply_icon(win)
        pan = self.panel(win, title, "Audit entries matching this record.")
        pan.pack(fill="both", expand=True, padx=14, pady=14)
        tree = self.tree(pan.body, ("Time","Action","Actor","Role","Page","Status","Details"), 14)
        like = f"%{target}%"
        rows = self.db.all("SELECT * FROM audit WHERE target LIKE ? OR details LIKE ? ORDER BY timestamp DESC LIMIT 250", (like, like))
        self.fill(tree, [(pretty(r["timestamp"]), r["action"], r["actor"], r["actor_role"] if "actor_role" in r.keys() else "", r["page"] if "page" in r.keys() else "", r["status"] if "status" in r.keys() else "", r["details"]) for r in rows])
        self.bind_context_menu(tree, lambda t: [("Copy Row", lambda: self.copy_tree_row(t)), ("Close", win.destroy)])
        ttk.Button(pan.body, text="Close", command=win.destroy).pack(anchor="w", pady=8)

    def export_selected_notification_excel(self):
        note_id = self.selected_manager_notification_id()
        if not note_id:
            messagebox.showinfo("Select Notification", "Select a manager notification first.")
            return
        row = self.db.one("SELECT * FROM manager_notifications WHERE id=?", (note_id,))
        if not row:
            return
        headers = ["Date/Time","12-Hour Time","24-Hour Time","Event","Severity","User","Asset","Action","Handled By","Status","Notes","Reason","Computer"]
        data = [[row["timestamp"], row["time_12"], row["time_24"], row["event_type"], row["severity"], row["user_involved"], row["asset_involved"], row["action_taken"], row["handled_by"], row["status"], row["notes"], row["reason"], row["computer_name"]]]
        path = self.folder_setting("excel_export_folder", "Exports") / f"manager_notification_{note_id}_{stamp()}.xlsx"
        self.write_xlsx(path, [("Notification", headers, data)])
        self.db.audit("EXPORT SELECTED NOTIFICATION", self.actor(), note_id, str(path))
        messagebox.showinfo("Export Notification", f"Notification exported:\n{path}")

    def export_selected_log_row(self, tree):
        values = self.selected_tree_values(tree)
        if not values:
            messagebox.showinfo("Export Selected Log", "Select a log row first.")
            return
        path = self.folder_setting("excel_export_folder", "Exports") / f"selected_log_{stamp()}.xlsx"
        headers = list(tree["columns"])
        self.write_xlsx(path, [("Selected Log", headers, [list(values)])])
        self.db.audit("EXPORT SELECTED LOG", self.actor(), "Logs", str(path))
        messagebox.showinfo("Export Selected Log", f"Selected log exported:\n{path}")

    def view_selected_notification_full(self):
        note_id = self.selected_manager_notification_id()
        if not note_id:
            messagebox.showinfo("Select Notification", "Select a manager notification first.")
            return
        row = self.db.one("SELECT * FROM manager_notifications WHERE id=?", (note_id,))
        if not row:
            return
        text = "\n".join(f"{key}: {row[key]}" for key in row.keys())
        messagebox.showinfo("Manager Notification", text)

    def open_related_from_selected_notification(self):
        note_id = self.selected_manager_notification_id()
        if not note_id:
            messagebox.showinfo("Select Notification", "Select a manager notification first.")
            return
        row = self.db.one("SELECT * FROM manager_notifications WHERE id=?", (note_id,))
        if not row:
            return
        asset = self.db.find_asset(row["asset_involved"] or "")
        if asset:
            self.asset_profile(asset)
            return
        person = self.db.find_person(row["user_involved"] or "")
        if person:
            self.person_profile(person)
            return
        messagebox.showinfo("Related Record", "No matching user or asset record was found for this notification.")

    def view_related_logs_for_selected_notification(self):
        note_id = self.selected_manager_notification_id()
        if not note_id:
            messagebox.showinfo("Select Notification", "Select a manager notification first.")
            return
        note = self.db.one("SELECT * FROM manager_notifications WHERE id=?", (note_id,))
        if not note:
            return
        terms = [
            clean(note["asset_involved"]),
            clean(note["user_involved"]),
            clean(note["event_type"]),
            clean(note["action_taken"]),
        ]
        terms = [t.lower() for t in terms if t]
        if not terms:
            messagebox.showinfo("Related Logs", "This notification does not have enough details to match related logs.")
            return
        win = tk.Toplevel(self)
        win.title(f"Related Logs - Notification {note_id}")
        win.geometry("1100x560")
        self.apply_icon(win)
        pan = self.panel(win, "Related Logs", "Audit and error records connected to this notification.")
        pan.pack(fill="both", expand=True, padx=14, pady=14)
        tree = self.tree(pan.body, ("Time", "Type", "Action / Source", "Actor", "Target", "Details"), 15)

        def matches(row_text):
            text = row_text.lower()
            return any(term in text for term in terms)

        rows = []
        for r in self.db.all("SELECT * FROM audit ORDER BY id DESC LIMIT 1000"):
            text = " ".join(str(r[k] or "") for k in r.keys())
            if matches(text):
                rows.append((r["timestamp"], "Audit", r["action"], r["actor"], r["target"], r["details"]))
        for r in self.db.all("SELECT * FROM errors ORDER BY id DESC LIMIT 1000"):
            text = " ".join(str(r[k] or "") for k in r.keys())
            if matches(text):
                rows.append((r["timestamp"], "Error", r["source"], r["actor"] if "actor" in r.keys() else "", r["message"], r["details"]))
        rows.sort(key=lambda item: item[0], reverse=True)
        self.fill(tree, [(pretty(r[0]), r[1], r[2], r[3], r[4], r[5]) for r in rows[:250]])
        self.bind_context_menu(tree, lambda t: [
            ("View Full Log Entry", lambda: self.show_full_tree_row(t, "Related Log"), bool(self.selected_tree_values(t))),
            ("Copy Row", lambda: self.copy_tree_row(t), bool(self.selected_tree_values(t))),
            ("Close", win.destroy),
        ])
        ttk.Button(pan.body, text="Close", command=win.destroy).pack(anchor="w", pady=8)

    def people_context_actions(self, tree):
        person = self.selected_person_from_tree()
        can_edit = self.has("people_edit")
        inactive = bool(person and person["status"] not in ("Active", "Manual Review"))
        return [
            ("Open Profile", self.open_selected_person, bool(person)),
            ("Edit User", self.person_form, bool(person) and can_edit),
            None,
            ("Deactivate User", lambda: self.set_selected_person_status("Inactive"), bool(person) and can_edit and not inactive),
            ("Reactivate User", lambda: self.set_selected_person_status("Active"), bool(person) and can_edit and inactive),
            ("View Current Checked-Out Assets", self.current_out_for_selected_person, bool(person)),
            ("Add AP Alert", self.add_alert_for_selected_person, bool(person) and self.has("manager")),
            ("View AP Alerts", self.view_alerts_for_selected_person, bool(person) and self.has("manager")),
            ("Assign Group / Permissions", self.assign_selected_person_group, bool(person) and can_edit),
            ("View Audit Log", self.audit_for_selected_person, bool(person) and self.has("manager")),
            None,
            ("Copy Row", lambda: self.copy_tree_row(tree), bool(person)),
        ]

    def asset_context_actions(self, tree):
        asset = self.selected_asset_from_tree()
        open_log = self.db.open_log(asset["barcode"]) if asset else None
        can_edit = self.has("assets_edit")
        return [
            ("View Details", self.open_selected_asset, bool(asset)),
            ("Edit Asset", self.asset_form, bool(asset) and can_edit),
            ("Duplicate / Copy Asset", self.duplicate_selected_asset, bool(asset) and can_edit),
            ("Delete / Retire Asset", self.delete_selected_asset, bool(asset) and can_edit),
            None,
            ("Checkout Asset", self.use_selected_asset_frontdesk, bool(asset) and not open_log),
            ("Return Asset", lambda: (setattr(self, "selected_asset", asset), self.return_scanned()), bool(asset) and bool(open_log)),
            ("View Asset History", self.open_selected_asset, bool(asset)),
            ("View Audit Log", self.audit_for_selected_asset, bool(asset) and self.has("manager")),
            None,
            ("Export Selected Asset", self.export_selected_asset_excel, bool(asset)),
            ("Mark Damaged / Repair", lambda: self.set_selected_asset_status("Repair"), bool(asset) and can_edit),
            ("Mark Missing", lambda: self.set_selected_asset_status("Missing"), bool(asset) and can_edit),
            ("Mark Missing Accessory", lambda: self.add_asset_issue_alert("Missing item/accessory"), bool(asset) and self.has("manager")),
            ("Add Asset Alert", self.add_alert_for_selected_asset, bool(asset) and self.has("manager")),
            ("Review Alerts", self.view_alerts_for_selected_asset, bool(asset) and self.has("manager")),
            None,
            ("Copy Row", lambda: self.copy_tree_row(tree), bool(asset)),
        ]

    def current_context_actions(self, tree):
        log = self.selected_current_log(tree)
        return [
            ("Return Selected Asset", self.return_selected_current_asset, bool(log)),
            ("Open Asset Record", self.open_current_tree_asset, bool(log)),
            ("Open User Record", self.open_current_tree_user, bool(log)),
            ("View Current Checkout", lambda: self.show_full_tree_row(tree, "Current Checkout"), bool(log)),
            ("Add Return Notes", self.add_notes_to_selected_current_log, bool(log)),
            ("View Log History", self.open_current_tree_logs, bool(log)),
            None,
            ("Refresh", self.refresh_current),
            ("Copy Row", lambda: self.copy_tree_row(tree), bool(log)),
        ]

    def manager_notification_context_actions(self, tree):
        has_row = bool(self.selected_tree_values(tree))
        return [
            ("Mark Reviewed", lambda: self.update_selected_notification("Reviewed"), has_row),
            ("Mark Resolved", lambda: self.update_selected_notification("Resolved"), has_row),
            ("Add Manager Note", self.add_note_to_selected_notification, has_row),
            ("View Full Notification", self.view_selected_notification_full, has_row),
            ("View Related User / Asset", self.open_related_from_selected_notification, has_row),
            ("View Related Logs", self.view_related_logs_for_selected_notification, has_row),
            None,
            ("Export Selected Notification", self.export_selected_notification_excel, has_row),
            ("Export All Notifications", self.export_manager_notifications_excel),
            ("Copy Row", lambda: self.copy_tree_row(tree), has_row),
        ]

    def logs_context_actions(self, tree):
        has_row = bool(self.selected_tree_values(tree))
        return [
            ("View Full Log Entry", lambda: self.show_full_tree_row(tree, "Log Entry"), has_row),
            ("Export Selected Log Row", lambda: self.export_selected_log_row(tree), has_row),
            ("Export Visible Logs", self.export_logs_excel),
            ("Copy Row / Details", lambda: self.copy_tree_row(tree), has_row),
            ("Open Log Folder", lambda: self.open_folder_path(self.folder_setting("system_log_folder", "Logs"))),
            None,
            ("Archive / Clear Old Logs", self.clear_old_logs, self.has("admin")),
        ]

    def validate_person_form_live(self, vars, widgets):
        bad = False
        checks = {
            "employee_id": lambda v: clean(v).upper().startswith("F"),
            "badge": lambda v: clean(v).startswith("88"),
            "first_name": lambda v: bool(clean(v)),
            "last_name": lambda v: bool(clean(v)),
            "role": lambda v: bool(clean(v)),
        }
        for key, fn in checks.items():
            ok = fn(vars[key].get())
            self.mark_required(widgets.get(key), not ok)
            bad = bad or not ok
        return not bad

    def normalize_asset_record(self, source):
        d = {k: clean(source.get(k)) for k in ASSET_FIELDS}
        if d["asset_type"] not in ASSET_TYPES:
            d["asset_type"] = "Item"
        if d["status"] not in ASSET_STATUS:
            d["status"] = "Available"
        if d["barcode"]:
            d["barcode"] = d["barcode"].upper()
        if d["controlled_key_number"]:
            d["controlled_key_number"] = d["controlled_key_number"].upper()
        if d["serial_number"]:
            d["serial_number"] = d["serial_number"].upper()
        apply_asset_type_field_map(d)
        if not d["name"] and d["barcode"]:
            prefix = ASSET_TYPE_GUIDANCE.get(d["asset_type"], ASSET_TYPE_GUIDANCE["Item"])["name_prefix"]
            d["name"] = f"{prefix} {d['barcode']}"
        details = normalize_asset_details(d["asset_type"], source)
        d["asset_details"] = asset_details_to_json(details)
        if d["asset_type"] == "Key":
            d["controlled_key_number"] = d["controlled_key_number"] or details.get("key_set_number", "").upper()
            d["location"] = d["location"] or details.get("key_ring_location", "")
        elif d["asset_type"] == "Radio":
            d["serial_number"] = d["serial_number"] or details.get("radio_serial_number", "").upper() or details.get("factory_serial_number", "").upper()
            d["location"] = d["location"] or details.get("radio_location", "")
        elif d["asset_type"] == "Tablet":
            d["serial_number"] = d["serial_number"] or details.get("tablet_factory_serial_number", "").upper()
            d["location"] = d["location"] or details.get("tablet_location", "")
        elif d["asset_type"] == "Temp Badge":
            d["location"] = d["location"] or details.get("badge_location", "")
        elif d["asset_type"] == "Scanner":
            d["serial_number"] = d["serial_number"] or details.get("scanner_serial_number", "").upper()
            d["location"] = d["location"] or details.get("scanner_location", "")
        return d

    def asset_duplicate_message(self, d, current_id=None):
        checks = [
            ("barcode", d.get("barcode"), "barcode"),
            ("controlled_key_number", d.get("controlled_key_number"), "controlled key number"),
            ("serial_number", d.get("serial_number"), "serial number"),
        ]
        for field, value, label in checks:
            if not value:
                continue
            row = self.db.one(f"SELECT id, barcode, name FROM assets WHERE lower({field})=lower(?)", (value,))
            if row and row["id"] != current_id:
                return f"Another asset already uses this {label}: {value}\n\nExisting asset: {row['barcode']} - {row['name']}"
        return ""

    def validate_asset_record(self, d, current_id=None):
        problems = []
        if not d.get("barcode"):
            problems.append("Barcode is required.")
        if not d.get("name"):
            problems.append("Name is required.")
        if d.get("asset_type") not in ASSET_TYPES:
            problems.append("Choose a valid asset type.")
        if d.get("status") not in ASSET_STATUS:
            problems.append("Choose a valid status.")
        duplicate = self.asset_duplicate_message(d, current_id)
        if duplicate:
            problems.append(duplicate)
        details = parse_asset_details(d.get("asset_details"))
        schema = asset_detail_schema(d.get("asset_type"))
        for spec in schema["fields"]:
            value = clean(details.get(spec["key"], ""))
            if spec.get("required") and not value:
                problems.append(f"{spec['label']} is required for {d.get('asset_type')}.")
            if spec.get("type") == "int":
                if not value.isdigit():
                    problems.append(f"{spec['label']} must be a number.")
                elif int(value) < spec.get("min", 0):
                    problems.append(f"{spec['label']} must be at least {spec.get('min')}.")
            if spec.get("unique") and value:
                for row in self.db.all("SELECT id, barcode, name, asset_type, asset_details FROM assets"):
                    if current_id and row["id"] == current_id:
                        continue
                    other = parse_asset_details(row["asset_details"] if "asset_details" in row.keys() else "")
                    if clean(other.get(spec["key"], "")).lower() == value.lower():
                        problems.append(f"{spec['label']} already exists on {row['barcode']} - {row['name']}.")
                        break
        if d.get("asset_type") == "Key":
            key_rows = details.get("keys", [])
            for idx, item in enumerate(key_rows, start=1):
                if not clean(item.get("serial")):
                    problems.append(f"Key {idx} serial/key number is required.")
                if not clean(item.get("access")):
                    problems.append(f"Key {idx} opens/access description is required.")
        return problems

    def validate_asset_form_live(self, vars, widgets):
        bad = False
        asset_type = vars["asset_type"].get() if "asset_type" in vars else "Item"
        rule = asset_type_rule(asset_type)
        checks = {
            "barcode": lambda v: bool(clean(v)),
            "name": lambda v: bool(clean(v)),
        }
        for key, fn in checks.items():
            ok = fn(vars[key].get())
            self.mark_required(widgets.get(key), not ok)
            bad = bad or not ok
        # Recommended fields by type
        for key in ("controlled_key_number", "serial_number"):
            self.mark_required(widgets.get(key), key in rule["recommended"] and not bool(clean(vars[key].get())))
        return not bad



    def normalize_person_id_badge(self, d, ask=True):
        """v5.2.0 rule: badge starts with 88, employee ID starts with F."""
        emp_id = clean(d.get("employee_id", "")).upper()
        badge = clean(d.get("badge", "")).upper()

        # If they are entered backwards, correct them.
        if emp_id.startswith("88") and badge.startswith("F"):
            if (not ask) or messagebox.askyesno(
                "Badge / ID Look Reversed",
                "It looks like Employee ID and Badge were entered backwards.\\n\\n"
                "Badge should start with 88.\\n"
                "Employee ID should start with F.\\n\\n"
                "Do you want me to swap them automatically?"
            ):
                emp_id, badge = badge, emp_id

        d["employee_id"] = emp_id
        d["badge"] = badge
        return d

    def validate_person_id_badge(self, d):
        badge = clean(d.get("badge", "")).upper()
        emp_id = clean(d.get("employee_id", "")).upper()
        problems = []
        if not badge.startswith("88"):
            problems.append("Badge should start with 88.")
        if not emp_id.startswith("F"):
            problems.append("Employee ID should start with F.")
        if problems:
            messagebox.showwarning(
                "Check Badge / Employee ID",
                "\\n".join(problems) + "\\n\\nExample:\\nBadge: 88984717\\nEmployee ID: F984717"
            )
            return False
        return True



    def is_training(self):
        return bool(self.training_mode.get())

    def training_prefix(self):
        return "[TRAINING MODE] " if self.is_training() else ""

    def toggle_training_mode(self):
        self.training_mode.set(not self.training_mode.get())
        if self.is_training():
            self.status.set("TRAINING MODE ON - actions are marked as training/demo.")
            self.db.audit("TRAINING MODE ON", self.actor(), "System", "Training/demo mode enabled.")
        else:
            self.status.set("Training Mode off.")
            self.db.audit("TRAINING MODE OFF", self.actor(), "System", "Training/demo mode disabled.")
        self.update_training_badge()
        self.refresh_scan_mode()

    def mark_training_notes(self, notes):
        notes = clean(notes)
        if self.is_training():
            return ("[TRAINING / DEMO] " + notes).strip()
        return notes

    def today_backup_ok(self):
        last = self.db.setting("last_backup", "")
        return last.startswith(dt.datetime.now().strftime("%Y-%m-%d"))

    def auto_backup_enabled(self):
        return self.db.setting("auto_daily_backup", "Yes") != "No"

    def scan_only_frontdesk_enabled(self):
        return self.db.setting("frontdesk_scan_only_mode", "No") == "Yes"

    def auto_daily_backup(self):
        if not self.auto_backup_enabled() or self.today_backup_ok():
            return
        try:
            out = self.backup_with_logs(show_message=False, label="AUTO DAILY BACKUP")
            self.status.set(f"Daily backup saved: {out.name}")
        except Exception:
            try:
                self.db.error("Auto Daily Backup", "Automatic backup failed", traceback.format_exc())
            except Exception:
                pass

    def dashboard_health_lines(self):
        c = self.db.counts()
        db_ok = "OK" if Path(self.db.path).exists() else "Missing"
        backup = "OK today" if self.today_backup_ok() else "Auto backup pending"
        return [
            ("Database", db_ok),
            ("Last backup", backup),
            ("Auto backup", "On" if self.auto_backup_enabled() else "Off"),
            ("Late returns", c.get("late_returns", 0)),
            ("Errors today", c.get("errors", 0)),

        ]


    def role(self):
        return self.operator_role.get() or "Guest"

    def actor(self):
        current = self.operator.get() or "Unknown"
        if self.scan_only_frontdesk_enabled() and current == "No operator signed in":
            return "Scan-Only Front Desk"
        return current

    def has(self, right):
        if right == "frontdesk" and self.scan_only_frontdesk_enabled():
            return True
        return right in self.db.group_rights(self.role())

    def require(self, right):
        if not self.has(right):
            messagebox.showwarning("Access Denied", f"Your role does not have access to this section.\n\nCurrent role: {self.role()}")
            return False
        return True

    def operator_login(self):
        scan = clean(self.operator_scan.get())
        if not scan:
            messagebox.showwarning("Missing Badge", "Scan or type operator badge/ID first.")
            return
        p = self.db.find_person(scan)
        if not p or p["status"] not in ("Active", "Manual Review"):
            messagebox.showwarning("Operator Not Found", "Operator was not found or is not active.")
            self.db.error("Operator Login", "Operator not found/inactive", scan)
            self.db.notify_manager("Failed login attempt", "Warning", scan, "", "Login blocked", "System", "Operator not found or inactive.")
            return
        if p["badge"] == "88984717" or p["employee_id"] == "F984717":
            self.db.run("UPDATE people SET role='Admin', status='Active', updated_at=? WHERE id=?", (now_iso(), p["id"]))
            p = self.db.find_person(scan)
        self.operator.set(self.db.person_name(p))
        self.operator_role.set(p["role"] or "Employee")
        self.update_operator_badges()
        self.operator_scan.set("")
        self.db.set_setting("last_operator_scan", "")  # v5.1.4: do not remember operator for startup auto-login
        self.db.set_setting("current_operator_name", self.actor())
        self.db.set_setting("current_operator_role", self.role())
        self.db.audit("OPERATOR LOGIN", self.actor(), "Operator", self.role())
        self.db.notify_manager("Operator login", "Info", self.actor(), "", "Login success", self.actor(), f"Role: {self.role()}", status="Reviewed")
        self.status.set(f"Operator signed in: {self.actor()} - Role: {self.role()}")
        self.refresh()

    def operator_logout(self):
        if messagebox.askyesno("Sign Out", "Sign out current operator?"):
            signed_out = self.actor()
            self.db.audit("OPERATOR LOGOUT", signed_out, "Operator", "Operator signed out.")
            self.operator.set("No operator signed in")
            self.operator_role.set("Guest")
            self.update_operator_badges()
            self.db.set_setting("last_operator_scan", "")
            self.db.set_setting("current_operator_name", "No operator signed in")
            self.db.set_setting("current_operator_role", "Guest")
            self.status.set("Signed out.")
            self.show("Dashboard")

    def restore_operator(self):
        # v5.1.4: safer shared-computer default.
        # Do not auto-login the last saved operator when the app starts.
        self.operator.set("No operator signed in")
        self.operator_role.set("Guest")
        self.update_operator_badges()
        self.operator_scan.set("")
        self.db.set_setting("current_operator_name", "No operator signed in")
        self.db.set_setting("current_operator_role", "Guest")
        self.status.set("No operator signed in. Scan or type operator badge at top-right.")
        try:
            self.db.audit("STARTUP NO AUTO LOGIN", "System", "Operator", "Startup auto-login disabled. Operator must sign in manually.")
        except Exception:
            pass

    def show(self, name):
        aliases = {
            "Checkout / Return": "Front Desk",
            "Users": "People / Users",
            "Groups / Permissions": "System",
            "Settings": "System",
            "Admin Tools": "System",
        }
        page_name = aliases.get(name, name)
        needed = {"Dashboard":"dashboard","Front Desk":"frontdesk","People / Users":"people_view","Assets":"assets_view","Search":"search","Manager":"manager","Reports":"reports","Logs":"manager","Groups / Permissions":"system","Settings":"system","Admin Tools":"admin","System":"system"}.get(page_name, "dashboard")
        if name != "Dashboard" and not self.has(needed):
            messagebox.showwarning("Access Denied", f"{name} requires access right: {needed}\n\nCurrent role: {self.role()}")
            return
        for p in self.pages.values():
            p.pack_forget()
        self.pages[page_name].pack(fill="both", expand=True)
        self.current_page = page_name
        try:
            self.db.set_setting("current_page", page_name)
        except Exception:
            pass
        nav_name = {"People / Users": "Users"}.get(page_name, page_name)
        self.active_nav_name = nav_name
        for n in self.nav:
            self.set_nav_button_state(n, n == nav_name)
        self.update_operator_badges()
        self.refresh_page(page_name)

    def dashboard_page(self):
        p = self.page("Dashboard")

        top = self.panel(p, "Dashboard", "Current AP activity, late returns, asset availability, and backup status.")
        top.pack(fill="x", pady=(0,12))
        self.dashboard_status_vars = {
            "database": tk.StringVar(value="Checking"),
            "backup": tk.StringVar(value="Checking"),
            "auto_refresh": tk.StringVar(value="Checking"),
            "late": tk.StringVar(value="0"),
            "errors": tk.StringVar(value="0"),
        }
        status_row = tk.Frame(top.body, bg=self.colors["panel"])
        status_row.pack(fill="x")
        self.status_badge(status_row, "Database", self.dashboard_status_vars["database"], "#198754")
        self.status_badge(status_row, "Backup", self.dashboard_status_vars["backup"], "#198754")
        self.status_badge(status_row, "Auto Refresh", self.dashboard_status_vars["auto_refresh"], "#0d6efd")
        self.status_badge(status_row, "Late Returns", self.dashboard_status_vars["late"], self.colors["red"])
        self.status_badge(status_row, "Errors Today", self.dashboard_status_vars["errors"], self.colors["red"])

        metrics = tk.Frame(p, bg=self.colors["bg"])
        metrics.pack(fill="x", pady=(0,14))
        self.metric_vars = {k:tk.StringVar(value="0") for k in ["out","keys_out","radios_out","tablets_out","temp_badges_out","late_returns","manager_alerts","errors","backup_status","available_assets","inactive_assets"]}
        metric_groups = [
            ("Operations", [
                ("All Out", "out", "All assets currently checked out.", "OUT", "#0d6efd", "#f2f7ff"),
                ("Keys Out", "keys_out", "Controlled keys currently checked out.", "KEY", "#3d6f8e", "#f1f7fb"),
                ("Radios Out", "radios_out", "Radios currently checked out.", "RAD", "#6f42c1", "#f7f1ff"),
                ("Tablets Out", "tablets_out", "Tablets currently checked out.", "TAB", "#0f766e", "#effcf9"),
                ("Temp Badges", "temp_badges_out", "Temp badges currently checked out.", "ID", "#0d6efd", "#f2f7ff"),
            ]),
            ("Alerts / Risk", [
                ("Late Returns", "late_returns", "Open items past their due-back time.", "LATE", self.colors["red"], "#fff0f1"),
                ("Active Alerts", "manager_alerts", "Manager notifications marked New.", "ALRT", "#b36b00", "#fff7ea"),
                ("Errors Today", "errors", "Errors logged today.", "ERR", self.colors["red"], "#fff0f1"),
            ]),
            ("System", [
                ("Backup", "backup_status", "Today backup status.", "BKUP", "#198754", "#f0fff6"),
            ]),
        ]
        card_index = 0
        for group_title, items in metric_groups:
            title = tk.Label(metrics, text=group_title, bg=self.colors["bg"], fg=self.colors["muted"], font=("Segoe UI", 9, "bold"), anchor="w")
            title.grid(row=card_index, column=0, columnspan=5, sticky="ew", padx=6, pady=(2,0))
            card_index += 1
            for i, (label, key, desc, icon, accent, bg) in enumerate(items):
                row_idx = card_index
                col_idx = i
                card = tk.Frame(metrics, bg=bg, highlightbackground="#d5dbe5", highlightthickness=1, cursor="hand2")
                card.grid(row=row_idx, column=col_idx, sticky="nsew", padx=6, pady=6)
                metrics.grid_columnconfigure(col_idx, weight=1, uniform="dashcards")
                top_line = tk.Frame(card, bg=bg)
                top_line.pack(fill="x", padx=14, pady=(12,0))
                tk.Label(top_line, text=icon, bg=accent, fg="white", font=("Segoe UI", 8, "bold"), padx=7, pady=2).pack(side="left")
                tk.Label(top_line, text=label, bg=bg, fg="#17202a", font=("Segoe UI", 10, "bold")).pack(side="left", padx=(8,0))
                tk.Label(card, textvariable=self.metric_vars[key], bg=bg, fg=accent, font=("Segoe UI", 26, "bold")).pack(anchor="w", padx=14, pady=(3,8))
                self.tip(card, desc)
                card.bind("<Button-1>", lambda e, k=key, l=label: self.dashboard_detail_popup(k, l))
                card.bind("<Button-3>", lambda e, k=key, l=label: self.dashboard_card_context(e, k, l))
                card.bind("<Enter>", lambda e, c=card, a=accent: c.configure(highlightbackground=a, highlightthickness=2), add="+")
                card.bind("<Leave>", lambda e, c=card: c.configure(highlightbackground="#d5dbe5", highlightthickness=1), add="+")
                for child in card.winfo_children():
                    child.configure(cursor="hand2")
                    child.bind("<Button-1>", lambda e, k=key, l=label: self.dashboard_detail_popup(k, l))
                    child.bind("<Button-3>", lambda e, k=key, l=label: self.dashboard_card_context(e, k, l))
                    child.bind("<Enter>", lambda e, c=card, a=accent: c.configure(highlightbackground=a, highlightthickness=2), add="+")
                    child.bind("<Leave>", lambda e, c=card: c.configure(highlightbackground="#d5dbe5", highlightthickness=1), add="+")
                    for sub in child.winfo_children():
                        sub.configure(cursor="hand2")
                        sub.bind("<Button-1>", lambda e, k=key, l=label: self.dashboard_detail_popup(k, l))
                        sub.bind("<Button-3>", lambda e, k=key, l=label: self.dashboard_card_context(e, k, l))
                        sub.bind("<Enter>", lambda e, c=card, a=accent: c.configure(highlightbackground=a, highlightthickness=2), add="+")
                        sub.bind("<Leave>", lambda e, c=card: c.configure(highlightbackground="#d5dbe5", highlightthickness=1), add="+")
            card_index += 1
        for col_idx in range(5):
            metrics.grid_columnconfigure(col_idx, weight=1, uniform="dashcards")

        recent = self.panel(p, "Recent Activity", "Latest checkouts and returns.")
        recent.pack(fill="both", expand=True)
        self.dashboard_tree = self.tree(recent.body, ("Time","Status","Type","Asset","Employee","Operator","Due/Returned"), 14)
        self.bind_context_menu(self.dashboard_tree, lambda t: [
            ("Open Details", lambda: self.show_full_tree_row(t, "Recent Activity"), bool(self.selected_tree_values(t))),
            ("Refresh", self.refresh_dashboard),
            ("Copy Row", lambda: self.copy_tree_row(t), bool(self.selected_tree_values(t))),
        ])


    def frontdesk_page(self):
        p = self.page("Front Desk")
        top = self.panel(p, "Front Desk", "Scan an employee, scan an asset, then check out or return.")
        top.pack(fill="x", pady=(0,10))
        self.scan_var = tk.StringVar()
        self.fd_message = tk.StringVar(value="Ready for employee scan.")
        self.workflow_var = tk.StringVar(value="[  ] Mode  >  [  ] Employee  >  [  ] Asset  >  [  ] Due Time  >  [  ] Complete")
        self.asset_warning_var = tk.StringVar()
        self.due_var = tk.StringVar(value=self.default_due_iso())
        try:
            self.due_var.trace_add("write", lambda *_: self.update_frontdesk_workflow())
        except Exception:
            pass
        self.condition_var = tk.StringVar(value="Good")
        self.notes_var = tk.StringVar()
        guide = tk.Frame(top.body, bg=self.colors["panel"])
        guide.pack(fill="x", pady=(0,6))
        self.frontdesk_mode_label = tk.Label(guide, textvariable=self.frontdesk_mode_var, bg="#d1e7dd", fg="#0f5132", font=("Segoe UI", 10, "bold"), padx=10, pady=5)
        self.frontdesk_mode_label.pack(side="left", padx=(0,8))
        tk.Label(guide, text="One scan box handles employees, assets, key numbers, serials, and device tags.", bg=self.colors["panel"], fg=self.colors["muted"], font=("Segoe UI", 10, "bold")).pack(side="left", fill="x", expand=True, anchor="w")
        self.frontdesk_next_step_label = tk.Label(top.body, textvariable=self.frontdesk_next_step_var, bg="#e7f1ff", fg="#084298", font=("Segoe UI", 15, "bold"), anchor="w", justify="left", padx=12, pady=10, wraplength=1250)
        self.frontdesk_next_step_label.pack(fill="x", pady=(0,8))
        entry = ttk.Entry(top.body, textvariable=self.scan_var, font=("Segoe UI", 22))
        entry.pack(fill="x", ipady=4, pady=(2,6))
        self.bind_enter(entry, self.scan)
        self.tip(entry, "Scan or type an employee badge, asset barcode, key number, or radio serial number. Press Enter to process it.")
        self.scan_entry = entry
        tk.Label(top.body, textvariable=self.scan_mode, bg="#f8fafc", fg="#334155", font=("Segoe UI", 10, "bold"), padx=8, pady=5).pack(fill="x", pady=(0,6))
        flow = tk.Frame(top.body, bg=self.colors["panel"])
        flow.pack(fill="x", pady=(0,8))
        self.workflow_step_widgets = []
        for step in ("Mode", "Employee", "Asset", "Due Time", "Complete"):
            step_label = tk.Label(flow, text=f"[  ] {step}", bg="#e9edf3", fg="#4b5563", font=("Segoe UI", 9, "bold"), padx=10, pady=6)
            step_label.pack(side="left", fill="x", expand=True, padx=(0,5))
            self.workflow_step_widgets.append(step_label)
        tk.Label(top.body, textvariable=self.asset_warning_var, bg="#fff3cd", fg="#664d03", font=("Segoe UI", 10, "bold"), padx=8, pady=5).pack(fill="x", pady=(0,6))
        ttk.Label(top.body, textvariable=self.fd_message, style="Panel.TLabel", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        quick = tk.Frame(top.body, bg="#f8fafc", highlightbackground=self.colors["line"], highlightthickness=1)
        quick.pack(fill="x", pady=(8,0))
        tk.Label(quick, text="Quick Scan", bg="#f8fafc", fg="#17202a", font=("Segoe UI", 10, "bold"), padx=10, pady=8).pack(side="left")
        self.tip(ttk.Button(quick, text="New Checkout", style="Green.TButton", command=lambda: self.set_frontdesk_quick_mode("checkout")), "Clear the selected asset and guide the screen for a normal checkout.").pack(side="left", padx=3, pady=6)
        self.tip(ttk.Button(quick, text="Return Item", style="Secondary.TButton", command=lambda: self.set_frontdesk_quick_mode("return")), "Clear selections and guide the screen for a return.").pack(side="left", padx=3, pady=6)
        self.same_employee_button = self.tip(ttk.Button(quick, text="Next Item Same Employee", style="Secondary.TButton", command=lambda: self.set_frontdesk_quick_mode("same_employee")), "Keep the selected employee and scan another asset.")
        self.same_employee_button.pack(side="left", padx=3, pady=6)
        self.tip(ttk.Button(quick, text="Quick Guide", command=self.export_quick_scan_guide), "Save a printable one-page scanner and keyboard guide.").pack(side="left", padx=3, pady=6)
        tk.Label(quick, text="Keyboard: F2 checkout, F3 return, F4 next item, F8 check out, F9 return.", bg="#f8fafc", fg=self.colors["muted"], font=("Segoe UI", 9, "bold"), padx=10).pack(side="left", fill="x", expand=True, anchor="w")
        btns = tk.Frame(top.body, bg=self.colors["panel"])
        btns.pack(fill="x", pady=8)
        ttk.Label(btns, text="Actions:", style="Panel.TLabel").pack(side="left", padx=(0,4))
        fd_tips = {
            "Scan / Enter": "Processes the manual scan box the same as pressing Enter.",
            "Check Out Item": "Checks the selected asset out to the selected employee.",
            "Return Item": "Returns the selected checked-out item. For best logging, scan the person returning it first.",
            "Start Over": "Clears selected employee, selected asset, and return-by person."
        }
        action_buttons = [
            ("Scan / Enter", self.scan, "Secondary.TButton"),
            ("Check Out Item", self.checkout, "Green.TButton"),
            ("Return Item", self.return_scanned, "Red.TButton"),
            ("Start Over", self.clear_frontdesk, "Secondary.TButton"),
        ]
        for label, cmd, style_name in action_buttons:
            btn = self.tip(ttk.Button(btns, text=label, command=cmd, style=style_name), fd_tips.get(label, label))
            btn.pack(side="left", padx=4)
            if label == "Scan / Enter":
                self.scan_button = btn
            elif label == "Check Out Item":
                self.checkout_button = btn
            elif label == "Return Item":
                self.return_button = btn
        due = tk.Frame(top.body, bg=self.colors["panel"])
        due.pack(fill="x", pady=(2,0))
        presets = tk.Frame(due, bg=self.colors["panel"])
        presets.pack(fill="x", pady=(0,6))
        ttk.Label(presets, text="Due presets:", style="Panel.TLabel").pack(side="left", padx=(0,4))
        for label, val in [("6:30 PM", lambda: self.due_var.set(due_today(18,30))), ("4:30 PM", lambda: self.due_var.set(due_today(16,30))), ("1 Hour", lambda: self.due_var.set(plus_hours(1))), ("3 Hours", lambda: self.due_var.set(plus_hours(3))), ("More...", self.set_custom_due_time)]:
            self.tip(ttk.Button(presets, text=label, command=val), f"Set due-back time to {label}.").pack(side="left", padx=3)
        detail_row = tk.Frame(due, bg=self.colors["panel"])
        detail_row.pack(fill="x")
        ttk.Label(detail_row, text="Due Time", style="Panel.TLabel").grid(row=0, column=0, sticky="w", padx=(0,6))
        ttk.Label(detail_row, text="Condition", style="Panel.TLabel").grid(row=0, column=1, sticky="w", padx=(8,6))
        ttk.Label(detail_row, text="Notes", style="Panel.TLabel").grid(row=0, column=2, sticky="w", padx=(8,0))
        ttk.Entry(detail_row, textvariable=self.due_var, width=25).grid(row=1, column=0, sticky="ew", padx=(0,6), pady=(2,0), ipady=2)
        ttk.Entry(detail_row, textvariable=self.condition_var, width=16).grid(row=1, column=1, sticky="ew", padx=(8,6), pady=(2,0), ipady=2)
        ttk.Entry(detail_row, textvariable=self.notes_var).grid(row=1, column=2, sticky="ew", padx=(8,0), pady=(2,0), ipady=2)
        detail_row.grid_columnconfigure(0, weight=1)
        detail_row.grid_columnconfigure(1, weight=0)
        detail_row.grid_columnconfigure(2, weight=3)
        cards = tk.Frame(p, bg=self.colors["bg"])
        cards.pack(fill="x", pady=(0,10))
        self.sel_person_var = tk.StringVar(value="No employee selected.")
        self.sel_asset_var = tk.StringVar(value="No asset selected.")
        return_by_text = "Return by: scan badge required for scan-only returns." if self.scan_only_frontdesk_enabled() else "Return by: signed-in operator unless another badge is scanned."
        self.return_person_var = tk.StringVar(value=return_by_text)
        def profile_body(panel, avatar_text, color, var):
            row = tk.Frame(panel.body, bg=self.colors["panel"])
            row.pack(fill="x")
            tk.Label(row, text=avatar_text, bg=color, fg="white", font=("Segoe UI", 10, "bold"), width=8, height=3).pack(side="left", padx=(0,10))
            tk.Label(row, textvariable=var, bg=self.colors["panel"], fg="#17202a", font=("Segoe UI", 11, "bold"), anchor="w", justify="left", wraplength=360).pack(side="left", fill="x", expand=True)
        c1 = self.panel(cards, "Selected Employee")
        c1.pack(side="left", fill="x", expand=True, padx=(0,5))
        profile_body(c1, "USER", "#0d6efd", self.sel_person_var)
        c2 = self.panel(cards, "Selected Asset")
        c2.pack(side="left", fill="x", expand=True, padx=5)
        profile_body(c2, "ITEM", "#6f42c1", self.sel_asset_var)
        c3 = self.panel(cards, "Return By")
        c3.pack(side="left", fill="x", expand=True, padx=(5,0))
        profile_body(c3, "BY", "#198754", self.return_person_var)
        self.update_frontdesk_workflow()
        current = self.panel(p, "Current Out")
        current.pack(fill="both", expand=True)
        self.current_tree = self.tree(current.body, ("Status","Type","Asset","Employee","Due"), 12)
        self.bind_context_menu(self.current_tree, self.current_context_actions)
        self.bind_frontdesk_shortcuts()

    def dashboard_detail_popup(self, key, title):
        win = tk.Toplevel(self)
        win.title(title)
        win.geometry("980x620")
        self.apply_icon(win)
        pan = self.panel(win, title, "Filtered dashboard details.")
        pan.pack(fill="both", expand=True, padx=14, pady=14)
        search_var = tk.StringVar()
        filter_var = tk.StringVar(value="All")
        bar = tk.Frame(pan.body, bg=self.colors["panel"])
        bar.pack(fill="x", pady=(0,8))
        ttk.Label(bar, text="Search:", style="Panel.TLabel").pack(side="left")
        entry = ttk.Entry(bar, textvariable=search_var)
        entry.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Label(bar, text="Filter:", style="Panel.TLabel").pack(side="left", padx=(6,2))
        ttk.Combobox(bar, textvariable=filter_var, values=["All","OUT","DUE SOON","OVERDUE","MISSING","REVIEW","RETURNED","Info","Warning","Critical","Active","New","Reviewed","Resolved","Backup"], state="readonly", width=13).pack(side="left", padx=4)
        ttk.Button(bar, text="Close", command=win.destroy).pack(side="right", padx=4)
        tree = self.tree(pan.body, ("Time","Status","Type","Asset/User","Details"), 14)

        def rows_for_key():
            if key in ("keys_out", "radios_out", "tablets_out", "temp_badges_out"):
                asset_type = {"keys_out":"Key", "radios_out":"Radio", "tablets_out":"Tablet", "temp_badges_out":"Temp Badge"}[key]
                return self.db.all("SELECT checked_out_at timestamp, status, asset_type, asset_barcode, asset_name, employee_name, due_back_at FROM activity WHERE status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') AND asset_type=? ORDER BY due_back_at", (asset_type,))
            if key == "out":
                return self.db.all("SELECT checked_out_at timestamp, status, asset_type, asset_barcode, asset_name, employee_name, due_back_at FROM activity WHERE status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') ORDER BY due_back_at")
            if key == "late_returns":
                return self.db.all("SELECT checked_out_at timestamp, status, asset_type, asset_barcode, asset_name, employee_name, due_back_at FROM activity WHERE status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') AND due_back_at < ? ORDER BY due_back_at", (now_iso(),))
            if key == "manager_alerts":
                rows = []
                for n in self.db.all("SELECT timestamp, status, severity, event_type, user_involved, asset_involved, notes FROM manager_notifications WHERE status='New' ORDER BY timestamp DESC"):
                    rows.append({"timestamp": n["timestamp"], "status": n["status"], "asset_type": n["severity"], "asset_barcode": n["event_type"], "asset_name": n["user_involved"], "employee_name": n["asset_involved"], "due_back_at": n["notes"]})
                for a in self.db.active_ap_alerts():
                    rows.append({"timestamp": a["timestamp"], "status": a["status"], "asset_type": a["severity"], "asset_barcode": a["alert_type"], "asset_name": a["target_label"], "employee_name": a["target_type"], "due_back_at": a["note"]})
                return sorted(rows, key=lambda r: r["timestamp"], reverse=True)
            if key == "errors":
                return self.db.all("SELECT timestamp, message status, source asset_type, source asset_barcode, message asset_name, '' employee_name, details due_back_at FROM errors WHERE timestamp LIKE ? ORDER BY timestamp DESC", (dt.datetime.now().strftime("%Y-%m-%d")+"%",))
            if key == "backup_status":
                last = self.db.setting("last_backup", "")
                return [{"timestamp": last or now_iso(), "status": "OK today" if self.today_backup_ok() else "Needs Backup", "asset_type": "Backup", "asset_barcode": "Last Backup", "asset_name": pretty(last) if last else "No backup found", "employee_name": "Auto backup " + ("On" if self.auto_backup_enabled() else "Off"), "due_back_at": self.db.setting("backup_folder", str(self.app_dir / "Backups"))}]
            return []

        current_rows = []
        def refresh():
            nonlocal current_rows
            q = clean(search_var.get()).lower()
            filter_text = filter_var.get()
            out = []
            for r in rows_for_key():
                values = (pretty(r["timestamp"]), r["status"], r["asset_type"], f"{r['asset_barcode']} / {r['asset_name']}", f"{r['employee_name']} | {pretty(r['due_back_at'])}")
                if filter_text != "All" and filter_text.lower() not in " ".join(str(x).lower() for x in values):
                    continue
                if not q or q in " ".join(str(x).lower() for x in values):
                    out.append(values)
            current_rows = out
            self.fill(tree, out)

        def export_popup_rows():
            if not current_rows:
                messagebox.showinfo("Export", "No rows to export.")
                return
            path = self.folder_setting("excel_export_folder", "Exports") / f"dashboard_{key}_{stamp()}.xlsx"
            try:
                self.write_xlsx(path, [(title[:31], ["Time", "Status", "Type", "Asset/User", "Details"], current_rows)])
                self.db.audit("EXPORT DASHBOARD DETAIL", self.actor(), title, f"Rows: {len(current_rows)}; File: {path}")
                self.db.notify_manager("Export created", "Info", self.actor(), "", f"Dashboard export: {title}", self.actor(), str(path), status="Reviewed")
                messagebox.showinfo("Export Complete", f"Dashboard details exported:\n{path}")
            except Exception as e:
                self.db.error("Export Dashboard Detail", str(e), traceback.format_exc())
                messagebox.showerror("Export Failed", f"Could not export dashboard details:\n{e}")
        self.bind_enter(entry, refresh)
        ttk.Button(bar, text="Apply", command=refresh).pack(side="left", padx=4)
        ttk.Button(bar, text="Export", command=export_popup_rows).pack(side="left", padx=4)
        self.bind_context_menu(tree, lambda t: [
            ("Open Details", lambda: self.show_full_tree_row(t, title), bool(self.selected_tree_values(t))),
            ("View Related Asset / User", lambda: self.open_related_from_detail_tree(t), bool(self.selected_tree_values(t))),
            ("Open Related Logs", lambda: self.open_logs_for_detail_tree(t), bool(self.selected_tree_values(t))),
            ("Export Visible List", export_popup_rows),
            ("Refresh", refresh),
            ("Copy Selected Row", lambda: self.copy_tree_row(t), bool(self.selected_tree_values(t))),
            ("Close", win.destroy),
        ])
        refresh()


    def make_count_bar(self, parent, items, var_store_name):
        bar = tk.Frame(parent, bg=self.colors["panel"])
        bar.pack(fill="x", pady=(0, 8))
        store = {}
        for i, (label, key) in enumerate(items):
            box = tk.Frame(bar, bg="#f8f9fa", highlightbackground="#d0d0d0", highlightthickness=1)
            box.grid(row=0, column=i, sticky="nsew", padx=3, pady=2)
            bar.grid_columnconfigure(i, weight=1, uniform=var_store_name)
            tk.Label(box, text=label, bg="#f8f9fa", fg="#333", font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=7, pady=(5,0))
            store[key] = tk.StringVar(value="0")
            tk.Label(box, textvariable=store[key], bg="#f8f9fa", fg="#111", font=("Segoe UI", 15, "bold")).pack(anchor="w", padx=7, pady=(0,5))
            self.tip(box, f"Current count for {label}.")
        setattr(self, var_store_name, store)


    def refresh_people_counts(self):
        c = self.db.counts()
        if hasattr(self, "people_count_vars"):
            for k, v in self.people_count_vars.items():
                v.set(str(c.get(k, 0)))

    def refresh_asset_counts(self):
        c = self.db.counts()
        if hasattr(self, "asset_count_vars"):
            for k, v in self.asset_count_vars.items():
                v.set(str(c.get(k, 0)))

    def refresh_manager_counts(self):
        c = self.db.counts()
        if hasattr(self, "manager_count_vars"):
            for k, v in self.manager_count_vars.items():
                v.set(str(c.get(k, 0)))



    def apply_default_tooltips(self):
        """Adds generic hover descriptions to common controls that do not already have custom text."""
        descriptions = {
            "Add / Edit Person": "Add a new person or edit the selected person.",
            "Open Profile": "Open the selected person's profile.",
            "Export People Template": "Save a CSV template for adding people and importing them later.",
            "Retire Asset": "Delete an unused asset or retire an asset that has history.",
            "Add Asset": "Add a new asset.",
            "Edit Selected": "Edit the selected asset.",
            "Duplicate": "Start a new asset using the selected asset as a template.",
            "Import Assets": "Import assets from a completed CSV template with preview first.",
            "Export Selected": "Export the selected asset to an Excel profile.",
            "Export All": "Export every asset record to a formatted Excel workbook.",
            "Export Type": "Export only one selected asset type to a formatted Excel workbook.",
            "Export Excel": "Export assets to a formatted Excel workbook.",
            "Export Profile": "Export this asset profile to Excel.",
            "Print Profile": "Save and print this asset profile.",
            "Save Excel": "Save the current report to Excel.",
            "Use At Front Desk": "Send the selected asset to Front Desk for checkout or return.",
            "Set": "Apply the selected status to the selected asset.",
            "Apply": "Apply the selected sorting/filter option.",
            "Search": "Run the current search.",
            "Audit": "Show audit records in Search.",
            "Clear": "Clear search filters and results.",
            "Current Out": "Show all items currently checked out.",
            "End of Shift": "Generate a detailed end-of-shift report.",
            "Detailed Audit Report": "Show a detailed report of audit records.",
            "Detailed Error Report": "Show saved application errors in detail.",
            "Export CSV Bundle": "Export people, assets, activity, audit, errors, and settings as CSV files.",
            "Backup with Logs": "Create a backup ZIP with database and logs.",
            "Print Report": "Print or save the current report text.",
            "Show Data Location": "Show the current SQLite .db data file path and config file path.",
            "Open Data Folder": "Open the folder containing the current .db data file.",
            "Change Data File Location": "Copy the current database to a new .db location and use it after restart.",
            "Backup Now": "Create a backup ZIP with the database and logs.",
            "Open Backup Folder": "Open the folder where backups are saved.",
            "Restore Backup": "Restore a database from a backup ZIP or DB file.",
            "Issue Report ZIP": "Create a support backup package.",
            "People Template": "Save a people import template CSV.",
            "Asset Template": "Save an asset import template CSV.",
            "All Templates": "Save both people and asset import templates.",
            "Import CSV": "Import people or assets from a completed template CSV.",
            "Settings": "Set store name, default due time, backup folder, and review asset categories.",
            "Role Test": "Show current operator role and permissions.",
            "Scanner Diagnostics": "Test a scan and see what the system detects.",
            "Health Check": "Show database and system status.",
            "Self-Test Info": "Show how to run the no-GUI self-test.",
            "Export System Report": "Save a system status report.",
            "Clear Errors": "Clear saved errors. Admin only.",
            "About": "Show app information.",
        }
        def walk(w):
            try:
                text = w.cget("text")
                if text in descriptions:
                    self.tip(w, descriptions[text])
            except Exception:
                pass
            try:
                if isinstance(w, ttk.Treeview):
                    self.tip(w, "Click a row to select it. Double-click rows where supported. Click column headers to sort.")
                elif isinstance(w, ttk.Combobox):
                    self.tip(w, "Choose an option from this list.")
                elif isinstance(w, ttk.Entry):
                    self.tip(w, "Type information here. Press Enter where supported.")
            except Exception:
                pass
            for child in w.winfo_children():
                walk(child)
        for page in self.pages.values():
            walk(page)


    def people_page(self):
        p = self.page("People / Users")
        top = self.panel(p, "People / Users", "Employees and operators. Admin users control roles.")
        top.pack(fill="both", expand=True)
        self.make_count_bar(top.body, [
            ("Total Users", "people"),
            ("Employees", "employee_users"),
            ("Front Desk", "front_desk_users"),
            ("AP Operators", "ap_operators"),
            ("Managers", "managers"),
            ("Admins", "admins"),
            ("Inactive", "inactive_people"),
        ], "people_count_vars")
        bar = tk.Frame(top.body, bg=self.colors["panel"])
        bar.pack(fill="x", pady=(0,8))
        ttk.Label(bar, text="View:", style="Panel.TLabel").pack(side="left", padx=(0,4))
        self.people_filter_buttons = {}
        for lab in ["All"] + self.db.group_names() + ["Inactive"]:
            chip = self.chip_button(bar, lab, command=lambda l=lab:self.set_people_filter(l), active=(lab == "All"))
            chip.pack(side="left", padx=3)
            self.people_filter_buttons[lab] = chip
        people_actions = tk.Frame(bar, bg=self.colors["panel"])
        people_actions.pack(side="right")
        ttk.Label(people_actions, text="Actions:", style="Panel.TLabel").pack(side="left", padx=(12,4))
        ttk.Button(people_actions, text="Add User", style="Green.TButton", command=lambda: self.popup_person(None) if self.require("people_edit") else None).pack(side="left", padx=3)
        ttk.Button(people_actions, text="Edit Selected", command=self.person_form).pack(side="left", padx=3)
        self.more_button(people_actions, [
            ("Open Profile", self.open_selected_person),
            ("People Template", self.export_people_import_template),
            ("View Current Checked-Out Assets", self.current_out_for_selected_person),
            None,
            ("Deactivate Selected", lambda: self.set_selected_person_status("Inactive")),
            ("Reactivate Selected", lambda: self.set_selected_person_status("Active")),
        ]).pack(side="left", padx=3)
        self.people_filter = "All"
        self.people_tree = self.tree(top.body, ("ID","Badge","Name","Role","Status","Department","Shift"), 20)
        self.people_tree.bind("<Double-1>", lambda e:self.open_selected_person())
        self.bind_context_menu(self.people_tree, self.people_context_actions)

    def assets_page(self):
        p = self.page("Assets")
        top = self.panel(p, "Assets", "Find, add, edit, retire, and review tracked assets.")
        top.pack(fill="both", expand=True)
        self.make_count_bar(top.body, [
            ("Total Assets", "assets"),
            ("Keys", "keys_total"),
            ("Radios", "radios_total"),
            ("Items", "items_total"),
            ("Available", "available_assets"),
            ("Inactive/Issue", "inactive_assets"),
        ], "asset_count_vars")
        search_bar = tk.Frame(top.body, bg=self.colors["panel"])
        search_bar.pack(fill="x", pady=(0,8))
        ttk.Label(search_bar, text="Find:", style="Panel.TLabel").pack(side="left", padx=(0,4))
        asset_search = ttk.Entry(search_bar, textvariable=self.asset_search_var, width=30)
        asset_search.pack(side="left", fill="x", expand=True, padx=(0,8))
        self.bind_enter(asset_search, self.refresh_assets)
        self.tip(asset_search, "Search barcode, name, key number, serial number, location, holder, status, or notes.")
        ttk.Label(search_bar, text="Type:", style="Panel.TLabel").pack(side="left", padx=(4,4))
        self.asset_filter_var = tk.StringVar(value=self.current_asset_filter)
        type_combo = ttk.Combobox(search_bar, textvariable=self.asset_filter_var, values=["All"] + ASSET_TYPES, state="readonly", width=12)
        type_combo.pack(side="left", padx=(0,8))
        type_combo.bind("<<ComboboxSelected>>", lambda e: self.set_asset_filter(self.asset_filter_var.get()))
        ttk.Label(search_bar, text="Status:", style="Panel.TLabel").pack(side="left", padx=(4,4))
        status_combo = ttk.Combobox(search_bar, textvariable=self.asset_status_filter, values=["All Status"] + ASSET_STATUS, state="readonly", width=13)
        status_combo.pack(side="left", padx=(0,8))
        status_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_assets())
        ttk.Label(search_bar, text="Sort:", style="Panel.TLabel").pack(side="left", padx=(4,4))
        ttk.Combobox(search_bar, textvariable=self.asset_sort_by, values=["Type","Name","Barcode","Status","Location","Key Number","Serial / Tag","Holder"], state="readonly", width=13).pack(side="left")
        ttk.Combobox(search_bar, textvariable=self.asset_sort_dir, values=["A-Z","Z-A"], state="readonly", width=6).pack(side="left", padx=4)
        ttk.Button(search_bar, text="Apply", command=self.refresh_assets).pack(side="left", padx=3)
        ttk.Button(search_bar, text="Clear", command=self.clear_asset_filters).pack(side="left", padx=3)

        action_bar = tk.Frame(top.body, bg=self.colors["panel"])
        action_bar.pack(fill="x", pady=(0,8))
        primary = tk.Frame(action_bar, bg=self.colors["panel"])
        primary.pack(side="left")
        ttk.Label(primary, text="Manage:", style="Panel.TLabel").pack(side="left", padx=(0,4))
        ttk.Button(primary, text="Add Asset", style="Green.TButton", command=self.quick_asset_form).pack(side="left", padx=3)
        ttk.Button(primary, text="Edit Selected", command=self.asset_form).pack(side="left", padx=3)

        workflow = tk.Frame(action_bar, bg=self.colors["panel"])
        workflow.pack(side="left", padx=(14,0))
        ttk.Label(workflow, text="Workflow:", style="Panel.TLabel").pack(side="left", padx=(0,4))
        self.more_button(workflow, [
            ("Use At Front Desk", self.use_selected_asset_frontdesk),
            ("Open Profile", self.open_selected_asset),
            ("Duplicate", self.duplicate_selected_asset),
            ("Export Selected", self.export_selected_asset_excel),
            ("Import Assets", self.import_from_template),
            None,
            ("Retire Asset", self.delete_selected_asset),
        ]).pack(side="left", padx=3)

        status_tools = tk.Frame(action_bar, bg=self.colors["panel"])
        status_tools.pack(side="left", padx=(14,0))
        ttk.Label(status_tools, text="Status:", style="Panel.TLabel").pack(side="left", padx=(0,4))
        ttk.Combobox(status_tools, textvariable=self.asset_quick_status_var, values=["Available","Repair","Missing","Retired"], state="readonly", width=10).pack(side="left", padx=3)
        ttk.Button(status_tools, text="Set", command=lambda: self.set_selected_asset_status(self.asset_quick_status_var.get())).pack(side="left", padx=3)

        tools = tk.Frame(action_bar, bg=self.colors["panel"])
        tools.pack(side="right")
        ttk.Button(tools, text="Export Excel", command=lambda: self.export_assets_excel(None if self.asset_export_type_var.get()=="All" else self.asset_export_type_var.get())).pack(side="right", padx=3)
        ttk.Combobox(tools, textvariable=self.asset_export_type_var, values=["All"] + ASSET_TYPES, state="readonly", width=12).pack(side="right", padx=3)
        ttk.Label(tools, text="Excel:", style="Panel.TLabel").pack(side="right", padx=(8,2))
        self.more_button(tools, [
            ("Export All Assets", lambda: self.export_assets_excel(None)),
            ("Export By Type", self.export_assets_by_type),
            ("Export Selected Asset", self.export_selected_asset_excel),
            ("Import Assets", self.import_from_template),
        ]).pack(side="right", padx=3)

        selected = tk.Label(top.body, textvariable=self.asset_selected_var, bg="#f8f9fa", fg="#111111", anchor="w", justify="left", font=("Segoe UI", 10, "bold"), padx=10, pady=7, wraplength=1200)
        selected.pack(fill="x", pady=(0,8))
        self.asset_tree = self.tree(top.body, ("Type","Barcode","Name","Status","Location","Key Number","Serial","Holder"), 20)
        self.asset_tree.bind("<Double-1>", lambda e:self.open_selected_asset())
        self.asset_tree.bind("<<TreeviewSelect>>", lambda e:self.update_asset_selection_summary())
        self.bind_context_menu(self.asset_tree, self.asset_context_actions)

    def search_page(self):
        p = self.page("Search")
        top = self.panel(p, "Search", "Find people, assets, activity, audit records, and errors.")
        top.pack(fill="both", expand=True)
        self.search_var = tk.StringVar()
        self.search_type_var = tk.StringVar(value="All")
        self.search_status_var = tk.StringVar(value="Any Status")
        bar = tk.Frame(top.body, bg=self.colors["panel"])
        bar.pack(fill="x", pady=(0,8))
        search_entry = ttk.Entry(bar, textvariable=self.search_var, font=("Segoe UI", 14))
        search_entry.pack(side="left", fill="x", expand=True, ipady=3, padx=(0,8))
        self.bind_enter(search_entry, self.search)
        self.tip(search_entry, "Type a search term and press Enter. Use filters to narrow results.")
        ttk.Combobox(bar, textvariable=self.search_type_var, values=["All","People","Assets","Activity","Audit","Errors","AP Alerts"], state="readonly", width=12).pack(side="left", padx=4)
        ttk.Combobox(bar, textvariable=self.search_status_var, values=["Any Status","Active","Available","Checked Out","Late","Returned","Inactive"], state="readonly", width=14).pack(side="left", padx=4)
        ttk.Button(bar, text="Search", style="Green.TButton", command=self.search).pack(side="left", padx=4)
        ttk.Button(bar, text="Clear", command=self.clear_search).pack(side="left", padx=4)
        modes = tk.Frame(top.body, bg=self.colors["panel"])
        modes.pack(fill="x", pady=(0,8))
        ttk.Label(modes, text="Quick modes:", style="Panel.TLabel").pack(side="left", padx=(0,4))
        self.search_mode_buttons = {}
        for mode in ["People", "Assets", "Logs", "Audit", "Errors"]:
            target = "Activity" if mode == "Logs" else mode
            chip = self.chip_button(modes, mode, command=lambda m=target: self.set_search_mode(m), active=False)
            chip.pack(side="left", padx=3)
            self.search_mode_buttons[target] = chip
        self.search_tree = self.tree(top.body, ("Type","Primary","Name/Title","Status/Action","Details"), 20)
        self.bind_context_menu(self.search_tree, lambda t: [
            ("View Full Row", lambda: self.show_full_tree_row(t, "Search Result"), bool(self.selected_tree_values(t))),
            ("Copy Row", lambda: self.copy_tree_row(t), bool(self.selected_tree_values(t))),
            ("Search Again", self.search),
        ])


    def manager_page(self):
        p = self.page("Manager")
        top = self.panel(p, "Manager", "Daily oversight for open assets, issue assets, backups, and reports.")
        top.pack(fill="both", expand=True)
        self.make_count_bar(top.body, [
            ("Open", "out"),
            ("Late", "late_returns"),
            ("Issues", "inactive_assets"),
            ("Available", "available_assets"),
            ("Active Alerts", "manager_alerts"),
            ("Errors Today", "errors"),
        ], "manager_count_vars")

        priority = tk.Frame(top.body, bg=self.colors["panel"])
        priority.pack(fill="x", pady=(0,8))
        tk.Label(priority, text="Today's Priorities", bg=self.colors["panel"], fg="#17202a", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=2, pady=(0,4))
        priority_cards = tk.Frame(priority, bg=self.colors["panel"])
        priority_cards.pack(fill="x")
        self.manager_focus_cards = {}
        priority_items = [
            ("Backup", "backup", "Backup Now", self.backup_with_logs),
            ("Late Returns", "late", "Overdue Report", self.report_overdue),
            ("Alerts", "alerts", "Review Alerts", self.ap_alerts_popup),
            ("Open Assets", "open", "Current Out", self.report_current_out),
        ]
        for label, key, button_text, command in priority_items:
            card = tk.Frame(priority_cards, bg="#f8fafc", highlightbackground=self.colors["line"], highlightthickness=1)
            card.pack(side="left", fill="x", expand=True, padx=(0,8))
            tk.Label(card, text=label, bg="#f8fafc", fg=self.colors["muted"], font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10, pady=(8,0))
            tk.Label(card, textvariable=self.manager_focus_vars[key], bg="#f8fafc", fg="#17202a", font=("Segoe UI", 12, "bold"), anchor="w", justify="left", wraplength=250).pack(anchor="w", padx=10, pady=(1,6))
            self.tip(ttk.Button(card, text=button_text, command=command), f"Open {button_text}.").pack(anchor="w", padx=10, pady=(0,9))
            self.manager_focus_cards[key] = card

        status_band = tk.Frame(top.body, bg=self.colors["panel"], padx=0, pady=0)
        status_band.pack(fill="x", pady=(0,8))
        status_items = [
            ("Last Backup Date", self.manager_backup_date_var),
            ("Last Backup Time", self.manager_backup_time_var),
            ("Backup Status", self.manager_backup_status_var),
            ("Auto Refresh", self.manager_refresh_countdown_var),
        ]
        self.manager_status_cards = {}
        self.manager_status_value_labels = {}
        for label, var in status_items:
            box = tk.Frame(status_band, bg=self.colors["panel_alt"], highlightbackground=self.colors["line"], highlightthickness=1)
            box.pack(side="left", fill="x", expand=True, padx=(0,10))
            tk.Label(box, text=label, bg=self.colors["panel_alt"], fg=self.colors["muted"], font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10, pady=(8,0))
            value_label = tk.Label(box, textvariable=var, bg=self.colors["panel_alt"], fg="#111111", font=("Segoe UI", 12, "bold"), wraplength=260, justify="left")
            value_label.pack(anchor="w", padx=10, pady=(0,8))
            self.manager_status_cards[label] = box
            self.manager_status_value_labels[label] = value_label
        self.tip(ttk.Button(status_band, text="Refresh Now", command=self.manual_manager_refresh), "Refresh Manager counts, backup status, and open asset lists now.").pack(side="right", padx=(4,0))

        actions = tk.Frame(top.body, bg=self.colors["panel"])
        actions.pack(fill="x", pady=(0,8))
        ttk.Label(actions, text="Daily:", style="Panel.TLabel").pack(side="left", padx=(0,4))
        self.tip(ttk.Button(actions, text="Backup Now", command=self.backup_with_logs, style="Green.TButton"), "Create a backup now.").pack(side="left", padx=3)
        self.tip(ttk.Button(actions, text="Closeout Checklist", command=self.closeout_checklist), "Open the daily closeout checklist.").pack(side="left", padx=3)
        self.tip(ttk.Button(actions, text="End of Shift", command=self.report_end_shift), "Generate the end-of-shift report.").pack(side="left", padx=3)
        ttk.Label(actions, text="Reports:", style="Panel.TLabel").pack(side="left", padx=(18,4))
        self.tip(ttk.Button(actions, text="Current Out", command=self.report_current_out), "Show all assets currently checked out.").pack(side="left", padx=3)
        self.tip(ttk.Button(actions, text="Overdue Report", command=self.report_overdue), "Show late returns that need attention.").pack(side="left", padx=3)
        ttk.Label(actions, text="View:", style="Panel.TLabel").pack(side="left", padx=(18,4))
        self.tip(ttk.Checkbutton(actions, text="Only show problems", variable=self.manager_show_problems_only, command=self.refresh_manager), "Hide normal open rows and show late, warning, missing, repair, retired, critical, or unresolved items.").pack(side="left", padx=3)
        ttk.Label(actions, text="More:", style="Panel.TLabel").pack(side="left", padx=(18,4))
        self.more_button(actions, [
            ("Refresh Manager", self.manual_manager_refresh),
            ("Log Viewer", lambda: self.show("Logs")),
            ("Audit Review", self.report_audit),
            ("Error Review", self.report_errors),
            None,
            ("Overdue", self.report_overdue),
            ("Asset Issues", self.report_asset_issues),
            ("AP Alerts", self.ap_alerts_popup),
            ("Backup History", self.backup_history_popup),
            ("Export History", self.export_history_popup),
            None,
            ("System Hub", lambda: self.show("System")),
        ]).pack(side="left", padx=3)

        status = tk.Label(top.body, textvariable=self.data_status, bg="#f8f9fa", fg="#111111", anchor="w", justify="left", font=("Segoe UI", 10, "bold"), padx=10, pady=7)
        status.pack(fill="x", pady=(0,8))
        health = tk.Label(top.body, textvariable=self.manager_health_var, bg="#f8f9fa", fg="#111111", anchor="w", justify="left", font=("Segoe UI", 10), padx=10, pady=7, wraplength=1280)
        health.pack(fill="x", pady=(0,8))

        alerts = self.panel(top.body, "Manager Notifications", "New warnings, blocked actions, overrides, errors, and review items.")
        alerts.pack(fill="x", pady=(0,8))
        alert_buttons = tk.Frame(alerts.body, bg=self.colors["panel"])
        alert_buttons.pack(fill="x", pady=(0,6))
        ttk.Label(alert_buttons, text="Status:", style="Panel.TLabel").pack(side="left", padx=(0,2))
        ttk.Combobox(alert_buttons, textvariable=self.manager_notification_status_var, values=["All","New","Reviewed","Resolved"], state="readonly", width=10).pack(side="left", padx=3)
        ttk.Label(alert_buttons, text="Severity:", style="Panel.TLabel").pack(side="left", padx=(6,2))
        ttk.Combobox(alert_buttons, textvariable=self.manager_notification_severity_var, values=["All","Info","Warning","Critical"], state="readonly", width=10).pack(side="left", padx=3)
        ttk.Label(alert_buttons, text="Find:", style="Panel.TLabel").pack(side="left", padx=(6,2))
        entry = ttk.Entry(alert_buttons, textvariable=self.manager_notification_search_var, width=18)
        entry.pack(side="left", padx=3)
        self.bind_enter(entry, self.refresh_manager)
        ttk.Button(alert_buttons, text="Mark Reviewed", command=lambda: self.update_selected_notification("Reviewed")).pack(side="left", padx=3)
        ttk.Button(alert_buttons, text="Mark Resolved", command=lambda: self.update_selected_notification("Resolved")).pack(side="left", padx=3)
        ttk.Button(alert_buttons, text="Add Note", command=self.add_note_to_selected_notification).pack(side="left", padx=3)
        self.more_button(alert_buttons, [
            ("Export Notifications", self.export_manager_notifications_excel),
            ("Refresh Notifications", self.manual_manager_refresh),
            ("Open Logs", lambda: self.show("Logs")),
        ]).pack(side="left", padx=3)
        self.manager_notifications_tree = self.tree(alerts.body, ("ID","Time","Severity","Event","User","Asset","Status","Notes"), 6)
        self.bind_context_menu(self.manager_notifications_tree, self.manager_notification_context_actions)

        grids = tk.Frame(top.body, bg=self.colors["panel"])
        grids.pack(fill="both", expand=True)
        left = self.panel(grids, "Open Assets", "Assets currently checked out or under review.")
        left.pack(side="left", fill="both", expand=True, padx=(0,5))
        right = self.panel(grids, "Issue Assets", "Assets marked missing, repair, retired, or checked out.")
        right.pack(side="left", fill="both", expand=True, padx=(5,0))
        self.manager_open_tree = self.tree(left.body, ("Status","Type","Asset","Holder","Due"), 12)
        self.manager_issue_tree = self.tree(right.body, ("Status","Type","Asset","Location","Holder"), 12)


    def closeout_checklist(self):
        key = "closeout_checklist"
        if self.focus_existing(key):
            return
        win = tk.Toplevel(self)
        win.title("Daily Closeout Checklist")
        win.geometry("760x620")
        self.apply_icon(win)
        self.track_window(key, win)
        pan = self.panel(win, "Daily Closeout Checklist", "Complete these before generating the End of Shift report.")
        pan.pack(fill="both", expand=True, padx=14, pady=14)
        items = [
            "Keys reviewed",
            "Radios reviewed",
            "Temp badges reviewed",
            "Late returns checked",
            "Errors reviewed",
            "Backup completed",
            "Notes added",
        ]
        vars = []
        for item in items:
            v = tk.BooleanVar(value=False)
            vars.append((item, v))
            cb = ttk.Checkbutton(pan.body, text=item, variable=v)
            cb.pack(anchor="w", pady=4)
            self.tip(cb, f"Mark complete: {item}.")
        notes_var = tk.StringVar()
        ttk.Label(pan.body, text="Closeout notes:", style="Panel.TLabel").pack(anchor="w", pady=(12,2))
        notes = tk.Text(pan.body, height=6, wrap="word")
        notes.pack(fill="both", expand=True)
        def save_and_report():
            completed = [item for item, v in vars if v.get()]
            missing = [item for item, v in vars if not v.get()]
            note_text = notes.get("1.0", "end").strip()
            detail = json.dumps({"completed": completed, "missing": missing, "notes": note_text}, indent=2)
            self.db.audit("DAILY CLOSEOUT CHECKLIST", self.actor(), "Reports", detail)
            if missing and not messagebox.askyesno("Checklist Incomplete", "Some checklist items are not complete. Generate End of Shift report anyway?"):
                return
            win.destroy()
            self.report_end_shift(extra_closeout=detail)
        self.tip(ttk.Button(pan.body, text="Save Checklist + Generate End of Shift", style="Green.TButton", command=save_and_report), "Save this checklist to audit and generate the End of Shift report.").pack(anchor="w", pady=10)



    def reports_page(self):
        p = self.page("Reports")
        top = self.panel(p, "Reports", "Current out, overdue assets, issue lists, history, exports, and end-of-shift summaries.")
        top.pack(fill="both", expand=True)
        bar = tk.Frame(top.body, bg=self.colors["panel"])
        bar.pack(fill="x", pady=(0,8))
        report_groups = [
            ("Daily:", [
                ("Current Out", self.report_current_out),
                ("Overdue", self.report_overdue),
                ("Asset Issues", self.report_asset_issues),
                ("Closeout Checklist", self.closeout_checklist),
                ("End of Shift", self.report_end_shift),
            ]),
            ("History:", [
                ("Employee History", self.report_employee_history),
                ("Asset History", self.report_asset_history),
                ("Alert History", self.report_alert_history),
                ("Group History", self.report_group_history),
                ("Detailed Audit Report", self.report_audit),
                ("Detailed Error Report", self.report_errors),
            ]),
            ("Export:", [
                ("Export CSV Bundle", self.export_xlsx),
                ("Backup with Logs", self.backup_with_logs),
            ]),
            ("Output:", [
                ("Copy Report", self.copy_report),
                ("Save TXT", self.save_report_txt),
                ("Save HTML", self.save_report_html),
                ("Save Excel", self.save_report_excel),
                ("Print", self.print_report),
            ]),
        ]
        tips = {
                "Current Out": "Show all assets currently checked out with holder and due details.",
                "Overdue": "Show open items that are past their due-back time.",
                "Asset Issues": "Show missing, repair, retired, and checked-out assets.",
                "Employee History": "Prompt for an employee ID or badge and show checkout history.",
                "Asset History": "Prompt for an asset barcode, key number, or serial and show history.",
                "Alert History": "Show AP alerts and Manager notifications in one report.",
                "Group History": "Show group membership and permission-change history.",
                "Closeout Checklist": "Complete a daily checklist and then generate the end-of-shift report.",
                "End of Shift": "Generate a detailed shift summary with open items, errors, and audit entries.",
                "Detailed Audit Report": "Show full audit history details.",
                "Detailed Error Report": "Show saved application errors and details.",
                "Export CSV Bundle": "Export people, assets, activity, audit, errors, and settings CSV files.",
                "Backup with Logs": "Create a backup ZIP with the database and logs.",
                "Copy Report": "Copy the current report text to the clipboard.",
                "Save TXT": "Save the current report as a text file.",
                "Save HTML": "Save the current report as a simple printable HTML file.",
                "Save Excel": "Save the current report as a simple Excel workbook.",
                "Print": "Print or save the report currently shown below.",
        }
        for row_index, (title, buttons) in enumerate(report_groups):
            group = tk.Frame(bar, bg=self.colors["panel"])
            group.grid(row=row_index, column=0, sticky="ew", pady=2)
            ttk.Label(group, text=title, style="Panel.TLabel").pack(side="left", padx=(0,4))
            for label, cmd in buttons:
                self.tip(ttk.Button(group, text=label, command=cmd), tips.get(label, label)).pack(side="left", padx=3)
        bar.grid_columnconfigure(0, weight=1)
        report_wrap = tk.Frame(top.body, bg=self.colors["panel"])
        report_wrap.pack(fill="both", expand=True)
        self.report_text = tk.Text(report_wrap, wrap="word", font=("Segoe UI", 10), bg="white", fg="#111", padx=14, pady=12)
        report_scroll = ttk.Scrollbar(report_wrap, orient="vertical", command=self.report_text.yview)
        self.report_text.configure(yscrollcommand=report_scroll.set)
        self.report_text.grid(row=0, column=0, sticky="nsew")
        report_scroll.grid(row=0, column=1, sticky="ns")
        report_wrap.grid_rowconfigure(0, weight=1)
        report_wrap.grid_columnconfigure(0, weight=1)
        self.report_text.insert("1.0", "Select a report above. Results will appear here.")

    def logs_page(self):
        p = self.page("Logs")
        top = self.panel(p, "Logs", "Audit records, error records, AP alerts, and Manager notifications.")
        top.pack(fill="both", expand=True)
        controls = tk.Frame(top.body, bg=self.colors["panel"])
        controls.pack(fill="x", pady=(0,8))
        self.log_filter_var = tk.StringVar(value="All Logs")
        self.log_search_var = tk.StringVar()
        self.log_user_var = tk.StringVar()
        self.log_asset_var = tk.StringVar()
        self.log_action_var = tk.StringVar()
        self.log_from_var = tk.StringVar()
        self.log_to_var = tk.StringVar()
        ttk.Label(controls, text="Type:", style="Panel.TLabel").pack(side="left")
        ttk.Combobox(controls, textvariable=self.log_filter_var, values=["All Logs","Audit","Errors","Manager Notifications","AP Alerts"], state="readonly", width=22).pack(side="left", padx=4)
        ttk.Label(controls, text="Find:", style="Panel.TLabel").pack(side="left", padx=(8,2))
        entry = ttk.Entry(controls, textvariable=self.log_search_var)
        entry.pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(controls, text="Search", command=self.refresh_logs_page).pack(side="left", padx=3)
        ttk.Button(controls, text="Export Logs", command=self.export_logs_excel).pack(side="left", padx=3)
        self.more_button(controls, [
            ("Export Notifications", self.export_manager_notifications_excel),
            ("Open Log Folder", lambda: self.open_folder_path(self.folder_setting("system_log_folder", "Logs"))),
            None,
            ("Archive / Clear Old Logs", self.clear_old_logs),
        ]).pack(side="left", padx=3)
        filters = tk.Frame(top.body, bg=self.colors["panel"])
        filters.pack(fill="x", pady=(0,8))
        for label, var, width, tip_text in [
            ("User:", self.log_user_var, 18, "Filter by user, actor, source, or alert target."),
            ("Asset:", self.log_asset_var, 18, "Filter by asset barcode, asset name, or target."),
            ("Action:", self.log_action_var, 18, "Filter by action, event, or alert type."),
            ("From:", self.log_from_var, 12, "Optional start date using YYYY-MM-DD."),
            ("To:", self.log_to_var, 12, "Optional end date using YYYY-MM-DD."),
        ]:
            ttk.Label(filters, text=label, style="Panel.TLabel").pack(side="left", padx=(8,2))
            f_entry = ttk.Entry(filters, textvariable=var, width=width)
            f_entry.pack(side="left", padx=(0,4))
            self.tip(f_entry, tip_text)
            self.bind_enter(f_entry, self.refresh_logs_page)
        self.logs_tree = self.tree(top.body, ("Time","Type","User/Source","Action","Target","Details"), 18)
        self.bind_enter(entry, self.refresh_logs_page)
        self.bind_context_menu(self.logs_tree, self.logs_context_actions)

    def groups_permissions_page(self):
        p = self.page("Groups / Permissions")
        top = self.panel(p, "Groups / Permissions", "Manage groups, rights, and assigned users.")
        top.pack(fill="both", expand=True)
        tk.Label(top.body, text="Group tools open in a focused management window so edits can be reviewed before saving.", bg=self.colors["panel"], fg="#333333", font=("Segoe UI", 11), anchor="w", justify="left").pack(fill="x", pady=(0,8))
        ttk.Button(top.body, text="Open Group Management", style="Green.TButton", command=self.group_management).pack(anchor="w", pady=4)

    def settings_page(self):
        p = self.page("Settings")
        top = self.panel(p, "Settings", "Data file, backup folders, export folders, due times, and app preferences.")
        top.pack(fill="both", expand=True)
        actions = [
            ("App Settings", self.app_settings),
            ("Show Data Location", self.show_data_location),
            ("Change Data File Location", self.change_data_location),
            ("Open Data Folder", self.open_data_folder),
            ("Open Backup Folder", self.open_backup_folder),
            ("Health Check", self.health_check),
        ]
        for label, cmd in actions:
            ttk.Button(top.body, text=label, command=cmd).pack(anchor="w", pady=4)

    def admin_tools_page(self):
        p = self.page("Admin Tools")
        top = self.panel(p, "Admin Tools", "Restricted tools for diagnostics and cleanup.")
        top.pack(fill="both", expand=True)
        tools = [
            ("Scanner Diagnostics", self.scanner_diagnostics),
            ("Role Test", self.role_test),
            ("Self-Test Info", self.selftest_info),
            ("Export System Report", self.export_system_report),
            ("Clear Errors", self.clear_errors),
            ("About", self.about),
        ]
        for label, cmd in tools:
            ttk.Button(top.body, text=label, command=cmd).pack(anchor="w", pady=4)

    def open_folder_path(self, folder):
        folder = Path(folder)
        folder.mkdir(parents=True, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(folder)
        else:
            messagebox.showinfo("Folder", str(folder))

    def log_rows_for_view(self):
        if not hasattr(self, "logs_tree"):
            return []
        q = clean(self.log_search_var.get()).lower()
        user_filter = clean(self.log_user_var.get()).lower() if hasattr(self, "log_user_var") else ""
        asset_filter = clean(self.log_asset_var.get()).lower() if hasattr(self, "log_asset_var") else ""
        action_filter = clean(self.log_action_var.get()).lower() if hasattr(self, "log_action_var") else ""
        date_from = clean(self.log_from_var.get()) if hasattr(self, "log_from_var") else ""
        date_to = clean(self.log_to_var.get()) if hasattr(self, "log_to_var") else ""
        mode = self.log_filter_var.get()
        rows = []
        if mode in ("All Logs", "Audit"):
            for r in self.db.all("SELECT timestamp, actor, action, target, details FROM audit ORDER BY id DESC LIMIT 500"):
                rows.append((r["timestamp"], "Audit", r["actor"], r["action"], r["target"], r["details"]))
        if mode in ("All Logs", "Errors"):
            for r in self.db.all("SELECT timestamp, source, message, details FROM errors ORDER BY id DESC LIMIT 500"):
                rows.append((r["timestamp"], "Error", r["source"], r["message"], "", r["details"]))
        if mode in ("All Logs", "Manager Notifications"):
            for r in self.db.all("SELECT timestamp, handled_by, event_type, user_involved, asset_involved, notes, status, severity FROM manager_notifications ORDER BY id DESC LIMIT 500"):
                target = " | ".join(x for x in [r["user_involved"], r["asset_involved"]] if x)
                rows.append((r["timestamp"], "Manager", r["handled_by"], r["event_type"], target, f"{r['severity']} | {r['status']} | {r['notes']}"))
        if mode in ("All Logs", "AP Alerts"):
            for r in self.db.all("SELECT timestamp, target_label, target_type, target_key, alert_type, severity, note, required_action, created_by, status FROM ap_alerts ORDER BY id DESC LIMIT 500"):
                target = f"{r['target_type']} {r['target_key']} | {r['target_label']}"
                details = f"{r['severity']} | {r['status']} | {r['note']} | Required: {r['required_action']}"
                rows.append((r["timestamp"], "AP Alert", r["created_by"], r["alert_type"], target, details))
        rows.sort(key=lambda r: r[0], reverse=True)
        shown = []
        for r in rows[:500]:
            raw_date = str(r[0] or "")[:10]
            if date_from and raw_date < date_from:
                continue
            if date_to and raw_date > date_to:
                continue
            raw_values = [str(v or "") for v in r]
            raw_text = " ".join(raw_values).lower()
            if user_filter and user_filter not in (str(r[2] or "") + " " + str(r[5] or "")).lower():
                continue
            if asset_filter and asset_filter not in (str(r[4] or "") + " " + str(r[5] or "")).lower():
                continue
            if action_filter and action_filter not in (str(r[3] or "") + " " + str(r[5] or "")).lower():
                continue
            values = (pretty(r[0]), r[1], r[2], r[3], r[4], r[5])
            if not q or q in raw_text or q in " ".join(str(v).lower() for v in values):
                shown.append(values)
        return shown

    def refresh_logs_page(self):
        if not hasattr(self, "logs_tree"):
            return
        shown = self.log_rows_for_view()
        self.fill(self.logs_tree, shown)

    def export_logs_excel(self):
        rows = self.log_rows_for_view()
        path = self.folder_setting("excel_export_folder", "Exports") / f"logs_{safe_filename_part(self.log_filter_var.get().lower().replace(' ', '_'), 'all_logs')}_{stamp()}.xlsx"
        headers = ["Time", "Type", "User/Source", "Action", "Target", "Details"]
        try:
            self.write_xlsx(path, [("Logs", headers, rows)])
            detail = f"Mode: {self.log_filter_var.get()}; Search: {self.log_search_var.get()}; User: {self.log_user_var.get()}; Asset: {self.log_asset_var.get()}; Action: {self.log_action_var.get()}; From: {self.log_from_var.get()}; To: {self.log_to_var.get()}; Rows: {len(rows)}; File: {path}"
            self.db.audit("EXPORT LOGS EXCEL", self.actor(), "Logs", detail)
            self.db.notify_manager("Log exported", "Info", self.actor(), "", "Filtered logs exported", self.actor(), detail, status="Reviewed")
            messagebox.showinfo("Export Logs", f"Logs exported:\n{path}")
        except Exception as e:
            self.db.error("Export Logs Excel", str(e), traceback.format_exc())
            messagebox.showerror("Export Logs", f"Could not export logs:\n{e}")

    def clear_old_logs(self):
        if not self.require("admin"):
            return
        title = "Archive / Clear Old Logs"
        days = simpledialog.askinteger(title, "Archive and clear log entries older than how many days?", parent=self, minvalue=30, maxvalue=3650)
        if not days:
            return
        reason = simpledialog.askstring(title, "Reason for archiving and clearing old logs:", parent=self)
        if not clean(reason):
            messagebox.showwarning("Reason Required", "A reason is required before logs can be cleared.")
            return
        cutoff = (dt.datetime.now() - dt.timedelta(days=days)).isoformat(timespec="seconds")
        counts = {
            "audit": self.db.one("SELECT COUNT(*) c FROM audit WHERE timestamp < ?", (cutoff,))["c"],
            "errors": self.db.one("SELECT COUNT(*) c FROM errors WHERE timestamp < ?", (cutoff,))["c"],
            "manager notifications": self.db.one("SELECT COUNT(*) c FROM manager_notifications WHERE timestamp < ?", (cutoff,))["c"],
            "resolved AP alerts": self.db.one("SELECT COUNT(*) c FROM ap_alerts WHERE timestamp < ? AND status!='Active'", (cutoff,))["c"],
        }
        total = sum(counts.values())
        if not total:
            messagebox.showinfo(title, "No old log records matched that cutoff.")
            return
        summary = "\n".join(f"{name}: {count}" for name, count in counts.items())
        if not messagebox.askyesno("Confirm Log Archive", f"This will archive and then clear records older than {days} days.\n\n{summary}\n\nContinue?"):
            return
        try:
            archive_path = self.folder_setting("excel_export_folder", "Exports") / f"archived_logs_before_{cutoff[:10]}_{stamp()}.xlsx"
            archive_sheets = []
            audit_rows = self.db.all("SELECT * FROM audit WHERE timestamp < ? ORDER BY timestamp", (cutoff,))
            error_rows = self.db.all("SELECT * FROM errors WHERE timestamp < ? ORDER BY timestamp", (cutoff,))
            note_rows = self.db.all("SELECT * FROM manager_notifications WHERE timestamp < ? ORDER BY timestamp", (cutoff,))
            alert_rows = self.db.all("SELECT * FROM ap_alerts WHERE timestamp < ? AND status!='Active' ORDER BY timestamp", (cutoff,))
            if audit_rows:
                headers = [h for h in self.table_columns("audit")]
                archive_sheets.append(("Audit", headers, [[r[h] for h in headers] for r in audit_rows]))
            if error_rows:
                headers = [h for h in self.table_columns("errors")]
                archive_sheets.append(("Errors", headers, [[r[h] for h in headers] for r in error_rows]))
            if note_rows:
                headers = [h for h in self.table_columns("manager_notifications")]
                archive_sheets.append(("Manager Notifications", headers, [[r[h] for h in headers] for r in note_rows]))
            if alert_rows:
                headers = [h for h in self.table_columns("ap_alerts")]
                archive_sheets.append(("AP Alerts", headers, [[r[h] for h in headers] for r in alert_rows]))
            if archive_sheets:
                self.write_xlsx(archive_path, archive_sheets)
            self.db.run("DELETE FROM audit WHERE timestamp < ?", (cutoff,))
            self.db.run("DELETE FROM errors WHERE timestamp < ?", (cutoff,))
            self.db.run("DELETE FROM manager_notifications WHERE timestamp < ?", (cutoff,))
            self.db.run("DELETE FROM ap_alerts WHERE timestamp < ? AND status!='Active'", (cutoff,))
            self.db.audit("ARCHIVE AND CLEAR OLD LOGS", self.actor(), "Logs", f"Cutoff={cutoff}; Counts={counts}; Archive={archive_path}; Reason={clean(reason)}")
            self.db.notify_manager("Log cleanup completed", "Warning", self.actor(), "", "Old logs archived and cleared", self.actor(), f"Cutoff={cutoff}; Counts={counts}; Archive={archive_path}", clean(reason), status="Reviewed")
            self.refresh_logs_page()
            messagebox.showinfo(title, f"Old logs archived and cleared.\n\nArchive:\n{archive_path}\n\n{summary}")
        except Exception as e:
            self.db.error(title, str(e), traceback.format_exc())
            messagebox.showerror(title, f"Could not archive/clear old logs:\n{e}")


    def system_page(self):
        p = self.page("System")
        top = self.panel(p, "System", "Data location, backups, imports, settings, diagnostics, and support tools.")
        top.pack(fill="both", expand=True)
        buttons = tk.Frame(top.body, bg=self.colors["panel"])
        buttons.pack(fill="x", pady=(0,10))
        sections = [
            ("Data & Storage", [
                ("Show Data Location", self.show_data_location),
                ("Open Data Folder", self.open_data_folder),
                ("Change Data File Location", self.change_data_location)
            ]),
            ("Backup & Restore", [
                ("Backup Now", self.backup_with_logs),
                ("Open Backup Folder", self.open_backup_folder),
                ("Restore Backup", self.restore_backup),
                ("Issue Report ZIP", self.issue_report)
            ]),
            ("Templates & Imports", [
                ("People Template", self.export_people_import_template),
                ("Asset Template", self.export_asset_import_template),
                ("All Templates", self.export_import_templates),
                ("Import CSV", self.import_from_template)
            ]),
            ("App & Security", [
                ("Settings", self.app_settings),
                ("Group Management", self.group_management),
                ("Reset Demo Data", self.reset_training_demo_data),
                ("Scanner Diagnostics", self.scanner_diagnostics),
                ("Role Test", self.role_test)
            ]),
            ("Diagnostics", [
                ("Health Check", self.health_check),
                ("Self-Test Info", self.selftest_info),
                ("Export System Report", self.export_system_report)
            ]),
            ("Advanced", [
                ("Clear Errors", self.clear_errors),
                ("About", self.about)
            ]),
        ]
        for idx, (title, btns) in enumerate(sections):
            card = tk.Frame(buttons, bg=self.colors["panel_alt"], highlightbackground=self.colors["line"], highlightthickness=1)
            card.grid(row=idx // 2, column=idx % 2, sticky="nsew", padx=6, pady=6)
            buttons.grid_columnconfigure(idx % 2, weight=1, uniform="system_cards")
            tk.Label(card, text=title, bg=self.colors["panel_alt"], fg="#17202a", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=12, pady=(10,4))
            grid = tk.Frame(card, bg=self.colors["panel_alt"])
            grid.pack(fill="x", padx=8, pady=(0,10))
            for i, (label, cmd) in enumerate(btns):
                style_name = "Green.TButton" if label in ("Backup Now", "Settings", "Group Management") else "TButton"
                ttk.Button(grid, text=label, command=cmd, style=style_name).grid(row=i // 2, column=i % 2, padx=4, pady=4, sticky="ew")
                grid.grid_columnconfigure(i % 2, weight=1)
        system_text_wrap = tk.Frame(top.body, bg=self.colors["panel"])
        system_text_wrap.pack(fill="both", expand=True)
        self.system_text = tk.Text(system_text_wrap, wrap="word", font=("Consolas", 11), bg="white", fg="#111", padx=10, pady=10)
        system_y = ttk.Scrollbar(system_text_wrap, orient="vertical", command=self.system_text.yview)
        system_x = ttk.Scrollbar(system_text_wrap, orient="horizontal", command=self.system_text.xview)
        self.system_text.configure(yscrollcommand=system_y.set, xscrollcommand=system_x.set)
        self.system_text.grid(row=0, column=0, sticky="nsew")
        system_y.grid(row=0, column=1, sticky="ns")
        system_x.grid(row=1, column=0, sticky="ew")
        system_text_wrap.grid_rowconfigure(0, weight=1)
        system_text_wrap.grid_columnconfigure(0, weight=1)
        self.system_text.insert("1.0", "Choose a System action above. Details will appear here.")

    def refresh_page(self, name=None):
        self.update_data_status()
        name = name or self.current_page
        if name == "Dashboard":
            self.refresh_dashboard()
        elif name == "Front Desk":
            self.refresh_current()
            self.refresh_scan_mode()
        elif name == "People / Users":
            self.refresh_people()
        elif name == "Assets":
            self.refresh_assets()
        elif name == "Manager":
            self.refresh_manager()
        elif name == "Logs":
            self.refresh_logs_page()
        elif name == "System":
            self.show_data_location()

    def refresh(self):
        self.refresh_page(self.current_page)
        self.update_operator_badges()

    def auto_refresh(self):
        try:
            self.auto_daily_backup()
            self.refresh_page(self.current_page)
        finally:
            self.next_auto_refresh_at = time.time() + (self.refresh_interval_ms() / 1000)
            self.after(self.refresh_interval_ms(), self.auto_refresh)

    def tick_refresh_countdown(self):
        try:
            remaining = max(0, int(self.next_auto_refresh_at - time.time()))
            minutes, seconds = divmod(remaining, 60)
            self.manager_refresh_countdown_var.set(f"Refresh in {minutes}:{seconds:02d}")
        except Exception:
            pass
        self.after(1000, self.tick_refresh_countdown)

    def manual_manager_refresh(self):
        self.next_auto_refresh_at = time.time() + (self.refresh_interval_ms() / 1000)
        self.update_data_status()
        self.refresh_manager()
        self.status.set("Manager page refreshed manually.")

    def set_manager_status_card(self, label, color):
        try:
            card = self.manager_status_cards.get(label)
            value = self.manager_status_value_labels.get(label)
            if card:
                card.configure(highlightbackground=color, highlightthickness=2)
            if value:
                value.configure(fg=color)
        except Exception:
            pass

    def set_manager_focus_card(self, key, color):
        try:
            card = self.manager_focus_cards.get(key)
            if card:
                card.configure(highlightbackground=color, highlightthickness=2)
                for child in card.winfo_children():
                    if isinstance(child, tk.Label) and str(child.cget("textvariable")):
                        child.configure(fg=color)
        except Exception:
            pass

    def refresh_dashboard(self):
        c = self.db.counts()
        for k,v in c.items():
            if k in self.metric_vars:
                self.metric_vars[k].set(str(v))
        if "backup_status" in self.metric_vars:
            self.metric_vars["backup_status"].set("OK" if self.today_backup_ok() else "Due")
        if hasattr(self, "dashboard_status_vars"):
            self.dashboard_status_vars["database"].set("OK" if Path(self.db.path).exists() else "Missing")
            self.dashboard_status_vars["backup"].set("OK today" if self.today_backup_ok() else "Due")
            self.dashboard_status_vars["auto_refresh"].set("On" if self.auto_backup_enabled() else "Off")
            self.dashboard_status_vars["late"].set(str(c.get("late_returns", 0)))
            self.dashboard_status_vars["errors"].set(str(c.get("errors", 0)))
        rows = self.db.all("SELECT * FROM activity ORDER BY COALESCE(returned_at,checked_out_at) DESC LIMIT 50")
        self.fill(self.dashboard_tree, [
            (
                pretty(r["returned_at"] or r["checked_out_at"]),
                r["status"],
                r["asset_type"],
                f"{r['asset_barcode']} / {r['asset_name']}",
                r["employee_name"],
                r["return_operator"] or r["checkout_operator"],
                pretty(r["returned_at"] or r["due_back_at"])
            )
            for r in rows
        ])


    def refresh_current(self):
        rows = self.db.all("SELECT * FROM activity WHERE status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') ORDER BY due_back_at")
        self.fill(self.current_tree, [(r["status"], r["asset_type"], f"{r['asset_barcode']} / {r['asset_name']}", r["employee_name"], pretty(r["due_back_at"])) for r in rows])

    def refresh_manager(self):
        self.refresh_manager_counts()
        c = self.db.counts()
        last_backup = self.db.setting("last_backup", "")
        if last_backup:
            pretty_backup = pretty(last_backup)
            parts = pretty_backup.split()
            self.manager_backup_date_var.set(parts[0] if parts else pretty_backup)
            self.manager_backup_time_var.set(" ".join(parts[1:]) if len(parts) > 1 else "")
            self.manager_backup_status_var.set("Backed up today" if self.today_backup_ok() else "Backup needed today")
        else:
            self.manager_backup_date_var.set("No backup found")
            self.manager_backup_time_var.set("")
            self.manager_backup_status_var.set("Backup needed")
        backup_color = self.colors["green"] if self.today_backup_ok() else self.colors["red"]
        self.set_manager_status_card("Last Backup Date", backup_color if last_backup else self.colors["red"])
        self.set_manager_status_card("Last Backup Time", backup_color if last_backup else self.colors["red"])
        self.set_manager_status_card("Backup Status", backup_color)
        self.set_manager_status_card("Auto Refresh", "#0d6efd")
        if hasattr(self, "manager_focus_vars"):
            self.manager_focus_vars["backup"].set("OK today" if self.today_backup_ok() else "Backup needed today")
            self.manager_focus_vars["late"].set(f"{c.get('late_returns', 0)} late return(s)")
            self.manager_focus_vars["alerts"].set(f"{c.get('manager_alerts', 0)} active alert(s)")
            self.manager_focus_vars["open"].set(f"{c.get('out', 0)} asset(s) currently out")
        if hasattr(self, "manager_focus_cards"):
            self.set_manager_focus_card("backup", self.colors["green"] if self.today_backup_ok() else self.colors["red"])
            self.set_manager_focus_card("late", self.colors["red"] if c.get("late_returns", 0) else self.colors["green"])
            self.set_manager_focus_card("alerts", "#b36b00" if c.get("manager_alerts", 0) else self.colors["green"])
            self.set_manager_focus_card("open", "#0d6efd" if c.get("out", 0) else self.colors["green"])
        open_rows = self.db.all("SELECT * FROM activity WHERE status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') ORDER BY due_back_at")
        issue_rows = self.db.all("SELECT * FROM assets WHERE status IN ('Checked Out','Missing','Repair','Retired') ORDER BY status, asset_type, name")
        problems_only = bool(getattr(self, "manager_show_problems_only", tk.BooleanVar(value=False)).get())
        if problems_only:
            def is_late_or_problem(row):
                try:
                    is_late = clean(row["due_back_at"]) and dt.datetime.fromisoformat(row["due_back_at"]) < dt.datetime.now()
                except Exception:
                    is_late = False
                return is_late or row["status"] in ("DUE SOON", "OVERDUE", "MISSING", "REVIEW")
            open_rows = [r for r in open_rows if is_late_or_problem(r)]
            issue_rows = [a for a in issue_rows if a["status"] in ("Missing", "Repair", "Retired")]
        if hasattr(self, "manager_open_tree"):
            self.fill(self.manager_open_tree, [
                (r["status"], r["asset_type"], f"{r['asset_barcode']} / {r['asset_name']}", r["employee_name"], pretty(r["due_back_at"]))
                for r in open_rows
            ])
        if hasattr(self, "manager_issue_tree"):
            self.fill(self.manager_issue_tree, [
                (a["status"], a["asset_type"], f"{a['barcode']} / {a['name']}", a["location"], a["current_holder_name"])
                for a in issue_rows
            ])
        if hasattr(self, "manager_health_var"):
            self.manager_health_var.set(" | ".join(f"{label}: {value}" for label, value in self.dashboard_health_lines()))
        if hasattr(self, "manager_notifications_tree"):
            clauses = []
            params = []
            if self.manager_notification_status_var.get() != "All":
                clauses.append("status=?")
                params.append(self.manager_notification_status_var.get())
            if self.manager_notification_severity_var.get() != "All":
                clauses.append("severity=?")
                params.append(self.manager_notification_severity_var.get())
            search = clean(self.manager_notification_search_var.get()).lower()
            sql = "SELECT * FROM manager_notifications"
            if clauses:
                sql += " WHERE " + " AND ".join(clauses)
            sql += " ORDER BY timestamp DESC LIMIT 200"
            notes = self.db.all(sql, tuple(params))
            if search:
                notes = [
                    n for n in notes
                    if search in " ".join(str(n[k] or "").lower() for k in n.keys())
                ]
            if problems_only:
                notes = [n for n in notes if n["status"] == "New" or n["severity"] in ("Warning", "Critical")]
            self.fill(self.manager_notifications_tree, [
                (n["id"], pretty(n["timestamp"]), n["severity"], n["event_type"], n["user_involved"], n["asset_involved"], n["status"], n["notes"])
                for n in notes
            ])
        prefix = "Manager problems view" if problems_only else "Manager view"
        self.status.set(f"{prefix} refreshed. Open: {len(open_rows)} | Issues: {len(issue_rows)}")

    def update_selected_notification(self, status):
        if not hasattr(self, "manager_notifications_tree"):
            return
        sel = self.manager_notifications_tree.selection()
        if not sel:
            messagebox.showinfo("Select Notification", "Select a manager notification first.")
            return
        note_id = self.manager_notifications_tree.item(sel[0], "values")[0]
        self.db.run("UPDATE manager_notifications SET status=?, handled_by=? WHERE id=?", (status, self.actor(), note_id))
        self.db.audit("MANAGER NOTIFICATION UPDATED", self.actor(), note_id, f"Status set to {status}")
        self.refresh_manager()

    def selected_manager_notification_id(self):
        if not hasattr(self, "manager_notifications_tree"):
            return None
        sel = self.manager_notifications_tree.selection()
        if not sel:
            return None
        return self.manager_notifications_tree.item(sel[0], "values")[0]

    def add_note_to_selected_notification(self):
        note_id = self.selected_manager_notification_id()
        if not note_id:
            messagebox.showinfo("Select Notification", "Select a manager notification first.")
            return
        note = simpledialog.askstring("Manager Note", "Add note to this notification:", parent=self)
        if not clean(note):
            return
        row = self.db.one("SELECT notes FROM manager_notifications WHERE id=?", (note_id,))
        existing = row["notes"] if row else ""
        combined = (existing + "\n" if existing else "") + f"{pretty(now_iso())} {self.actor()}: {clean(note)}"
        self.db.run("UPDATE manager_notifications SET notes=?, handled_by=? WHERE id=?", (combined, self.actor(), note_id))
        self.db.audit("MANAGER NOTIFICATION NOTE", self.actor(), note_id, clean(note))
        self.refresh_manager()

    def export_manager_notifications_excel(self):
        rows = self.db.all("SELECT * FROM manager_notifications ORDER BY timestamp DESC")
        headers = ["Date/Time","12-Hour Time","24-Hour Time","Event","Severity","User","Asset","Action","Handled By","Status","Notes","Reason","Computer"]
        data = [[r["timestamp"], r["time_12"], r["time_24"], r["event_type"], r["severity"], r["user_involved"], r["asset_involved"], r["action_taken"], r["handled_by"], r["status"], r["notes"], r["reason"], r["computer_name"]] for r in rows]
        path = self.folder_setting("excel_export_folder", "Exports") / f"manager_notifications_{stamp()}.xlsx"
        try:
            self.write_xlsx(path, [("Manager Notifications", headers, data)])
            self.db.audit("EXPORT MANAGER NOTIFICATIONS", self.actor(), "Manager", str(path))
            self.db.notify_manager("Log exported", "Info", self.actor(), "", "Manager notifications exported", self.actor(), str(path), status="Reviewed")
            messagebox.showinfo("Export Complete", f"Manager notifications exported:\n{path}")
        except Exception as e:
            self.db.error("Export Manager Notifications", str(e), traceback.format_exc())
            messagebox.showerror("Export Failed", f"Could not export manager notifications:\n{e}")

    def add_ap_alert_popup(self, target_type, target_key, target_label):
        if not self.require("manager"):
            return
        win = tk.Toplevel(self)
        win.title(f"Add AP Alert - {target_label}")
        win.geometry("720x520")
        self.apply_icon(win)
        pan = self.panel(win, "Add AP Alert", target_label)
        pan.pack(fill="both", expand=True, padx=14, pady=14)
        alert_type = tk.StringVar(value="Manager approval required")
        severity = tk.StringVar(value="Warning")
        required_action = tk.StringVar(value="Review before checkout")
        row = tk.Frame(pan.body, bg=self.colors["panel"])
        row.pack(fill="x", pady=(0,8))
        ttk.Label(row, text="Alert Type:", style="Panel.TLabel").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        ttk.Combobox(row, textvariable=alert_type, values=[
            "Security note", "Manager approval required", "Restricted checkout", "Damaged asset",
            "Missing pieces", "Prior late returns", "Inactive user", "Information only"
        ], state="readonly", width=28).grid(row=1, column=0, sticky="ew", padx=4)
        ttk.Label(row, text="Severity:", style="Panel.TLabel").grid(row=0, column=1, sticky="w", padx=4, pady=4)
        ttk.Combobox(row, textvariable=severity, values=["Info", "Warning", "Critical"], state="readonly", width=16).grid(row=1, column=1, sticky="ew", padx=4)
        ttk.Label(row, text="Required Action:", style="Panel.TLabel").grid(row=0, column=2, sticky="w", padx=4, pady=4)
        ttk.Combobox(row, textvariable=required_action, values=[
            "Information only", "Review before checkout", "Manager approval required", "Block checkout until resolved"
        ], width=28).grid(row=1, column=2, sticky="ew", padx=4)
        for col in range(3):
            row.grid_columnconfigure(col, weight=1)
        ttk.Label(pan.body, text="Alert Note:", style="Panel.TLabel").pack(anchor="w", pady=(8,2))
        note_text = tk.Text(pan.body, height=8, wrap="word", font=("Segoe UI", 10), bg="white", fg="#111")
        note_text.pack(fill="both", expand=True)
        self.tip(note_text, "Write the clear warning operators should see before checkout.")
        buttons = tk.Frame(pan.body, bg=self.colors["panel"])
        buttons.pack(fill="x", pady=(10,0))

        def save():
            note = clean(note_text.get("1.0", "end"))
            if not note:
                messagebox.showwarning("Missing Note", "Enter the alert note before saving.")
                return
            self.db.add_ap_alert(target_type, target_key, target_label, alert_type.get(), severity.get(), note, required_action.get(), self.actor())
            self.status.set("AP alert saved.")
            self.refresh()
            win.destroy()

        ttk.Button(buttons, text="Save Alert", style="Green.TButton", command=save).pack(side="left", padx=4)
        ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="left", padx=4)

    def ap_alerts_popup(self, target_type=None, target_key=None, target_label="All AP Alerts"):
        if not self.require("manager"):
            return
        title = "AP Alerts" if not target_type else f"AP Alerts - {target_label}"
        win = tk.Toplevel(self)
        win.title(title)
        win.geometry("1120x660")
        self.apply_icon(win)
        pan = self.panel(win, title, "Active, reviewed, and resolved AP alerts.")
        pan.pack(fill="both", expand=True, padx=14, pady=14)
        controls = tk.Frame(pan.body, bg=self.colors["panel"])
        controls.pack(fill="x", pady=(0,8))
        status_var = tk.StringVar(value="Active" if target_type else "All")
        ttk.Label(controls, text="Status:", style="Panel.TLabel").pack(side="left")
        ttk.Combobox(controls, textvariable=status_var, values=["All", "Active", "Reviewed", "Resolved"], state="readonly", width=14).pack(side="left", padx=4)
        ttk.Button(controls, text="Refresh", command=lambda: refresh()).pack(side="left", padx=3)
        ttk.Button(controls, text="Mark Reviewed", command=lambda: update_status("Reviewed")).pack(side="left", padx=3)
        ttk.Button(controls, text="Mark Resolved", command=lambda: update_status("Resolved")).pack(side="left", padx=3)
        ttk.Button(controls, text="Export", command=lambda: export_alerts()).pack(side="left", padx=3)
        if target_type and target_key:
            ttk.Button(controls, text="Add Alert", command=lambda: self.add_ap_alert_popup(target_type, target_key, target_label)).pack(side="left", padx=3)
        ttk.Button(controls, text="Close", command=win.destroy).pack(side="right", padx=3)
        tree = self.tree(pan.body, ("ID","Time","Target","Severity","Type","Status","Required Action","Note","Created By"), 16)
        current_rows = []

        def load_rows():
            sql = "SELECT * FROM ap_alerts"
            params = []
            clauses = []
            if target_type and target_key:
                clauses.append("target_type=?")
                params.append(target_type)
                clauses.append("lower(target_key)=lower(?)")
                params.append(target_key)
            if status_var.get() != "All":
                clauses.append("status=?")
                params.append(status_var.get())
            if clauses:
                sql += " WHERE " + " AND ".join(clauses)
            sql += " ORDER BY CASE severity WHEN 'Critical' THEN 0 WHEN 'Warning' THEN 1 ELSE 2 END, timestamp DESC"
            return self.db.all(sql, tuple(params))

        def refresh():
            nonlocal current_rows
            rows = load_rows()
            current_rows = [
                (r["id"], pretty(r["timestamp"]), f"{r['target_type']} {r['target_key']} | {r['target_label']}", r["severity"], r["alert_type"], r["status"], r["required_action"], r["note"], r["created_by"])
                for r in rows
            ]
            self.fill(tree, current_rows)

        def selected_id():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("Select Alert", "Select an AP alert first.")
                return None
            return tree.item(sel[0], "values")[0]

        def update_status(status):
            alert_id = selected_id()
            if not alert_id:
                return
            reason = ""
            if status == "Resolved":
                reason = simpledialog.askstring("Resolve AP Alert", "Reason/resolution note:", parent=win)
                if not clean(reason):
                    messagebox.showwarning("Reason Required", "A resolution reason is required.")
                    return
            self.db.update_ap_alert_status(alert_id, status, self.actor(), clean(reason))
            refresh()
            self.refresh()

        def export_alerts():
            if not current_rows:
                messagebox.showinfo("Export Alerts", "No AP alert rows to export.")
                return
            path = self.folder_setting("excel_export_folder", "Exports") / f"ap_alerts_{stamp()}.xlsx"
            try:
                self.write_xlsx(path, [("AP Alerts", ["ID","Time","Target","Severity","Type","Status","Required Action","Note","Created By"], current_rows)])
                self.db.audit("EXPORT AP ALERTS", self.actor(), "AP Alerts", f"Rows: {len(current_rows)}; File: {path}")
                self.db.notify_manager("Log exported", "Info", self.actor(), "", "AP alerts exported", self.actor(), str(path), status="Reviewed")
                messagebox.showinfo("Export Alerts", f"AP alerts exported:\n{path}")
            except Exception as e:
                self.db.error("Export AP Alerts", str(e), traceback.format_exc())
                messagebox.showerror("Export Alerts", f"Could not export AP alerts:\n{e}")

        self.bind_context_menu(tree, lambda t: [
            ("View Full Alert", lambda: self.show_full_tree_row(t, "AP Alert"), bool(self.selected_tree_values(t))),
            ("Mark Reviewed", lambda: update_status("Reviewed"), bool(self.selected_tree_values(t))),
            ("Mark Resolved", lambda: update_status("Resolved"), bool(self.selected_tree_values(t))),
            ("Export Visible Alerts", export_alerts),
            ("Copy Row", lambda: self.copy_tree_row(t), bool(self.selected_tree_values(t))),
        ])
        refresh()

    def audit_history_popup(self, title, pattern):
        win = tk.Toplevel(self)
        win.title(title)
        win.geometry("980x560")
        self.apply_icon(win)
        pan = self.panel(win, title, "Recent matching audit records.")
        pan.pack(fill="both", expand=True, padx=14, pady=14)
        tree = self.tree(pan.body, ("Time","Action","Actor","Target","Details"), 14)
        rows = self.db.all("SELECT * FROM audit WHERE action LIKE ? ORDER BY id DESC LIMIT 200", (pattern,))
        self.fill(tree, [(pretty(r["timestamp"]), r["action"], r["actor"], r["target"], r["details"]) for r in rows])
        self.bind_context_menu(tree, lambda t: [
            ("View Full Audit Entry", lambda: self.show_full_tree_row(t, "Audit Entry"), bool(self.selected_tree_values(t))),
            ("Copy Row", lambda: self.copy_tree_row(t), bool(self.selected_tree_values(t))),
        ])
        ttk.Button(pan.body, text="Close", command=win.destroy).pack(anchor="w", pady=8)

    def backup_history_popup(self):
        self.audit_history_popup("Backup History", "%BACKUP%")

    def export_history_popup(self):
        self.audit_history_popup("Export History", "%EXPORT%")

    def export_quick_scan_guide(self):
        folder = self.folder_setting("report_export_folder", "Reports")
        path = folder / f"Quick_Scan_Guide_{stamp()}.html"
        html = f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Macy's AP Quick Scan Guide</title>
<style>
body {{ font-family: Segoe UI, Arial, sans-serif; color:#111; margin:28px; }}
h1 {{ margin-bottom:4px; }}
h2 {{ margin-top:22px; border-bottom:2px solid #e21a2c; padding-bottom:4px; }}
.step {{ border:1px solid #ccc; padding:10px 12px; margin:8px 0; }}
.key {{ display:inline-block; min-width:64px; padding:3px 8px; margin-right:8px; background:#111; color:white; font-weight:bold; }}
.note {{ background:#fff3cd; padding:10px 12px; border:1px solid #ffe08a; }}
</style>
</head>
<body>
<h1>Macy's AP Quick Scan Guide</h1>
<p>Generated {pretty(now_iso())}</p>
<div class="note"><b>Scan-Only Mode:</b> If Admin turned this on in Settings, Front Desk does not need operator login. Checkout is employee badge, then item. Return is return-by badge and checked-out item.</div>
<h2>Checkout</h2>
<div class="step"><b>1.</b> Sign in as operator, unless Scan-Only Mode is on.</div>
<div class="step"><b>2.</b> Scan employee badge.</div>
<div class="step"><b>3.</b> Scan asset barcode, key number, serial, or device tag.</div>
<div class="step"><b>4.</b> Pick due time and condition.</div>
<div class="step"><b>5.</b> Click <b>Check Out Item</b> or press <b>F8</b>.</div>
<h2>Return</h2>
<div class="step"><b>1.</b> Press <b>F3</b> or click <b>Return Item</b> in Quick Scan.</div>
<div class="step"><b>2.</b> Scan the person returning it.</div>
<div class="step"><b>3.</b> Scan the checked-out item.</div>
<div class="step"><b>4.</b> Click <b>Return Item</b> or press <b>F9</b>.</div>
<h2>Keyboard Shortcuts</h2>
<p><span class="key">F2</span>New Checkout</p>
<p><span class="key">F3</span>Return Mode</p>
<p><span class="key">F4</span>Next Item Same Employee</p>
<p><span class="key">F8</span>Check Out Item</p>
<p><span class="key">F9</span>Return Item</p>
<p><span class="key">Esc</span>Start Over</p>
<div class="note">Use sample data only in screenshots and training. Do not print live employee, AP alert, or incident details.</div>
</body>
</html>"""
        path.write_text(html, encoding="utf-8")
        self.db.audit("EXPORT QUICK SCAN GUIDE", self.actor(), "Front Desk", str(path))
        messagebox.showinfo("Quick Scan Guide Saved", f"Printable guide saved:\n{path}")
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
        except Exception:
            pass

    def reset_training_demo_data(self):
        if not self.has("manager"):
            messagebox.showwarning("Manager Required", "Demo reset requires Manager/Admin access.")
            return
        if not self.is_training():
            if not messagebox.askyesno("Training Mode Required", "Demo reset only runs in Training Mode.\n\nTurn Training Mode on and create/reset demo records?"):
                return
            self.training_mode.set(True)
            self.update_operator_badges()
        if not messagebox.askyesno("Reset Demo Data", "This will create/reset DEMO sample users and assets only.\n\nIt will not delete live people or live assets. Continue?"):
            return
        try:
            self.backup_with_logs(show_message=False, label="TRAINING DEMO RESET BACKUP")
        except Exception:
            pass
        demo_people = [
            {"employee_id": "DEMO001", "badge": "880001", "first_name": "Demo", "last_name": "Associate", "department": "Asset Protection", "status": "Active", "shift": "Training", "role": "Employee", "notes": "[TRAINING / DEMO] Sample checkout user."},
            {"employee_id": "DEMO002", "badge": "880002", "first_name": "Demo", "last_name": "Manager", "department": "Asset Protection", "status": "Active", "shift": "Training", "role": "Manager", "notes": "[TRAINING / DEMO] Sample manager user."},
        ]
        for person in demo_people:
            existing = self.db.find_person(person["employee_id"])
            if existing:
                self.db.update_person(existing["id"], person, self.actor())
            else:
                self.db.add_person(person, self.actor())
        demo_assets = [
            {"asset_type": "Key", "barcode": "DEMO-KEY-001", "name": "Demo Key Ring 001", "status": "Available", "location": "Training", "controlled_key_number": "DEMO-K1", "serial_number": "", "asset_details": asset_details_to_json({"key_set_number": "DEMO-K1", "key_ring_serial_number": "DEMO-RING-001", "number_of_keys": "2", "key_ring_location": "Training", "key_ring_use": "Training checkout practice", "keys": [{"serial": "D1", "access": "Training Door"}, {"serial": "D2", "access": "Training Cabinet"}]}), "notes": "[TRAINING / DEMO] Sample key ring."},
            {"asset_type": "Radio", "barcode": "DEMO-RADIO-001", "name": "Demo Radio 001", "status": "Available", "location": "Training", "controlled_key_number": "", "serial_number": "DEMO-RAD-001", "asset_details": asset_details_to_json({"radio_number": "DEMO-RADIO-001", "factory_serial_number": "DEMO-RAD-001", "radio_serial_number": "DEMO-RAD-001", "radio_location": "Training", "assigned_area": "Training"}), "notes": "[TRAINING / DEMO] Sample radio."},
            {"asset_type": "Tablet", "barcode": "DEMO-TAB-001", "name": "Demo Tablet 001", "status": "Available", "location": "Training", "controlled_key_number": "", "serial_number": "DEMO-TAB-001", "asset_details": asset_details_to_json({"tablet_number": "DEMO-TAB-001", "tablet_serial_number": "DEMO-TAB-001", "imei_or_license": "DEMO-LICENSE", "tablet_location": "Training", "accessories": "Case, Charger"}), "notes": "[TRAINING / DEMO] Sample tablet."},
        ]
        demo_barcodes = [a["barcode"] for a in demo_assets]
        for asset in demo_assets:
            d = self.normalize_asset_record(asset)
            existing = self.db.find_asset(d["barcode"])
            if existing:
                self.db.update_asset(existing["id"], d, self.actor())
            else:
                self.db.add_asset(d, self.actor())
        placeholders = ",".join("?" for _ in demo_barcodes)
        self.db.run(f"DELETE FROM activity WHERE asset_barcode IN ({placeholders}) OR employee_id LIKE 'DEMO%'", tuple(demo_barcodes))
        self.db.run(f"UPDATE assets SET status='Available', current_holder_id='', current_holder_name='', updated_at=? WHERE barcode IN ({placeholders})", (now_iso(), *demo_barcodes))
        self.db.audit("TRAINING DEMO RESET", self.actor(), "Training", "Demo users/assets reset; live records preserved.")
        self.db.notify_manager("Training demo reset", "Info", self.actor(), "", "Demo sample records reset", self.actor(), "DEMO001, DEMO002, DEMO-KEY-001, DEMO-RADIO-001, DEMO-TAB-001", status="Reviewed")
        self.refresh()
        messagebox.showinfo("Demo Data Ready", "Training demo users and assets are ready.\n\nTry scanning badge 880001, then asset DEMO-KEY-001.")

    def set_people_filter(self, f):
        self.people_filter = f
        self.update_people_filter_chips()
        self.refresh_people()

    def update_people_filter_chips(self):
        if not hasattr(self, "people_filter_buttons"):
            return
        for label, button in self.people_filter_buttons.items():
            self.set_chip_active(button, label == getattr(self, "people_filter", "All"))

    def refresh_people(self):
        self.refresh_people_counts()
        self.update_people_filter_chips()
        rows = self.db.all("SELECT * FROM people ORDER BY last_name, first_name")
        out = []
        for r in rows:
            if self.people_filter == "All":
                ok = True
            elif self.people_filter == "Inactive":
                ok = r["status"] not in ("Active","Manual Review")
            else:
                ok = r["role"] == self.people_filter
            if ok:
                out.append((r["employee_id"], r["badge"], self.db.person_name(r), r["role"], r["status"], r["department"], r["shift"]))
        self.fill(self.people_tree, out)

    def set_asset_filter(self, f):
        self.current_asset_filter = f
        if hasattr(self, "asset_filter_var"):
            self.asset_filter_var.set(f)
        self.refresh_assets()

    def clear_asset_filters(self):
        self.asset_search_var.set("")
        self.asset_status_filter.set("All Status")
        self.set_asset_filter("All")

    def selected_asset_from_tree(self):
        if not hasattr(self, "asset_tree"):
            return None
        sel = self.asset_tree.selection()
        if not sel:
            return None
        barcode = self.asset_tree.item(sel[0], "values")[1]
        return self.db.find_asset(barcode)

    def update_asset_selection_summary(self):
        a = self.selected_asset_from_tree()
        if not a:
            self.asset_selected_var.set("Select an asset to view details and actions.")
            return
        status_note = ""
        open_log = self.db.open_log(a["barcode"])
        if open_log:
            status_note = f" | Out to {open_log['employee_name']} | Due {pretty(open_log['due_back_at'])}"
        self.asset_selected_var.set(
            f"Selected: {a['asset_type']} | {a['barcode']} | {a['name']} | {a['status']} | "
            f"Location: {a['location'] or 'None'} | Holder: {a['current_holder_name'] or 'None'}{status_note}"
        )

    def refresh_assets(self):
        self.refresh_asset_counts()
        sort_map = {"Type":"asset_type","Name":"name","Barcode":"barcode","Status":"status","Location":"location","Key Number":"controlled_key_number","Serial / Tag":"serial_number","Holder":"current_holder_name"}
        col = sort_map.get(self.asset_sort_by.get(), "asset_type")
        direction = "DESC" if self.asset_sort_dir.get() == "Z-A" else "ASC"
        rows = self.db.all(f"SELECT * FROM assets ORDER BY {col} {direction}, name ASC")
        out = []
        q = clean(self.asset_search_var.get()).lower()
        status_filter = self.asset_status_filter.get()
        for r in rows:
            if self.current_asset_filter != "All" and r["asset_type"] != self.current_asset_filter:
                continue
            if status_filter != "All Status" and r["status"] != status_filter:
                continue
            details_text = asset_detail_text(r["asset_type"], r["asset_details"] if "asset_details" in r.keys() else "")
            haystack = " ".join(str(r[k] or "") for k in ["asset_type","barcode","name","status","location","controlled_key_number","serial_number","notes","current_holder_name"]).lower()
            haystack = f"{haystack} {details_text.lower()}"
            if q and q not in haystack:
                continue
            out.append((r["asset_type"], r["barcode"], r["name"], r["status"], r["location"], r["controlled_key_number"], r["serial_number"], r["current_holder_name"]))
        self.fill(self.asset_tree, out)
        self.status.set(f"Assets shown: {len(out)}")
        self.update_asset_selection_summary()

    def set_selected_asset_status(self, status):
        if not self.require("assets_edit"):
            return
        a = self.selected_asset_from_tree()
        if not a:
            messagebox.showinfo("Select Asset", "Select an asset first.")
            return
        if self.db.open_log(a["barcode"]) and status != "Checked Out":
            messagebox.showwarning("Asset Checked Out", "Return this asset before changing its status.")
            return
        d = {k: a[k] for k in ASSET_DB_FIELDS if k in a.keys()}
        d["status"] = status
        self.db.update_asset(a["id"], d, self.actor())
        self.db.audit("ASSET STATUS UPDATE", self.actor(), a["barcode"], f"Status set to {status}")
        self.refresh_assets()

    def set_asset_status_by_id(self, asset_id, status, win=None):
        if not self.require("assets_edit"):
            return
        a = self.db.one("SELECT * FROM assets WHERE id=?", (asset_id,))
        if not a:
            messagebox.showwarning("Asset Not Found", "This asset could not be found.")
            return
        if self.db.open_log(a["barcode"]) and status != "Checked Out":
            messagebox.showwarning("Asset Checked Out", "Return this asset before changing its status.")
            return
        d = {k: a[k] for k in ASSET_DB_FIELDS if k in a.keys()}
        d["status"] = status
        self.db.update_asset(a["id"], d, self.actor())
        self.db.audit("ASSET STATUS UPDATE", self.actor(), a["barcode"], f"Status set to {status}")
        self.refresh_assets()
        if win:
            try:
                win.destroy()
            except Exception:
                pass

    def use_selected_asset_frontdesk(self):
        a = self.selected_asset_from_tree()
        if not a:
            messagebox.showinfo("Select Asset", "Select an asset first.")
            return
        self.use_asset_frontdesk(a)

    def duplicate_selected_asset(self):
        if not self.require("assets_edit"):
            return
        a = self.selected_asset_from_tree()
        if not a:
            messagebox.showinfo("Select Asset", "Select an asset to duplicate.")
            return
        prefill = {k: a[k] for k in ASSET_DB_FIELDS if k in a.keys()}
        prefill["barcode"] = ""
        prefill["name"] = f"Copy of {a['name']}"
        prefill["status"] = "Available"
        self.popup_asset(None, prefill=prefill)

    def export_selected_asset_excel(self):
        a = self.selected_asset_from_tree()
        if not a:
            messagebox.showinfo("Select Asset", "Select an asset to export.")
            return
        details = parse_asset_details(a["asset_details"] if "asset_details" in a.keys() else "")
        rows = [
            ["Type", a["asset_type"]],
            ["Barcode", a["barcode"]],
            ["Name", a["name"]],
            ["Status", a["status"]],
            ["Location", a["location"]],
            ["Holder", a["current_holder_name"]],
            ["Controlled Key Number", a["controlled_key_number"]],
            ["Serial / Tag", a["serial_number"]],
            ["Notes", a["notes"]],
            ["Details", asset_detail_text(a["asset_type"], details)],
        ]
        path = self.folder_setting("excel_export_folder", "Exports") / f"asset_{safe_filename_part(a['barcode'], 'asset')}_{stamp()}.xlsx"
        self.write_xlsx(path, [("Asset Profile", ["Field", "Value"], rows)])
        self.db.audit("EXPORT SELECTED ASSET", self.actor(), a["barcode"], str(path))
        messagebox.showinfo("Export Complete", f"Selected asset exported:\n{path}")

    def use_asset_frontdesk(self, a):
        self.selected_asset = a
        if hasattr(self, "sel_asset_var"):
            self.sel_asset_var.set(f"{a['asset_type']} | {a['barcode']} | {a['name']} | {a['status']}")
        if hasattr(self, "fd_message"):
            open_log = self.db.open_log(a["barcode"])
            self.fd_message.set("Asset is out. Scan return-by badge, then Return Item." if open_log else "Asset selected. Scan/select employee, then Check Out Item.")
        self.show("Front Desk")
        self.refresh_scan_mode()

    def scan(self):
        s = clean(self.scan_var.get())
        self.scan_var.set("")
        if not s:
            return
        scan_only = self.scan_only_frontdesk_enabled()
        p = self.db.find_person(s)
        if p:
            # If an out asset is already selected, this person is logged as the return-by person.
            if self.selected_asset and self.db.open_log(self.selected_asset["barcode"]):
                self.returning_person = p
                self.return_person_var.set(f"Return by: {self.db.person_name(p)} | {p['badge']}")
                self.fd_message.set("Return person selected. Click Return Item.")
                self.last_frontdesk_scan_kind = "person"
                self.refresh_scan_mode()
                self.frontdesk_feedback("return", "Return-by person selected. Click Return Item.")
            elif scan_only and getattr(self, "frontdesk_preferred_mode", "checkout") == "return":
                self.returning_person = p
                self.selected_person = None
                self.sel_person_var.set("No checkout employee selected.")
                self.return_person_var.set(f"Return by: {self.db.person_name(p)} | {p['badge']}")
                self.fd_message.set("Return-by person selected. Scan the checked-out item.")
                self.last_frontdesk_scan_kind = "person"
                self.refresh_scan_mode()
                self.frontdesk_feedback("return", "Return-by person selected. Scan the checked-out item.")
            else:
                self.frontdesk_preferred_mode = "checkout"
                self.selected_person = p
                self.sel_person_var.set(f"{self.db.person_name(p)} | {p['badge']} | {p['role']}")
                self.fd_message.set("Employee selected. Scan item next.")
                self.last_frontdesk_scan_kind = "person"
                self.refresh_scan_mode()
                self.frontdesk_feedback("success", "Employee selected. Scan the asset next.")
            return
        a = self.db.find_asset(s)
        if a:
            self.selected_asset = a
            self.sel_asset_var.set(f"{a['asset_type']} | {a['barcode']} | {a['name']} | {a['status']}")
            open_log = self.db.open_log(a["barcode"])
            self.frontdesk_preferred_mode = "return" if open_log else "checkout"
            if open_log and scan_only and self.selected_person and not self.returning_person and self.last_frontdesk_scan_kind == "person":
                self.returning_person = self.selected_person
                self.return_person_var.set(f"Return by: {self.db.person_name(self.returning_person)} | {self.returning_person['badge']}")
                self.selected_person = None
                self.sel_person_var.set("No checkout employee selected.")
                self.fd_message.set("Checked-out item selected. Return-by badge already scanned. Click Return Item.")
                feedback_message = "Checked-out item selected. Click Return Item."
            else:
                self.fd_message.set("Item is currently OUT. Scan returning person badge, then click Return Item." if open_log else "Asset selected. Click Check Out Item.")
                feedback_message = "Asset is out. Scan the return-by person." if open_log else "Asset selected. Ready for checkout when employee and due time are ready."
            self.last_frontdesk_scan_kind = "asset"
            self.refresh_scan_mode()
            self.frontdesk_feedback("return" if open_log else "success", feedback_message)
            return
        self.frontdesk_feedback("blocked", f"No employee or asset matched: {s}")
        self.unknown_person_prompt(s)

    def checkout(self):
        if not self.require("frontdesk"):
            return
        if not self.selected_person:
            messagebox.showwarning("Missing Employee", "Scan/select employee first.")
            self.frontdesk_feedback("blocked", "Missing employee. Scan/select employee first.")
            return
        if not self.selected_asset:
            messagebox.showwarning("Missing Asset", "Scan/select asset first.")
            self.frontdesk_feedback("blocked", "Missing asset. Scan/select asset first.")
            return
        if self.db.open_log(self.selected_asset["barcode"]):
            messagebox.showwarning("Already Out", "This asset is already checked out.")
            self.frontdesk_feedback("blocked", "Asset is already out. Use Return Mode.")
            self.db.notify_manager("Blocked checkout attempt", "Warning", self.db.person_name(self.selected_person), self.selected_asset["barcode"], "Asset already checked out", self.actor(), "Checkout blocked before save.")
            return
        alert_lines = []
        hard_block = False
        saved_alerts = list(self.db.active_ap_alerts("Person", self.selected_person["employee_id"])) + list(self.db.active_ap_alerts("Asset", self.selected_asset["barcode"]))
        for alert in saved_alerts:
            required_action = clean(alert["required_action"])
            if alert["severity"] == "Critical" or "block" in required_action.lower():
                hard_block = True
            alert_lines.append(
                f"{alert['target_type']} AP Alert [{alert['severity']}] {alert['alert_type']}: {alert['note']} "
                f"(Created by {alert['created_by']} at {pretty(alert['timestamp'])}; Required: {required_action or 'Review'})"
            )
        person_notes = clean(self.selected_person["notes"]).lower()
        asset_notes = clean(self.selected_asset["notes"]).lower()
        if self.selected_person["status"] != "Active":
            hard_block = True
            alert_lines.append(f"Employee status is {self.selected_person['status']}.")
        if any(token in person_notes for token in ("ap alert", "restricted", "manager approval", "no keys", "blocked")):
            alert_lines.append(f"Employee alert note: {self.selected_person['notes']}")
            if any(token in person_notes for token in ("restricted", "manager approval", "no keys", "blocked")):
                hard_block = True
        if any(token in asset_notes for token in ("ap alert", "restricted", "manager approval", "damaged", "missing pieces", "blocked")):
            alert_lines.append(f"Asset alert note: {self.selected_asset['notes']}")
            if any(token in asset_notes for token in ("restricted", "manager approval", "blocked")):
                hard_block = True
        if self.selected_asset["asset_type"] == "Key" and "no keys" in person_notes:
            hard_block = True
            alert_lines.append("Employee note indicates keys are not allowed.")
        overdue = self.db.all("SELECT * FROM activity WHERE employee_id=? AND status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') AND due_back_at < ?", (self.selected_person["employee_id"], now_iso()))
        if overdue:
            alert_lines.append(f"Employee has {len(overdue)} overdue item(s).")
        if alert_lines:
            detail = "\n".join(alert_lines)
            self.db.audit("AP ALERT SHOWN", self.actor(), self.selected_asset["barcode"], detail)
            self.db.notify_manager("Checkout warning shown", "Warning", self.db.person_name(self.selected_person), self.selected_asset["barcode"], "Operator warning before checkout", self.actor(), detail)
            if hard_block and not self.has("manager"):
                messagebox.showwarning("Checkout Blocked", detail + "\n\nA Manager or Admin must review this before checkout can continue.")
                self.frontdesk_feedback("blocked", "Checkout blocked. Manager/Admin review required.")
                self.db.notify_manager("Blocked checkout attempt", "Critical", self.db.person_name(self.selected_person), self.selected_asset["barcode"], "AP alert blocked checkout", self.actor(), detail)
                self.db.audit("CHECKOUT BLOCKED", self.actor(), self.selected_asset["barcode"], detail)
                return
            if not messagebox.askyesno("AP Alert / Checkout Warning", detail + "\n\nContinue checkout?"):
                self.db.notify_manager("Blocked checkout attempt", "Critical", self.db.person_name(self.selected_person), self.selected_asset["barcode"], "Operator cancelled after warning", self.actor(), detail)
                self.db.audit("CHECKOUT BLOCKED", self.actor(), self.selected_asset["barcode"], detail)
                return
            if hard_block:
                self.db.notify_manager("Checkout override", "Critical", self.db.person_name(self.selected_person), self.selected_asset["barcode"], "Manager/Admin continued after AP alert", self.actor(), detail)
        if self.selected_asset["status"] in ("Missing", "Repair", "Retired"):
            status_detail = f"This asset is marked {self.selected_asset['status']} and needs Manager/Admin approval before checkout."
            if not self.has("manager"):
                messagebox.showwarning("Checkout Blocked", status_detail)
                self.frontdesk_feedback("blocked", "Checkout blocked by asset status.")
                self.db.notify_manager("Blocked checkout attempt", "Critical", self.db.person_name(self.selected_person), self.selected_asset["barcode"], f"Asset status {self.selected_asset['status']} requires manager", self.actor(), status_detail)
                self.db.audit("CHECKOUT BLOCKED", self.actor(), self.selected_asset["barcode"], status_detail)
                return
            if not messagebox.askyesno("Asset Status Warning", f"This asset is marked {self.selected_asset['status']}.\n\nContinue checkout anyway?"):
                self.db.notify_manager("Blocked checkout attempt", "Warning", self.db.person_name(self.selected_person), self.selected_asset["barcode"], f"Asset status {self.selected_asset['status']} blocked checkout", self.actor())
                return
            self.db.notify_manager("Checkout override", "Warning", self.db.person_name(self.selected_person), self.selected_asset["barcode"], f"Asset status {self.selected_asset['status']} override", self.actor())
        try:
            dt.datetime.fromisoformat(self.due_var.get())
        except Exception:
            messagebox.showwarning("Invalid Due Time", "Due time must use a valid date/time. Use one of the quick due buttons if unsure.")
            self.frontdesk_feedback("blocked", "Invalid due time. Use a due preset or enter a valid date/time.")
            return
        checkout_notes = self.notes_var.get()
        if self.scan_only_frontdesk_enabled() and (self.operator.get() or "") == "No operator signed in":
            checkout_notes = (checkout_notes + " | " if checkout_notes else "") + "Scan-only Front Desk checkout"
        log_id = self.db.checkout(self.selected_person, self.selected_asset, self.due_var.get(), self.actor(), self.condition_var.get(), self.mark_training_notes(checkout_notes))
        if not log_id:
            messagebox.showwarning("Already Checked Out", "Another computer checked this asset out before this save completed. The screen will refresh now.")
            self.frontdesk_feedback("blocked", "Another computer checked this asset out first.")
            self.refresh()
            return
        messagebox.showinfo("Checked Out", f"{self.selected_asset['name']} checked out to {self.db.person_name(self.selected_person)}.")
        self.selected_asset = None
        self.sel_asset_var.set("No asset selected. Scan next item if this employee needs more than one.")
        self.fd_message.set("Checkout complete. Scan another item for the same employee or scan a different employee.")
        self.refresh_scan_mode()
        self.frontdesk_feedback("success", "Checkout complete. Scan another item or a different employee.")
        self.refresh_current()

    def return_scanned(self):
        if not self.require("frontdesk"):
            return
        if not self.selected_asset:
            messagebox.showwarning("Missing Asset", "Scan/select item to return first.")
            self.frontdesk_feedback("blocked", "Missing asset. Scan/select the item being returned.")
            return
        log = self.db.open_log(self.selected_asset["barcode"])
        if not log:
            messagebox.showwarning("Not Out", "This item is not currently out.")
            self.frontdesk_feedback("blocked", "This asset is not currently checked out.")
            return
        return_by = self.returning_person
        if return_by is None:
            if self.scan_only_frontdesk_enabled():
                messagebox.showwarning("Return Badge Required", "Scan the badge of the person returning this item before returning it.")
                self.frontdesk_feedback("blocked", "Scan the return-by badge before returning this item.")
                return
            if not messagebox.askyesno("Return Person Not Scanned", "No returning person badge was scanned. Continue using the signed-in operator as the return-by person?"):
                return
        return_by_name = self.db.person_name(return_by) if return_by else self.actor()
        notes = self.mark_training_notes(self.notes_var.get())
        if self.scan_only_frontdesk_enabled() and (self.operator.get() or "") == "No operator signed in":
            notes = (notes + " | " if notes else "") + "Scan-only Front Desk return"
        if return_by:
            notes = (notes + " | " if notes else "") + f"Returned by badge: {return_by['badge']} / {return_by_name}"
            if clean(return_by["employee_id"]).lower() != clean(log["employee_id"]).lower():
                detail = (
                    f"Returned by different badge: {return_by_name} / {return_by['badge']} | "
                    f"Checked out to: {log['employee_name']} / {log['employee_badge']}"
                )
                notes = (notes + " | " if notes else "") + detail
                self.db.audit("RETURN BY DIFFERENT PERSON", self.actor(), self.selected_asset["barcode"], detail)
        self.db.return_asset(log, self.actor(), self.condition_var.get(), notes)
        self.db.audit("RETURN BY PERSON", self.actor(), self.selected_asset["barcode"], f"Returned by: {return_by_name} | Original checkout holder: {log['employee_name']}")
        messagebox.showinfo("Returned", f"{self.selected_asset['name']} returned.\n\nReturned by: {return_by_name}")
        self.frontdesk_preferred_mode = "checkout"
        self.clear_frontdesk()
        self.frontdesk_feedback("success", "Return complete. Ready for the next scan.")
        self.refresh_scan_mode()
        self.refresh_current()

    def clear_frontdesk(self):
        self.selected_person = None
        self.selected_asset = None
        self.returning_person = None
        self.frontdesk_preferred_mode = "checkout"
        self.last_frontdesk_scan_kind = ""
        self.sel_person_var.set("No employee selected.")
        self.sel_asset_var.set("No asset selected.")
        self.return_person_var.set("Return by: scan badge required for scan-only returns." if self.scan_only_frontdesk_enabled() else "Return by: signed-in operator unless another badge is scanned.")
        self.fd_message.set("Scan-only ready. Scan employee badge for checkout, or return-by badge and item for return." if self.scan_only_frontdesk_enabled() else "Ready for employee scan.")
        self.refresh_scan_mode()
        self.scan_entry.focus_set()

    def person_form(self, person=None):
        if not self.require("people_edit"):
            return
        if person is None:
            sel = self.people_tree.selection()
            if sel:
                emp_id = self.people_tree.item(sel[0],"values")[0]
                person = self.db.one("SELECT * FROM people WHERE employee_id=?", (emp_id,))
        self.popup_person(person)

    def popup_person(self, person=None):
        key = "person:new" if not person else f"person:{person['id']}"
        if self.focus_existing(key):
            return
        win = tk.Toplevel(self)
        win.title("Add/Edit Person")
        win.geometry("900x700")
        self.apply_icon(win)
        self.track_window(key, win)
        p = self.panel(win, "Add/Edit Person", "Create or update a user. Role changes are limited by the signed-in operator.")
        p.pack(fill="both", expand=True, padx=14, pady=14)
        fields = ["employee_id","badge","first_name","last_name","department","status","shift","role","notes"]
        vars = {k: tk.StringVar(value=(person[k] if person and k in person.keys() and person[k] is not None else "")) for k in fields}
        if not person and hasattr(self, "popup_person_prefill_value"):
            scan_value = clean(getattr(self, "popup_person_prefill_value", ""))
            if scan_value.upper().startswith("F"):
                vars["employee_id"].set(scan_value.upper())
                vars["badge"].set("")
            elif scan_value.startswith("88"):
                vars["badge"].set(scan_value)
                vars["employee_id"].set("")
            else:
                vars["badge"].set(scan_value)
                vars["employee_id"].set("")
            vars["notes"].set("")
            try:
                delattr(self, "popup_person_prefill_value")
            except Exception:
                pass
        if not vars["status"].get(): vars["status"].set("Active")
        field_widgets = {}
        if not vars["role"].get(): vars["role"].set("Employee")
        form = tk.Frame(p.body, bg=self.colors["panel"])
        form.pack(fill="both", expand=True)
        def field(label,k,r,c,values=None):
            b = tk.Frame(form,bg=self.colors["panel"]); b.grid(row=r,column=c,sticky="ew",padx=6,pady=6); form.grid_columnconfigure(c, weight=1)
            ttk.Label(b,text=label,style="Panel.TLabel").pack(anchor="w")
            e = ttk.Combobox(b,textvariable=vars[k],values=values,state="readonly") if values else ttk.Entry(b,textvariable=vars[k])
            e.pack(fill="x")
            field_widgets[k] = e
            vars[k].trace_add("write", lambda *args: self.validate_person_form_live(vars, field_widgets))
            self.tip(e, f"Enter {label}. Required fields highlight yellow until valid.")
        roles = self.db.group_names() if self.role()=="Admin" else [r for r in self.db.group_names() if ROLE_RANK.get(r, 1) <= ROLE_RANK.get(self.role(), 1) and r!="Admin"]
        field("Employee ID / F-number","employee_id",0,0); field("Badge / 88-number","badge",0,1); field("First","first_name",1,0); field("Last","last_name",1,1)
        field("Department","department",2,0); field("Status","status",2,1,PEOPLE_STATUS); field("Shift","shift",3,0); field("Role","role",3,1,roles); field("Notes","notes",4,0)
        self.after(100, lambda: self.validate_person_form_live(vars, field_widgets))
        row=tk.Frame(p.body,bg=self.colors["panel"]); row.pack(fill="x",pady=10)
        def save():
            d={k:vars[k].get().strip() for k in fields}
            d = self.normalize_person_id_badge(d, ask=True)
            if not d["employee_id"] or not d["badge"] or not d["first_name"] or not d["last_name"]:
                messagebox.showwarning("Missing Info","ID, badge, first, and last are required."); return
            if not self.validate_person_form_live(vars, field_widgets):
                messagebox.showwarning("Missing / Invalid Fields", "Highlighted fields need attention before saving.")
                return
            if not self.validate_person_id_badge(d):
                return
            old_role = person["role"] if person else "Front Desk"
            new_role = d["role"]
            if not self.can_assign_role(old_role, new_role, person):
                return
            try:
                if person: self.db.update_person(person["id"], d, self.actor())
                else: self.db.add_person(d, self.actor())
                win.destroy(); self.refresh_people()
            except sqlite3.IntegrityError as e:
                messagebox.showerror("Duplicate", str(e))
        win.bind("<Control-Return>", lambda e: save())
        self.tip(ttk.Button(row,text="Save",style="Green.TButton",command=save), "Save this person. You can also press Ctrl+Enter.").pack(side="left",padx=4)
        self.tip(ttk.Button(row,text="Cancel",command=win.destroy), "Close without saving.").pack(side="left",padx=4)

    def can_assign_role(self, old_role, new_role, person=None):
        if self.role() != "Admin" and (old_role == "Admin" or new_role == "Admin"):
            messagebox.showwarning("Admin Protected", "Only Admin can create/change/modify Admin users.")
            self.db.error("Role Blocked", "Non-admin attempted Admin role change", f"{self.actor()} old={old_role} new={new_role}")
            return False
        if ROLE_RANK[new_role] > ROLE_RANK[self.role()]:
            messagebox.showwarning("Role Not Allowed", "You cannot assign a role higher than your own.")
            return False
        if person and self.db.person_name(person) == self.actor() and ROLE_RANK[new_role] > ROLE_RANK[old_role]:
            messagebox.showwarning("Self-Promotion Blocked", "You cannot promote yourself.")
            return False
        return True

    def open_selected_person(self):
        sel = self.people_tree.selection()
        if not sel: return
        emp_id = self.people_tree.item(sel[0],"values")[0]
        p = self.db.one("SELECT * FROM people WHERE employee_id=?", (emp_id,))
        if p: self.person_profile(p)

    def person_profile(self, p):
        key=f"person_profile:{p['id']}"
        if self.focus_existing(key): return
        win=tk.Toplevel(self); win.title(f"Person Profile - {self.db.person_name(p)}"); win.geometry("1000x700"); self.apply_icon(win); self.track_window(key,win)
        pan=self.panel(win, f"Person Profile - {self.db.person_name(p)}", "Current items and recent history.")
        pan.pack(fill="both",expand=True,padx=14,pady=14)
        info=tk.Label(pan.body,text=f"ID: {p['employee_id']} | Badge: {p['badge']} | Role: {p['role']} | Status: {p['status']} | Dept: {p['department']}",bg=self.colors["panel"],font=("Segoe UI",11,"bold"))
        info.pack(anchor="w", pady=(0,8))
        actions=tk.Frame(pan.body,bg=self.colors["panel"])
        actions.pack(fill="x",pady=(0,8))
        person_label=f"{self.db.person_name(p)} | {p['employee_id']} | {p['badge']}"
        ttk.Button(actions,text="Edit Person",command=lambda:self.popup_person(p)).pack(side="left",padx=4)
        ttk.Button(actions,text="Add AP Alert",command=lambda:self.add_ap_alert_popup("Person",p["employee_id"],person_label)).pack(side="left",padx=4)
        ttk.Button(actions,text="View AP Alerts",command=lambda:self.ap_alerts_popup("Person",p["employee_id"],person_label)).pack(side="left",padx=4)
        t=self.tree(pan.body,("Status","Asset","Type","Out","Due","Returned"),14)
        rows=self.db.all("SELECT * FROM activity WHERE employee_id=? ORDER BY checked_out_at DESC LIMIT 200",(p["employee_id"],))
        self.fill(t,[(r["status"],f"{r['asset_barcode']} / {r['asset_name']}",r["asset_type"],pretty(r["checked_out_at"]),pretty(r["due_back_at"]),pretty(r["returned_at"])) for r in rows])

    def asset_form(self, asset=None):
        if not self.require("assets_edit"):
            return
        if asset is None:
            sel = self.asset_tree.selection()
            if sel:
                barcode = self.asset_tree.item(sel[0],"values")[1]
                asset = self.db.find_asset(barcode)
            else:
                messagebox.showinfo("Select Asset", "Select an asset to edit, or use Add Asset.")
                return
        self.popup_asset(asset)

    def quick_asset_form(self):
        if not self.require("assets_edit"):
            return
        self.popup_asset(None)

    def popup_asset(self, asset=None, prefill=None):
        key="asset:new" if not asset else f"asset_edit:{asset['barcode']}"
        if self.focus_existing(key): return
        win=tk.Toplevel(self); win.title("Add/Edit Asset"); win.geometry("1060x820"); self.apply_icon(win); self.track_window(key,win)
        pan=self.panel(win,"Add/Edit Asset","Create one asset or use Save + New for bulk entry.")
        pan.pack(fill="both",expand=True,padx=14,pady=14)
        fields=ASSET_FIELDS
        source_asset = prefill or asset
        vars={k:tk.StringVar(value=(source_asset[k] if source_asset and k in source_asset.keys() and source_asset[k] is not None else "")) for k in fields}
        if not vars["asset_type"].get(): vars["asset_type"].set("Item")
        if not vars["status"].get(): vars["status"].set("Available")
        existing_details = parse_asset_details(source_asset["asset_details"] if source_asset and "asset_details" in source_asset.keys() else "")
        detail_vars = {}
        detail_widgets = {}
        accessory_vars = {}
        key_row_frame = None
        field_widgets = {}
        guidance_var = tk.StringVar()
        saved_recent = []
        saved_recent_var = tk.StringVar(value="Recently saved: none yet")
        form=tk.Frame(pan.body,bg=self.colors["panel"]); form.pack(fill="x")
        def field(label,k,r,c,values=None):
            b=tk.Frame(form,bg=self.colors["panel"]); b.grid(row=r,column=c,sticky="ew",padx=6,pady=6); form.grid_columnconfigure(c,weight=1)
            ttk.Label(b,text=label,style="Panel.TLabel").pack(anchor="w")
            e=ttk.Combobox(b,textvariable=vars[k],values=values,state="readonly") if values else ttk.Entry(b,textvariable=vars[k])
            e.pack(fill="x")
            field_widgets[k] = e
            vars[k].trace_add("write", lambda *args: self.validate_asset_form_live(vars, field_widgets))
            self.tip(e, f"Enter {label}. Required/recommended fields highlight yellow until complete.")
        field("Type","asset_type",0,0,ASSET_TYPES); field("Barcode / Scan ID","barcode",0,1); field("Asset Name","name",1,0); field("Status","status",1,1,ASSET_STATUS)
        field("Location","location",2,0); field("Notes","notes",2,1)

        guide=tk.Label(pan.body,textvariable=guidance_var,bg=self.colors["panel"],fg="#555555",anchor="w",justify="left",wraplength=860)
        guide.pack(fill="x",pady=(2,8))
        detail_panel = self.panel(pan.body, "Asset Type Details", "Only fields for the selected asset type are shown.")
        detail_panel.pack(fill="both", expand=True, pady=(0,8))

        def detail_source():
            source = {k: vars[k].get() for k in fields}
            source["accessories"] = [name for name, v in accessory_vars.items() if v.get()]
            for k, v in detail_vars.items():
                source[k] = v.get()
            return source

        def add_detail_field(parent, spec, row, col):
            box = tk.Frame(parent, bg=self.colors["panel"])
            box.grid(row=row, column=col, sticky="ew", padx=6, pady=5)
            parent.grid_columnconfigure(col, weight=1)
            label_text = spec["label"] + (" *" if spec.get("required") else "")
            ttk.Label(box, text=label_text, style="Panel.TLabel").pack(anchor="w")
            var = tk.StringVar(value=detail_vars.get(spec["key"], tk.StringVar(value=existing_details.get(spec["key"], spec.get("default", "")))).get())
            detail_vars[spec["key"]] = var
            entry = ttk.Entry(box, textvariable=var)
            entry.pack(fill="x")
            detail_widgets[spec["key"]] = entry
            if spec.get("type") == "int":
                var.trace_add("write", lambda *args, v=var: v.set("".join(ch for ch in v.get() if ch.isdigit())) if any(not ch.isdigit() for ch in v.get()) else None)
            self.tip(entry, f"Enter {spec['label']}.")
            return entry

        def rebuild_key_rows(*args):
            nonlocal key_row_frame
            if key_row_frame is None:
                return
            for child in key_row_frame.winfo_children():
                child.destroy()
            count_text = detail_vars.get("number_of_keys", tk.StringVar(value="1")).get()
            count = max(1, min(50, int(count_text) if count_text.isdigit() else 1))
            existing_keys = existing_details.get("keys", [])
            ttk.Label(key_row_frame, text="Keys on this ring:", style="Panel.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", padx=6, pady=(6,2))
            for idx in range(count):
                previous = existing_keys[idx] if idx < len(existing_keys) and isinstance(existing_keys[idx], dict) else {}
                s_key = f"key_{idx+1}_serial"
                a_key = f"key_{idx+1}_access"
                if s_key not in detail_vars:
                    detail_vars[s_key] = tk.StringVar(value=previous.get("serial", ""))
                if a_key not in detail_vars:
                    detail_vars[a_key] = tk.StringVar(value=previous.get("access", ""))
                add_detail_field(key_row_frame, {"key": s_key, "label": f"Key {idx+1} serial/key number", "required": True}, idx+1, 0)
                add_detail_field(key_row_frame, {"key": a_key, "label": f"Key {idx+1} opens/access description", "required": True}, idx+1, 1)

        def rebuild_detail_form(*args):
            nonlocal key_row_frame
            asset_type = vars["asset_type"].get() if vars["asset_type"].get() in ASSET_TYPES else "Item"
            schema = asset_detail_schema(asset_type)
            for child in detail_panel.body.winfo_children():
                child.destroy()
            detail_widgets.clear()
            accessory_vars.clear()
            tk.Label(detail_panel.body, text=schema["title"], bg=self.colors["panel"], fg="#111111", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=6, pady=(0,4))
            grid = tk.Frame(detail_panel.body, bg=self.colors["panel"])
            grid.pack(fill="x")
            for i, spec in enumerate(schema["fields"]):
                add_detail_field(grid, spec, i // 2, i % 2)
            if asset_type == "Tablet":
                accessory_box = tk.Frame(detail_panel.body, bg=self.colors["panel"])
                accessory_box.pack(fill="x", padx=6, pady=(8,2))
                ttk.Label(accessory_box, text="Accessories:", style="Panel.TLabel").grid(row=0, column=0, sticky="w")
                selected = existing_details.get("accessories", [])
                for idx, name in enumerate(TABLET_ACCESSORIES, start=1):
                    var = tk.BooleanVar(value=name in selected)
                    accessory_vars[name] = var
                    cb = ttk.Checkbutton(accessory_box, text=name, variable=var)
                    cb.grid(row=idx // 4, column=idx % 4, sticky="w", padx=8, pady=3)
                    self.tip(cb, f"Mark if this tablet includes: {name}.")
            if asset_type == "Key":
                key_row_frame = tk.Frame(detail_panel.body, bg=self.colors["panel"])
                key_row_frame.pack(fill="x", pady=(8,0))
                if "number_of_keys" in detail_vars:
                    detail_vars["number_of_keys"].trace_add("write", rebuild_key_rows)
                rebuild_key_rows()
            else:
                key_row_frame = None
            refresh_guidance()

        def refresh_guidance(*args):
            d = self.normalize_asset_record(detail_source())
            guidance = ASSET_TYPE_GUIDANCE.get(d["asset_type"], ASSET_TYPE_GUIDANCE["Item"])
            rule = asset_type_rule(d["asset_type"])
            recommended = ", ".join(label for key, label in [
                ("controlled_key_number", "controlled key number"),
                ("serial_number", "serial number / device tag"),
            ] if key in rule["recommended"])
            parts = [guidance["notes"], rule["field_notes"]]
            if recommended:
                parts.append(f"Recommended: {recommended}.")
            duplicate = self.asset_duplicate_message(d, asset["id"] if asset else None)
            if duplicate:
                parts.append(duplicate.replace("\n\n", " "))
            guidance_var.set(" ".join(parts))
            self.validate_asset_form_live(vars, field_widgets)

        for v in vars.values():
            v.trace_add("write", refresh_guidance)
        vars["asset_type"].trace_add("write", rebuild_detail_form)

        presets=tk.Frame(pan.body,bg=self.colors["panel"]); presets.pack(fill="x",pady=(0,8))
        ttk.Label(presets,text="Presets:",style="Panel.TLabel").pack(side="left",padx=(4,8))
        def use_preset(asset_type):
            vars["asset_type"].set(asset_type)
            if not vars["location"].get():
                vars["location"].set("AP" if asset_type in ("Key", "Radio") else "Front Desk")
            if not vars["status"].get():
                vars["status"].set("Available")
            field_widgets["barcode"].focus_set()
        for asset_type in ASSET_TYPES:
            self.tip(ttk.Button(presets,text=asset_type,command=lambda t=asset_type: use_preset(t)), f"Set this form up for a {asset_type}.").pack(side="left",padx=2)

        def autofill_name():
            d = self.normalize_asset_record(detail_source())
            vars["barcode"].set(d["barcode"])
            vars["name"].set(d["name"])

        quick=tk.Frame(pan.body,bg=self.colors["panel"]); quick.pack(fill="x",pady=(0,8))
        self.tip(ttk.Button(quick,text="Auto-Fill Name",command=autofill_name), "Build a clean name from the asset type and barcode.").pack(side="left",padx=4)
        self.tip(ttk.Button(quick,text="Clear For New",command=lambda: clear_form()), "Clear the window so you can add the next asset.").pack(side="left",padx=4)
        tk.Label(pan.body,textvariable=saved_recent_var,bg=self.colors["panel"],fg="#198754",anchor="w",font=("Segoe UI",10,"bold")).pack(fill="x",pady=(0,4))
        self.after(100, lambda: (rebuild_detail_form(), refresh_guidance(), self.validate_asset_form_live(vars, field_widgets)))
        row=tk.Frame(pan.body,bg=self.colors["panel"]); row.pack(fill="x",pady=10)
        def clear_form():
            for k in fields:
                vars[k].set("")
            vars["asset_type"].set("Item")
            vars["status"].set("Available")
            existing_details.clear()
            detail_vars.clear()
            rebuild_detail_form()
            field_widgets["barcode"].focus_set()

        def save(close_after=True):
            d=self.normalize_asset_record(detail_source())
            for k in fields:
                vars[k].set(d[k])
            problems = self.validate_asset_record(d, asset["id"] if asset else None)
            if problems:
                messagebox.showwarning("Asset Needs Attention", "\n\n".join(problems)); return False
            if not self.validate_asset_form_live(vars, field_widgets):
                messagebox.showwarning("Missing / Recommended Fields", "Highlighted fields need attention before saving."); return False
            try:
                if asset: self.db.update_asset(asset["id"], d, self.actor())
                else: self.db.add_asset(d, self.actor())
                self.refresh_assets()
                if not asset:
                    saved_recent.insert(0, f"{d['barcode']} - {d['name']}")
                    del saved_recent[5:]
                    saved_recent_var.set("Recently saved: " + " | ".join(saved_recent))
                if close_after or asset:
                    win.destroy()
                else:
                    clear_form()
                    self.status.set("Asset saved. Ready for the next asset.")
                return True
            except sqlite3.IntegrityError as e:
                messagebox.showerror("Duplicate", str(e))
                return False
        rebuild_detail_form()
        refresh_guidance()
        win.bind("<Control-Return>", lambda e: save())
        self.tip(ttk.Button(row,text="Save",style="Green.TButton",command=save), "Save this asset. You can also press Ctrl+Enter.").pack(side="left",padx=4)
        if not asset:
            self.tip(ttk.Button(row,text="Save + New",style="Green.TButton",command=lambda: save(False)), "Save this asset and keep the window ready for the next one.").pack(side="left",padx=4)
        ttk.Button(row,text="Cancel",command=win.destroy).pack(side="left",padx=4)

    def open_selected_asset(self):
        sel=self.asset_tree.selection()
        if not sel: return
        barcode=self.asset_tree.item(sel[0],"values")[1]
        a=self.db.find_asset(barcode)
        if a: self.asset_profile(a)

    def asset_profile(self, a):
        key=f"asset_profile:{a['barcode']}"
        if self.focus_existing(key): return
        win=tk.Toplevel(self); win.title(f"Asset Profile - {a['barcode']}"); win.geometry("1050x720"); self.apply_icon(win); self.track_window(key,win)
        pan=self.panel(win,f"Asset Profile - {a['barcode']}","View asset status and history.")
        pan.pack(fill="both",expand=True,padx=14,pady=14)
        open_log=self.db.open_log(a["barcode"])
        info=f"Type: {a['asset_type']} | Name: {a['name']} | Status: {a['status']} | Location: {a['location']} | Holder: {a['current_holder_name'] or 'None'} | Key#: {a['controlled_key_number']} | Serial: {a['serial_number']}"
        tk.Label(pan.body,text=info,bg=self.colors["panel"],font=("Segoe UI",11,"bold"),wraplength=950,justify="left").pack(anchor="w",pady=(0,8))
        detail_line = asset_detail_text(a["asset_type"], a["asset_details"] if "asset_details" in a.keys() else "")
        if detail_line:
            tk.Label(pan.body,text=detail_line,bg="#f8f9fa",fg="#111111",font=("Segoe UI",10),wraplength=950,justify="left",anchor="w",padx=10,pady=7).pack(fill="x",pady=(0,8))
        row=tk.Frame(pan.body,bg=self.colors["panel"]); row.pack(fill="x",pady=(0,8))
        asset_label=f"{a['barcode']} | {a['name']} | {a['asset_type']}"
        ttk.Button(row,text="Edit Asset",command=lambda:self.popup_asset(a)).pack(side="left",padx=4)
        ttk.Button(row,text="Use At Front Desk",command=lambda:self.use_asset_frontdesk(a)).pack(side="left",padx=4)
        ttk.Button(row,text="Add AP Alert",command=lambda:self.add_ap_alert_popup("Asset",a["barcode"],asset_label)).pack(side="left",padx=4)
        ttk.Button(row,text="View AP Alerts",command=lambda:self.ap_alerts_popup("Asset",a["barcode"],asset_label)).pack(side="left",padx=4)
        ttk.Button(row,text="Export Profile",command=lambda:self.export_asset_profile_excel(a)).pack(side="left",padx=4)
        ttk.Button(row,text="Print Profile",command=lambda:self.print_asset_profile(a)).pack(side="left",padx=4)
        profile_status_var = tk.StringVar(value="Repair")
        ttk.Label(row, text="Status:", style="Panel.TLabel").pack(side="left", padx=(12,4))
        ttk.Combobox(row, textvariable=profile_status_var, values=["Available","Repair","Missing","Retired"], state="readonly", width=10).pack(side="left", padx=4)
        ttk.Button(row,text="Set",command=lambda:self.set_asset_status_by_id(a["id"], profile_status_var.get(), win)).pack(side="left",padx=4)
        ttk.Button(row,text="Close",command=win.destroy).pack(side="left",padx=4)
        t=self.tree(pan.body,("Time","Status","Employee","Operator","Condition","Notes"),14)
        rows=self.db.all("SELECT * FROM activity WHERE lower(asset_barcode)=lower(?) ORDER BY checked_out_at DESC LIMIT 200",(a["barcode"],))
        self.fill(t,[(pretty(r["checked_out_at"]),r["status"],r["employee_name"],r["return_operator"] or r["checkout_operator"],r["condition_in"] or r["condition_out"],r["return_notes"] or r["checkout_notes"]) for r in rows])

    def asset_profile_rows(self, a):
        return [
            ["Type", a["asset_type"]],
            ["Barcode", a["barcode"]],
            ["Name", a["name"]],
            ["Status", a["status"]],
            ["Location", a["location"]],
            ["Holder", a["current_holder_name"] or "None"],
            ["Controlled Key Number", a["controlled_key_number"]],
            ["Serial / Tag", a["serial_number"]],
            ["Details", asset_detail_text(a["asset_type"], a["asset_details"] if "asset_details" in a.keys() else "")],
            ["Notes", a["notes"] or ""],
        ]

    def export_asset_profile_excel(self, a):
        path = self.folder_setting("excel_export_folder", "Exports") / f"asset_profile_{safe_filename_part(a['barcode'], 'asset')}_{stamp()}.xlsx"
        self.write_xlsx(path, [("Asset Profile", ["Field", "Value"], self.asset_profile_rows(a))])
        self.db.audit("EXPORT ASSET PROFILE", self.actor(), a["barcode"], str(path))
        messagebox.showinfo("Export Complete", f"Asset profile exported:\n{path}")

    def print_asset_profile(self, a):
        text = "\n".join(f"{label}: {value}" for label, value in self.asset_profile_rows(a))
        path = self.folder_setting("report_export_folder", "Reports") / f"ASSET_PROFILE_{safe_filename_part(a['barcode'], 'asset')}_{stamp()}.txt"
        path.write_text(text, encoding="utf-8")
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path), "print")
        except Exception:
            pass
        self.db.audit("PRINT ASSET PROFILE", self.actor(), a["barcode"], str(path))
        messagebox.showinfo("Print Asset Profile", f"Profile saved for printing:\n{path}")

    def focus_existing(self, key):
        win = self.open_windows.get(key)
        try:
            if win and win.winfo_exists():
                win.lift(); win.focus_force(); return True
        except Exception:
            pass
        return False

    def track_window(self, key, win):
        self.open_windows[key]=win
        try:
            win.minsize(680, 460)
            win.transient(self)
        except Exception:
            pass
        def close():
            self.open_windows.pop(key,None)
            try: win.destroy()
            except Exception: pass
        win.protocol("WM_DELETE_WINDOW", close)

    def go_search(self):
        if hasattr(self, "quick_search"):
            self.search_var.set(self.quick_search.get())
        self.show("Search")
        self.search()

    def set_search_mode(self, mode):
        self.search_type_var.set(mode)
        self.update_search_mode_chips()
        self.search()

    def update_search_mode_chips(self):
        if not hasattr(self, "search_mode_buttons"):
            return
        selected = self.search_type_var.get() if hasattr(self, "search_type_var") else "All"
        for mode, button in self.search_mode_buttons.items():
            self.set_chip_active(button, mode == selected)

    def clear_search(self):
        self.search_var.set("")
        self.search_type_var.set("All")
        self.search_status_var.set("Any Status")
        if hasattr(self, "search_tree"):
            self.search_tree.delete(*self.search_tree.get_children())
        self.update_search_mode_chips()

    def search(self):
        q = clean(self.search_var.get())
        search_type = self.search_type_var.get() if hasattr(self, "search_type_var") else "All"
        status_filter = self.search_status_var.get() if hasattr(self, "search_status_var") else "Any Status"
        self.update_search_mode_chips()
        if not q and search_type not in ("Audit","Errors","AP Alerts") and status_filter == "Any Status":
            return
        like = f"%{q}%"
        res = []

        def include(kind):
            return search_type in ("All", kind)

        if include("People"):
            rows = self.db.all("SELECT * FROM people WHERE employee_id LIKE ? OR badge LIKE ? OR first_name LIKE ? OR last_name LIKE ? OR department LIKE ? OR role LIKE ? OR status LIKE ? ORDER BY last_name LIMIT 200", (like,like,like,like,like,like,like))
            for p in rows:
                if status_filter == "Active" and p["status"] != "Active":
                    continue
                if status_filter == "Inactive" and p["status"] in ("Active","Manual Review"):
                    continue
                if status_filter not in ("Any Status","Active","Inactive"):
                    continue
                res.append(("Person", p["employee_id"], self.db.person_name(p), p["role"], f"Badge {p['badge']} | {p['department']} | {p['status']}"))

        if include("Assets"):
            rows = self.db.all("SELECT * FROM assets WHERE barcode LIKE ? OR name LIKE ? OR controlled_key_number LIKE ? OR serial_number LIKE ? OR location LIKE ? OR status LIKE ? OR asset_type LIKE ? ORDER BY name LIMIT 200", (like,like,like,like,like,like,like))
            for a in rows:
                if status_filter == "Available" and a["status"] != "Available":
                    continue
                if status_filter == "Checked Out" and a["status"] != "Checked Out":
                    continue
                if status_filter == "Inactive" and a["status"] not in ("Retired","Repair","Missing"):
                    continue
                if status_filter not in ("Any Status","Available","Checked Out","Inactive"):
                    continue
                res.append(("Asset", a["barcode"], a["name"], a["status"], f"{a['asset_type']} | Key {a['controlled_key_number']} | Serial {a['serial_number']} | Holder {a['current_holder_name'] or 'None'}"))

        if include("Activity"):
            rows = self.db.all("SELECT * FROM activity WHERE log_id LIKE ? OR asset_barcode LIKE ? OR asset_name LIKE ? OR employee_name LIKE ? OR status LIKE ? OR checkout_operator LIKE ? OR return_operator LIKE ? ORDER BY checked_out_at DESC LIMIT 250", (like,like,like,like,like,like,like))
            for r in rows:
                is_late = r["status"] in ("OUT","DUE SOON","OVERDUE","MISSING","REVIEW") and r["due_back_at"] and r["due_back_at"] < now_iso()
                if status_filter == "Late" and not is_late:
                    continue
                if status_filter == "Returned" and r["status"] != "RETURNED":
                    continue
                if status_filter == "Checked Out" and r["status"] not in ("OUT","DUE SOON","OVERDUE","MISSING","REVIEW"):
                    continue
                if status_filter not in ("Any Status","Late","Returned","Checked Out"):
                    continue
                res.append(("Activity", r["log_id"], f"{r['asset_barcode']} / {r['asset_name']}", "LATE" if is_late else r["status"], f"{r['employee_name']} | Due {pretty(r['due_back_at'])} | Out by {r['checkout_operator']} | Return by {r['return_operator'] or ''}"))

        if include("Audit"):
            rows = self.db.all("SELECT * FROM audit WHERE action LIKE ? OR actor LIKE ? OR target LIKE ? OR details LIKE ? ORDER BY timestamp DESC LIMIT 250", (like,like,like,like))
            for r in rows:
                res.append(("Audit", pretty(r["timestamp"]), r["action"], r["actor"], r["details"]))

        if include("Errors"):
            rows = self.db.all("SELECT * FROM errors WHERE source LIKE ? OR message LIKE ? OR details LIKE ? ORDER BY timestamp DESC LIMIT 250", (like,like,like))
            for r in rows:
                res.append(("Error", pretty(r["timestamp"]), r["source"], r["message"], r["details"][:180] if r["details"] else ""))

        if include("AP Alerts"):
            rows = self.db.all("SELECT * FROM ap_alerts WHERE target_label LIKE ? OR target_key LIKE ? OR alert_type LIKE ? OR severity LIKE ? OR note LIKE ? OR required_action LIKE ? OR status LIKE ? ORDER BY timestamp DESC LIMIT 250", (like,like,like,like,like,like,like))
            for r in rows:
                res.append(("AP Alert", pretty(r["timestamp"]), f"{r['target_type']} {r['target_key']}", r["status"], f"{r['severity']} | {r['alert_type']} | {r['note']} | Required: {r['required_action']}"))

        self.fill(self.search_tree, res)
        self.status.set(f"Search found {len(res)} result(s).")




    def clean_report_text(self, text):
        """Convert escaped report text into readable real lines."""
        if text is None:
            return ""
        text = str(text)
        return text.replace("\\r\\n", "\n").replace("\\n", "\n").replace("\\t", "    ")

    def clean_detail_text(self, text):
        """Clean saved audit/error details for readable reports."""
        text = self.clean_report_text(text)
        stripped = text.strip()
        if not stripped:
            return "(no details saved)"
        try:
            obj = json.loads(stripped)
            return json.dumps(obj, indent=2, ensure_ascii=False)
        except Exception:
            return stripped


    def report_header(self, title):
        store_name = self.db.setting("store_name", "Macy's AP")
        return [
            title,
            "=" * 72,
            f"Location: {store_name}",
            f"Generated: {pretty(now_iso())}",
            f"Generated by: {self.actor()}",
            f"Role: {self.role()}",
            "",
        ]

    def report_section(self, title):
        return ["", title, "-" * 72]

    def report_kv(self, label, value):
        return f"{label}: {value}"

    def report_blank_if_empty(self, rows, message):
        return [message] if not rows else []


    def show_report(self, text):
        """Display report text safely and switch to Reports page when needed."""
        try:
            text = self.clean_report_text(text)
            if self.current_page != "Reports":
                self.show("Reports")
            self.report_text.delete("1.0", "end")
            self.report_text.insert("1.0", text)
            self.status.set("Report generated.")
        except Exception:
            try:
                self.db.error("Show Report", "Failed to display report", traceback.format_exc())
            except Exception:
                pass
            messagebox.showerror("Report Error", "The report was created but could not be displayed.")


    def report_current_out(self):
        rows = self.db.all("SELECT * FROM activity WHERE status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') ORDER BY due_back_at")
        lines = self.report_header("CURRENT OUT REPORT")
        lines += [
            self.report_kv("Open items", len(rows)),
        ]
        lines += self.report_section("Open Items")
        if not rows:
            lines.append("No items currently out.")
        for idx, r in enumerate(rows, start=1):
            late = "YES" if r["due_back_at"] and r["due_back_at"] < now_iso() else "NO"
            lines += [
                f"Item #{idx}",
                self.report_kv("Log ID", r["log_id"]),
                self.report_kv("Status", f"{r['status']}   Late: {late}"),
                self.report_kv("Asset", f"{r['asset_type']} | {r['asset_barcode']} | {r['asset_name']}"),
                self.report_kv("Checked out to", f"{r['employee_name']}"),
                self.report_kv("Employee ID", r["employee_id"]),
                self.report_kv("Employee badge", r["employee_badge"]),
                self.report_kv("Checked out at", pretty(r["checked_out_at"])),
                self.report_kv("Due back", pretty(r["due_back_at"])),
                self.report_kv("Checkout operator", r["checkout_operator"]),
                self.report_kv("Condition out", r["condition_out"]),
                self.report_kv("Notes", r["checkout_notes"] or ""),
                "",
            ]
        self.db.audit("VIEW CURRENT OUT REPORT", self.actor(), "Reports", f"Displayed open rows: {len(rows)}")
        self.show_report("\n".join(lines))

    def report_overdue(self):
        rows = self.db.all("SELECT * FROM activity WHERE status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') AND due_back_at < ? ORDER BY due_back_at", (now_iso(),))
        lines = self.report_header("OVERDUE ASSETS REPORT")
        lines.append(self.report_kv("Overdue items", len(rows)))
        lines += self.report_section("Overdue Items")
        if not rows:
            lines.append("No overdue assets right now.")
        for idx, r in enumerate(rows, start=1):
            lines += [
                f"Overdue Item #{idx}",
                self.report_kv("Asset", f"{r['asset_type']} | {r['asset_barcode']} | {r['asset_name']}"),
                self.report_kv("Holder", f"{r['employee_name']} | {r['employee_badge']}"),
                self.report_kv("Due back", pretty(r["due_back_at"])),
                self.report_kv("Checked out at", pretty(r["checked_out_at"])),
                self.report_kv("Checkout operator", r["checkout_operator"]),
                "",
            ]
        self.db.audit("VIEW OVERDUE REPORT", self.actor(), "Reports", f"Displayed overdue rows: {len(rows)}")
        self.show_report("\n".join(lines))

    def report_asset_issues(self):
        rows = self.db.all("SELECT * FROM assets WHERE status IN ('Checked Out','Missing','Repair','Retired') ORDER BY status, asset_type, name")
        lines = self.report_header("ASSET ISSUE REPORT")
        lines.append(self.report_kv("Issue assets", len(rows)))
        lines += self.report_section("Assets Requiring Attention")
        if not rows:
            lines.append("No checked-out, missing, repair, or retired assets found.")
        for idx, a in enumerate(rows, start=1):
            lines += [
                f"Asset #{idx}",
                self.report_kv("Status", a["status"]),
                self.report_kv("Asset", f"{a['asset_type']} | {a['barcode']} | {a['name']}"),
                self.report_kv("Location", a["location"]),
                self.report_kv("Holder", a["current_holder_name"] or "None"),
                self.report_kv("Key number", a["controlled_key_number"]),
                self.report_kv("Serial", a["serial_number"]),
                self.report_kv("Notes", a["notes"] or ""),
                "",
            ]
        self.db.audit("VIEW ASSET ISSUE REPORT", self.actor(), "Reports", f"Displayed assets: {len(rows)}")
        self.show_report("\n".join(lines))

    def report_employee_history(self):
        scan = simpledialog.askstring("Employee History", "Enter employee ID or badge:")
        if not scan:
            return
        person = self.db.find_person(scan)
        if not person:
            messagebox.showwarning("Employee Not Found", "No employee matched that ID or badge.")
            return
        rows = self.db.all("SELECT * FROM activity WHERE employee_id=? OR employee_badge=? ORDER BY checked_out_at DESC LIMIT 500", (person["employee_id"], person["badge"]))
        lines = self.report_header("EMPLOYEE CHECKOUT HISTORY")
        lines += [
            self.report_kv("Employee", self.db.person_name(person)),
            self.report_kv("Employee ID", person["employee_id"]),
            self.report_kv("Badge", person["badge"]),
            self.report_kv("Rows displayed", len(rows)),
        ]
        lines += self.report_section("History")
        if not rows:
            lines.append("No checkout history found for this employee.")
        for idx, r in enumerate(rows, start=1):
            lines += [
                f"Record #{idx}",
                self.report_kv("Status", r["status"]),
                self.report_kv("Asset", f"{r['asset_type']} | {r['asset_barcode']} | {r['asset_name']}"),
                self.report_kv("Checked out", pretty(r["checked_out_at"])),
                self.report_kv("Due back", pretty(r["due_back_at"])),
                self.report_kv("Returned", pretty(r["returned_at"])),
                "",
            ]
        self.db.audit("VIEW EMPLOYEE HISTORY REPORT", self.actor(), person["employee_id"], f"Rows: {len(rows)}")
        self.show_report("\n".join(lines))

    def report_asset_history(self):
        scan = simpledialog.askstring("Asset History", "Enter asset barcode, key number, or serial:")
        if not scan:
            return
        asset = self.db.find_asset(scan)
        if not asset:
            messagebox.showwarning("Asset Not Found", "No asset matched that barcode, key number, or serial.")
            return
        rows = self.db.all("SELECT * FROM activity WHERE lower(asset_barcode)=lower(?) ORDER BY checked_out_at DESC LIMIT 500", (asset["barcode"],))
        lines = self.report_header("ASSET CHECKOUT HISTORY")
        lines += [
            self.report_kv("Asset", f"{asset['asset_type']} | {asset['barcode']} | {asset['name']}"),
            self.report_kv("Status", asset["status"]),
            self.report_kv("Location", asset["location"]),
            self.report_kv("Holder", asset["current_holder_name"] or "None"),
            self.report_kv("Rows displayed", len(rows)),
        ]
        lines += self.report_section("History")
        if not rows:
            lines.append("No checkout history found for this asset.")
        for idx, r in enumerate(rows, start=1):
            lines += [
                f"Record #{idx}",
                self.report_kv("Status", r["status"]),
                self.report_kv("Employee", f"{r['employee_name']} | {r['employee_badge']}"),
                self.report_kv("Checked out", pretty(r["checked_out_at"])),
                self.report_kv("Due back", pretty(r["due_back_at"])),
                self.report_kv("Returned", pretty(r["returned_at"])),
                "",
            ]
        self.db.audit("VIEW ASSET HISTORY REPORT", self.actor(), asset["barcode"], f"Rows: {len(rows)}")
        self.show_report("\n".join(lines))

    def report_alert_history(self):
        alerts = self.db.all("SELECT * FROM ap_alerts ORDER BY timestamp DESC LIMIT 500")
        notes = self.db.all("SELECT * FROM manager_notifications ORDER BY timestamp DESC LIMIT 500")
        lines = self.report_header("ALERT HISTORY REPORT")
        lines += [
            self.report_kv("AP alerts shown", len(alerts)),
            self.report_kv("Manager notifications shown", len(notes)),
        ]
        lines += self.report_section("AP Alerts")
        if not alerts:
            lines.append("No AP alerts found.")
        for idx, a in enumerate(alerts, start=1):
            lines += [
                f"AP Alert #{idx}",
                self.report_kv("Time", pretty(a["timestamp"])),
                self.report_kv("Target", f"{a['target_type']} {a['target_key']} | {a['target_label']}"),
                self.report_kv("Severity", a["severity"]),
                self.report_kv("Type", a["alert_type"]),
                self.report_kv("Status", a["status"]),
                self.report_kv("Required action", a["required_action"]),
                self.report_kv("Created by", a["created_by"]),
                self.report_kv("Note", a["note"]),
                "",
            ]
        lines += self.report_section("Manager Notifications")
        if not notes:
            lines.append("No Manager notifications found.")
        for idx, n in enumerate(notes, start=1):
            lines += [
                f"Manager Notification #{idx}",
                self.report_kv("Time", pretty(n["timestamp"])),
                self.report_kv("Severity", n["severity"]),
                self.report_kv("Event", n["event_type"]),
                self.report_kv("Status", n["status"]),
                self.report_kv("User", n["user_involved"]),
                self.report_kv("Asset", n["asset_involved"]),
                self.report_kv("Action", n["action_taken"]),
                self.report_kv("Handled by", n["handled_by"]),
                self.report_kv("Notes", n["notes"]),
                self.report_kv("Reason", n["reason"]),
                "",
            ]
        self.db.audit("VIEW ALERT HISTORY REPORT", self.actor(), "Reports", f"AP alerts={len(alerts)} Manager notifications={len(notes)}")
        self.show_report("\n".join(lines))

    def report_group_history(self):
        groups = self.db.all("SELECT g.*, (SELECT COUNT(*) FROM people p WHERE p.role=g.name AND p.status!='Deleted') user_count FROM groups g ORDER BY CASE protected WHEN 'Yes' THEN 0 ELSE 1 END, name")
        audit_rows = self.db.all("SELECT * FROM audit WHERE action LIKE '%GROUP%' OR action LIKE '%PERMISSION%' ORDER BY timestamp DESC LIMIT 500")
        lines = self.report_header("GROUP HISTORY REPORT")
        lines += [
            self.report_kv("Groups shown", len(groups)),
            self.report_kv("History rows shown", len(audit_rows)),
        ]
        lines += self.report_section("Current Groups")
        if not groups:
            lines.append("No groups found.")
        for idx, g in enumerate(groups, start=1):
            lines += [
                f"Group #{idx}",
                self.report_kv("Name", g["name"]),
                self.report_kv("Protected", g["protected"]),
                self.report_kv("Assigned users", g["user_count"]),
                self.report_kv("Rights", g["rights"]),
                self.report_kv("Notes", g["notes"]),
                "",
            ]
        lines += self.report_section("Group / Permission History")
        if not audit_rows:
            lines.append("No group or permission audit history found.")
        for idx, r in enumerate(audit_rows, start=1):
            lines += [
                f"History #{idx}",
                self.report_kv("Time", pretty(r["timestamp"])),
                self.report_kv("Action", r["action"]),
                self.report_kv("Actor", r["actor"]),
                self.report_kv("Role", r["actor_role"] if "actor_role" in r.keys() else ""),
                self.report_kv("Target", r["target"]),
                self.report_kv("Status", r["status"] if "status" in r.keys() else ""),
                self.report_kv("Details", r["details"]),
                "",
            ]
        self.db.audit("VIEW GROUP HISTORY REPORT", self.actor(), "Reports", f"Groups={len(groups)} History={len(audit_rows)}")
        self.show_report("\n".join(lines))


    def report_end_shift(self, extra_closeout=None):
        c = self.db.counts()
        open_rows = self.db.all("SELECT * FROM activity WHERE status IN ('OUT','DUE SOON','OVERDUE','MISSING','REVIEW') ORDER BY due_back_at")
        error_rows = self.db.all("SELECT * FROM errors ORDER BY timestamp DESC LIMIT 20")
        audit_rows = self.db.all("SELECT * FROM audit ORDER BY timestamp DESC LIMIT 20")

        lines = self.report_header("MACY'S AP END OF SHIFT REPORT")

        lines += self.report_section("Summary")
        for label, key in [
            ("Keys out", "keys_out"),
            ("Radios out", "radios_out"),
            ("Items out", "items_out"),
            ("Late returns", "late_returns"),
            ("Total users", "people"),
            ("Employees", "employee_users"),
            ("Front Desk users", "front_desk_users"),
            ("AP Operators", "ap_operators"),
            ("Managers", "managers"),
            ("Admins", "admins"),
            ("Total assets", "assets"),
            ("Active alerts", "manager_alerts"),
            ("Errors today", "errors"),
        ]:
            lines.append(self.report_kv(label, c.get(key, 0)))

        if extra_closeout:
            lines += self.report_section("Daily Closeout Checklist")
            lines.append(self.clean_detail_text(extra_closeout))

        lines += self.report_section("Open Items / Late Returns")
        if not open_rows:
            lines.append("No open items.")
        for idx, r in enumerate(open_rows, start=1):
            late = "YES" if r["due_back_at"] and r["due_back_at"] < now_iso() else "NO"
            lines += [
                f"Open Item #{idx}",
                self.report_kv("Status", f"{r['status']}   Late: {late}"),
                self.report_kv("Asset", f"{r['asset_type']} | {r['asset_barcode']} | {r['asset_name']}"),
                self.report_kv("Holder", r["employee_name"]),
                self.report_kv("Due back", pretty(r["due_back_at"])),
                "",
            ]

        lines += self.report_section("Recent Errors")
        if not error_rows:
            lines.append("No recent errors.")
        for idx, e in enumerate(error_rows, start=1):
            lines += [
                f"Error #{idx}",
                self.report_kv("Time", pretty(e["timestamp"])),
                self.report_kv("Source", e["source"]),
                self.report_kv("Message", e["message"]),
                "",
            ]

        lines += self.report_section("Recent Audit Entries")
        if not audit_rows:
            lines.append("No recent audit entries.")
        for idx, a in enumerate(audit_rows, start=1):
            lines += [
                f"Audit #{idx}",
                self.report_kv("Time", pretty(a["timestamp"])),
                self.report_kv("Action", a["action"]),
                self.report_kv("Actor", a["actor"]),
                self.report_kv("Target", a["target"]),
                self.report_kv("Details", self.clean_detail_text(a["details"])[:360]),
                "",
            ]

        self.db.audit("VIEW END OF SHIFT REPORT", self.actor(), "Reports", f"Open={len(open_rows)} ErrorsShown={len(error_rows)} AuditShown={len(audit_rows)}")
        self.show_report("\n".join(lines))


    def report_audit(self):
        rows = self.db.all("SELECT * FROM audit ORDER BY timestamp DESC LIMIT 1000")
        lines = self.report_header("DETAILED AUDIT REPORT")
        lines += [
            self.report_kv("Database", self.db.path),
            self.report_kv("Total displayed", len(rows)),
            "",
            "This report shows who did what, when it happened, what record was affected, and the full saved details.",
        ]
        lines += self.report_section("Audit Records")
        if not rows:
            lines.append("No audit entries found.")
        for idx, r in enumerate(rows, start=1):
            lines += [
                f"Audit Record #{idx}",
                self.report_kv("Audit ID", r["id"]),
                self.report_kv("Time", pretty(r["timestamp"])),
                self.report_kv("Action", r["action"]),
                self.report_kv("Actor", r["actor"]),
                self.report_kv("Target", r["target"]),
                "Details:",
                self.clean_detail_text(r["details"]),
                "",
            ]
        self.db.audit("VIEW DETAILED AUDIT REPORT", self.actor(), "Reports", f"Displayed audit rows: {len(rows)}")
        self.show_report("\n".join(lines))


    def report_errors(self):
        rows = self.db.all("SELECT * FROM errors ORDER BY timestamp DESC LIMIT 1000")
        lines = self.report_header("DETAILED ERROR REPORT")
        lines += [
            self.report_kv("Database", self.db.path),
            self.report_kv("Total displayed", len(rows)),
            "",
            "This report shows application errors in detail, including saved traceback/details when available.",
        ]
        lines += self.report_section("Error Records")
        if not rows:
            lines.append("No error entries found.")
        for idx, r in enumerate(rows, start=1):
            lines += [
                f"Error Record #{idx}",
                self.report_kv("Error ID", r["id"]),
                self.report_kv("Time", pretty(r["timestamp"])),
                self.report_kv("Source", r["source"]),
                self.report_kv("Message", r["message"]),
                "Details:",
                self.clean_detail_text(r["details"]),
                "",
            ]
        self.db.audit("VIEW DETAILED ERROR REPORT", self.actor(), "Reports", f"Displayed error rows: {len(rows)}")
        self.show_report("\n".join(lines))



    def current_report_text(self):
        try:
            return self.report_text.get("1.0", "end").strip()
        except Exception:
            return ""

    def copy_report(self):
        text = self.current_report_text()
        if not text:
            messagebox.showinfo("Copy Report", "No report text to copy.")
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self.db.audit("COPY REPORT", self.actor(), "Reports", f"Copied report text length: {len(text)}")
        messagebox.showinfo("Copy Report", "Report copied to clipboard.")

    def save_report_txt(self):
        text = self.current_report_text()
        if not text:
            messagebox.showinfo("Save Report", "No report text to save.")
            return
        path = filedialog.asksaveasfilename(title="Save report as TXT", defaultextension=".txt", initialdir=str(self.folder_setting("report_export_folder", "Reports")), initialfile=f"AP_Report_{stamp()}.txt", filetypes=[("Text files","*.txt"),("All files","*.*")])
        if not path:
            return
        Path(path).write_text(text, encoding="utf-8")
        self.db.audit("SAVE REPORT TXT", self.actor(), "Reports", str(path))
        messagebox.showinfo("Save Report", f"Report saved:\\n{path}")

    def save_report_html(self):
        text = self.current_report_text()
        if not text:
            messagebox.showinfo("Save Report", "No report text to save.")
            return
        path = filedialog.asksaveasfilename(title="Save report as HTML", defaultextension=".html", initialdir=str(self.folder_setting("report_export_folder", "Reports")), initialfile=f"AP_Report_{stamp()}.html", filetypes=[("HTML files","*.html"),("All files","*.*")])
        if not path:
            return
        safe = (text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))
        html = f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Macy's AP Report</title>
<style>
body {{ font-family: Segoe UI, Arial, sans-serif; margin: 32px; color: #111; }}
pre {{ white-space: pre-wrap; line-height: 1.4; font-size: 14px; }}
h1 {{ color: #c40000; }}
</style>
</head>
<body>
<h1>Macy's AP Report</h1>
<pre>{safe}</pre>
</body>
</html>"""
        Path(path).write_text(html, encoding="utf-8")
        self.db.audit("SAVE REPORT HTML", self.actor(), "Reports", str(path))
        messagebox.showinfo("Save Report", f"HTML report saved:\\n{path}")

    def save_report_excel(self):
        text = self.current_report_text()
        if not text:
            messagebox.showinfo("Save Report", "No report text to save.")
            return
        path = filedialog.asksaveasfilename(title="Save report as Excel", defaultextension=".xlsx", initialdir=str(self.folder_setting("report_export_folder", "Reports")), initialfile=f"AP_Report_{stamp()}.xlsx", filetypes=[("Excel Workbook","*.xlsx"),("All files","*.*")])
        if not path:
            return
        try:
            self.write_xlsx(Path(path), [("Report", ["Report Lines"], [[line] for line in text.splitlines()])])
            self.db.audit("SAVE REPORT EXCEL", self.actor(), "Reports", str(path))
            messagebox.showinfo("Save Report", f"Excel report saved:\n{path}")
        except Exception as e:
            self.db.error("Save Report Excel", str(e), traceback.format_exc())
            messagebox.showerror("Save Report", f"Could not save Excel report:\n{e}")



    def print_report(self):
        text = self.report_text.get("1.0", "end").strip()
        if not text:
            messagebox.showinfo("Print Report", "No report text to print.")
            return
        path = self.folder_setting("report_export_folder", "Reports") / f"PRINT_REPORT_{stamp()}.txt"
        path.write_text(text, encoding="utf-8")
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path), "print")
                messagebox.showinfo("Print Report", f"Sent to default printer.\\n\\nSaved copy:\\n{path}")
            else:
                messagebox.showinfo("Print Report", f"Report saved for printing:\\n{path}")
        except Exception:
            messagebox.showinfo("Print Report", f"Could not send directly to printer. Saved report here:\\n{path}")
        self.db.audit("PRINT REPORT", self.actor(), "Report", str(path))


    def export_xlsx(self):
        # Lightweight Excel-compatible CSV bundle instead of external dependencies.
        folder = filedialog.askdirectory(title="Choose Export Folder")
        if not folder: return
        tables = ["people","assets","activity","audit","errors","manager_notifications","ap_alerts","groups","settings"]
        for table in tables:
            self.write_csv(Path(folder)/f"{table}.csv", self.db.all(f"SELECT * FROM {table} ORDER BY 1"), self.table_columns(table))
        self.db.audit("EXPORT CSV BUNDLE", self.actor(), "Reports", f"Export folder: {folder}; Tables: {', '.join(tables)}")
        messagebox.showinfo("Export Complete", f"CSV files exported to:\n{folder}")

    def export_assets_by_type(self):
        asset_type = self.current_asset_filter if self.current_asset_filter in ASSET_TYPES else None
        if not asset_type:
            asset_type = simpledialog.askstring("Export By Type", "Enter asset type to export:\n" + ", ".join(ASSET_TYPES), parent=self)
            asset_type = clean(asset_type)
        if asset_type not in ASSET_TYPES:
            messagebox.showwarning("Choose Asset Type", "Choose a valid asset type before exporting.")
            return
        self.export_assets_excel(asset_type)

    def asset_export_table(self, asset_type):
        rows = self.db.all("SELECT * FROM assets WHERE asset_type=? ORDER BY name, barcode", (asset_type,))
        schema = asset_detail_schema(asset_type)
        headers = ["Asset Type", "Barcode", "Name", "Status", "Location", "Holder", "Controlled Key Number", "Serial / Tag", "Notes"]
        headers += [spec["label"] for spec in schema["fields"]]
        if asset_type == "Tablet":
            headers += ["Accessories"]
        if asset_type == "Key":
            headers += ["Keys on Ring"]
        data = []
        for row in rows:
            details = parse_asset_details(row["asset_details"] if "asset_details" in row.keys() else "")
            values = [row["asset_type"], row["barcode"], row["name"], row["status"], row["location"], row["current_holder_name"], row["controlled_key_number"], row["serial_number"], row["notes"]]
            values += [details.get(spec["key"], "") for spec in schema["fields"]]
            if asset_type == "Tablet":
                accessories = details.get("accessories", [])
                if details.get("other_accessory"):
                    accessories = list(accessories) + [details["other_accessory"]]
                values.append(", ".join(accessories))
            if asset_type == "Key":
                keys = []
                for idx, item in enumerate(details.get("keys", []), start=1):
                    keys.append(f"Key {idx}: {item.get('serial','')} - {item.get('access','')}".strip(" -"))
                values.append("; ".join(keys))
            data.append(values)
        return headers, data

    def export_assets_excel(self, asset_type=None):
        if asset_type and asset_type not in ASSET_TYPES:
            messagebox.showwarning("Choose Asset Type", "Choose a valid asset type before exporting.")
            return
        default_name = f"assets_{safe_filename_part((asset_type or 'all').lower().replace(' ', '_'), 'all')}_{stamp()}.xlsx"
        path = filedialog.asksaveasfilename(title="Save Asset Excel Export", defaultextension=".xlsx", initialdir=str(self.folder_setting("excel_export_folder", "Exports")), initialfile=default_name, filetypes=[("Excel Workbook", "*.xlsx")])
        if not path:
            return
        sheets = []
        export_types = [asset_type] if asset_type else ASSET_TYPES
        for t in export_types:
            headers, data = self.asset_export_table(t)
            if data or asset_type:
                sheets.append((t[:31], headers, data))
        if not sheets:
            sheets.append(("Assets", ["Message"], [["No asset records found."]]))
        try:
            self.write_xlsx(Path(path), sheets)
            detail = f"Asset type: {asset_type or 'All'}; Sheets: {len(sheets)}; File: {path}"
            self.db.audit("EXPORT ASSETS EXCEL", self.actor(), "Assets", detail)
            self.db.notify_manager("Export created", "Info", self.actor(), "", "Asset Excel export", self.actor(), detail, status="Reviewed")
            messagebox.showinfo("Excel Export Complete", f"Asset Excel workbook saved:\n{path}")
        except Exception as e:
            self.db.error("Export Assets Excel", str(e), traceback.format_exc())
            self.db.notify_manager("Export failed", "Warning", self.actor(), "", "Asset Excel export failed", self.actor(), str(e))
            messagebox.showerror("Excel Export Failed", f"Could not save Excel export:\n{e}")

    def export_context_rows(self):
        rows = [("Current page", getattr(self, "current_page", ""))]
        if hasattr(self, "asset_filter_var"):
            rows.extend([
                ("Asset type filter", self.asset_filter_var.get()),
                ("Asset status filter", self.asset_status_filter.get()),
                ("Asset search", self.asset_search_var.get()),
                ("Asset export dropdown", self.asset_export_type_var.get()),
            ])
        if hasattr(self, "log_filter_var"):
            rows.extend([
                ("Log type filter", self.log_filter_var.get()),
                ("Log search", self.log_search_var.get()),
                ("Log user filter", self.log_user_var.get()),
                ("Log asset filter", self.log_asset_var.get()),
                ("Log action filter", self.log_action_var.get()),
                ("Log from date", self.log_from_var.get()),
                ("Log to date", self.log_to_var.get()),
            ])
        if hasattr(self, "manager_notification_status_var"):
            rows.extend([
                ("Manager notification status", self.manager_notification_status_var.get()),
                ("Manager notification severity", self.manager_notification_severity_var.get()),
                ("Manager notification search", self.manager_notification_search_var.get()),
            ])
        return [(label, clean(value)) for label, value in rows if clean(value)]

    def write_xlsx(self, path, sheets):
        if self is not None:
            now = dt.datetime.now()
            info_rows = [
                ["Title", Path(path).stem],
                ["Export date", now.strftime("%Y-%m-%d")],
                ["Export time 12-hour", now.strftime("%I:%M:%S %p")],
                ["Export time 24-hour", now.strftime("%H:%M:%S")],
                ["Exported by", self.actor()],
                ["Role", self.role()],
                ["Worksheets", ", ".join(str(sheet[0]) for sheet in sheets)],
            ]
            for label, value in self.export_context_rows():
                info_rows.append([label, value])
            sheets = [("Export Info", ["Field", "Value"], info_rows)] + list(sheets)
        write_xlsx_workbook(path, sheets)

    def table_columns(self, table):
        if table not in {"people", "assets", "activity", "audit", "errors", "manager_notifications", "ap_alerts", "groups", "settings"}:
            return []
        rows = self.db.all(f"PRAGMA table_info({table})")
        return [r["name"] for r in rows]

    def write_csv(self, path, rows, headers=None):
        headers = list(headers or (rows[0].keys() if rows else []))
        with open(path,"w",newline="",encoding="utf-8-sig") as f:
            w=csv.writer(f)
            if headers:
                w.writerow(headers)
            for r in rows:
                w.writerow([r[k] for k in headers])

    def backup_with_logs(self, show_message=True, label="BACKUP WITH LOGS"):
        folder=Path(self.db.setting("backup_folder", str(self.app_dir/"Backups")))
        folder.mkdir(parents=True, exist_ok=True)
        temp=folder/f"backup_{stamp()}"
        temp.mkdir(exist_ok=True)
        self.db.conn.commit()
        shutil.copy2(self.db.path, temp/self.db.path.name)
        backup_tables = ["people","assets","activity","audit","errors","manager_notifications","ap_alerts","groups","settings"]
        for table in backup_tables:
            self.write_csv(temp/f"{table}.csv", self.db.all(f"SELECT * FROM {table} ORDER BY 1"), self.table_columns(table))
        (temp/"BACKUP_README.txt").write_text(
            f"{APP_TITLE}\n{APP_VERSION}\nCreated: {now_iso()}\nOperator: {self.actor()}\nRole: {self.role()}\n"
            f"Data file: {self.db.path}\nIncludes: database and CSV files for {', '.join(backup_tables)}\n"
            f"People: {self.db.one('SELECT COUNT(*) c FROM people')['c']}\n"
            f"Assets: {self.db.one('SELECT COUNT(*) c FROM assets')['c']}\n"
            f"Activity: {self.db.one('SELECT COUNT(*) c FROM activity')['c']}\n"
            f"Audit: {self.db.one('SELECT COUNT(*) c FROM audit')['c']}\n"
            f"Errors: {self.db.one('SELECT COUNT(*) c FROM errors')['c']}\n",
            encoding="utf-8")
        prefix = "MACYS_AP_AUTO_BACKUP" if label.startswith("AUTO") else "MACYS_AP_BACKUP_WITH_LOGS"
        out=folder/f"{prefix}_{stamp()}.zip"
        try:
            with zipfile.ZipFile(out,"w",zipfile.ZIP_DEFLATED) as z:
                for f in temp.iterdir(): z.write(f,arcname=f.name)
        finally:
            shutil.rmtree(temp, ignore_errors=True)
        self.db.set_setting("last_backup", now_iso())
        self.db.audit(label, self.actor(), "Backup", str(out))
        if hasattr(self, "manager_backup_status_var"):
            self.refresh_manager()
        if show_message:
            messagebox.showinfo("Backup Complete", f"Backup saved:\n{out}")
        return out

    def open_backup_folder(self):
        folder=Path(self.db.setting("backup_folder", str(self.app_dir/"Backups")))
        folder.mkdir(parents=True, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(folder)
        else:
            messagebox.showinfo("Backup Folder", str(folder))

    def issue_report(self):
        self.backup_with_logs()

    def group_management(self):
        if not self.require("system"):
            return
        key = "group_management"
        if self.focus_existing(key):
            return
        win = tk.Toplevel(self)
        win.title("Group Management")
        win.geometry("1100x740")
        self.apply_icon(win)
        self.track_window(key, win)
        pan = self.panel(win, "Group Management", "Manage groups, permissions, and assigned users.")
        pan.pack(fill="both", expand=True, padx=14, pady=14)
        top = tk.Frame(pan.body, bg=self.colors["panel"])
        top.pack(fill="x", pady=(0,8))
        name_var = tk.StringVar()
        ttk.Label(top, text="Group:", style="Panel.TLabel").pack(side="left")
        ttk.Entry(top, textvariable=name_var, width=24).pack(side="left", padx=4)
        rights_vars = {right: tk.BooleanVar(value=False) for right in ALL_RIGHTS}
        body = tk.Frame(pan.body, bg=self.colors["panel"])
        body.pack(fill="both", expand=True)
        left = tk.Frame(body, bg=self.colors["panel"])
        left.pack(side="left", fill="both", expand=True, padx=(0,8))
        right = tk.Frame(body, bg=self.colors["panel"])
        right.pack(side="left", fill="both", expand=True)
        self.group_tree = self.tree(left, ("Group","Protected","Users","Rights"), 14)
        rights_box = self.panel(right, "Rights", "Check the rights this group should have.")
        rights_box.pack(fill="x", pady=(0,8))
        right_groups = [
            ("Access", ["dashboard", "search", "reports"]),
            ("People", ["people_view", "people_edit"]),
            ("Assets", ["assets_view", "assets_edit", "frontdesk"]),
            ("Admin", ["manager", "system", "admin"]),
        ]
        for col, (group_title, rights) in enumerate(right_groups):
            group_frame = tk.Frame(rights_box.body, bg=self.colors["panel_alt"], highlightbackground=self.colors["line"], highlightthickness=1)
            group_frame.grid(row=0, column=col, sticky="nsew", padx=5, pady=4)
            ttk.Label(group_frame, text=group_title, style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=8, pady=(8,3))
            for right in rights:
                if right not in rights_vars:
                    continue
                cb = ttk.Checkbutton(group_frame, text=right, variable=rights_vars[right])
                cb.pack(anchor="w", padx=8, pady=2)
            rights_box.body.grid_columnconfigure(col, weight=1)
        assigned = self.panel(right, "Assigned Users", "Users currently assigned to the selected group.")
        assigned.pack(fill="both", expand=True)
        self.group_users_tree = self.tree(assigned.body, ("Employee ID","Badge","Name","Status"), 8)

        def selected_group():
            sel = self.group_tree.selection()
            if not sel:
                return None
            return self.group_tree.item(sel[0], "values")[0]

        def open_group_user_profile():
            values = self.selected_tree_values(self.group_users_tree)
            if not values:
                messagebox.showinfo("Open User", "Select an assigned user first.")
                return
            person = self.db.find_person(values[0])
            if person:
                self.person_profile(person)

        def export_group_users():
            group = selected_group()
            if not group:
                messagebox.showinfo("Export Group Users", "Select a group first.")
                return
            rows = self.db.all("SELECT * FROM people WHERE role=? AND status!='Deleted' ORDER BY last_name, first_name", (group,))
            data = [[u["employee_id"], u["badge"], self.db.person_name(u), u["status"], u["department"], u["shift"]] for u in rows]
            path = self.folder_setting("excel_export_folder", "Exports") / f"group_users_{safe_filename_part(group.replace(' ', '_'), 'group')}_{stamp()}.xlsx"
            self.write_xlsx(path, [("Group Users", ["Employee ID","Badge","Name","Status","Department","Shift"], data)])
            self.db.audit("EXPORT GROUP USERS", self.actor(), group, f"Rows: {len(data)}; File: {path}")
            messagebox.showinfo("Export Group Users", f"Group users exported:\n{path}")

        def refresh_groups():
            rows = self.db.all("SELECT g.*, (SELECT COUNT(*) FROM people p WHERE p.role=g.name AND p.status!='Deleted') user_count FROM groups g ORDER BY CASE protected WHEN 'Yes' THEN 0 ELSE 1 END, name")
            self.fill(self.group_tree, [(r["name"], r["protected"], r["user_count"], r["rights"]) for r in rows])

        def load_group(event=None):
            group = selected_group()
            if not group:
                return
            row = self.db.one("SELECT * FROM groups WHERE name=?", (group,))
            name_var.set(group)
            rights = {x for x in str(row["rights"] or "").split(",") if x}
            for right, var in rights_vars.items():
                var.set(right in rights)
            users = self.db.all("SELECT * FROM people WHERE role=? AND status!='Deleted' ORDER BY last_name, first_name", (group,))
            self.fill(self.group_users_tree, [(u["employee_id"], u["badge"], self.db.person_name(u), u["status"]) for u in users])

        def save_group():
            group = clean(name_var.get())
            if not group:
                messagebox.showwarning("Group Name Required", "Enter a group name.")
                return
            rights = ",".join([right for right, var in rights_vars.items() if var.get()])
            existing = self.db.one("SELECT * FROM groups WHERE name=?", (group,))
            if existing:
                reason = simpledialog.askstring("Reason Required", f"Why update rights for group '{group}'?", parent=win)
                if not clean(reason):
                    messagebox.showwarning("Reason Required", "A reason is required.")
                    return
                old_rights = existing["rights"] or ""
                self.db.run("UPDATE groups SET rights=?, updated_at=? WHERE name=?", (rights, now_iso(), group))
                action = "GROUP RIGHTS UPDATED"
                detail = f"Old rights: {old_rights}; New rights: {rights}; Reason: {clean(reason)}"
            else:
                self.db.run("INSERT INTO groups(name,rights,protected,notes,created_at,updated_at) VALUES(?,?,?,?,?,?)", (group, rights, "No", "", now_iso(), now_iso()))
                action = "GROUP ADDED"
                detail = f"Rights: {rights}"
            self.db.audit(action, self.actor(), group, detail)
            self.db.notify_manager("Permission changed", "Info", self.actor(), "", action, self.actor(), f"Group: {group}; {detail}", status="Reviewed")
            refresh_groups()

        def rename_group():
            old = selected_group()
            new = clean(name_var.get())
            if not old or not new or old == new:
                messagebox.showinfo("Rename Group", "Select a group and enter a new name.")
                return
            row = self.db.one("SELECT * FROM groups WHERE name=?", (old,))
            if row and row["protected"] == "Yes":
                messagebox.showwarning("Protected Group", "Default groups cannot be renamed.")
                return
            reason = simpledialog.askstring("Reason Required", f"Why rename group '{old}' to '{new}'?", parent=win)
            if not clean(reason):
                messagebox.showwarning("Reason Required", "A reason is required.")
                return
            affected = self.db.one("SELECT COUNT(*) c FROM people WHERE role=?", (old,))["c"]
            self.db.run("UPDATE groups SET name=?, updated_at=? WHERE name=?", (new, now_iso(), old))
            self.db.run("UPDATE people SET role=?, updated_at=? WHERE role=?", (new, now_iso(), old))
            detail = f"Old: {old}; New: {new}; Users affected: {affected}; Reason: {clean(reason)}"
            self.db.audit("GROUP RENAMED", self.actor(), new, detail)
            self.db.notify_manager("Group renamed", "Warning", self.actor(), "", "Group renamed", self.actor(), detail, clean(reason))
            refresh_groups()

        def delete_group():
            group = selected_group()
            if not group:
                messagebox.showinfo("Delete Group", "Select a group first.")
                return
            row = self.db.one("SELECT * FROM groups WHERE name=?", (group,))
            if row and row["protected"] == "Yes":
                messagebox.showwarning("Protected Group", "Default Admin, Manager, and system groups cannot be deleted.")
                return
            affected = self.db.one("SELECT COUNT(*) c FROM people WHERE role=?", (group,))["c"]
            replacement = ""
            if affected:
                if not messagebox.askyesno("Users Assigned", f"{affected} user(s) are assigned to this group.\n\nChoose a replacement group before deleting?"):
                    return
                replacement = clean(simpledialog.askstring("Replacement Group", "Move assigned users to which group?\n\nUse Employee to remove the custom group assignment.", initialvalue="Employee", parent=win))
                if replacement not in self.db.group_names() or replacement == group:
                    messagebox.showwarning("Replacement Required", "Choose an existing replacement group that is not the group being deleted.")
                    return
            reason = simpledialog.askstring("Reason Required", f"Why delete group '{group}'?", parent=win)
            if not clean(reason):
                messagebox.showwarning("Reason Required", "A reason is required.")
                return
            if affected:
                self.db.run("UPDATE people SET role=?, updated_at=? WHERE role=?", (replacement, now_iso(), group))
            self.db.run("DELETE FROM groups WHERE name=?", (group,))
            detail = f"Deleted group: {group}; Users affected: {affected}; Replacement group: {replacement or 'None'}; Reason: {clean(reason)}"
            self.db.audit("GROUP DELETED", self.actor(), group, detail)
            self.db.notify_manager("Group deleted", "Critical", self.actor(), "", "Group deleted", self.actor(), detail, clean(reason))
            refresh_groups()

        def assign_user_to_group():
            group = selected_group()
            if not group:
                messagebox.showinfo("Assign User", "Select a group first.")
                return
            scan = simpledialog.askstring("Assign User", "Enter employee ID or badge to assign to this group:", parent=win)
            if not clean(scan):
                return
            person = self.db.find_person(scan)
            if not person:
                messagebox.showwarning("User Not Found", "No user matched that employee ID or badge.")
                return
            old_group = person["role"] or "Employee"
            if old_group == group:
                messagebox.showinfo("Assign User", "That user is already assigned to this group.")
                return
            if not self.can_assign_role(old_group, group, person):
                return
            reason = simpledialog.askstring("Reason Required", f"Why assign {self.db.person_name(person)} from {old_group} to {group}?", parent=win)
            if not clean(reason):
                messagebox.showwarning("Reason Required", "A reason is required.")
                return
            self.db.run("UPDATE people SET role=?, updated_at=? WHERE id=?", (group, now_iso(), person["id"]))
            detail = f"User: {self.db.person_name(person)}; Employee ID: {person['employee_id']}; Old group: {old_group}; New group: {group}; Reason: {clean(reason)}"
            self.db.audit("USER GROUP ASSIGNED", self.actor(), person["employee_id"], detail)
            self.db.notify_manager("Permission changed", "Info", self.db.person_name(person), "", "User group assigned", self.actor(), detail, clean(reason), status="Reviewed")
            load_group()
            refresh_groups()
            self.refresh_people()

        self.group_tree.bind("<<TreeviewSelect>>", load_group)
        self.bind_context_menu(self.group_tree, lambda t: [
            ("Load Group", load_group, bool(self.selected_tree_values(t))),
            ("Add / Save Group", save_group),
            ("Assign User", assign_user_to_group, bool(selected_group())),
            ("Export Group Users", export_group_users, bool(selected_group())),
            ("Rename Selected", rename_group, bool(selected_group())),
            ("Delete Selected", delete_group, bool(selected_group())),
            ("Copy Row", lambda: self.copy_tree_row(t), bool(self.selected_tree_values(t))),
        ])
        self.bind_context_menu(self.group_users_tree, lambda t: [
            ("Open User Profile", open_group_user_profile, bool(self.selected_tree_values(t))),
            ("Assign User To Selected Group", assign_user_to_group, bool(selected_group())),
            ("Export Group Users", export_group_users, bool(selected_group())),
            ("Copy Row", lambda: self.copy_tree_row(t), bool(self.selected_tree_values(t))),
        ])
        buttons = tk.Frame(pan.body, bg=self.colors["panel"])
        buttons.pack(fill="x", pady=8)
        ttk.Button(buttons, text="Add / Save Group", style="Green.TButton", command=save_group).pack(side="left", padx=4)
        ttk.Button(buttons, text="Assign User", command=assign_user_to_group).pack(side="left", padx=4)
        ttk.Button(buttons, text="Export Users", command=export_group_users).pack(side="left", padx=4)
        ttk.Button(buttons, text="Rename Selected", command=rename_group).pack(side="left", padx=4)
        ttk.Button(buttons, text="Delete Selected", command=delete_group).pack(side="left", padx=4)
        ttk.Button(buttons, text="Refresh", command=refresh_groups).pack(side="left", padx=4)
        ttk.Button(buttons, text="Close", command=win.destroy).pack(side="left", padx=4)
        refresh_groups()

    def app_settings(self):
        if not self.require("system"):
            return
        key = "app_settings"
        if self.focus_existing(key):
            return
        win = tk.Toplevel(self)
        win.title("App Settings")
        win.geometry("900x680")
        self.apply_icon(win)
        self.track_window(key, win)
        pan = self.panel(win, "App Settings", "Store name, default due time, backups, and asset categories.")
        pan.pack(fill="both", expand=True, padx=14, pady=14)

        vars = {
            "store_name": tk.StringVar(value=self.db.setting("store_name", "Macy's AP")),
            "default_due_time": tk.StringVar(value=self.db.setting("default_due_time", "18:30")),
            "backup_folder": tk.StringVar(value=self.db.setting("backup_folder", str(self.app_dir / "Backups"))),
            "excel_export_folder": tk.StringVar(value=self.db.setting("excel_export_folder", str(self.app_dir / "Exports"))),
            "report_export_folder": tk.StringVar(value=self.db.setting("report_export_folder", str(self.app_dir / "Reports"))),
            "error_log_folder": tk.StringVar(value=self.db.setting("error_log_folder", str(self.app_dir / "Logs"))),
            "audit_log_folder": tk.StringVar(value=self.db.setting("audit_log_folder", str(self.app_dir / "Logs"))),
            "system_log_folder": tk.StringVar(value=self.db.setting("system_log_folder", str(self.app_dir / "Logs"))),
            "auto_daily_backup": tk.StringVar(value=self.db.setting("auto_daily_backup", "Yes")),
            "frontdesk_scan_only_mode": tk.StringVar(value=self.db.setting("frontdesk_scan_only_mode", "No")),
            "default_export_type": tk.StringVar(value=self.db.setting("default_export_type", "All")),
            "refresh_seconds": tk.StringVar(value=clean(self.db.setting("refresh_seconds", "")) or "300"),
            "display_density": tk.StringVar(value=self.db.setting("display_density", "Comfortable")),
        }
        data_file_var = tk.StringVar(value=str(self.current_data_file()))
        path_entries = []
        form = tk.Frame(pan.body, bg=self.colors["panel"])
        form.pack(fill="x")
        data_row = tk.Frame(form, bg=self.colors["panel"])
        data_row.grid(row=0, column=0, columnspan=4, sticky="ew", padx=6, pady=(0,8))
        data_row.grid_columnconfigure(1, weight=1)
        ttk.Label(data_row, text="Main Data File Location", style="Panel.TLabel").grid(row=0, column=0, sticky="w", padx=(0,6))
        data_entry = ttk.Entry(data_row, textvariable=data_file_var, state="readonly")
        data_entry.grid(row=0, column=1, sticky="ew", padx=4)
        self.tip(data_entry, "This is the active database file. Use Change to move to a shared/network database; restart after changing.")
        ttk.Button(data_row, text="Change", command=self.change_data_location).grid(row=0, column=2, padx=3)
        ttk.Button(data_row, text="Open", command=self.open_data_folder).grid(row=0, column=3, padx=3)
        rows = [
            ("Store / Location Name", "store_name", "Shown on generated reports."),
            ("Default Due Time (HH:MM)", "default_due_time", "Used when the Front Desk screen opens."),
            ("Backup Folder", "backup_folder", "Where backup ZIP files are saved."),
            ("Excel Export Folder", "excel_export_folder", "Default folder for Excel asset/log exports."),
            ("Report Export Folder", "report_export_folder", "Default folder for saved TXT/HTML reports."),
            ("Error Log Folder", "error_log_folder", "Default folder for exported error logs."),
            ("Audit Log Folder", "audit_log_folder", "Default folder for exported audit logs."),
            ("System Log Folder", "system_log_folder", "Default folder for system reports and logs."),
            ("Automatic Daily Backup", "auto_daily_backup", "Yes keeps one backup per day without interrupting users."),
            ("Front Desk Scan-Only Mode", "frontdesk_scan_only_mode", "Admin only. Yes lets Front Desk checkout/return by scanning employee badge and item without operator login."),
            ("Default Asset Export Type", "default_export_type", "Default asset type selected in the Asset page Excel export dropdown."),
            ("Refresh Timer Seconds", "refresh_seconds", "Auto-refresh interval. Default is 300 seconds / 5 minutes. Use 5 to 600 seconds."),
            ("Display Density", "display_density", "Comfortable is easier to read. Compact shows more rows on smaller screens."),
        ]
        for r, (label, key_name, tip) in enumerate(rows):
            grid_row = r + 1
            ttk.Label(form, text=label, style="Panel.TLabel").grid(row=grid_row, column=0, sticky="w", padx=6, pady=6)
            if key_name in ("auto_daily_backup", "frontdesk_scan_only_mode"):
                entry = ttk.Combobox(form, textvariable=vars[key_name], values=["Yes", "No"], state="readonly")
            elif key_name == "default_export_type":
                entry = ttk.Combobox(form, textvariable=vars[key_name], values=["All"] + ASSET_TYPES, state="readonly")
            elif key_name == "display_density":
                entry = ttk.Combobox(form, textvariable=vars[key_name], values=["Comfortable", "Compact", "Spacious"], state="readonly")
            elif key_name == "refresh_seconds":
                entry = ttk.Combobox(form, textvariable=vars[key_name], values=["60", "120", "300", "600"])
            else:
                entry = ttk.Entry(form, textvariable=vars[key_name])
            entry.grid(row=grid_row, column=1, sticky="ew", padx=6, pady=6)
            self.tip(entry, tip)
            if key_name.endswith("_folder"):
                path_entries.append((entry, key_name))
                ttk.Button(form, text="Browse", command=lambda k=key_name: choose_folder(k)).grid(row=grid_row, column=2, padx=3, pady=6)
                ttk.Button(form, text="Open", command=lambda k=key_name: open_folder(k)).grid(row=grid_row, column=3, padx=3, pady=6)
        form.grid_columnconfigure(1, weight=1)

        ttk.Label(pan.body, text="Allowed Asset Types:", style="Panel.TLabel").pack(anchor="w", padx=6, pady=(14,2))
        tk.Label(pan.body, text=", ".join(ASSET_TYPES), bg=self.colors["panel"], fg="#333333", anchor="w", justify="left", wraplength=650).pack(fill="x", padx=6)

        def choose_folder(key_name):
            folder = filedialog.askdirectory(title="Choose folder")
            if folder:
                vars[key_name].set(folder)

        def open_folder(key_name):
            folder = Path(clean(vars[key_name].get()) or self.app_dir)
            folder.mkdir(parents=True, exist_ok=True)
            if sys.platform.startswith("win"):
                os.startfile(folder)
            else:
                messagebox.showinfo("Folder", str(folder))

        folder_defaults = {
            "backup_folder": self.app_dir / "Backups",
            "excel_export_folder": self.app_dir / "Exports",
            "report_export_folder": self.app_dir / "Reports",
            "error_log_folder": self.app_dir / "Logs",
            "audit_log_folder": self.app_dir / "Logs",
            "system_log_folder": self.app_dir / "Logs",
        }

        def reset_defaults():
            for key_name, folder in folder_defaults.items():
                vars[key_name].set(str(folder))
            vars["refresh_seconds"].set("300")
            vars["frontdesk_scan_only_mode"].set("No")

        def reset_one_path(key_name):
            if key_name in folder_defaults:
                vars[key_name].set(str(folder_defaults[key_name]))

        def validate_folder(key_name):
            folder = Path(clean(vars[key_name].get()))
            if not folder:
                messagebox.showwarning("Validate Path", "Folder path is blank.")
                return
            if not folder.exists():
                if not messagebox.askyesno("Create Folder?", f"This folder does not exist:\n{folder}\n\nCreate it now?"):
                    return
                folder.mkdir(parents=True, exist_ok=True)
            test = folder / f".write_test_{uuid.uuid4().hex}.tmp"
            try:
                test.write_text("ok", encoding="utf-8")
                test.unlink(missing_ok=True)
                messagebox.showinfo("Validate Path", f"Folder is writable:\n{folder}")
                self.db.audit("VALIDATE SETTINGS PATH", self.actor(), key_name, str(folder))
            except Exception as e:
                messagebox.showerror("Validate Path", f"Folder is not writable:\n{folder}\n\n{e}")

        def show_path_menu(event, key_name):
            menu = tk.Menu(win, tearoff=0)
            menu.add_command(label="Browse Folder", command=lambda: choose_folder(key_name))
            menu.add_command(label="Open Folder", command=lambda: open_folder(key_name))
            menu.add_command(label="Reset This Path", command=lambda: reset_one_path(key_name))
            menu.add_command(label="Validate Path", command=lambda: validate_folder(key_name))
            menu.tk_popup(event.x_root, event.y_root)
            return "break"

        def show_data_file_menu(event):
            menu = tk.Menu(win, tearoff=0)
            menu.add_command(label="Change Data File Location", command=self.change_data_location)
            menu.add_command(label="Open Data Folder", command=self.open_data_folder)
            menu.add_command(label="Copy Data File Path", command=lambda: (self.clipboard_clear(), self.clipboard_append(data_file_var.get()), self.status.set("Data file path copied.")))
            menu.tk_popup(event.x_root, event.y_root)
            return "break"

        data_entry.bind("<Button-3>", show_data_file_menu, add="+")
        for entry, key_name in path_entries:
            entry.bind("<Button-3>", lambda e, k=key_name: show_path_menu(e, k), add="+")

        def save():
            due_value = clean(vars["default_due_time"].get())
            try:
                hour, minute = [int(x) for x in due_value.split(":", 1)]
                if not (0 <= hour <= 23 and 0 <= minute <= 59):
                    raise ValueError
            except Exception:
                messagebox.showwarning("Invalid Due Time", "Default due time must be HH:MM, like 18:30.")
                return
            refresh_value = clean(vars["refresh_seconds"].get())
            if not refresh_value.isdigit() or not (5 <= int(refresh_value) <= 600):
                messagebox.showwarning("Invalid Refresh Timer", "Refresh timer must be a number from 5 to 600 seconds.")
                return
            if vars["default_export_type"].get() not in (["All"] + ASSET_TYPES):
                messagebox.showwarning("Invalid Export Type", "Choose a valid default export type.")
                return
            if vars["display_density"].get() not in ("Comfortable", "Compact", "Spacious"):
                messagebox.showwarning("Invalid Display Density", "Choose Comfortable, Compact, or Spacious.")
                return
            if vars["frontdesk_scan_only_mode"].get() not in ("Yes", "No"):
                messagebox.showwarning("Invalid Front Desk Mode", "Choose Yes or No for Front Desk Scan-Only Mode.")
                return
            old_scan_only = self.db.setting("frontdesk_scan_only_mode", "No")
            if vars["frontdesk_scan_only_mode"].get() != old_scan_only and self.role() != "Admin":
                messagebox.showwarning("Admin Required", "Only Admin can change Front Desk Scan-Only Mode.")
                return
            for key_name, var in vars.items():
                if key_name.endswith("_folder"):
                    folder = Path(clean(var.get()))
                    if not folder.exists():
                        if not messagebox.askyesno("Create Folder?", f"This folder does not exist:\n{folder}\n\nCreate it now?"):
                            return
                        try:
                            folder.mkdir(parents=True, exist_ok=True)
                        except Exception as e:
                            messagebox.showerror("Folder Error", f"Could not create folder:\n{folder}\n\n{e}")
                            return
                self.db.set_setting(key_name, clean(var.get()))
            self.db.audit("UPDATE APP SETTINGS", self.actor(), "Settings", json.dumps({k: v.get() for k, v in vars.items()}))
            self.db.notify_manager("Settings path changed", "Info", self.actor(), "", "Settings saved", self.actor(), json.dumps({k: v.get() for k, v in vars.items()}))
            self.asset_export_type_var.set(vars["default_export_type"].get())
            self.next_auto_refresh_at = time.time() + (self.refresh_interval_ms() / 1000)
            self.style()
            self.update_operator_badges()
            self.refresh_scan_mode()
            self.status.set("Settings saved.")
            win.destroy()

        row = tk.Frame(pan.body, bg=self.colors["panel"])
        row.pack(fill="x", pady=14)
        ttk.Button(row, text="Reset to Defaults", command=reset_defaults).pack(side="left", padx=4)
        ttk.Button(row, text="Save Settings", style="Green.TButton", command=save).pack(side="left", padx=4)
        ttk.Button(row, text="Cancel", command=win.destroy).pack(side="left", padx=4)

    def health_check(self):
        c=self.db.counts()
        text=[f"{APP_TITLE} HEALTH CHECK", "="*80, f"Version: {APP_VERSION}", f"Database: {self.db.path}", f"Operator: {self.actor()} - {self.role()}", ""]
        for k,v in c.items(): text.append(f"{k:<20} {v}")
        self.system_text.delete("1.0","end"); self.system_text.insert("1.0","\n".join(text))

    def selftest_info(self):
        self.system_text.delete("1.0","end")
        self.system_text.insert("1.0","Close the app and run Run_Self_Test_No_GUI.bat to test database, checkout/return, and reports.")

    def export_system_report(self):
        self.health_check()
        path=self.folder_setting("system_log_folder", "Logs")/f"SYSTEM_REPORT_{stamp()}.txt"
        path.write_text(self.system_text.get("1.0","end"),encoding="utf-8")
        messagebox.showinfo("Saved", str(path))

    def search_audit(self):
        self.show("Search")
        if hasattr(self, "search_status_var"):
            self.search_status_var.set("Any Status")
        self.search_var.set("")
        if hasattr(self, "search_type_var"):
            self.set_search_mode("Audit")

    def clear_errors(self):
        if self.role()!="Admin":
            messagebox.showwarning("Admin Only","Only Admin can clear errors.")
            return
        if messagebox.askyesno("Clear Errors","Clear all error log entries?"):
            self.db.run("DELETE FROM errors")
            self.db.audit("CLEAR ERRORS", self.actor(), "Errors", "Error log cleared.")
            self.health_check()

    def about(self):
        key = "about"
        if self.focus_existing(key):
            return
        win = tk.Toplevel(self)
        win.title("About")
        win.geometry("760x560")
        self.apply_icon(win)
        self.track_window(key, win)
        pan = self.panel(win, "About Macy's AP", "Asset accountability, checkout tracking, reports, backups, and shared-data support.")
        pan.pack(fill="both", expand=True, padx=14, pady=14)
        text = tk.Text(pan.body, wrap="word", font=("Segoe UI", 10), bg="white", fg="#111", padx=12, pady=12, height=16)
        text.pack(fill="both", expand=True)
        lines = [
            f"{APP_TITLE}",
            f"Version: {APP_VERSION}",
            "",
            "Purpose",
            "Track AP keys, radios, temp badges, scanners, tablets, and other issued assets.",
            "",
            "Daily workflow",
            "1. Sign in as the operator.",
            "2. Open Front Desk.",
            "3. Scan the employee badge.",
            "4. Scan the asset barcode, key number, or serial number.",
            "5. Check out or return the asset.",
            "",
            "Shared data",
            "Use System > Change Data File Location to point every computer to the same shared .db file.",
            "A UNC path such as \\\\SERVER\\Share\\Macys_AP_Data\\macys_ap_data.db is preferred.",
            "",
            "Backups",
            "Automatic daily backup is on by default and can be changed in System > Settings.",
            "",
            "Roles",
            "Employee < Front Desk < AP Operator < Manager < Admin",
        ]
        text.insert("1.0", "\n".join(lines))
        text.configure(state="disabled")
        ttk.Button(pan.body, text="Close", command=win.destroy).pack(anchor="w", pady=(10,0))


    # -----------------------------
    # v5.1 additions: role test, scanner diagnostics, backup restore, import templates, delete asset
    # -----------------------------
    def v51_apply_enhancements(self):
        try:
            self.db.audit("V5.1 ENABLED", "System", "Startup", "Import/restore/diagnostics/delete asset tools enabled.")
        except Exception:
            pass

    def role_test(self):
        lines = [
            "ROLE / RIGHTS TEST",
            "=" * 80,
            f"Current operator : {self.actor()}",
            f"Current role     : {self.role()}",
            "",
            "Access:",
            f"Dashboard        : {'Yes' if self.has('dashboard') else 'No'}",
            f"Front Desk       : {'Yes' if self.has('frontdesk') else 'No'}",
            f"People view      : {'Yes' if self.has('people_view') else 'No'}",
            f"People edit      : {'Yes' if self.has('people_edit') else 'No'}",
            f"Assets view      : {'Yes' if self.has('assets_view') else 'No'}",
            f"Assets edit      : {'Yes' if self.has('assets_edit') else 'No'}",
            f"Search           : {'Yes' if self.has('search') else 'No'}",
            f"Manager          : {'Yes' if self.has('manager') else 'No'}",
            f"Reports          : {'Yes' if self.has('reports') else 'No'}",
            f"System           : {'Yes' if self.has('system') else 'No'}",
            f"Admin tools      : {'Yes' if self.has('admin') else 'No'}",
            "",
            "Role order:",
            "Employee < Front Desk < AP Operator < Manager < Admin",
            "",
            "Rules:",
            "- System access is controlled by role.",
            "- Admin and Manager can open System.",
            "- Managers cannot create/change/modify Admin users.",
            "- Only Admin can clear errors and perform highest-level admin changes.",
        ]
        self.system_text.delete("1.0", "end")
        self.system_text.insert("1.0", "\n".join(lines))
        self.db.audit("ROLE TEST", self.actor(), self.role(), "Role test viewed.")

    def scanner_diagnostics(self):
        key = "scanner_diag"
        if self.focus_existing(key):
            return
        win = tk.Toplevel(self)
        win.title("Scanner Diagnostics")
        win.geometry("820x520")
        self.apply_icon(win)
        self.track_window(key, win)
        pan = self.panel(win, "Scanner Diagnostics", "Scan or type anything below to see what the app detects.")
        pan.pack(fill="both", expand=True, padx=14, pady=14)
        scan_var = tk.StringVar()
        entry = ttk.Entry(pan.body, textvariable=scan_var, font=("Segoe UI", 18))
        entry.pack(fill="x", ipady=4, pady=(0, 8))
        out = tk.Text(pan.body, wrap="word", font=("Consolas", 11), bg="white", fg="#111")
        out.pack(fill="both", expand=True)

        def run_test(event=None):
            s = clean(scan_var.get())
            out.delete("1.0", "end")
            if not s:
                out.insert("1.0", "No scan entered.")
                return
            p = self.db.find_person(s)
            a = self.db.find_asset(s)
            lines = [f"Last scan: {s}", ""]
            if p:
                lines += [
                    "Detected as: Person / Employee",
                    f"Name       : {self.db.person_name(p)}",
                    f"Employee ID: {p['employee_id']}",
                    f"Badge      : {p['badge']}",
                    f"Role       : {p['role']}",
                    f"Status     : {p['status']}",
                    "Suggested action: scan an asset next for checkout.",
                ]
            elif a:
                open_log = self.db.open_log(a["barcode"])
                lines += [
                    "Detected as: Asset",
                    f"Type       : {a['asset_type']}",
                    f"Barcode    : {a['barcode']}",
                    f"Name       : {a['name']}",
                    f"Status     : {a['status']}",
                    f"Key number : {a['controlled_key_number']}",
                    f"Serial     : {a['serial_number']}",
                    f"Details    : {asset_detail_text(a['asset_type'], a['asset_details'] if 'asset_details' in a.keys() else '')}",
                    f"Holder     : {a['current_holder_name'] or 'None'}",
                    "Suggested action: return item." if open_log else "Suggested action: check out item after employee scan.",
                ]
            else:
                lines += ["Detected as: Unknown", "Suggested action: add this as a person or asset if needed."]
            out.insert("1.0", "\n".join(lines))
            scan_var.set("")
            entry.focus_set()

        entry.bind("<Return>", run_test)
        row = tk.Frame(pan.body, bg=self.colors["panel"])
        row.pack(fill="x", pady=8)
        ttk.Button(row, text="Test Scan", style="Green.TButton", command=run_test).pack(side="left", padx=4)
        ttk.Button(row, text="Close", command=win.destroy).pack(side="left", padx=4)
        entry.focus_set()

    def asset_import_template_rows(self):
        headers = [
            "asset_type","barcode","name","status","location","controlled_key_number","serial_number","notes",
            "key_set_number","key_ring_serial_number","number_of_keys","key_ring_location","key_ring_use","key_1_serial","key_1_access","key_2_serial","key_2_access",
            "radio_number","factory_serial_number","radio_serial_number","radio_location","assigned_area",
            "tablet_number","tablet_factory_serial_number","imei_number","windows_license_key","tablet_location","accessories","other_accessory",
            "badge_number","badge_location",
            "scanner_number","scanner_serial_number","scanner_location",
            "item_number","item_location","item_use",
            "asset_number","asset_location","asset_use",
        ]
        examples = [
            {
                "asset_type":"Key", "barcode":"KEY-001", "name":"Key Ring 001", "status":"Available", "location":"AP",
                "controlled_key_number":"001", "notes":"Example key", "key_set_number":"001",
                "key_ring_serial_number":"KEYRING001", "number_of_keys":"2", "key_ring_location":"AP",
                "key_ring_use":"AP office/control rooms", "key_1_serial":"K001", "key_1_access":"AP Office",
                "key_2_serial":"K002", "key_2_access":"Control Room",
            },
            {
                "asset_type":"Radio", "barcode":"RADIO-001", "name":"Radio 001", "status":"Available", "location":"AP",
                "serial_number":"RAD001", "notes":"Example radio", "radio_number":"RADIO-001",
                "factory_serial_number":"FAC-RAD001", "radio_serial_number":"RAD001", "radio_location":"AP",
                "assigned_area":"Floor AP",
            },
            {
                "asset_type":"Temp Badge", "barcode":"TEMP-001", "name":"Temp Badge 001", "status":"Available",
                "location":"Front Desk", "notes":"Example temp badge", "badge_number":"TEMP-001",
                "badge_location":"Front Desk", "assigned_area":"Front Desk",
            },
            {
                "asset_type":"Scanner", "barcode":"SCAN-001", "name":"Scanner 001", "status":"Available",
                "location":"Front Desk", "serial_number":"SCAN001", "notes":"Example scanner",
                "scanner_number":"SCAN-001", "scanner_serial_number":"SCAN001", "scanner_location":"Front Desk",
            },
            {
                "asset_type":"Tablet", "barcode":"TAB-001", "name":"Tablet 001", "status":"Available",
                "location":"AP", "serial_number":"TAB001", "notes":"Example tablet",
                "tablet_number":"TAB-001", "tablet_factory_serial_number":"TAB-FAC001", "imei_number":"IMEI001",
                "windows_license_key":"WIN-KEY", "tablet_location":"AP", "assigned_area":"AP",
                "accessories":"Tablet cover;Charging cable",
            },
            {
                "asset_type":"Item", "barcode":"ITEM-001", "name":"Equipment Pouch 001", "status":"Available",
                "location":"Front Desk", "notes":"Example item", "item_number":"ITEM-001",
                "item_location":"Front Desk", "item_use":"Equipment storage",
            },
            {
                "asset_type":"Other", "barcode":"OTHER-001", "name":"Other Asset 001", "status":"Available",
                "location":"AP", "notes":"Example other asset", "asset_number":"OTHER-001",
                "asset_location":"AP", "asset_use":"Special asset",
            },
        ]
        return [headers] + [[example.get(header, "") for header in headers] for example in examples]


    def export_people_import_template(self):
        folder = filedialog.askdirectory(title="Choose folder to save People import template")
        if not folder:
            return
        path = Path(folder) / "people_import_template.csv"
        rows = [
            ["employee_id","badge","first_name","last_name","department","status","shift","role","notes"],
            ["F984717","88984717","Christopher","Schumacher","Asset Protection","Active","Default","Admin","Example row - keep Admin only if importing as Admin"],
            ["F1001","881001","Jane","Doe","Receiving","Active","Day","Employee","Example regular employee"],
            ["F1002","881002","Alex","Frontdesk","Asset Protection","Active","Day","Front Desk","Example front desk user"],
            ["F1003","881003","John","Smith","Receiving","Active","Night","AP Operator","Example operator"],
        ]
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerows(rows)
        messagebox.showinfo("People Template Saved", f"People import template saved to:\n{path}\n\nFill this out, then use System > Import CSV.")
        self.db.audit("EXPORT PEOPLE TEMPLATE", self.actor(), "People Template", str(path))

    def export_asset_import_template(self):
        folder = filedialog.askdirectory(title="Choose folder to save Asset import template")
        if not folder:
            return
        path = Path(folder) / "assets_import_template.csv"
        rows = self.asset_import_template_rows()
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerows(rows)
        messagebox.showinfo("Asset Template Saved", f"Asset import template saved to:\n{path}\n\nFill this out, then use System > Import CSV.")
        self.db.audit("EXPORT ASSET TEMPLATE", self.actor(), "Asset Template", str(path))


    def export_import_templates(self):
        folder = filedialog.askdirectory(title="Choose folder to save import templates")
        if not folder:
            return
        folder = Path(folder)
        people_rows = [
            ["employee_id","badge","first_name","last_name","department","status","shift","role","notes"],
            ["88984717","F984717","Christopher","Schumacher","Asset Protection","Active","Default","Admin","Example row"],
            ["1001","BADGE1001","Jane","Doe","Asset Protection","Active","Day","Front Desk","Example employee"],
        ]
        asset_rows = self.asset_import_template_rows()
        def write(path, rows):
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerows(rows)
        write(folder / "people_import_template.csv", people_rows)
        write(folder / "assets_import_template.csv", asset_rows)
        messagebox.showinfo("Templates Saved", f"Import templates saved to:\n{folder}")
        self.db.audit("EXPORT IMPORT TEMPLATES", self.actor(), "Templates", str(folder))

    def preview_import_rows(self, headers, rows):
        addable = 0
        skipped = 0
        preview_errors = []
        kind = "Unknown"
        if "asset_type" in headers and "barcode" in headers:
            kind = "Assets"
            for i, row in enumerate(rows, start=2):
                d = self.normalize_asset_record(row)
                problems = self.validate_asset_record(d)
                if problems:
                    skipped += 1
                    preview_errors.append(f"Row {i}: {'; '.join(problems)}")
                else:
                    addable += 1
        elif "employee_id" in headers and "badge" in headers:
            kind = "People"
            for i, row in enumerate(rows, start=2):
                badge = clean(row.get("badge"))
                empid = clean(row.get("employee_id"))
                if not badge or not empid or self.db.find_person(badge) or self.db.find_person(empid):
                    skipped += 1
                    preview_errors.append(f"Row {i}: missing ID/badge or already exists.")
                else:
                    addable += 1
        else:
            return kind, "CSV does not match the people or assets template headers.", False
        lines = [
            f"Import type: {kind}",
            f"Rows found: {len(rows)}",
            f"Ready to add: {addable}",
            f"Will skip: {skipped}",
        ]
        if preview_errors:
            lines += ["", "First issues:"] + preview_errors[:8]
        lines += ["", "Continue with this import?"]
        return kind, "\n".join(lines), addable > 0

    def import_from_template(self):
        if not self.require("system"):
            return
        path = filedialog.askopenfilename(title="Choose people or assets CSV import file", filetypes=[("CSV files","*.csv"),("All files","*.*")])
        if not path:
            return
        path = Path(path)
        added = 0
        skipped = 0
        errors = []
        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                headers = [h.strip() for h in (reader.fieldnames or [])]
                rows = list(reader)
                kind, preview, can_import = self.preview_import_rows(headers, rows)
                if not can_import:
                    messagebox.showwarning("Import Preview", preview)
                    return
                if not messagebox.askyesno("Import Preview", preview):
                    return
                if "asset_type" in headers and "barcode" in headers:
                    for i, row in enumerate(rows, start=2):
                        try:
                            d = self.normalize_asset_record(row)
                            problems = self.validate_asset_record(d)
                            if problems:
                                skipped += 1
                                errors.append(f"Row {i}: {'; '.join(problems)}")
                                continue
                            self.db.add_asset(d, self.actor())
                            added += 1
                        except Exception as e:
                            errors.append(f"Row {i}: {e}")
                elif "employee_id" in headers and "badge" in headers:
                    for i, row in enumerate(rows, start=2):
                        try:
                            badge = clean(row.get("badge"))
                            empid = clean(row.get("employee_id"))
                            if not badge or not empid or self.db.find_person(badge) or self.db.find_person(empid):
                                skipped += 1
                                continue
                            role = clean(row.get("role")) or "Employee"
                            d = {
                                "employee_id": empid,
                                "badge": badge,
                                "first_name": clean(row.get("first_name")),
                                "last_name": clean(row.get("last_name")),
                                "department": clean(row.get("department")),
                                "status": clean(row.get("status")) or "Active",
                                "shift": clean(row.get("shift")),
                                "role": role if role in self.db.group_names() else "Employee",
                                "notes": clean(row.get("notes")),
                            }
                            d = self.normalize_person_id_badge(d, ask=False)
                            if not d["badge"].startswith("88") or not d["employee_id"].startswith("F"):
                                skipped += 1
                                errors.append(f"Row {i}: skipped. Badge must start with 88 and Employee ID must start with F.")
                                continue
                            if d["role"] == "Admin" and self.role() != "Admin":
                                skipped += 1
                                errors.append(f"Row {i}: Admin role skipped. Only Admin can import Admin users.")
                                continue
                            self.db.add_person(d, self.actor())
                            added += 1
                        except Exception as e:
                            errors.append(f"Row {i}: {e}")
                else:
                    messagebox.showerror("Import Failed", "CSV does not match the people or assets template headers.")
                    return
            msg = f"Import complete.\n\nAdded: {added}\nSkipped: {skipped}\nErrors: {len(errors)}"
            if errors:
                msg += "\n\nFirst errors:\n" + "\n".join(errors[:8])
            messagebox.showinfo("Import Complete", msg)
            self.db.audit("IMPORT CSV", self.actor(), str(path), msg)
            self.refresh()
        except Exception:
            self.log_exception("Import From Template")

    def restore_backup(self):
        if not self.require("system"):
            return
        path = filedialog.askopenfilename(title="Choose backup ZIP or database to restore", filetypes=[("Backup ZIP or DB","*.zip *.db"),("All files","*.*")])
        if not path:
            return
        path = Path(path)
        if not messagebox.askyesno("Restore Backup", "This will replace the current database.\n\nThe current database will be backed up first.\n\nContinue?"):
            return
        temp = None
        try:
            # backup current DB first
            backup_folder = Path(self.db.setting("backup_folder", str(self.app_dir / "Backups")))
            backup_folder.mkdir(parents=True, exist_ok=True)
            current_backup = backup_folder / f"PRE_RESTORE_CURRENT_DB_{stamp()}.db"
            current_db = self.current_data_file()
            self.db.conn.commit()
            shutil.copy2(current_db, current_backup)

            restore_db = None
            temp = self.app_dir / f"_restore_temp_{stamp()}"
            if path.suffix.lower() == ".zip":
                temp.mkdir(exist_ok=True)
                with zipfile.ZipFile(path) as z:
                    z.extractall(temp)
                dbs = list(temp.rglob("*.db"))
                if dbs:
                    restore_db = dbs[0]
            elif path.suffix.lower() == ".db":
                restore_db = path
            if not restore_db or not restore_db.exists():
                messagebox.showerror("Restore Failed", "No .db file found in selected backup.")
                return

            self.db.audit("RESTORE BACKUP START", self.actor(), str(path), f"Restoring database from: {restore_db}; Current database backup: {current_backup}")
            self.db.conn.close()
            shutil.copy2(restore_db, current_db)
            messagebox.showinfo("Restore Complete", f"Backup restored to:\n{current_db}\n\nPrevious database backup saved to:\n{current_backup}\n\nRestart the app now.")
            self.destroy()
        except Exception:
            try:
                self.db = DB(self.app_dir / DB_FILE)
            except Exception:
                pass
            self.log_exception("Restore Backup")
        finally:
            if temp:
                shutil.rmtree(temp, ignore_errors=True)

    def delete_selected_asset(self):
        if not self.require("assets_edit"):
            return
        sel = self.asset_tree.selection()
        if not sel:
            messagebox.showinfo("Delete Asset", "Select an asset first.")
            return
        barcode = self.asset_tree.item(sel[0], "values")[1]
        a = self.db.find_asset(barcode)
        if not a:
            return
        if self.db.open_log(a["barcode"]):
            messagebox.showwarning("Cannot Delete", "This asset is currently checked out. Return it before deleting.")
            return
        count = self.db.one("SELECT COUNT(*) c FROM activity WHERE lower(asset_barcode)=lower(?)", (a["barcode"],))["c"]
        if count:
            if not messagebox.askyesno("Retire Asset", f"This asset has history, so it will be marked Retired instead of permanently deleted.\n\nRetire {a['barcode']}?"):
                return
            self.db.run("UPDATE assets SET status='Retired', updated_at=? WHERE id=?", (now_iso(), a["id"]))
            self.db.audit("RETIRE ASSET", self.actor(), a["barcode"], "Asset retired because history exists.")
        else:
            if not messagebox.askyesno("Delete Asset", f"Permanently delete asset {a['barcode']}?\n\nNo history exists."):
                return
            self.db.run("DELETE FROM assets WHERE id=?", (a["id"],))
            self.db.audit("DELETE ASSET", self.actor(), a["barcode"], "Asset permanently deleted. No history existed.")
        self.refresh_assets()

    def patch_assets_delete_button(self):
        # v5.1.1: button is built directly on Assets toolbar. Do not add a duplicate.
        self._v51_delete_button_added = True
        return

    def tk_error(self, exc, val, tb):
        details="".join(traceback.format_exception(exc,val,tb))
        try: self.db.error("Tkinter Runtime",str(val),details)
        except Exception: pass
        messagebox.showerror("Application Error", f"{val}\n\nSaved to Error Log.")

    def log_exception(self, source):
        details=traceback.format_exc()
        try: self.db.error(source,"Exception",details)
        except Exception: pass
        messagebox.showerror("Error", f"An error occurred in {source}.")

def run_self_test():
    test_dir = Path(__file__).resolve().parent
    test_db = test_dir / "SELF_TEST_v5.db"
    if test_db.exists(): test_db.unlink()
    results = []
    def ok(name, cond):
        results.append((name, bool(cond)))
    try:
        db=DB(test_db)
        ok("create_database", test_db.exists())
        admin=db.find_person("F984717")
        ok("default_admin", admin and admin["role"]=="Admin")
        ok("groups_seeded", db.one("SELECT COUNT(*) c FROM groups")["c"] >= 5 and "admin" in {x.lower() for x in db.group_rights("Admin")})
        a=db.find_asset("ITEM-001")
        ok("default_asset", bool(a))
        mapped_key = apply_asset_type_field_map({"asset_type":"Key","controlled_key_number":"K-900","serial_number":"RAD-900"})
        mapped_radio = apply_asset_type_field_map({"asset_type":"Radio","controlled_key_number":"K-901","serial_number":"RAD-901"})
        mapped_badge = apply_asset_type_field_map({"asset_type":"Temp Badge","controlled_key_number":"K-902","serial_number":"RAD-902"})
        ok("asset_mapping_key", mapped_key["controlled_key_number"]=="K-900" and mapped_key["serial_number"]=="")
        ok("asset_mapping_radio", mapped_radio["controlled_key_number"]=="" and mapped_radio["serial_number"]=="RAD-901")
        ok("asset_mapping_badge", mapped_badge["controlled_key_number"]=="" and mapped_badge["serial_number"]=="")
        test_asset = {
            "asset_type": "Tablet",
            "barcode": "TAB-SELFTEST",
            "name": "Tablet Self Test",
            "status": "Available",
            "location": "AP",
            "controlled_key_number": "",
            "serial_number": "TAB-SELFTEST-SN",
            "asset_details": asset_details_to_json({"tablet_number":"TAB-SELFTEST","tablet_factory_serial_number":"TAB-SELFTEST-SN","imei_number":"123456789012345","tablet_location":"AP","assigned_area":"Self Test","accessories":["Tablet cover","Charging cable"]}),
            "notes": "Self-test asset workflow",
        }
        db.add_asset(test_asset, "Self Test")
        ok("asset_add_reload", db.find_asset("TAB-SELFTEST")["serial_number"] == "TAB-SELFTEST-SN")
        ok("asset_details_reload", parse_asset_details(db.find_asset("TAB-SELFTEST")["asset_details"])["tablet_number"] == "TAB-SELFTEST")
        log=db.checkout(admin,a,plus_hours(1),"Self Test")
        ok("checkout", bool(db.open_log("ITEM-001")))
        ok("double_checkout_guard", db.checkout(admin,a,plus_hours(1),"Self Test") is None)
        db.return_asset(db.open_log("ITEM-001"),"Self Test")
        ok("return", not db.open_log("ITEM-001"))
        key_asset = db.find_asset("KEY-001")
        d = {k: key_asset[k] for k in ASSET_DB_FIELDS if k in key_asset.keys()}
        d["status"] = "Repair"
        db.update_asset(key_asset["id"], d, "Self Test")
        ok("asset_status_update", db.find_asset("KEY-001")["status"] == "Repair")
        test_asset_row = db.find_asset("TAB-SELFTEST")
        edited = {k: test_asset_row[k] for k in ASSET_DB_FIELDS if k in test_asset_row.keys()}
        edited["location"] = "Manager Office"
        db.update_asset(test_asset_row["id"], edited, "Self Test")
        ok("asset_edit_reload", db.find_asset("TAB-SELFTEST")["location"] == "Manager Office")
        xlsx_path = test_dir / "SELF_TEST_ASSET_EXPORT.xlsx"
        if xlsx_path.exists():
            xlsx_path.unlink()
        App.write_xlsx(None, xlsx_path, [("Assets", ["Barcode", "Name"], [["TAB-SELFTEST", "Tablet Self Test"]])])
        ok("excel_export_write", xlsx_path.exists() and xlsx_path.stat().st_size > 1000)
        xlsx_path.unlink(missing_ok=True)
        counts = db.counts()
        ok("counts", counts["assets"] >= 4 and counts["inactive_assets"] >= 1)
        ok("dashboard_type_counts", counts["keys_total"] >= 1 and counts["radios_total"] >= 1 and counts["items_total"] >= 1)
        db.notify_manager("Self Test Notification", "Info", "Self Test", "TAB-SELFTEST", "Validate manager notification table", "Self Test", "OK", status="Reviewed")
        ok("manager_notifications", db.one("SELECT COUNT(*) c FROM manager_notifications")["c"] >= 1)
        db.add_ap_alert("Person", "F984717", "Christopher Schumacher | F984717", "Security note", "Warning", "Self-test AP alert", "Review before checkout", "Self Test")
        alert_row = db.one("SELECT * FROM ap_alerts WHERE target_key='F984717' ORDER BY id DESC LIMIT 1")
        ok("ap_alert_saved", bool(alert_row) and db.counts()["manager_alerts"] >= 1)
        db.update_ap_alert_status(alert_row["id"], "Resolved", "Self Test", "Self-test cleanup")
        ok("ap_alert_resolved", db.one("SELECT status FROM ap_alerts WHERE id=?", (alert_row["id"],))["status"] == "Resolved")
        db.audit("TEST","Self Test","Audit","OK")
        db.error("TEST","Error","OK")
        ok("audit_error", db.one("SELECT COUNT(*) c FROM audit")["c"] > 0 and db.one("SELECT COUNT(*) c FROM errors")["c"] > 0)
        rich_audit = db.one("SELECT actor_role,page,status,computer_name FROM audit WHERE action='TEST' ORDER BY id DESC LIMIT 1")
        rich_error = db.one("SELECT actor_role,page,status,computer_name FROM errors WHERE source='TEST' ORDER BY id DESC LIMIT 1")
        ok("rich_log_fields", rich_audit["status"] == "Success" and rich_error["status"] == "Error" and "actor_role" in rich_audit.keys())
        db.conn.close()
        test_db.unlink(missing_ok=True)
        ok("cleanup", not test_db.exists())
    except Exception:
        results.append(("exception", False))
    passed = all(v for _,v in results)
    txt = ["Macy's AP v5 Self Test", "="*72, f"Overall: {'PASS' if passed else 'FAIL'}", ""]
    for name, value in results:
        txt.append(("PASS  " if value else "FAIL  ") + name)
    (test_dir/"SELF_TEST_RESULTS.txt").write_text("\n".join(txt), encoding="utf-8")
    (test_dir/"SELF_TEST_RESULTS.json").write_text(json.dumps({"passed":passed,"results":results}, indent=2), encoding="utf-8")
    print("\n".join(txt))
    return 0 if passed else 1

def main():
    if "--self-test" in sys.argv:
        raise SystemExit(run_self_test())
    app=App()
    app.mainloop()

if __name__ == "__main__":
    main()
