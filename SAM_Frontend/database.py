"""database.py: SAM's database, kept in an Excel workbook for now.

The team plan says "for now, an Excel sheet" with each student's user ID,
password and For you info. This is that sheet: data/sam_database.xlsx

    Users      One row per student: how they sign in, plus their survey answers
    Activity   Which websites students open and what subjects they ask SAM about.
               Used to rank the Popular near you page.
    Resources  Links students share on the Resources page
    Saved      Places each student saved to My plan, with the day they plan to go
    Languages  Made-up programming languages students create in the Code translator

Passwords are never saved. The "Password (hashed)" column holds a one-way
PBKDF2 hash, so nobody who opens the file can read a student's password.

You can open the file in Excel to look, but don't edit it while SAM is running:
SAM rewrites the whole file each time it saves. If you change it while SAM is
stopped, SAM reads your changes the next time it starts.

Online (for example on Vercel), files don't last: each request can run on a
fresh server. So when DATABASE_URL (or POSTGRES_URL) is set, the same sheets are
kept in Postgres instead, one table row per sheet row. The rest of SAM calls only
the functions at the bottom of this file, so it works the same either way.
"""

import json
import os
import re
import secrets
import threading
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font, PatternFill

# On Vercel (which sets VERCEL=1) only /tmp can be written to, and it's wiped whenever Vercel
# restarts SAM. With DATABASE_URL set, student data lives in Postgres and only the map cache
# goes in /tmp. Without it, accounts there last only until the next restart.
_DEFAULT_DATA_DIR = Path("/tmp/sam-data") if os.environ.get("VERCEL") else Path(__file__).resolve().parent / "data"
DATA_DIR = Path(os.environ.get("SAM_DATA_DIR") or _DEFAULT_DATA_DIR)
DB_PATH = DATA_DIR / "sam_database.xlsx"
MAX_TEXT = 2000  # Longest text SAM stores in one cell

# Each sheet's columns: (key used in code, header shown in Excel, column width)
SHEETS = {
    "Users": [
        ("user_id", "User ID", 10),
        ("username", "Username", 16),
        ("name", "First name", 14),
        ("email", "Email", 30),
        ("password_hash", "Password (hashed)", 22),
        ("provider", "Signs in with", 15),
        ("provider_id", "Google ID", 22),
        ("created", "Created", 17),
        ("last_sign_in", "Last sign-in", 17),
        ("age", "Age", 6),
        ("loc", "Where they live", 18),
        ("gender", "Gender", 12),
        ("plays_sports", "Interested in sports", 11),
        ("sports", "Sports", 30),
        ("sport_other", "Other sport", 16),
        ("vol", "Volunteering", 34),
        ("vol_other", "Other volunteering", 18),
        ("survey_saved", "Survey saved", 17),
        ("radius", "Travel distance (miles)", 12),
        ("lat", "Latitude", 11),
        ("lon", "Longitude", 11),
        ("place", "Place on the map", 22),
    ],
    "Activity": [
        ("time", "Time", 17),
        ("user_id", "User ID", 10),
        ("area", "Area", 14),
        ("category", "Category", 11),
        ("topic", "Topic", 22),
        ("title", "Title", 32),
        ("url", "Link", 44),
    ],
    "Resources": [
        ("posted", "Posted", 17),
        ("url", "Link", 44),
        ("posted_by", "Posted by", 16),
        ("age", "Age", 6),
        ("user_id", "User ID", 10),
    ],
    "Saved": [
        ("saved", "Saved", 17),
        ("user_id", "User ID", 10),
        ("kind", "Kind", 8),
        ("item_id", "Item ID", 20),
        ("title", "Title", 32),
        ("detail", "Detail", 24),
        ("address", "Address", 34),
        ("url", "Link", 44),
        ("lat", "Latitude", 11),
        ("lon", "Longitude", 11),
        ("date", "Planned day", 13),
        ("time", "Planned time", 12),
    ],
    "Languages": [
        ("saved", "Saved", 17),
        ("user_id", "User ID", 10),
        ("name", "Language", 18),
        ("spec", "Keywords (JSON)", 60),
    ],
}

def _find_postgres_url(env):
    """The Postgres address from Vercel's Neon or Supabase integration, or "" when there isn't one.

    The integrations set DATABASE_URL (or POSTGRES_URL). When the database was connected with a
    custom prefix, the names start with it instead, for example SAM_DATA_DATABASE_URL.
    """
    for name in ("DATABASE_URL", "POSTGRES_URL"):
        if (env.get(name) or "").strip():
            return env[name].strip()
    for suffix in ("_DATABASE_URL", "_POSTGRES_URL"):
        for name in sorted(env):
            if name.endswith(suffix) and (env[name] or "").strip():
                return env[name].strip()
    return ""


# Postgres, when one is set up. Vercel's Neon and Supabase integrations set these for you.
POSTGRES_URL = _find_postgres_url(os.environ)
USING_POSTGRES = bool(POSTGRES_URL)

# Hold `lock` while checking something and then writing, so two requests can't both pass the check.
lock = threading.RLock()
_cache = {"wb": None, "mtime": None}


# ---------- Reading and writing the workbook ----------

def _header(sheet, key):
    return next(h for k, h, _ in SHEETS[sheet] if k == key)


def _set_up(wb):
    """Make sure every sheet and column exists, with a bold, frozen header row."""
    for sheet, columns in SHEETS.items():
        if sheet not in wb.sheetnames:
            wb.create_sheet(sheet)
        ws = wb[sheet]
        have = {c.value for c in ws[1] if c.value}
        next_col = len([c for c in ws[1] if c.value]) + 1
        for _, header, width in columns:
            if header in have:
                continue
            cell = ws.cell(row=1, column=next_col, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="0F766E")
            ws.column_dimensions[cell.column_letter].width = width
            next_col += 1
        ws.freeze_panes = "A2"
    # A new workbook starts with an empty "Sheet"; drop it.
    if "Sheet" in wb.sheetnames and "Sheet" not in SHEETS and wb["Sheet"].max_row == 1 and wb["Sheet"]["A1"].value is None:
        del wb["Sheet"]


def _open():
    """The workbook, read again if the file changed since SAM last read it."""
    mtime = DB_PATH.stat().st_mtime_ns if DB_PATH.exists() else None
    if _cache["wb"] is None or mtime != _cache["mtime"]:
        wb = load_workbook(DB_PATH) if mtime is not None else Workbook()
        _set_up(wb)
        _cache.update(wb=wb, mtime=mtime)
    return _cache["wb"]


def _save(wb):
    """Write to a temporary file first, so a crash mid-save can't corrupt the database."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    tmp = DB_PATH.with_name(DB_PATH.stem + ".saving.xlsx")
    wb.save(tmp)
    os.replace(tmp, DB_PATH)
    _cache["mtime"] = DB_PATH.stat().st_mtime_ns


def _columns(ws):
    return {c.value: c.column for c in ws[1] if c.value}


def _put(cell, value):
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub("", value).strip()[:MAX_TEXT]
        cell.value = value or None
        if value:
            # Always text, never a formula, even if a student types "=SUM(...)".
            cell.data_type = "s"
    else:
        cell.value = value
        if isinstance(value, datetime):
            cell.number_format = "yyyy-mm-dd hh:mm"


def _row_dict(sheet, values, cols):
    row = {}
    for key, header, _ in SHEETS[sheet]:
        i = cols.get(header)
        value = values[i - 1] if i and i - 1 < len(values) else None
        row[key] = "" if value is None else value
    return row


def now():
    return datetime.now().replace(microsecond=0)


# ---------- The Excel workbook (on your own computer) ----------

def _xl_rows(sheet):
    """Every row of a sheet, as dicts keyed by the names in SHEETS."""
    with lock:
        ws = _open()[sheet]
        cols = _columns(ws)
        return [
            _row_dict(sheet, values, cols)
            for values in ws.iter_rows(min_row=2, values_only=True)
            if any(v not in (None, "") for v in values)
        ]


def _xl_add_row(sheet, fields):
    with lock:
        wb = _open()
        ws = wb[sheet]
        cols = _columns(ws)
        r = ws.max_row + 1
        for key, value in fields.items():
            _put(ws.cell(row=r, column=cols[_header(sheet, key)]), value)
        _save(wb)


def _matching_rows(ws, cols, sheet, where):
    """Row numbers whose columns equal every value in `where`, for example {"user_id": "U0001"}."""
    wanted = [(cols[_header(sheet, key)], str(value)) for key, value in where.items()]
    return [r for r in range(2, ws.max_row + 1)
            if all(str(ws.cell(row=r, column=col).value or "") == value for col, value in wanted)]


def _xl_update_rows(sheet, where, fields):
    """Change every row that matches `where`. Returns how many rows changed."""
    with lock:
        wb = _open()
        ws = wb[sheet]
        cols = _columns(ws)
        found = _matching_rows(ws, cols, sheet, where)
        for r in found:
            for key, value in fields.items():
                _put(ws.cell(row=r, column=cols[_header(sheet, key)]), value)
        if found:
            _save(wb)
        return len(found)


def _xl_delete_rows(sheet, where):
    """Remove every row that matches `where`. Returns how many rows were removed."""
    with lock:
        wb = _open()
        ws = wb[sheet]
        found = _matching_rows(ws, _columns(ws), sheet, where)
        for r in reversed(found):  # Bottom up, so the row numbers above stay right
            ws.delete_rows(r)
        if found:
            _save(wb)
        return len(found)


# ---------- Postgres (online) ----------
# One table holds every sheet: each row is (sheet name, the row's fields as JSON).
# Dates are stored as {"$dt": "2026-11-27T14:05:00"} and come back as datetimes, as from Excel.

_pg_ready = False


def _pg():
    """A new connection. Serverless functions shouldn't hold one open between requests."""
    import psycopg  # Only needed online, so SAM still runs on a laptop without it
    return psycopg.connect(POSTGRES_URL, autocommit=True, connect_timeout=10, prepare_threshold=None)


def _pg_set_up():
    global _pg_ready
    if _pg_ready:
        return
    with _pg() as conn:
        conn.execute("""CREATE TABLE IF NOT EXISTS sam_rows (
            id BIGSERIAL PRIMARY KEY,
            sheet TEXT NOT NULL,
            data JSONB NOT NULL)""")
        conn.execute("CREATE INDEX IF NOT EXISTS sam_rows_sheet ON sam_rows (sheet, id)")
        conn.execute("CREATE TABLE IF NOT EXISTS sam_settings (name TEXT PRIMARY KEY, value TEXT NOT NULL)")
    _pg_ready = True


def _clean(value):
    """The same rules as an Excel cell: no control characters, trimmed, at most MAX_TEXT long."""
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub("", value).strip()[:MAX_TEXT]
        return value or None
    if isinstance(value, datetime):
        return {"$dt": value.isoformat()}
    return value


def _unclean(value):
    if isinstance(value, dict) and "$dt" in value:
        return datetime.fromisoformat(value["$dt"])
    return "" if value is None else value


def _check_keys(sheet, keys):
    known = {k for k, _, _ in SHEETS[sheet]}
    for key in keys:
        if key not in known:
            raise KeyError(f"{sheet} has no column {key!r}")


def _where_sql(sheet, where):
    """WHERE text and values matching every value in `where`, compared as text like the Excel version."""
    _check_keys(sheet, where)
    sql, args = "sheet = %s", [sheet]
    for key, value in where.items():
        sql += " AND COALESCE(data->>%s, '') = %s"
        args += [key, str(value)]
    return sql, args


def _pg_rows(sheet):
    _pg_set_up()
    with _pg() as conn:
        found = conn.execute("SELECT data FROM sam_rows WHERE sheet = %s ORDER BY id", [sheet]).fetchall()
    return [{key: _unclean(data.get(key)) for key, _, _ in SHEETS[sheet]} for (data,) in found]


def _pg_add_row(sheet, fields, conn=None):
    _check_keys(sheet, fields)
    data = json.dumps({key: _clean(value) for key, value in fields.items()})
    if conn is not None:
        conn.execute("INSERT INTO sam_rows (sheet, data) VALUES (%s, %s::jsonb)", [sheet, data])
        return
    _pg_set_up()
    with _pg() as c:
        c.execute("INSERT INTO sam_rows (sheet, data) VALUES (%s, %s::jsonb)", [sheet, data])


def _pg_update_rows(sheet, where, fields):
    _check_keys(sheet, fields)
    _pg_set_up()
    sql, args = _where_sql(sheet, where)
    patch = json.dumps({key: _clean(value) for key, value in fields.items()})
    with _pg() as conn:
        return conn.execute(f"UPDATE sam_rows SET data = data || %s::jsonb WHERE {sql}", [patch, *args]).rowcount


def _pg_delete_rows(sheet, where):
    _pg_set_up()
    sql, args = _where_sql(sheet, where)
    with _pg() as conn:
        return conn.execute(f"DELETE FROM sam_rows WHERE {sql}", args).rowcount


def _pg_add_user(fields):
    """Pick the next user ID and add the row in one transaction. The lock stops two
    servers handing out the same ID at the same moment."""
    _pg_set_up()
    with _pg() as conn:
        with conn.transaction():
            conn.execute("SELECT pg_advisory_xact_lock(4515001)")
            found = conn.execute(
                "SELECT data->>'user_id' FROM sam_rows WHERE sheet = 'Users'").fetchall()
            numbers = [int(m.group(1)) for (uid,) in found if (m := re.fullmatch(r"U(\d+)", uid or ""))]
            user_id = f"U{max(numbers, default=0) + 1:04d}"
            _pg_add_row("Users", {**fields, "user_id": user_id, "created": now()}, conn=conn)
    return user_id


def shared_secret(name="session_key"):
    """A random secret every server shares, made the first time it's asked for. Online only."""
    _pg_set_up()
    with _pg() as conn:
        conn.execute("INSERT INTO sam_settings (name, value) VALUES (%s, %s) ON CONFLICT (name) DO NOTHING",
                     [name, secrets.token_hex(32)])
        return conn.execute("SELECT value FROM sam_settings WHERE name = %s", [name]).fetchone()[0]


# ---------- What the rest of SAM calls ----------

def rows(sheet):
    """Every row of a sheet, as dicts keyed by the names in SHEETS."""
    return _pg_rows(sheet) if USING_POSTGRES else _xl_rows(sheet)


def add_row(sheet, fields):
    return _pg_add_row(sheet, fields) if USING_POSTGRES else _xl_add_row(sheet, fields)


def update_rows(sheet, where, fields):
    """Change every row that matches `where`. Returns how many rows changed."""
    return _pg_update_rows(sheet, where, fields) if USING_POSTGRES else _xl_update_rows(sheet, where, fields)


def update_row(sheet, key, match, fields):
    """Change the rows whose `key` column equals `match`. Returns False if none does."""
    return update_rows(sheet, {key: match}, fields) > 0


def delete_rows(sheet, where):
    """Remove every row that matches `where`. Returns how many rows were removed."""
    return _pg_delete_rows(sheet, where) if USING_POSTGRES else _xl_delete_rows(sheet, where)


def add_user(fields):
    """Add a student and return their new user ID (U0001, U0002, ...)."""
    with lock:
        if USING_POSTGRES:
            return _pg_add_user(fields)
        numbers = [int(m.group(1)) for u in rows("Users") if (m := re.fullmatch(r"U(\d+)", str(u["user_id"])))]
        user_id = f"U{max(numbers, default=0) + 1:04d}"
        add_row("Users", {**fields, "user_id": user_id, "created": now()})
        return user_id
