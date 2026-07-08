"""
Data loading + helpers for the Athlete Dashboard.
Ported from the original R/Shiny app.

IMPORTANT: the master workbook is formula-driven — athlete-name columns are
`=Bio!A2` style links and some metrics (e.g. Arm Care Total Score) are formulas.
We therefore NEVER write to the master file. New data entered in the app is
appended to append-only CSVs under entries/ and merged in here at load time.
"""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd

import sheets

# Master workbook + companion entries folder (both next to this file).
EXCEL_PATH = Path(__file__).parent / "Master Evaluation Sheet.xlsx"
ENTRIES_DIR = Path(__file__).parent / "entries"

# Map each sheet to the column used to identify an athlete.
_SHEETS = {
    "bio": ("Bio", "Player Name"),
    "motor": ("Motor Preferences", "Athlete"),
    "power": ("Power Testing", "Name"),
    "pitching": ("Pitching", "Athlete"),
    "armcare": ("Arm Care", "Athlete"),
    "context": ("Context", "Athlete"),
    "injuries": ("Injuries", "Athlete"),
    "mss": ("MSSPosture", "Athlete"),
    "plan": ("Athlete Plan", "Athlete"),
    "notes": ("Coaches Notes", "Name"),
}

# Long-format sheets that accept appended entries: key -> (csv name, id column).
ENTRY_SHEETS = {
    "power": ("power.csv", "Name"),
    "armcare": ("armcare.csv", "Athlete"),
    "pitching": ("pitching.csv", "Athlete"),
}
NOTES_CSV = "notes.csv"

# When backed by Google Sheets, appended entries live in these append-only tabs
# (the master tabs stay formula-driven and are never written to).
ENTRY_TABS = {
    "power": "Entries - Power",
    "armcare": "Entries - Arm Care",
    "pitching": "Entries - Pitching",
}
NOTES_TAB = "Entries - Notes"


def _clean_names(df: pd.DataFrame) -> pd.DataFrame:
    """Trim whitespace from column names (mirrors R clean_names)."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _read_entries_csv(fname: str) -> "pd.DataFrame | None":
    f = ENTRIES_DIR / fname
    if not f.exists():
        return None
    try:
        df = pd.read_csv(f)
        return _clean_names(df) if not df.empty else None
    except Exception:
        return None


def _read_entries(key: str) -> "pd.DataFrame | None":
    """Appended entries for a key, from the Sheet tab (live) or local CSV."""
    if sheets.sheets_enabled():
        df = sheets.read_tab_optional(ENTRY_TABS[key])
        return _clean_names(df) if df is not None else None
    return _read_entries_csv(ENTRY_SHEETS[key][0])


def _read_notes_extra() -> "pd.DataFrame | None":
    if sheets.sheets_enabled():
        df = sheets.read_tab_optional(NOTES_TAB)
        return _clean_names(df) if df is not None else None
    return _read_entries_csv(NOTES_CSV)


def load_data(path: str | Path = EXCEL_PATH) -> dict[str, pd.DataFrame]:
    """Read every master sheet, then merge any appended entries.

    Reads from the live Google Sheet when configured (see sheets.py), else from
    the local Excel workbook + entries/ CSVs. Both paths are non-destructive:
    the formula-driven master is never written to.
    """
    use_sheets = sheets.sheets_enabled()
    xl = None if use_sheets else pd.ExcelFile(path)
    db: dict[str, pd.DataFrame] = {}
    for key, (sheet, id_col) in _SHEETS.items():
        if use_sheets:
            df = _clean_names(sheets.read_tab(sheet))
        else:
            df = _clean_names(pd.read_excel(xl, sheet_name=sheet))
        if id_col in df.columns:
            df = df[df[id_col].notna()].reset_index(drop=True)
        db[key] = df

    # Merge appended tabular entries (non-destructive; master never rewritten).
    for key, (fname, id_col) in ENTRY_SHEETS.items():
        extra = _read_entries(key)
        if extra is not None and id_col in extra.columns:
            merged = pd.concat([db[key], extra], ignore_index=True)
            merged = merged[merged[id_col].notna()].reset_index(drop=True)
            db[key] = merged

    # Companion coaching notes (long format), merged by get_notes().
    db["notes_extra"] = _read_notes_extra()
    return db


def athlete_names(db: dict[str, pd.DataFrame]) -> list[str]:
    return [str(n) for n in db["bio"]["Player Name"].dropna().tolist()]


def athlete_row(df: pd.DataFrame, name_col: str, athlete: str) -> pd.DataFrame:
    """Rows for one athlete (empty frame if none)."""
    if name_col not in df.columns:
        return df.iloc[0:0]
    return df[df[name_col].astype(str) == str(athlete)]


# ---------------------------------------------------------------------------
# Value helpers
# ---------------------------------------------------------------------------
_NA_STRINGS = {"", "na", "n/a", "nan", "nat", "#div/0!", "none"}


def is_blank(x) -> bool:
    """True for None / NaN / blank / N/A-style placeholders."""
    if x is None:
        return True
    try:
        if pd.isna(x):
            return True
    except (TypeError, ValueError):
        pass
    return str(x).strip().lower() in _NA_STRINGS


def clean_str(x) -> str | None:
    if is_blank(x):
        return None
    return str(x).strip()


def safe_num(x) -> float | None:
    """Numeric value or None (mirrors R safe_num)."""
    if is_blank(x):
        return None
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def cell(df: pd.DataFrame, col: str, default=None):
    """First cell of a column for the current (already-filtered) frame."""
    if df is None or df.empty or col not in df.columns:
        return default
    val = df[col].iloc[0]
    return default if is_blank(val) else val


def fmt_date(x, out_fmt: str = "%B %d, %Y") -> str | None:
    """Format a value that may be a datetime, Excel serial, or ISO string."""
    if is_blank(x):
        return None
    if isinstance(x, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(x).strftime(out_fmt)
    s = str(x).strip()
    # Numeric: Excel serial date, or a bare year
    try:
        num = float(s)
        if num > 40000:
            d = pd.Timestamp("1899-12-30") + pd.Timedelta(days=num)
            return d.strftime(out_fmt)
        if 1900 <= num <= 2100:  # year only -> show just the year
            return str(int(num))
    except (TypeError, ValueError):
        pass
    # ISO / general string
    parsed = pd.to_datetime(s[:10], errors="coerce")
    if pd.notna(parsed):
        return parsed.strftime(out_fmt)
    return s


def to_timestamp(x):
    """Best-effort parse to a real Timestamp for plotting (None if impossible)."""
    if is_blank(x):
        return None
    if isinstance(x, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(x)
    s = str(x).strip()
    try:
        num = float(s)
        if num > 40000:
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=num)
        if 1900 <= num <= 2100:  # bare year -> Jan 1 of that year
            return pd.Timestamp(int(num), 1, 1)
    except (TypeError, ValueError):
        pass
    if s.lower() == "current":
        return pd.Timestamp.today().normalize()
    parsed = pd.to_datetime(s[:10], errors="coerce")
    return parsed if pd.notna(parsed) else None


# ---------------------------------------------------------------------------
# Coaches Notes parsing — master sheet: col 1 Name, then (Date, Notes, Video)*
# ---------------------------------------------------------------------------
def parse_notes(df: pd.DataFrame, athlete: str) -> list[dict]:
    rows = df[df["Name"].astype(str) == str(athlete)]
    if rows.empty:
        return []
    row = rows.iloc[0]
    values = list(row.values)  # positional, matches R column iteration

    entries: list[dict] = []
    idx = 1  # skip Name column
    while idx + 2 <= len(values) - 1:
        raw = values[idx]
        date_val = fmt_date(raw)
        notes_val = clean_str(values[idx + 1])
        video_val = clean_str(values[idx + 2])
        if date_val is not None or notes_val is not None:
            entries.append({
                "date": date_val or "Date unknown",
                "notes": notes_val or "No notes recorded.",
                "video": video_val,
                "ts": to_timestamp(raw),
            })
        idx += 3
    return entries


def get_notes(db: dict, athlete: str) -> list[dict]:
    """Master notes + appended notes, sorted oldest -> newest."""
    entries = parse_notes(db["notes"], athlete)
    extra = db.get("notes_extra")
    if extra is not None and "Name" in extra.columns:
        sub = extra[extra["Name"].astype(str) == str(athlete)]
        for _, r in sub.iterrows():
            raw = r.get("Date")
            entries.append({
                "date": fmt_date(raw) or "Date unknown",
                "notes": clean_str(r.get("Notes")) or "No notes recorded.",
                "video": clean_str(r.get("Video")),
                "ts": to_timestamp(raw),
            })
    entries.sort(key=lambda e: (e.get("ts") is None, e.get("ts") or pd.Timestamp.min))
    return entries


# ---------------------------------------------------------------------------
# Video embedding — Google Drive share links or local filenames
# ---------------------------------------------------------------------------
def drive_embed_url(video_val: str):
    """Return a Google-Drive /preview embed URL, or None if not a Drive link."""
    if is_blank(video_val):
        return None
    v = str(video_val)
    if "drive.google.com" not in v.lower():
        return None
    m = re.search(r"/d/([^/?]+)", v) or re.search(r"id=([^&]+)", v)
    if m:
        return f"https://drive.google.com/file/d/{m.group(1)}/preview"
    return None
