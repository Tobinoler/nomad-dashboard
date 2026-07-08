"""
Google Sheets backend for the Athlete Dashboard.

This is the *live* data store used when deployed. It mirrors the existing
Excel + entries/CSV design exactly, just backed by a private Google Sheet:

  * Master tabs (Bio, Power Testing, ...) are read live (formulas -> values).
  * Append-only "entry" tabs receive new test rows / notes, so in-app data
    entry PERSISTS on a stateless host (unlike writing local files).

If credentials + a sheet id are not configured, everything here is inert and
data.py / actions.py fall back to the local Excel file + entries/ CSVs. That
keeps local development working with zero setup.

Configuration (environment variables):
  NOMAD_SHEET_ID                 - the Google Sheet's id (from its URL)
  GOOGLE_SERVICE_ACCOUNT_JSON    - the service-account key, as a JSON string
                                   (used on the host; stored as a secret)
  ...or a local file `service_account.json` next to this file (dev only,
  gitignored).
"""
from __future__ import annotations

import json
import os
from functools import lru_cache
from pathlib import Path

import pandas as pd

_KEY_FILE = Path(__file__).parent / "service_account.json"
_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def _sheet_id() -> str | None:
    return os.environ.get("NOMAD_SHEET_ID") or None


def sheets_enabled() -> bool:
    """True only if we have both a sheet id and some form of credentials."""
    if _sheet_id() is None:
        return False
    return bool(os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")) or _KEY_FILE.exists()


def _credentials():
    """Build service-account credentials from env JSON or a local key file."""
    from google.oauth2.service_account import Credentials

    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if raw:
        info = json.loads(raw)
        return Credentials.from_service_account_info(info, scopes=_SCOPES)
    return Credentials.from_service_account_file(str(_KEY_FILE), scopes=_SCOPES)


@lru_cache(maxsize=1)
def _spreadsheet():
    """Open (and cache) the configured spreadsheet handle."""
    import gspread

    gc = gspread.authorize(_credentials())
    return gc.open_by_key(_sheet_id())


def _mangle_columns(header: list) -> list[str]:
    """De-duplicate/blank-fill column names the way pandas.read_excel does.

    Blank -> 'Unnamed: N'; duplicate 'X' -> 'X', 'X.1', 'X.2', ... This keeps
    the wide, repeated-header Coaches Notes sheet readable positionally.
    """
    seen: dict[str, int] = {}
    out: list[str] = []
    for i, raw in enumerate(header):
        name = str(raw).strip()
        if name == "":
            name = f"Unnamed: {i}"
        if name in seen:
            seen[name] += 1
            name = f"{name}.{seen[name]}"
        else:
            seen[name] = 0
        out.append(name)
    return out


def read_tab(title: str) -> pd.DataFrame:
    """Read one worksheet into a DataFrame using UNFORMATTED values.

    Unformatted values keep numbers numeric and dates as serial numbers
    (same 1899-12-30 epoch as Excel), which data.py's helpers already handle.
    """
    ws = _spreadsheet().worksheet(title)
    values = ws.get_values(value_render_option="UNFORMATTED_VALUE")
    if not values:
        return pd.DataFrame()
    header = _mangle_columns(values[0])
    rows = values[1:]
    width = len(header)
    # pad/truncate each row so ragged trailing-blank rows line up with header
    norm = [(r + [None] * width)[:width] for r in rows]
    df = pd.DataFrame(norm, columns=header)
    # blank strings -> NA so is_blank()/safe_num() behave like the Excel read
    return df.replace("", pd.NA)


def read_tab_optional(title: str) -> "pd.DataFrame | None":
    """read_tab, but return None if the tab doesn't exist yet (empty entries)."""
    import gspread

    try:
        df = read_tab(title)
    except gspread.WorksheetNotFound:
        return None
    return df if not df.empty else None


def append_row_dict(title: str, row: dict) -> None:
    """Append one row-dict to an entry tab, creating it (with headers) if new.

    If the tab exists, values are aligned to its existing header order and any
    new keys extend the header, so appends never silently drop columns.
    """
    import gspread

    ss = _spreadsheet()
    try:
        ws = ss.worksheet(title)
        header = ws.row_values(1)
    except gspread.WorksheetNotFound:
        header = list(row.keys())
        ws = ss.add_worksheet(title=title, rows=100, cols=max(len(header), 1))
        ws.update([header], "A1")

    # extend header if this row introduces new columns
    new_cols = [k for k in row.keys() if k not in header]
    if new_cols:
        header = header + new_cols
        ws.update([header], "A1")

    values = ["" if row.get(col) is None else str(row.get(col)) for col in header]
    ws.append_row(values, value_input_option="USER_ENTERED")
