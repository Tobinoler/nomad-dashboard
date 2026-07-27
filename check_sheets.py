"""
Google Sheets connection checker for the Athlete Dashboard.

Run this AFTER doing the Google setup. It verifies each step and tells you
exactly what is wrong and how to fix it — no silent failures.

    # in PowerShell, from C:\\NomadDashboard:
    $env:NOMAD_SHEET_ID = "the-long-id-from-your-Google-Sheet-URL"
    python check_sheets.py
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

OK, BAD, WARN = "[ OK ] ", "[FAIL] ", "[warn] "
KEY_FILE = Path(__file__).parent / "service_account.json"

EXPECTED_TABS = ["Bio", "Motor Preferences", "Power Testing", "Pitching",
                 "Arm Care", "Context", "Injuries", "MSSPosture",
                 "Athlete Plan", "Coaches Notes"]


def fail(msg: str, fix: str) -> None:
    print(f"{BAD}{msg}")
    print(f"       -> fix: {fix}")
    sys.exit(1)


def main() -> None:
    print("\nChecking the Google Sheets connection...\n")

    # 1) Sheet id -----------------------------------------------------------
    sid = os.environ.get("NOMAD_SHEET_ID")
    if not sid:
        fail("NOMAD_SHEET_ID is not set.",
             'In PowerShell run:  $env:NOMAD_SHEET_ID = "the-id-from-your-sheet-URL"'
             "\n              (the id is the long code between /d/ and /edit in the URL)")
    print(f"{OK}NOMAD_SHEET_ID is set ({sid[:10]}...)")

    # 2) Credentials present ------------------------------------------------
    have_env = bool(os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON"))
    if not have_env and not KEY_FILE.exists():
        fail("No credentials found.",
             "Put the downloaded Google key file here (rename it exactly):\n"
             f"              {KEY_FILE}")
    print(f"{OK}Credentials found ({'env var' if have_env else KEY_FILE.name})")

    # 3) Libraries installed ------------------------------------------------
    try:
        import gspread  # noqa: F401
        from google.oauth2.service_account import Credentials  # noqa: F401
    except Exception as e:  # pragma: no cover
        fail(f"A required library is missing: {e}",
             "Run:  python -m pip install -r requirements.txt")
    print(f"{OK}gspread + google-auth installed")

    import gspread
    import sheets as S

    # 4) Credentials load + auth -------------------------------------------
    try:
        creds = S._credentials()
    except Exception as e:
        fail(f"The key file could not be read: {e}",
             "service_account.json looks invalid — re-download the JSON key "
             "from Google Cloud > Service Accounts > Keys.")
    sa_email = getattr(creds, "service_account_email", "(unknown)")
    print(f"{OK}Service account: {sa_email}")

    # 5) Open the spreadsheet ----------------------------------------------
    try:
        ss = S._spreadsheet()
        title = ss.title
    except gspread.SpreadsheetNotFound:
        fail("No Sheet found for that NOMAD_SHEET_ID.",
             "Recheck the id — it must be the code from the Sheet's own URL.")
    except gspread.exceptions.APIError as e:
        s = str(e)
        if "PERMISSION_DENIED" in s or "403" in s:
            fail("Permission denied opening the Sheet.",
                 f"Open the Sheet > Share, and add this as Editor:\n              {sa_email}")
        if "SERVICE_DISABLED" in s or "has not been used" in s or "API has not" in s:
            fail("The Google Sheets API is not enabled for your project.",
                 "In Google Cloud, search 'Google Sheets API' and click Enable.")
        fail(f"Google API error: {e}", "See the message above.")
    print(f"{OK}Opened spreadsheet: '{title}'")

    # 6) Expected tabs ------------------------------------------------------
    tabs = [w.title for w in ss.worksheets()]
    missing = [t for t in EXPECTED_TABS if t not in tabs]
    if missing:
        print(f"{WARN}Missing expected tab(s): {missing}")
        print(f"       found tabs: {tabs}")
        print("       -> fix: make sure you used File > Save as Google Sheets on the "
              ".xlsx so every tab carried over (names must match exactly).")
    else:
        print(f"{OK}All 10 master tabs present")

    # 7) Read test ----------------------------------------------------------
    try:
        df = S.read_tab("Bio")
        print(f"{OK}Read the 'Bio' tab: {len(df)} rows x {len(df.columns)} cols")
    except Exception as e:
        fail(f"Could not read a tab: {e}", "Check the sharing/permission step above.")

    # 8) Write test (create + delete a temp tab; no data is touched) --------
    try:
        ws = ss.add_worksheet(title="__conn_test__", rows=1, cols=1)
        ss.del_worksheet(ws)
        print(f"{OK}Write test passed (created + removed a temp tab)")
    except gspread.exceptions.APIError as e:
        fail(f"Reading works but WRITING failed: {e}",
             f"Re-share the Sheet with {sa_email} as EDITOR (not Viewer).")

    print("\n" + "=" * 60)
    print("SUCCESS - the app will now use the live Google Sheet.")
    print("Keep NOMAD_SHEET_ID set (and service_account.json in place),")
    print("then start the app:  python -m shiny run app.py")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()
