"""
Write-back, history, and export helpers for the Athlete Dashboard.

The master workbook is formula-driven (names are `=Bio!A2`, etc.), so we NEVER
write to it. New entries are appended to append-only CSVs under entries/, which
data.load_data() merges back in. This keeps the master's formulas intact and
makes data entry safe and reversible (just delete a CSV row).
"""
from __future__ import annotations

import io
from datetime import datetime

import pandas as pd

import data as D


# ---------------------------------------------------------------------------
# Appending entries (safe, non-destructive)
# ---------------------------------------------------------------------------
def _append_csv(fname: str, row: dict):
    """Append one row-dict to entries/<fname>, creating it with headers if new."""
    D.ENTRIES_DIR.mkdir(exist_ok=True)
    f = D.ENTRIES_DIR / fname
    new = pd.DataFrame([row])
    if f.exists():
        try:
            old = pd.read_csv(f)
            out = pd.concat([old, new], ignore_index=True)
        except Exception:
            out = new
    else:
        out = new
    out.to_csv(f, index=False)
    return f


def add_entry(kind: str, values: dict):
    """Append a test row for kind in {'power','armcare','pitching'}."""
    fname, _id = D.ENTRY_SHEETS[kind]
    # keep only non-empty values
    clean = {k: v for k, v in values.items()
             if v is not None and str(v).strip() != ""}
    return _append_csv(fname, clean)


def add_coaches_note(athlete: str, date_val, notes: str, video: str | None = None):
    return _append_csv(D.NOTES_CSV, {
        "Name": athlete,
        "Date": _iso(date_val),
        "Notes": notes,
        "Video": (video or ""),
    })


def armcare_total_score(ir, er, scaption, grip, body_weight):
    """Mirror the master's Arm Care formula: ((IR+ER+Scaption+Grip)/BW)*100."""
    vals = [D.safe_num(x) for x in (ir, er, scaption, grip)]
    bw = D.safe_num(body_weight)
    if any(v is None for v in vals) or not bw:
        return None
    return round(sum(vals) / bw * 100, 2)


def next_test_number(db: dict, athlete: str) -> int:
    p = D.athlete_row(db["power"], "Name", athlete)
    return int(len(p)) + 1


def _iso(x) -> str:
    """Normalise a date-ish value to an ISO string for storage."""
    if x is None:
        return ""
    if isinstance(x, (datetime, pd.Timestamp)):
        return pd.Timestamp(x).strftime("%Y-%m-%d")
    try:
        return x.strftime("%Y-%m-%d")  # datetime.date
    except AttributeError:
        return str(x)


# ---------------------------------------------------------------------------
# History (sorted by date) for trend charts + latest-value lookups
# ---------------------------------------------------------------------------
def athlete_history(df: pd.DataFrame, name_col: str, athlete: str,
                    date_col: str) -> pd.DataFrame:
    """All rows for an athlete, with a parsed `_date` column, sorted ascending."""
    sub = D.athlete_row(df, name_col, athlete).copy()
    if sub.empty:
        return sub
    if date_col in sub.columns:
        sub["_date"] = pd.to_datetime(sub[date_col], errors="coerce")
        sub = sub.sort_values("_date", na_position="first").reset_index(drop=True)
    return sub


def latest_row(df: pd.DataFrame, name_col: str, athlete: str,
               date_col: str) -> pd.DataFrame:
    """The most recent single row for an athlete (falls back to first row)."""
    h = athlete_history(df, name_col, athlete, date_col)
    return h.tail(1) if not h.empty else h


# ---------------------------------------------------------------------------
# PDF one-pager
# ---------------------------------------------------------------------------
def build_one_pager(db: dict, athlete: str) -> bytes:
    """Render a one-page PDF summary for an athlete; returns raw PDF bytes."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer,
                                    Table, TableStyle)

    PRIMARY = colors.HexColor("#3c8dbc")
    bio = D.athlete_row(db["bio"], "Player Name", athlete)
    pit = latest_row(db["pitching"], "Athlete", athlete, "Date")
    pwr = latest_row(db["power"], "Name", athlete, "Test Date")
    ac = latest_row(db["armcare"], "Athlete", athlete, "Date")
    plan = D.athlete_row(db["plan"], "Athlete", athlete)
    notes = D.get_notes(db, athlete)

    styles = getSampleStyleSheet()
    h1 = ParagraphStyle("h1", parent=styles["Heading1"], textColor=PRIMARY,
                        fontSize=20, spaceAfter=2)
    sub = ParagraphStyle("sub", parent=styles["Normal"], fontSize=10,
                         textColor=colors.HexColor("#666"))
    sec = ParagraphStyle("sec", parent=styles["Heading2"], textColor=PRIMARY,
                        fontSize=13, spaceBefore=10, spaceAfter=4)
    body = styles["Normal"]

    def g(df, col, default="—"):
        v = D.cell(df, col)
        return default if v is None else str(v)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, topMargin=0.6 * inch,
                            bottomMargin=0.6 * inch, leftMargin=0.7 * inch,
                            rightMargin=0.7 * inch, title=f"{athlete} — One-Pager")
    el = []

    el.append(Paragraph(str(athlete), h1))
    pos = " / ".join([x for x in [g(bio, "Position 1", ""), g(bio, "Position 2", "")]
                      if x and x != "—"])
    el.append(Paragraph(
        f"{g(bio, 'High School / College')} &nbsp;•&nbsp; Class of "
        f"{g(bio, 'Graduating Class')} &nbsp;•&nbsp; {pos or '—'}", sub))
    el.append(Spacer(1, 8))

    bio_rows = [
        ["Height", g(bio, "Height"), "Weight", f"{g(bio, 'Weight')} lbs"],
        ["Select Team", g(bio, "Select Team"), "FB Velo (Max)",
         f"{g(pit, 'FB Velo (Max)')} mph" if D.cell(pit, "FB Velo (Max)") else "—"],
    ]
    t = Table(bio_rows, colWidths=[1.4 * inch, 2.0 * inch, 1.4 * inch, 2.0 * inch])
    t.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("TEXTCOLOR", (0, 0), (0, -1), PRIMARY),
        ("TEXTCOLOR", (2, 0), (2, -1), PRIMARY),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LINEBELOW", (0, 0), (-1, -2), 0.4, colors.HexColor("#e0e0e0")),
    ]))
    el.append(t)

    el.append(Paragraph("Power Snapshot", sec))
    power_bits = []
    for label, col, unit in [("Body Weight", "Body Weight", " lbs"),
                             ("Vert (Regular)", "Vert- Regular", '"'),
                             ("Vert (CMJ)", "Vert- CMJ", '"'),
                             ("Vert (Step-In)", "Vert- Step In", '"')]:
        v = D.cell(pwr, col)
        if v is not None:
            power_bits.append(f"<b>{label}:</b> {v}{unit}")
    el.append(Paragraph("&nbsp;&nbsp;•&nbsp;&nbsp;".join(power_bits) or "No data", body))

    if D.cell(ac, "Total Score") is not None:
        el.append(Paragraph("Arm Care", sec))
        el.append(Paragraph(
            f"<b>Total Score:</b> {round(D.safe_num(D.cell(ac,'Total Score')),1)} "
            f"&nbsp;•&nbsp; <b>IR:</b> {g(ac,'IR')} &nbsp;•&nbsp; <b>ER:</b> {g(ac,'ER')} "
            f"&nbsp;•&nbsp; <b>Grip:</b> {g(ac,'Grip')} &nbsp;•&nbsp; "
            f"<b>Scaption:</b> {g(ac,'Scaption')}", body))

    el.append(Paragraph("Development Plan", sec))
    for label in ["Body", "Performance", "Pitching"]:
        el.append(Paragraph(f"<b>{label}:</b> {g(plan, label)}", body))
        el.append(Spacer(1, 2))

    if notes:
        el.append(Paragraph("Latest Coaching Note", sec))
        last = notes[-1]
        el.append(Paragraph(f"<b>{last['date']}:</b> {last['notes']}", body))

    el.append(Spacer(1, 14))
    el.append(Paragraph(
        f"Generated {datetime.now().strftime('%B %d, %Y')} — Athlete Dashboard",
        ParagraphStyle("f", parent=sub, fontSize=8)))

    doc.build(el)
    buf.seek(0)
    return buf.getvalue()
