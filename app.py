"""
=============================================================================
ATHLETE DASHBOARD  —  Shiny for Python
=============================================================================
Ported from the original R/Shiny app, plus:
  - Progress-over-time trend charts (Power, Arm Care, Pitching)
  - Data Entry tab (appends to append-only CSVs; master workbook never touched)
  - PDF one-pager export

Run locally:
    pip install -r requirements.txt
    shiny run --reload app.py
    # open http://127.0.0.1:8000

Data:
    "Master Evaluation Sheet.xlsx" must sit next to app.py.
    New entries are written to an entries/ folder created alongside it.
=============================================================================
"""
from __future__ import annotations

import os

import pandas as pd
import plotly.graph_objects as go
from shiny import App, reactive, render, ui
from shinywidgets import output_widget, render_widget

import actions as A
import data as D
import sheets

# On a hosted server the filesystem is ephemeral, so appended entries won't
# persist unless the live Google Sheets backend is on. Set NOMAD_HOSTED=1 in
# the host's environment to surface a heads-up banner on the Data Entry tab.
HOSTED = bool(os.environ.get("NOMAD_HOSTED"))

try:
    from faicons import icon_svg
except Exception:
    icon_svg = None


# =============================================================================
# DATA (initial load for UI choices)
# =============================================================================
_db0 = D.load_data()
ATHLETES = D.athlete_names(_db0)

PRIMARY = "#3c8dbc"
GREEN = "#00a65a"
RED = "#dd4b39"
ORANGE = "#f39c12"

POWER_METRICS = ["Body Weight", "Vert- Regular", "Vert- CMJ", "Vert- Step In",
                 "4lb MB Shot Put (Left)", "4lb MB Shot Put (Right)",
                 "Squat (1.00% BW)", "Trap Bar (1.25% BW)", "Triple Broad Jump"]
ARMCARE_METRICS = ["Total Score", "IR", "ER", "Scaption", "Grip", "Body Weight"]

# Athlete comparison — display label -> (db key, name column, date column, value column).
# Uses each athlete's most recent test for the metric.
COMPARE_METRICS = {
    "Body Weight (lbs)":        ("power", "Name", "Test Date", "Body Weight"),
    "Vertical – Regular (in)":  ("power", "Name", "Test Date", "Vert- Regular"),
    "Vertical – CMJ (in)":      ("power", "Name", "Test Date", "Vert- CMJ"),
    "Vertical – Step-In (in)":  ("power", "Name", "Test Date", "Vert- Step In"),
    "MB Shot Put – Left":       ("power", "Name", "Test Date", "4lb MB Shot Put (Left)"),
    "MB Shot Put – Right":      ("power", "Name", "Test Date", "4lb MB Shot Put (Right)"),
    "Squat (1.00% BW)":         ("power", "Name", "Test Date", "Squat (1.00% BW)"),
    "Trap Bar (1.25% BW)":      ("power", "Name", "Test Date", "Trap Bar (1.25% BW)"),
    "Triple Broad Jump":        ("power", "Name", "Test Date", "Triple Broad Jump"),
    "Arm Care – Total Score":   ("armcare", "Athlete", "Date", "Total Score"),
    "Arm Care – IR":            ("armcare", "Athlete", "Date", "IR"),
    "Arm Care – ER":            ("armcare", "Athlete", "Date", "ER"),
    "Arm Care – Grip":          ("armcare", "Athlete", "Date", "Grip"),
    "Arm Care – Scaption":      ("armcare", "Athlete", "Date", "Scaption"),
}


# =============================================================================
# SMALL UI HELPERS
# =============================================================================
def ico(name: str):
    if icon_svg is None:
        return None
    try:
        return icon_svg(name)
    except Exception:
        return None


def vbox(title, value, icon="circle", theme="primary"):
    return ui.value_box(title, str(value), showcase=ico(icon), theme=theme)


# Responsive column widths: stack on phones, spread on desktop.
CW4 = {"sm": [6, 6, 6, 6], "lg": [3, 3, 3, 3]}
CW3 = {"sm": [12, 6, 6], "lg": [4, 4, 4]}
CW_2 = {"sm": [12, 12], "lg": [6, 6]}
CW_57 = {"sm": [12, 12], "lg": [5, 7]}


def make_video_ui(video_val):
    embed = D.drive_embed_url(video_val)
    if embed:
        return ui.div(
            ui.tags.iframe(src=embed, width="100%", height="320", frameborder="0",
                           allow="autoplay; encrypted-media", allowfullscreen="true"),
            style="margin-top:10px;border-radius:6px;overflow:hidden;",
        )
    if not D.is_blank(video_val):
        return ui.div(
            ui.tags.video(
                ui.tags.source(src=f"videos/{video_val}", type="video/mp4"),
                controls="", width="100%", style="display:block;max-height:320px;",
            ),
            style="margin-top:10px;background:#000;border-radius:6px;overflow:hidden;",
        )
    return None


def empty_fig(title: str = "", note: str = "No data"):
    fig = go.Figure()
    fig.update_layout(
        title=title, xaxis={"visible": False}, yaxis={"visible": False},
        annotations=[dict(text=note, showarrow=False,
                          font=dict(size=15, color="#888"))],
        margin=dict(t=50, l=20, r=20, b=20), plot_bgcolor="white",
    )
    return fig


def trend_fig(hist: pd.DataFrame, metric: str, unit: str = ""):
    """Line/marker plot of a numeric metric over the history's _date column."""
    if hist.empty or metric not in hist.columns or "_date" not in hist.columns:
        return empty_fig(metric, "No dated history yet")
    d = hist.dropna(subset=["_date"]).copy()
    d["_y"] = pd.to_numeric(d[metric], errors="coerce")
    d = d.dropna(subset=["_y"])
    if d.empty:
        return empty_fig(f"{metric} over time", f"No {metric} data")
    fig = go.Figure(go.Scatter(
        x=d["_date"], y=d["_y"], mode="lines+markers",
        line=dict(color=PRIMARY, width=3), marker=dict(size=11, color=PRIMARY),
    ))
    if len(d) == 1:
        fig.add_annotation(x=d["_date"].iloc[0], y=d["_y"].iloc[0],
                           text="Log another test to draw the trend",
                           showarrow=True, arrowcolor="#bbb", ax=0, ay=-35,
                           font=dict(size=11, color="#888"))
    fig.update_layout(title=f"{metric} over time",
                      yaxis_title=(metric + (f" ({unit})" if unit else "")),
                      xaxis_title="", margin=dict(t=50, b=30), plot_bgcolor="white")
    return fig


# =============================================================================
# THEME + CSS
# =============================================================================
# Clean Bootswatch preset, brand blue as primary so charts + UI stay cohesive.
THEME = ui.Theme("cosmo").add_defaults(primary=PRIMARY)

APP_CSS = ui.tags.style("""
  :root { --brand:#3c8dbc; --brand-2:#00a65a; }
  body { background:#f4f6f9; -webkit-font-smoothing:antialiased;
         text-rendering:optimizeLegibility; }
  .athlete-bar { font-size:15px; color:#555; margin:-4px 0 14px 2px; }
  .athlete-bar b { color:var(--brand); }

  /* Cards: soft, elevated, subtle hover lift */
  .card { border:none !important; border-radius:12px !important;
          box-shadow:0 1px 3px rgba(16,24,40,.06),0 1px 2px rgba(16,24,40,.04);
          transition:box-shadow .18s ease, transform .18s ease; }
  .card:hover { box-shadow:0 8px 22px rgba(16,24,40,.10); }
  .card-header { background:transparent !important;
                 border-bottom:1px solid #eef0f3 !important;
                 font-weight:600; letter-spacing:.2px; }

  /* Value boxes */
  .bslib-value-box { border-radius:12px !important;
                     box-shadow:0 1px 3px rgba(16,24,40,.06) !important; }

  /* Nav pills */
  .nav-pills { gap:4px; margin-bottom:16px; }
  .nav-pills .nav-link { color:#54606d; border-radius:8px; font-weight:500;
                         padding:7px 14px; transition:background .15s,color .15s; }
  .nav-pills .nav-link:hover { background:#e9eef3; color:var(--brand); }
  .nav-pills .nav-link.active { background:var(--brand) !important; color:#fff;
                         box-shadow:0 2px 6px rgba(60,141,188,.35); }

  /* Sidebar */
  .btn-outline-primary { border-radius:8px; }

  .section-header { font-size:16px; font-weight:bold;
                    border-bottom:2px solid #3c8dbc;
                    margin-bottom:10px; padding-bottom:4px; }
  .plan-card { background:#f9f9f9; border-left:4px solid #3c8dbc;
               padding:12px; margin-bottom:12px; border-radius:8px; }
  .video-no-data { padding:40px 20px; color:#888; text-align:center;
                   background:#f5f5f5; border-radius:8px; margin-top:10px; }
  .entry-hint { color:#777; font-size:13px; margin-bottom:8px; }

  .cn-timeline { position:relative; padding:10px 0; }
  .cn-timeline::before { content:''; position:absolute; left:22px;
      top:0; bottom:0; width:3px;
      background:linear-gradient(to bottom,#3c8dbc,#00a65a); border-radius:3px; }
  .cn-entry { position:relative; margin-left:58px; margin-bottom:28px; }
  .cn-dot { position:absolute; left:-44px; top:14px; width:16px; height:16px;
      background:#3c8dbc; border:3px solid #fff; border-radius:50%;
      box-shadow:0 0 0 2px #3c8dbc; }
  .cn-card { background:#fff; border:1px solid #e0e0e0;
      border-left:4px solid #3c8dbc; border-radius:6px; padding:14px 16px;
      box-shadow:0 1px 4px rgba(0,0,0,.07); }
  .cn-date { font-size:12px; font-weight:700; text-transform:uppercase;
      letter-spacing:.6px; color:#3c8dbc; margin-bottom:6px; }
  .cn-notes { font-size:14px; line-height:1.55; color:#333; margin-bottom:0; }
  .cn-no-data { padding:40px 20px; color:#888; text-align:center;
      font-size:15px; background:#f9f9f9; border-radius:6px; }
""")


# =============================================================================
# UI
# =============================================================================
app_ui = ui.page_sidebar(
    ui.sidebar(
        ui.input_select("athlete", "Select Athlete", choices=ATHLETES,
                        selected=ATHLETES[0] if ATHLETES else None),
        ui.hr(),
        ui.download_button("dl_pdf", "Download One-Pager (PDF)",
                           class_="btn-outline-primary btn-sm"),
        width=260,
    ),
    APP_CSS,
    ui.busy_indicators.use(spinners=True, pulse=True),
    ui.output_ui("athlete_bar"),
    ui.navset_pill(

        # ---- OVERVIEW ------------------------------------------------------
        ui.nav_panel(
            "Overview",
            ui.output_ui("overview_boxes_1"),
            ui.output_ui("overview_boxes_2"),
            ui.card(ui.card_header("Quick Snapshot"), ui.output_ui("overview_summary")),
            ui.card(
                ui.card_header("Athlete Video"),
                ui.div(
                    ui.input_action_button("btn_behind", "Behind View",
                                           class_="btn-primary btn-sm me-2"),
                    ui.input_action_button("btn_side", "Side View",
                                           class_="btn-info btn-sm me-2"),
                    ui.input_action_button("btn_other", "Other View",
                                           class_="btn-secondary btn-sm"),
                ),
                ui.output_ui("video_player"),
            ),
            icon=ico("user"),
        ),

        # ---- COMPARE -------------------------------------------------------
        ui.nav_panel(
            "Compare",
            ui.card(
                ui.card_header("Compare Athletes"),
                ui.layout_columns(
                    ui.input_selectize(
                        "cmp_athletes", "Athletes to compare", choices=ATHLETES,
                        multiple=True,
                        selected=ATHLETES[:3] if len(ATHLETES) >= 3 else ATHLETES),
                    ui.input_select("cmp_metric", "Strength metric",
                                    choices=list(COMPARE_METRICS.keys()),
                                    selected="Arm Care – Total Score"),
                    col_widths={"sm": [12, 12], "lg": [7, 5]},
                ),
            ),
            ui.card(ui.card_header("Ranked Comparison"), output_widget("compare_chart")),
            ui.card(ui.card_header("All Strength Numbers (latest per athlete)"),
                    ui.output_data_frame("compare_table")),
            icon=ico("scale-balanced"),
        ),

        # ---- MOTOR ---------------------------------------------------------
        ui.nav_panel(
            "Motor Preferences",
            ui.layout_columns(
                ui.card(ui.card_header("Motor Profile"), ui.output_table("motor_table")),
                ui.card(ui.card_header("Motor Preference Spectrum"),
                        output_widget("motor_chart")),
                col_widths=CW_57,
            ),
            icon=ico("brain"),
        ),

        # ---- POWER ---------------------------------------------------------
        ui.nav_panel(
            "Power Testing",
            ui.output_ui("power_boxes"),
            ui.layout_columns(
                ui.card(ui.card_header("Medicine Ball — Shot Put"), output_widget("mb_chart")),
                ui.card(ui.card_header("Vertical Jump Comparison"), output_widget("vert_chart")),
                col_widths=CW_2,
            ),
            ui.card(
                ui.card_header("Progress Over Time"),
                ui.input_select("power_metric", "Metric", choices=POWER_METRICS,
                                selected="Vert- CMJ"),
                output_widget("power_trend"),
            ),
            ui.card(ui.card_header("Full Power Testing History"),
                    ui.output_data_frame("power_table")),
            icon=ico("dumbbell"),
        ),

        # ---- PITCHING ------------------------------------------------------
        ui.nav_panel(
            "Pitching",
            ui.output_ui("pitch_boxes"),
            ui.layout_columns(
                ui.card(ui.card_header("Pitch Repertoire"), ui.output_ui("pitch_arsenal")),
                ui.card(ui.card_header("Strengths & Weaknesses"), ui.output_ui("pitch_sw")),
                col_widths=CW_2,
            ),
            ui.card(ui.card_header("Goals"), ui.output_ui("pitch_goals")),
            ui.card(ui.card_header("Fastball Velocity Over Time"),
                    output_widget("pitch_trend")),
            ui.card(ui.card_header("Full Pitching History"),
                    ui.output_data_frame("pitch_table")),
            icon=ico("baseball"),
        ),

        # ---- CONTEXT -------------------------------------------------------
        ui.nav_panel(
            "Context",
            ui.card(ui.card_header("Athlete Background"), ui.output_ui("context_bullets")),
            icon=ico("book"),
        ),

        # ---- INJURIES ------------------------------------------------------
        ui.nav_panel(
            "Injuries",
            ui.card(ui.card_header("Injury History"), ui.output_data_frame("injury_table")),
            ui.card(ui.card_header("Injury Timeline"), output_widget("injury_timeline")),
            icon=ico("kit-medical"),
        ),

        # ---- MSS -----------------------------------------------------------
        ui.nav_panel(
            "MSS / Posture",
            ui.card(ui.card_header("Movement Screen & Posture Observations"),
                    ui.output_ui("mss_bullets")),
            icon=ico("person"),
        ),

        # ---- PLAN ----------------------------------------------------------
        ui.nav_panel(
            "Athlete Plan",
            ui.card(ui.card_header("Development Plan"), ui.output_ui("plan_cards")),
            icon=ico("clipboard-list"),
        ),

        # ---- ARM CARE ------------------------------------------------------
        ui.nav_panel(
            "Arm Care",
            ui.output_ui("armcare_boxes"),
            ui.layout_columns(
                ui.card(ui.card_header("Arm Care Metrics"), ui.output_table("armcare_table")),
                ui.card(ui.card_header("Metrics vs. Benchmarks"), output_widget("armcare_chart")),
                col_widths=CW_57,
            ),
            ui.card(
                ui.card_header("Progress Over Time"),
                ui.input_select("armcare_metric", "Metric", choices=ARMCARE_METRICS,
                                selected="Total Score"),
                output_widget("armcare_trend"),
            ),
            icon=ico("heart-pulse"),
        ),

        # ---- COACHES NOTES -------------------------------------------------
        ui.nav_panel(
            "Coaches Notes",
            ui.output_ui("cn_boxes"),
            ui.card(ui.card_header("Development Timeline"), ui.output_ui("coaches_timeline")),
            icon=ico("chalkboard-user"),
        ),

        # ---- DATA ENTRY ----------------------------------------------------
        ui.nav_panel(
            "Data Entry",
            ui.output_ui("entry_target"),
            ui.layout_columns(
                ui.card(
                    ui.card_header("Add Arm Care Test"),
                    ui.input_date("ac_date", "Date"),
                    ui.input_numeric("ac_bw", "Body Weight (lbs)", value=None),
                    ui.input_numeric("ac_ir", "IR", value=None),
                    ui.input_numeric("ac_er", "ER", value=None),
                    ui.input_numeric("ac_scap", "Scaption", value=None),
                    ui.input_numeric("ac_grip", "Grip", value=None),
                    ui.div("Total Score is computed automatically.", class_="entry-hint"),
                    ui.input_action_button("submit_armcare", "Save Arm Care Test",
                                           class_="btn-success btn-sm"),
                ),
                ui.card(
                    ui.card_header("Add Power Test"),
                    ui.input_date("pw_date", "Test Date"),
                    ui.input_numeric("pw_bw", "Body Weight (lbs)", value=None),
                    ui.input_numeric("pw_reg", 'Vert - Regular (")', value=None),
                    ui.input_numeric("pw_cmj", 'Vert - CMJ (")', value=None),
                    ui.input_numeric("pw_step", 'Vert - Step In (")', value=None),
                    ui.input_numeric("pw_mbl", "4lb MB Shot Put (Left)", value=None),
                    ui.input_numeric("pw_mbr", "4lb MB Shot Put (Right)", value=None),
                    ui.input_action_button("submit_power", "Save Power Test",
                                           class_="btn-success btn-sm"),
                ),
                col_widths=CW_2,
            ),
            ui.layout_columns(
                ui.card(
                    ui.card_header("Add Pitching Entry"),
                    ui.input_date("pt_date", "Date"),
                    ui.input_numeric("pt_max", "FB Velo (Max)", value=None),
                    ui.input_numeric("pt_avg", "FB Velo (Avg)", value=None),
                    ui.input_text("pt_p1", "Pitch 1"),
                    ui.input_text("pt_p2", "Pitch 2"),
                    ui.input_text("pt_p3", "Pitch 3"),
                    ui.input_action_button("submit_pitching", "Save Pitching Entry",
                                           class_="btn-success btn-sm"),
                ),
                ui.card(
                    ui.card_header("Add Coaching Note"),
                    ui.input_date("note_date", "Date"),
                    ui.input_text_area("note_text", "Notes", rows=5,
                                       placeholder="What was worked on / observed…"),
                    ui.input_text("note_video", "Video URL (optional)",
                                  placeholder="Google Drive share link"),
                    ui.input_action_button("submit_note", "Save Coaching Note",
                                           class_="btn-success btn-sm"),
                ),
                col_widths=CW_2,
            ),
            icon=ico("pen-to-square"),
        ),

        id="tabs",
    ),
    title="Athlete Dashboard",
    theme=THEME,
    fillable=False,
)


# =============================================================================
# SERVER
# =============================================================================
def server(input, output, session):

    # -- Reactive data layer: reloads (master + appended CSVs) after each write
    data_version = reactive.value(0)

    @reactive.calc
    def DB():
        data_version()  # dependency
        return D.load_data()

    # Single-athlete frames
    @reactive.calc
    def r_bio():
        return D.athlete_row(DB()["bio"], "Player Name", input.athlete())

    @reactive.calc
    def r_motor():
        return D.athlete_row(DB()["motor"], "Athlete", input.athlete())

    # Power: latest single row for snapshots, full history for tables/trends
    @reactive.calc
    def r_power():
        return A.latest_row(DB()["power"], "Name", input.athlete(), "Test Date")

    @reactive.calc
    def r_power_all():
        return A.athlete_history(DB()["power"], "Name", input.athlete(), "Test Date")

    @reactive.calc
    def r_pitching():
        return A.latest_row(DB()["pitching"], "Athlete", input.athlete(), "Date")

    @reactive.calc
    def r_pitching_all():
        return A.athlete_history(DB()["pitching"], "Athlete", input.athlete(), "Date")

    @reactive.calc
    def r_armcare():
        return A.latest_row(DB()["armcare"], "Athlete", input.athlete(), "Date")

    @reactive.calc
    def r_armcare_all():
        return A.athlete_history(DB()["armcare"], "Athlete", input.athlete(), "Date")

    @reactive.calc
    def r_context():
        return D.athlete_row(DB()["context"], "Athlete", input.athlete())

    @reactive.calc
    def r_injuries():
        return D.athlete_row(DB()["injuries"], "Athlete", input.athlete())

    @reactive.calc
    def r_mss():
        return D.athlete_row(DB()["mss"], "Athlete", input.athlete())

    @reactive.calc
    def r_plan():
        return D.athlete_row(DB()["plan"], "Athlete", input.athlete())

    @reactive.calc
    def r_notes():
        return D.get_notes(DB(), input.athlete())

    # -- Header bar + PDF download
    @render.ui
    def athlete_bar():
        return ui.div(ui.HTML(f"Viewing: <b>{input.athlete()}</b>"), class_="athlete-bar")

    @render.download(filename=lambda: f"{input.athlete()}_one_pager.pdf")
    def dl_pdf():
        yield A.build_one_pager(DB(), input.athlete())

    # =========================================================================
    # OVERVIEW
    # =========================================================================
    @render.ui
    def overview_boxes_1():
        b = r_bio()
        return ui.layout_columns(
            vbox("Athlete", input.athlete(), "user", "primary"),
            vbox("School", D.cell(b, "High School / College", "—"), "school", "primary"),
            vbox("Class of", D.cell(b, "Graduating Class", "—"), "graduation-cap", "info"),
            vbox("Team", D.cell(b, "Select Team", "—"), "users", "secondary"),
            col_widths=CW4,
        )

    @render.ui
    def overview_boxes_2():
        b = r_bio()
        return ui.layout_columns(
            vbox("Position 1", D.cell(b, "Position 1", "—"), "star", "primary"),
            vbox("Position 2", D.cell(b, "Position 2", "—"), "star-half-stroke", "info"),
            vbox("Height", D.cell(b, "Height", "—"), "ruler-vertical", "success"),
            vbox("Weight (lbs)", D.cell(b, "Weight", "—"), "weight-hanging", "warning"),
            col_widths=CW4,
        )

    @render.ui
    def overview_summary():
        pit, plan, inj = r_pitching(), r_plan(), r_injuries()
        fb = D.cell(pit, "FB Velo (Max)")
        fb_max = f"{fb} mph" if fb is not None else "No data"
        inj1 = D.cell(inj, "Injury 1") or "None reported"
        return ui.TagList(
            ui.div("At a Glance", class_="section-header"),
            ui.tags.ul(
                ui.tags.li(ui.tags.b("FB Velo (Max): "), fb_max),
                ui.tags.li(ui.tags.b("Last Injury: "), inj1),
                ui.tags.li(ui.tags.b("Body Plan: "), D.cell(plan, "Body", "—")),
                ui.tags.li(ui.tags.b("Performance Plan: "), D.cell(plan, "Performance", "—")),
                ui.tags.li(ui.tags.b("Pitching Plan: "), D.cell(plan, "Pitching", "—")),
            ),
        )

    # -- Video angle selection
    selected_video = reactive.value("behind")

    @reactive.effect
    @reactive.event(input.btn_behind)
    def _(): selected_video.set("behind")

    @reactive.effect
    @reactive.event(input.btn_side)
    def _(): selected_video.set("side")

    @reactive.effect
    @reactive.event(input.btn_other)
    def _(): selected_video.set("other")

    @reactive.effect
    @reactive.event(input.athlete)
    def _(): selected_video.set("behind")

    @render.ui
    def video_player():
        b = r_bio()
        angle = selected_video()
        col_map = {"behind": "Video Behind", "side": "Video Side", "other": "Video Other"}
        vid = make_video_ui(D.cell(b, col_map[angle]))
        if vid is None:
            return ui.div(ui.p(f"No {angle} view available for {input.athlete()}."),
                          class_="video-no-data")
        return vid

    # =========================================================================
    # COMPARE ATHLETES
    # =========================================================================
    def _metric_value(db, athlete, spec):
        """Latest numeric value of one COMPARE_METRICS spec for an athlete."""
        key, name_col, date_col, col = spec
        row = A.latest_row(db[key], name_col, athlete, date_col)
        return D.safe_num(D.cell(row, col))

    @render_widget
    def compare_chart():
        chosen = list(input.cmp_athletes() or [])
        label = input.cmp_metric()
        if not chosen:
            return empty_fig(label, "Pick one or more athletes to compare")
        spec = COMPARE_METRICS[label]
        pairs = [(a, _metric_value(DB(), a, spec)) for a in chosen]
        pairs = [(a, v) for a, v in pairs if v is not None]
        if not pairs:
            return empty_fig(label, f"No {label} recorded for the selected athletes")
        pairs.sort(key=lambda p: p[1], reverse=True)
        names = [a for a, _ in pairs]
        vals = [v for _, v in pairs]
        # highlight the currently-selected sidebar athlete, if in the set
        colors = [GREEN if a == input.athlete() else PRIMARY for a in names]
        fig = go.Figure(go.Bar(
            x=vals, y=names, orientation="h", marker_color=colors,
            text=[f"{v:g}" for v in vals], textposition="outside",
            hovertemplate="%{y}: %{x}<extra></extra>"))
        fig.update_layout(
            title=f"{label} — ranked", xaxis_title=label, yaxis_title="",
            yaxis=dict(autorange="reversed"), margin=dict(t=50, l=10, r=30, b=30),
            plot_bgcolor="white", height=max(260, 46 * len(names) + 90), showlegend=False)
        return fig

    @render.data_frame
    def compare_table():
        chosen = list(input.cmp_athletes() or [])
        if not chosen:
            return render.DataGrid(pd.DataFrame({"Message": ["Select athletes to compare"]}))
        db = DB()
        table = {"Metric": list(COMPARE_METRICS.keys())}
        for a in chosen:
            col = []
            for spec in COMPARE_METRICS.values():
                v = _metric_value(db, a, spec)
                col.append("—" if v is None else f"{v:g}")
            table[a] = col
        return render.DataGrid(pd.DataFrame(table), width="100%")

    # =========================================================================
    # MOTOR
    # =========================================================================
    @render.table
    def motor_table():
        m = r_motor()
        if m.empty:
            return pd.DataFrame({"Message": ["No data"]})
        cols = [c for c in m.columns if c != "Athlete"]
        s = m.iloc[0][cols]
        return pd.DataFrame({"Category": cols,
                             "Value": [("" if D.is_blank(v) else str(v)) for v in s.values]})

    MOTOR_SPECTRA = [
        ("Associated / Disassociated", "Associated/ Disassociated", "Associated", "Disassociated"),
        ("Vertical / Horizontal", "Vertical/ Horizontal", "Vertical", "Horizontal"),
        ("Supination / Pronation", "Supination / Pronation", "Supination", "Pronation"),
        ("Axial / Large", "Axial/Large", "Axial", "Large"),
        ("Focal / Global", "Focal/Global Vision", "Focal", "Global"),
        ("Breathe In / Out", "Breathe In / Out", "In", "Out"),
        ("Red / Blue", "Red/Blue", "Red", "Blue"),
    ]

    @render_widget
    def motor_chart():
        m = r_motor()
        if m.empty:
            return empty_fig("Motor Preferences")
        rows = []
        for disp, col, left, right in MOTOR_SPECTRA:
            if col not in m.columns:
                continue
            val = D.cell(m, col)
            if val is None:
                continue
            v = str(val).lower()
            if left.lower() in v:
                x = -1
            elif right.lower() in v:
                x = 1
            else:
                continue
            rows.append({"disp": disp, "left": left, "right": right, "x": x, "text": str(val)})
        if not rows:
            return empty_fig("Motor Preferences")
        ys = list(range(len(rows)))[::-1]
        fig = go.Figure()
        for y in ys:
            fig.add_shape(type="line", x0=-1, x1=1, y0=y, y1=y,
                          line=dict(color="#e3e3e3", width=4), layer="below")
        fig.add_trace(go.Scatter(
            x=[r["x"] for r in rows], y=ys, mode="markers",
            marker=dict(size=20, color=[PRIMARY if r["x"] < 0 else GREEN for r in rows],
                        line=dict(color="white", width=2)),
            text=[r["text"] for r in rows], hovertemplate="%{text}<extra></extra>",
            showlegend=False))
        for r, y in zip(rows, ys):
            fig.add_annotation(x=-1, y=y, text=r["left"], xanchor="right", xshift=-10,
                               showarrow=False, font=dict(size=11, color="#999"))
            fig.add_annotation(x=1, y=y, text=r["right"], xanchor="left", xshift=10,
                               showarrow=False, font=dict(size=11, color="#999"))
        fig.update_yaxes(tickvals=ys, ticktext=[r["disp"] for r in rows], tickfont=dict(size=12))
        fig.update_xaxes(range=[-1.9, 1.9], showticklabels=False, zeroline=True,
                         zerolinecolor="#ccc")
        fig.update_layout(title="Motor Preference Spectrum",
                          height=max(260, 58 * len(rows)),
                          margin=dict(l=150, r=110, t=50, b=20), plot_bgcolor="white")
        return fig

    # =========================================================================
    # POWER
    # =========================================================================
    @render.ui
    def power_boxes():
        p = r_power()
        bw = D.cell(p, "Body Weight")

        def vert(col):
            v = D.cell(p, col)
            return f'{v}"' if v is not None else "—"

        return ui.layout_columns(
            vbox("Body Weight", f"{bw} lbs" if bw is not None else "—", "weight-hanging", "warning"),
            vbox("Vert (Regular)", vert("Vert- Regular"), "arrow-up", "success"),
            vbox("Vert (CMJ)", vert("Vert- CMJ"), "arrow-up", "primary"),
            vbox("Vert (Step-In)", vert("Vert- Step In"), "arrow-up", "info"),
            col_widths=CW4,
        )

    @render_widget
    def mb_chart():
        p = r_power()
        left = D.safe_num(D.cell(p, "4lb MB Shot Put (Left)"))
        right = D.safe_num(D.cell(p, "4lb MB Shot Put (Right)"))
        if left is None or right is None:
            return empty_fig("4lb MB Shot Put")
        fig = go.Figure(go.Bar(x=["Left", "Right"], y=[left, right],
                               marker_color=[PRIMARY, GREEN]))
        fig.update_layout(title="4lb MB Shot Put", yaxis_title="Distance (ft)",
                          showlegend=False)
        return fig

    @render_widget
    def vert_chart():
        p = r_power()
        names, vals = [], []
        for label, col in [("Regular", "Vert- Regular"), ("CMJ", "Vert- CMJ"),
                           ("Step-In", "Vert- Step In")]:
            v = D.safe_num(D.cell(p, col))
            if v is not None:
                names.append(label)
                vals.append(v)
        if not vals:
            return empty_fig("Vertical Jump")
        fig = go.Figure(go.Bar(x=names, y=vals, marker_color=PRIMARY))
        fig.update_layout(title="Vertical Jump", yaxis_title="Height (inches)")
        return fig

    @render_widget
    def power_trend():
        return trend_fig(r_power_all(), input.power_metric())

    @render.data_frame
    def power_table():
        p = r_power_all()
        if p.empty:
            return render.DataGrid(pd.DataFrame({"Message": ["No data"]}))
        p = p.drop(columns=[c for c in ["_date"] if c in p.columns])
        p = p.dropna(axis=1, how="all").astype(str)
        return render.DataGrid(p, width="100%")

    # =========================================================================
    # PITCHING
    # =========================================================================
    def _pitches(p):
        vals = [D.cell(p, f"Pitch {i}") for i in range(1, 6)]
        return [str(v) for v in vals if v is not None]

    @render.ui
    def pitch_boxes():
        p = r_pitching()
        mx = D.cell(p, "FB Velo (Max)")
        av = D.cell(p, "FB Velo (Avg)")
        arsenal = " · ".join(_pitches(p)) or "—"
        return ui.layout_columns(
            vbox("FB Max", f"{mx} mph" if mx is not None else "—", "baseball", "danger"),
            vbox("FB Avg", f"{av} mph" if av is not None else "—", "baseball", "warning"),
            vbox("Arsenal", arsenal, "list", "primary"),
            col_widths=CW3,
        )

    @render.ui
    def pitch_arsenal():
        pitches = _pitches(r_pitching())
        if not pitches:
            return ui.p("No data")
        return ui.tags.ul(*[ui.tags.li(x) for x in pitches])

    @render.ui
    def pitch_sw():
        p = r_pitching()
        if p.empty:
            return ui.p("No data")
        return ui.TagList(
            ui.div(ui.tags.b("Strengths: "), D.cell(p, "Strengths", "—")),
            ui.hr(),
            ui.div(ui.tags.b("Weaknesses: "), D.cell(p, "Weaknesses", "—")),
        )

    @render.ui
    def pitch_goals():
        p = r_pitching()
        if p.empty:
            return ui.p("No data")
        return ui.p(D.cell(p, "Goals", "—"))

    @render_widget
    def pitch_trend():
        h = r_pitching_all()
        if h.empty or "_date" not in h.columns:
            return empty_fig("Fastball velocity", "No dated history yet")
        d = h.dropna(subset=["_date"]).copy()
        fig = go.Figure()
        for col, color, name in [("FB Velo (Max)", RED, "FB Max"),
                                 ("FB Velo (Avg)", ORANGE, "FB Avg")]:
            if col not in d.columns:
                continue
            d["_y"] = pd.to_numeric(d[col], errors="coerce")
            dd = d.dropna(subset=["_y"])
            if not dd.empty:
                fig.add_trace(go.Scatter(x=dd["_date"], y=dd["_y"], mode="lines+markers",
                                         name=name, line=dict(color=color, width=3),
                                         marker=dict(size=10)))
        if not fig.data:
            return empty_fig("Fastball velocity", "No velocity data")
        fig.update_layout(title="Fastball velocity over time", yaxis_title="mph",
                          xaxis_title="", margin=dict(t=50, b=30), plot_bgcolor="white",
                          legend=dict(orientation="h", y=-0.2))
        return fig

    @render.data_frame
    def pitch_table():
        p = r_pitching_all()
        if p.empty:
            return render.DataGrid(pd.DataFrame({"Message": ["No data"]}))
        p = p.drop(columns=[c for c in ["_date"] if c in p.columns])
        p = p.dropna(axis=1, how="all").astype(str)
        return render.DataGrid(p, width="100%")

    # =========================================================================
    # CONTEXT
    # =========================================================================
    @render.ui
    def context_bullets():
        c = r_context()
        if c.empty:
            return ui.p("No context data")
        cols = [col for col in c.columns if col.startswith("Context")]
        items = [D.cell(c, col) for col in cols]
        items = [x for x in items if x is not None]
        if not items:
            return ui.p("No context data")
        return ui.tags.ul(*[ui.tags.li(x) for x in items])

    # =========================================================================
    # INJURIES
    # =========================================================================
    @render.data_frame
    def injury_table():
        inj = r_injuries()
        if inj.empty:
            return render.DataGrid(pd.DataFrame({"Message": ["No injury data"]}))
        rows, i = [], 1
        while all(f"Injury {i}{suf}" in inj.columns for suf in ["", " Date", " Recovery"]):
            nm = D.cell(inj, f"Injury {i}")
            if nm is not None:
                dt = D.fmt_date(D.cell(inj, f"Injury {i} Date"), "%b %Y") or "—"
                rv = D.cell(inj, f"Injury {i} Recovery") or "—"
                rows.append({"Injury": str(nm), "Date": dt, "Recovery": str(rv)})
            i += 1
        if not rows:
            return render.DataGrid(pd.DataFrame({"Message": ["No injury data"]}))
        return render.DataGrid(pd.DataFrame(rows), width="100%")

    @render_widget
    def injury_timeline():
        inj = r_injuries()
        if inj.empty:
            return empty_fig("Injury Timeline")
        pts, i = [], 1
        while all(f"Injury {i}{s}" in inj.columns for s in ["", " Date", " Recovery"]):
            nm = D.cell(inj, f"Injury {i}")
            if nm is not None:
                ts = D.to_timestamp(D.cell(inj, f"Injury {i} Date"))
                if ts is not None:
                    pts.append({"date": ts, "name": str(nm),
                                "recovery": str(D.cell(inj, f"Injury {i} Recovery") or "—")})
            i += 1
        if not pts:
            return empty_fig("Injury Timeline", "No dated injuries")
        pts.sort(key=lambda p: p["date"])
        ys = [0.5 if k % 2 == 0 else -0.5 for k in range(len(pts))]
        xs = [p["date"] for p in pts]
        fig = go.Figure()
        fig.add_shape(type="line", x0=min(xs), x1=max(xs), y0=0, y1=0,
                      line=dict(color="#ccc", width=2), layer="below")
        for p, y in zip(pts, ys):
            fig.add_shape(type="line", x0=p["date"], x1=p["date"], y0=0, y1=y,
                          line=dict(color="#e0b3ad", width=2), layer="below")
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode="markers+text",
            marker=dict(size=16, color=RED, line=dict(color="white", width=2)),
            text=[p["name"] for p in pts],
            textposition=["top center" if y > 0 else "bottom center" for y in ys],
            customdata=[[p["recovery"], p["date"].strftime("%b %Y")] for p in pts],
            hovertemplate="<b>%{text}</b><br>%{customdata[1]}"
                          "<br>Recovery: %{customdata[0]}<extra></extra>",
            showlegend=False))
        fig.update_yaxes(visible=False, range=[-1.4, 1.4])
        fig.update_xaxes(title="")
        fig.update_layout(title="Injury Timeline", height=300,
                          margin=dict(t=50, b=30), plot_bgcolor="white")
        return fig

    # =========================================================================
    # MSS
    # =========================================================================
    @render.ui
    def mss_bullets():
        m = r_mss()
        if m.empty:
            return ui.p("No MSS/Posture data")
        cols = [col for col in m.columns if col.startswith("Observation")]
        items = [D.cell(m, col) for col in cols]
        items = [x for x in items if x is not None]
        if not items:
            return ui.p("No significant observations recorded")
        return ui.tags.ul(*[ui.tags.li(x) for x in items])

    # =========================================================================
    # PLAN
    # =========================================================================
    @render.ui
    def plan_cards():
        p = r_plan()
        if p.empty:
            return ui.p("No plan data")

        def card(title, content, color):
            return ui.div(
                ui.div(ui.tags.b(title)),
                ui.p(content if content is not None else "—"),
                class_="plan-card", style=f"border-left-color:{color};")

        return ui.TagList(
            card("Body", D.cell(p, "Body"), GREEN),
            card("Performance", D.cell(p, "Performance"), PRIMARY),
            card("Pitching", D.cell(p, "Pitching"), RED),
        )

    # =========================================================================
    # ARM CARE
    # =========================================================================
    @render.ui
    def armcare_boxes():
        a = r_armcare()
        date_str = D.fmt_date(D.cell(a, "Date"), "%b %d, %Y") or "No data"
        score = D.safe_num(D.cell(a, "Total Score"))
        if score is None:
            score_disp, score_theme = "No data", "secondary"
        else:
            score_disp = str(round(score, 1))
            score_theme = "success" if score >= 100 else "warning" if score >= 80 else "danger"
        bw = D.safe_num(D.cell(a, "Body Weight"))
        grip = D.safe_num(D.cell(a, "Grip"))
        return ui.layout_columns(
            vbox("Last Tested", date_str, "calendar", "primary"),
            vbox("Total Score", score_disp, "star", score_theme),
            vbox("Body Weight", f"{bw} lbs" if bw is not None else "No data",
                 "weight-hanging", "warning"),
            vbox("Grip Strength", f"{grip} lbs" if grip is not None else "No data",
                 "hand-fist", "info"),
            col_widths=CW4,
        )

    @render.table
    def armcare_table():
        a = r_armcare()
        ir = D.safe_num(D.cell(a, "IR"))
        if a.empty or ir is None:
            return pd.DataFrame({"Message": ["No arm care data recorded yet"]})

        def fmt(x):
            return str(round(x, 1)) if x is not None else "—"

        bw = D.safe_num(D.cell(a, "Body Weight"))
        er = D.safe_num(D.cell(a, "ER"))
        scap = D.safe_num(D.cell(a, "Scaption"))
        grip = D.safe_num(D.cell(a, "Grip"))
        score = D.safe_num(D.cell(a, "Total Score"))
        return pd.DataFrame({
            "Metric": ["Body Weight", "IR", "ER", "Scaption", "Grip", "Total Score"],
            "Value": [
                f"{bw} lbs" if bw is not None else "—",
                f"{fmt(ir)} lbs", f"{fmt(er)} lbs",
                f"{fmt(scap)} lbs", f"{fmt(grip)} lbs", fmt(score),
            ],
        })

    @render_widget
    def armcare_chart():
        a = r_armcare()
        ir = D.safe_num(D.cell(a, "IR"))
        if a.empty or ir is None:
            return empty_fig("Arm Care", "No arm care data recorded yet")
        metrics = ["IR", "ER", "Grip", "Scaption"]
        benchmarks = {"IR": 50, "ER": 40, "Grip": 35, "Scaption": 35}
        actuals = [D.safe_num(D.cell(a, m)) for m in metrics]
        if all(v is None for v in actuals):
            return empty_fig("Arm Care", "No arm care data recorded yet")
        actuals_plot = [v if v is not None else 0 for v in actuals]
        bench_plot = [benchmarks[m] for m in metrics]
        score = D.safe_num(D.cell(a, "Total Score"))
        fig = go.Figure()
        fig.add_bar(x=metrics, y=actuals_plot, name="Athlete", marker_color=PRIMARY)
        fig.add_bar(x=metrics, y=bench_plot, name="Benchmark", marker_color="#aaaaaa")
        title = input.athlete()
        if score is not None:
            title += f"<br><sup>Total Score: {round(score, 1)}</sup>"
        fig.update_layout(barmode="group", title=title, yaxis_title="lbs",
                          legend=dict(orientation="h", x=0.3, y=-0.15), margin=dict(t=80))
        return fig

    @render_widget
    def armcare_trend():
        return trend_fig(r_armcare_all(), input.armcare_metric())

    # =========================================================================
    # COACHES NOTES
    # =========================================================================
    @render.ui
    def cn_boxes():
        entries = r_notes()
        n = len(entries)
        first = entries[0]["date"] if n else "—"
        last = entries[-1]["date"] if n else "—"
        return ui.layout_columns(
            vbox("Total Sessions", n, "chalkboard-user", "primary" if n else "danger"),
            vbox("First Session", first, "calendar-plus", "success"),
            vbox("Latest Session", last, "calendar-check", "info"),
            col_widths=CW3,
        )

    @render.ui
    def coaches_timeline():
        entries = r_notes()
        if not entries:
            return ui.div(ui.p(f"No coaching sessions recorded yet for {input.athlete()}."),
                          class_="cn-no-data")
        cards = []
        for e in entries:
            vid = make_video_ui(e["video"])
            children = [ui.div(e["date"], class_="cn-date"),
                        ui.p(e["notes"], class_="cn-notes")]
            if vid is not None:
                children.append(vid)
            cards.append(ui.div(ui.div(class_="cn-dot"),
                                ui.div(*children, class_="cn-card"), class_="cn-entry"))
        return ui.div(*cards, class_="cn-timeline")

    # =========================================================================
    # DATA ENTRY
    # =========================================================================
    @render.ui
    def entry_target():
        if HOSTED and not sheets.sheets_enabled():
            return ui.div(
                ui.HTML("&#9888;&#65039; <b>Saving is turned off on the shared server.</b> "
                        "New entries won't persist yet — this turns on once live "
                        "Google&nbsp;Sheets sync is enabled. You can still browse "
                        "everything and generate PDFs."),
                class_="entry-hint",
                style="background:#fff3cd;border:1px solid #ffe69c;color:#664d03;"
                      "padding:11px 13px;border-radius:8px;font-size:13.5px;",
            )
        return ui.div(
            ui.HTML(f"New entries will be saved for <b>{input.athlete()}</b> "
                    "(switch athletes in the sidebar). The master workbook is never "
                    "modified — entries are stored in an <code>entries/</code> folder."),
            class_="entry-hint",
        )

    @reactive.effect
    @reactive.event(input.submit_armcare)
    def _add_armcare():
        a = input.athlete()
        score = A.armcare_total_score(input.ac_ir(), input.ac_er(),
                                      input.ac_scap(), input.ac_grip(), input.ac_bw())
        A.add_entry("armcare", {
            "Athlete": a, "Date": str(input.ac_date()),
            "Body Weight": input.ac_bw(), "IR": input.ac_ir(), "ER": input.ac_er(),
            "Scaption": input.ac_scap(), "Grip": input.ac_grip(), "Total Score": score,
        })
        data_version.set(data_version() + 1)
        ui.notification_show(f"Arm Care test saved for {a}.", type="message", duration=4)

    @reactive.effect
    @reactive.event(input.submit_power)
    def _add_power():
        a = input.athlete()
        A.add_entry("power", {
            "Name": a, "Test Date": str(input.pw_date()),
            "Test Number": A.next_test_number(DB(), a),
            "Body Weight": input.pw_bw(), "Vert- Regular": input.pw_reg(),
            "Vert- CMJ": input.pw_cmj(), "Vert- Step In": input.pw_step(),
            "4lb MB Shot Put (Left)": input.pw_mbl(),
            "4lb MB Shot Put (Right)": input.pw_mbr(),
        })
        data_version.set(data_version() + 1)
        ui.notification_show(f"Power test saved for {a}.", type="message", duration=4)

    @reactive.effect
    @reactive.event(input.submit_pitching)
    def _add_pitching():
        a = input.athlete()
        A.add_entry("pitching", {
            "Athlete": a, "Date": str(input.pt_date()),
            "FB Velo (Max)": input.pt_max(), "FB Velo (Avg)": input.pt_avg(),
            "Pitch 1": input.pt_p1(), "Pitch 2": input.pt_p2(), "Pitch 3": input.pt_p3(),
        })
        data_version.set(data_version() + 1)
        ui.notification_show(f"Pitching entry saved for {a}.", type="message", duration=4)

    @reactive.effect
    @reactive.event(input.submit_note)
    def _add_note():
        a = input.athlete()
        text = (input.note_text() or "").strip()
        if not text:
            ui.notification_show("Please enter some notes before saving.",
                                 type="warning", duration=4)
            return
        A.add_coaches_note(a, str(input.note_date()), text,
                           (input.note_video() or "").strip() or None)
        data_version.set(data_version() + 1)
        ui.notification_show(f"Coaching note saved for {a}.", type="message", duration=4)


app = App(app_ui, server)
