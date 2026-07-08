# Athlete Dashboard — Shiny for Python

A Python port of the original R/Shiny athlete dashboard.

## Setup

```bash
cd NomadDashboard
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Place `Master Evaluation Sheet.xlsx` in this folder (same directory as `app.py`).

## Run

```bash
shiny run --reload app.py
```

Then open http://127.0.0.1:8000 in your browser.

## Layout

- **app.py** — UI + server (all 10 tabs).
- **data.py** — Excel loading, cleaning, date parsing, notes parsing, video-link helpers.
- **www/videos/** — optional folder for local `.mp4` files referenced by filename
  (Google Drive share links are embedded automatically).

## Tabs

Overview, Motor Preferences, Power Testing, Pitching, Context, Injuries,
MSS / Posture, Athlete Plan, Arm Care, Coaches Notes — mirroring the R app.

## Notes on the port

- `shinydashboard` info boxes → Shiny for Python `value_box`.
- `renderPlotly` → `shinywidgets` + Plotly (same figures).
- `DTOutput` → `render.data_frame` / `DataGrid`.
- Google Drive `/view` links are converted to embeddable `/preview` URLs, same
  as the R `make_video_ui`.
- Excel dates already parse as real datetimes via pandas, so the R serial-date
  fallback is kept only for string/serial edge cases.
```
