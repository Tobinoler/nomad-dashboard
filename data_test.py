import re
from datetime import date, datetime
from pathlib import Path
import pandas as pd
EXCEL_PATH = Path("Master Evaluation Sheet.xlsx")
_SHEETS = {"bio":("Bio","Player Name"),"motor":("Motor Preferences","Athlete"),"power":("Power Testing","Name"),"pitching":("Pitching","Athlete"),"armcare":("Arm Care","Athlete"),"context":("Context","Athlete"),"injuries":("Injuries","Athlete"),"mss":("MSSPosture","Athlete"),"plan":("Athlete Plan","Athlete"),"notes":("Coaches Notes","Name")}
def _clean(df):
    df=df.copy(); df.columns=[str(c).strip() for c in df.columns]; return df
def load_data(path=EXCEL_PATH):
    xl=pd.ExcelFile(path); db={}
    for k,(s,idc) in _SHEETS.items():
        df=_clean(pd.read_excel(xl,sheet_name=s))
        if idc in df.columns: df=df[df[idc].notna()].reset_index(drop=True)
        db[k]=df
    return db
def athlete_names(db): return [str(n) for n in db["bio"]["Player Name"].dropna().tolist()]
def athlete_row(df,c,a):
    if c not in df.columns: return df.iloc[0:0]
    return df[df[c].astype(str)==str(a)]
_NA={"","na","n/a","nan","nat","#div/0!","none"}
def is_blank(x):
    if x is None: return True
    try:
        if pd.isna(x): return True
    except (TypeError,ValueError): pass
    return str(x).strip().lower() in _NA
def clean_str(x): return None if is_blank(x) else str(x).strip()
def safe_num(x):
    if is_blank(x): return None
    try: return float(x)
    except (TypeError,ValueError): return None
def cell(df,col,default=None):
    if df is None or df.empty or col not in df.columns: return default
    v=df[col].iloc[0]; return default if is_blank(v) else v
def fmt_date(x,out="%B %d, %Y"):
    if is_blank(x): return None
    if isinstance(x,(datetime,date,pd.Timestamp)): return pd.Timestamp(x).strftime(out)
    s=str(x).strip()
    try:
        n=float(s)
        if n>40000: return (pd.Timestamp("1899-12-30")+pd.Timedelta(days=n)).strftime(out)
        if 1900<=n<=2100: return str(int(n))
    except (TypeError,ValueError): pass
    p=pd.to_datetime(s[:10],errors="coerce")
    return p.strftime(out) if pd.notna(p) else s
def to_timestamp(x):
    if is_blank(x): return None
    if isinstance(x,(datetime,date,pd.Timestamp)): return pd.Timestamp(x)
    s=str(x).strip()
    try:
        n=float(s)
        if n>40000: return pd.Timestamp("1899-12-30")+pd.Timedelta(days=n)
        if 1900<=n<=2100: return pd.Timestamp(int(n),1,1)
    except (TypeError,ValueError): pass
    if s.lower()=="current": return pd.Timestamp.today().normalize()
    p=pd.to_datetime(s[:10],errors="coerce"); return p if pd.notna(p) else None
def parse_notes(df,a):
    r=df[df["Name"].astype(str)==str(a)]
    if r.empty: return []
    vals=list(r.iloc[0].values); e=[]; i=1
    while i+2<=len(vals)-1:
        d=fmt_date(vals[i]); n=clean_str(vals[i+1]); v=clean_str(vals[i+2])
        if d is not None or n is not None:
            e.append({"date":d or "Date unknown","notes":n or "No notes recorded.","video":v})
        i+=3
    return e
def drive_embed_url(v):
    if is_blank(v): return None
    v=str(v)
    if "drive.google.com" not in v.lower(): return None
    m=re.search(r"/d/([^/?]+)",v) or re.search(r"id=([^&]+)",v)
    return f"https://drive.google.com/file/d/{m.group(1)}/preview" if m else None
