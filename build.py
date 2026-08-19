#!/usr/bin/env python3
"""HOM tracker build: HOM.xlsx -> index.html (single source of truth: ~/hom/HOM.xlsx)"""
import pandas as pd, openpyxl, json, os, sys
from datetime import datetime, timedelta

def excel_weeknum(d):
    """Excel WEEKNUM(serial,1): weeks start SUNDAY, week 1 = week containing Jan 1 (matches their ChestData formula)."""
    d = d.date() if hasattr(d, "date") else d
    jan1 = datetime(d.year, 1, 1).date()
    sunday0 = jan1 - timedelta(days=(jan1.weekday() + 1) % 7)
    return 1 + (d - sunday0).days // 7

BASE = "/home/rus/hom"
SRC = os.path.join(BASE, "HOM.xlsx")
OUT = os.path.join(BASE, "index.html")

wb = openpyxl.load_workbook(SRC, data_only=True)
ws = wb["ChestData"]
rows = []
for r in ws.iter_rows(min_row=2, values_only=True):
    if r[0] is None:
        continue
    rows.append(r[:12])
cd = pd.DataFrame(rows, columns=["Date", "Clanmate", "ChestName", "ChestLevel", "ChestType", "week", "Type1", "Type2", "Points", "Valid", "name", "CorrName"])
cd["Clanmate"] = cd["CorrName"].fillna(cd["Clanmate"])
cd["Date"] = pd.to_datetime(cd["Date"], errors="coerce")
cd["Points"] = pd.to_numeric(cd["Points"], errors="coerce").fillna(0)
cd["ChestLevel"] = pd.to_numeric(cd["ChestLevel"], errors="coerce").fillna(0).astype(int)

ps = pd.DataFrame([r for r in wb["Point System"].iter_rows(min_row=2, max_col=4, values_only=True) if r[0] is not None],
                  columns=["Source", "Type", "Level", "Points"])
ps["Level"] = pd.to_numeric(ps["Level"], errors="coerce").fillna(0).astype(int)
ps["Points"] = pd.to_numeric(ps["Points"], errors="coerce").fillna(0).astype(int)

cm = pd.DataFrame([r for r in wb["ClanMates"].iter_rows(min_row=2, max_col=4, values_only=True) if r[1] is not None],
                  columns=["Status", "ClanMate", "Might", "G"])

v = cd[cd["Valid"] == 1].copy()
v["Week"] = v["Date"].map(excel_weeknum)
last_week = int(v["Week"].max())
prev_week = last_week - 1
cw = v[v["Week"] == last_week].copy()
w2 = v[v["Week"].isin([last_week, prev_week])].copy()

def bucket(t):
    t = str(t).strip().lower()
    if "common" in t: return "Common"
    if "epic" in t: return "Epic"
    if "rare" in t: return "Rare"
    if "citadel" in t: return "Citadel"
    if "hero" in t: return "Heroic"
    return "Other"
cw["Type"] = cw["ChestType"].map(bucket)
w2["Type"] = w2["ChestType"].map(bucket)

pw = int(cw["Points"].sum())
cw_chests = int(len(cw))
cw_active = int(cw["Clanmate"].nunique())
prev_points = int(v[v["Week"] == prev_week]["Points"].sum())
max_date = v["Date"].max()
week_start = max_date - pd.Timedelta(days=(max_date.weekday() + 1) % 7)  # Sunday
elapsed = (max_date - week_start).days + 1
forecast = round(pw / elapsed * 7)
fdelta = round((forecast - prev_points) / prev_points * 100) if prev_points else 0

weeks = sorted(v["Week"].unique())[-8:]
perf = v[v["Week"].isin(weeks)].groupby("Week").agg(points=("Points", "sum"), chests=("ChestName", "count"))
perf = [{"w": f"W{int(k)}", "points": int(r["points"]), "chests": int(r["chests"])} for k, r in perf.iterrows()]

cats = ["Common", "Citadel", "Rare", "Epic", "Heroic"]
cat_top = {}
for c in cats:
    sub = cw[cw["Type"] == c]
    top = sub.groupby("Clanmate").size().sort_values(ascending=False).head(5)
    cat_top[c] = [{"name": k, "n": int(n)} for k, n in top.items()]

w2_names = set(w2["Clanmate"])
lbv = cw[cw["Clanmate"].isin(w2_names)].groupby("Clanmate").agg(chests=("ChestName", "count"), points=("Points", "sum")).sort_values("points", ascending=False)
lb = [{"name": k, "chests": int(r["chests"]), "points": int(r["points"])} for k, r in lbv.iterrows()]

types = cw.groupby("Type").size().sort_values(ascending=False).reset_index()
types.columns = ["t", "n"]

gmap = dict(zip(cm["ClanMate"].astype(str).str.strip(), cm["G"].astype(str).str.strip()))
def glevel(n):
    g = gmap.get(n, "")
    return g if g in ("6", "7", "8", "9") else "Unknown"

mweeks = weeks[-6:][::-1]
wm = v[v["Week"].isin(mweeks)]
cw_pts = wm[wm["Week"] == last_week].groupby("Clanmate")["Points"].sum()
mat = []
for name in sorted(w2_names, key=lambda n: -cw_pts.get(n, 0)):
    rows_w = wm[wm["Clanmate"] == name]
    wdata = []
    for wk in mweeks:
        sub = rows_w[rows_w["Week"] == wk]
        wdata.append({"w": f"W{int(wk)}", "pts": int(sub["Points"].sum()), "chests": int(len(sub))})
    mat.append({"name": name, "g": glevel(name), "weeks": wdata})

rows_g = [{"name": n, "g": glevel(n)} for n in w2_names]
gc = pd.DataFrame(rows_g).groupby("g").size()
gc = gc.reindex(["9", "8", "7", "6", "Unknown"], fill_value=0)
g_dist = [{"g": k, "n": int(v)} for k, v in gc.items() if int(v) > 0]
unknowns = sorted([r["name"] for r in rows_g if r["g"] == "Unknown"])

out = {
    "generated": v["Date"].max().date().isoformat(),
    "last_week": last_week,
    "stats": {"points": pw, "chests": cw_chests, "active": cw_active, "prev_points": prev_points,
              "forecast": forecast, "fdelta": fdelta, "elapsed": elapsed},
    "perf": perf, "cat_top": cat_top, "leaderboard": lb, "types": types.to_dict("records"),
    "g_dist": g_dist, "unknowns": unknowns, "matrix": mat,
}
js = json.dumps(out)

# --- site template ---
TEMPLATE = os.path.join(BASE, "template.html")
html = open(TEMPLATE).read().replace("__DATA__", js)
open(OUT, "w").write(html)
print(f"built {OUT} ({len(html)} bytes) · {len(lb)} players · week {last_week}")
