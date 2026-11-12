# -*- coding: utf-8 -*-
"""Static management visuals (PNG) generated from the same simulated tables as the workbook.
Run: python3 -I build_charts.py"""
import os, sys
from collections import Counter, OrderedDict
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import config as C, build_data

OUT = os.path.join(ROOT, "Dashboard", "Visuals"); os.makedirs(OUT, exist_ok=True)
D = build_data.build(); H, A, R = D["hazards"], D["actions"], D["records"]

NAVY, SLATE, GREY, LIGHT = "#1F3A4D", "#5B7A99", "#8A949C", "#A9C3DB"
BC = {"Low": "#7FB77E", "Moderate": "#F2CB5C", "High": "#F0924A", "Extreme": "#C8342B"}
TEXTC = "#222222"
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9, "axes.edgecolor": "#BBBBBB", "axes.labelcolor": TEXTC,
                     "xtick.color": TEXTC, "ytick.color": TEXTC, "axes.spines.top": False, "axes.spines.right": False})
FOOT = "SIMULATED DATA: independent portfolio case study. Not Powergas records."


def save(fig, name, foot=True, fy=-0.05):
    if foot:
        fig.text(0.01, fy, FOOT, fontsize=6.5, color="#9B2C2C", ha="left", va="top")
    fig.savefig(os.path.join(OUT, name), dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def heat(kL, kS, title, name, note):
    cnt = Counter((h[kL], h[kS]) for h in H)
    fig, ax = plt.subplots(figsize=(6.4, 5.2))
    for l in range(1, 6):
        for s in range(1, 6):
            b = C.band(l * s)
            ax.add_patch(Rectangle((s - 1, l - 1), 1, 1, facecolor=BC[b], edgecolor="white", lw=2))
            ax.text(s - 0.5, l - 0.22, str(l * s), ha="center", va="center", fontsize=7, color="#444444")
            n = cnt.get((l, s), 0)
            if n:
                ax.text(s - 0.5, l - 0.55, str(n), ha="center", va="center", fontsize=17, fontweight="bold", color="#111111")
    ax.set_xlim(0, 5); ax.set_ylim(0, 5)
    ax.set_xticks([i + .5 for i in range(5)]); ax.set_xticklabels([f"{i}\n{C.SEVERITY[i][0]}" for i in range(1, 6)])
    ax.set_yticks([i + .5 for i in range(5)]); ax.set_yticklabels([f"{i} {C.LIKELIHOOD[i][0]}" for i in range(1, 6)])
    ax.set_xlabel("Severity"); ax.set_ylabel("Likelihood"); ax.tick_params(length=0)
    for sp in ax.spines.values(): sp.set_visible(False)
    ax.set_title(title, loc="left", fontsize=11, fontweight="bold", color=NAVY, pad=10)
    handles = [Rectangle((0, 0), 1, 1, facecolor=BC[b]) for b in C.BAND_ORDER]
    labs = [f"{b} ({lo}-{hi})" for b, lo, hi, _ in C.BANDS]
    ax.legend(handles, labs, ncol=4, loc="upper center", bbox_to_anchor=(0.5, -0.17), frameon=False, fontsize=8)
    fig.text(0.01, -0.07, note, fontsize=7, color="#555555", ha="left", va="top")
    save(fig, name, fy=-0.11)


heat("L", "S", "Initial risk: hazards by likelihood and severity (n=40)", "heatmap_initial.png",
     "Small grey figure = cell score (L x S). Large figure = number of simulated hazards in the cell.")
heat("rL", "rS", "Projected residual risk after recommended controls (n=40)", "heatmap_residual.png",
     "Projection only: valid when every recommended control is implemented and verified.")

# ---- risk distribution initial / current / residual
fig, ax = plt.subplots(figsize=(6.8, 3.8))
x = range(4); w = 0.26
series = [("Initial", "init_cat", NAVY), ("Current (control-verified)", "cur_cat", SLATE), ("Projected residual", "res_cat", LIGHT)]
for i, (lab, k, col) in enumerate(series):
    vals = [sum(1 for h in H if h[k] == b) for b in C.BAND_ORDER]
    bars = ax.bar([j + (i - 1) * w for j in x], vals, w, label=lab, color=col)
    for b_, v in zip(bars, vals): ax.text(b_.get_x() + b_.get_width() / 2, v + .3, str(v), ha="center", fontsize=8)
ax.set_xticks(list(x)); ax.set_xticklabels(C.BAND_ORDER); ax.set_ylabel("Number of hazards")
ax.set_title("Hazards by risk category: initial, current and projected residual", loc="left", fontsize=11, fontweight="bold", color=NAVY)
ax.legend(frameon=False, fontsize=8); ax.yaxis.grid(True, color="#EEEEEE"); ax.set_axisbelow(True)
save(fig, "risk_distribution.png")

# ---- hazards by category stacked by initial band
fig, ax = plt.subplots(figsize=(6.8, 3.8))
cats = C.CAT_NAMES
left = [0] * len(cats)
for b in C.BAND_ORDER:
    vals = [sum(1 for h in H if h["category"] == c and h["init_cat"] == b) for c in cats]
    ax.barh(cats, vals, left=left, color=BC[b], label=b, edgecolor="white")
    for i, v in enumerate(vals):
        if v: ax.text(left[i] + v / 2, i, str(v), ha="center", va="center", fontsize=8)
    left = [a + v for a, v in zip(left, vals)]
ax.invert_yaxis(); ax.set_xlabel("Number of hazards")
ax.set_title("Hazards by operational category and initial risk category", loc="left", fontsize=11, fontweight="bold", color=NAVY)
ax.legend(ncol=4, frameon=False, fontsize=8, loc="lower center", bbox_to_anchor=(0.4, -0.3))
save(fig, "hazards_by_category.png", fy=-0.14)

# ---- top 10 hazards initial vs residual
top = sorted(H, key=lambda h: (-h["init_score"], h["hazard_id"]))[:10]
fig, ax = plt.subplots(figsize=(7.4, 4.2))
import textwrap
labs = [h['hazard_id'] + "  " + "\n      ".join(textwrap.wrap(h['title'], 40)) for h in top]
y = range(len(top))
ax.barh([i - .2 for i in y], [h["init_score"] for h in top], .38, color=[BC[h["init_cat"]] for h in top], label="Initial score")
ax.barh([i + .2 for i in y], [h["res_score"] for h in top], .38, color="#C9D3DC", label="Projected residual")
for i, h in enumerate(top):
    ax.text(h["init_score"] + .3, i - .2, str(h["init_score"]), va="center", fontsize=8)
    ax.text(h["res_score"] + .3, i + .2, str(h["res_score"]), va="center", fontsize=8, color="#555555")
ax.set_yticks(list(y)); ax.set_yticklabels(labs, fontsize=8); ax.invert_yaxis(); ax.set_xlim(0, 24)
ax.set_xlabel("Risk score (likelihood x severity)")
ax.set_title("Ten highest initial risk scores (ties broken by hazard ID)", loc="left", fontsize=11, fontweight="bold", color=NAVY)
ax.axvline(10, color="#BBBBBB", lw=.8, ls="--"); ax.axvline(17, color="#BBBBBB", lw=.8, ls="--")
ax.text(10.2, 9.7, "High", fontsize=7, color="#777"); ax.text(17.2, 9.7, "Extreme", fontsize=7, color="#777")
ax.legend(frameon=False, fontsize=8, loc="lower right")
save(fig, "top10_hazards.png")

# ---- action status
order = ["Not Started", "In Progress", "Overdue", "Completed (Awaiting Verification)", "Closed (Verified)"]
colr = {"Not Started": GREY, "In Progress": "#E3B341", "Overdue": BC["Extreme"], "Completed (Awaiting Verification)": LIGHT, "Closed (Verified)": "#4E9A5B"}
cnt = Counter(a["status"] for a in A)
fig, ax = plt.subplots(figsize=(6.6, 3.2))
vals = [cnt.get(s, 0) for s in order]
ax.barh(order, vals, color=[colr[s] for s in order])
for i, v in enumerate(vals): ax.text(v + .3, i, str(v), va="center", fontsize=9)
ax.invert_yaxis(); ax.set_xlabel("Number of corrective actions")
ax.set_title(f"Corrective actions by status at {C.STATUS_DATE:%d %b %Y} (n={len(A)})", loc="left", fontsize=11, fontweight="bold", color=NAVY)
save(fig, "action_status.png")

# ---- overdue by role
od = Counter(a["role"] for a in A if a["status"] == "Overdue")
opn = Counter(a["role"] for a in A if a["is_open"])
roles = sorted(opn, key=lambda r: (-od.get(r, 0), -opn[r], r))
fig, ax = plt.subplots(figsize=(6.8, 3.8))
ax.barh(roles, [opn[r] for r in roles], color="#C9D3DC", label="Open actions")
ax.barh(roles, [od.get(r, 0) for r in roles], color=BC["Extreme"], label="of which overdue")
for i, r in enumerate(roles): ax.text(opn[r] + .15, i, f"{od.get(r,0)} of {opn[r]}", va="center", fontsize=8)
ax.invert_yaxis(); ax.set_xlabel("Number of open corrective actions"); ax.legend(frameon=False, fontsize=8, loc="lower right")
ax.set_title("Open and overdue actions by responsible role", loc="left", fontsize=11, fontweight="bold", color=NAVY)
save(fig, "overdue_by_role.png")

# ---- record trend by month (all types)
months = sorted({r["month"] for r in R})
types = ["Near miss", "Unsafe condition", "Unsafe act", "Inspection finding", "Incident (actual)"]
tc = {"Near miss": "#D97B29", "Unsafe condition": NAVY, "Unsafe act": GREY, "Inspection finding": "#6FA3C7", "Incident (actual)": BC["Extreme"]}
fig, ax = plt.subplots(figsize=(7.2, 3.8))
bottom = [0] * len(months)
for t in types:
    vals = [sum(1 for r in R if r["month"] == m and r["record_type"] == t) for m in months]
    ax.bar(months, vals, bottom=bottom, color=tc[t], label=t, edgecolor="white")
    bottom = [b + v for b, v in zip(bottom, vals)]
for i, b in enumerate(bottom): ax.text(i, b + .15, str(b), ha="center", fontsize=8)
ax.set_ylabel("Records per month"); ax.legend(ncol=3, frameon=False, fontsize=8, loc="upper center", bbox_to_anchor=(0.5, -0.12))
ax.set_title(f"Simulated safety records by month and type (n={len(R)})", loc="left", fontsize=11, fontweight="bold", color=NAVY)
plt.setp(ax.get_xticklabels(), rotation=0, fontsize=8)
save(fig, "record_trend.png", fy=-0.2)

# ---- stage concentration
stages = list(C.STAGES.values())
fig, ax = plt.subplots(figsize=(7.2, 3.8))
ini = [sum(h["init_score"] for h in H if h["stage"] == s) for s in stages]
res = [sum(h["res_score"] for h in H if h["stage"] == s) for s in stages]
ax.barh([i - .2 for i in range(len(stages))], ini, .38, color=NAVY, label="Sum of initial scores")
ax.barh([i + .2 for i in range(len(stages))], res, .38, color=LIGHT, label="Sum of projected residual scores")
for i in range(len(stages)):
    ax.text(ini[i] + .8, i - .2, str(ini[i]), va="center", fontsize=8); ax.text(res[i] + .8, i + .2, str(res[i]), va="center", fontsize=8)
ax.set_yticks(range(len(stages))); ax.set_yticklabels(stages, fontsize=8); ax.invert_yaxis(); ax.legend(frameon=False, fontsize=8, ncol=2, loc="upper center", bbox_to_anchor=(0.4, -0.16))
ax.set_xlabel("Sum of risk scores (not a probability or a loss estimate)")
ax.set_title("Risk concentration by operational stage", loc="left", fontsize=11, fontweight="bold", color=NAVY)
save(fig, "stage_concentration.png", fy=-0.2)

# ---- records by hazard category (for near-miss patterns)
cats = C.CAT_NAMES
fig, ax = plt.subplots(figsize=(7.2, 3.8))
left = [0] * len(cats)
for t in types:
    vals = [sum(1 for r in R if r["category"] == c and r["record_type"] == t) for c in cats]
    ax.barh(cats, vals, left=left, color=tc[t], label=t, edgecolor="white")
    left = [a + v for a, v in zip(left, vals)]
for i, v in enumerate(left): ax.text(v + .2, i, str(v), va="center", fontsize=8)
ax.invert_yaxis(); ax.set_xlabel("Simulated records"); ax.legend(ncol=3, frameon=False, fontsize=8, loc="upper center", bbox_to_anchor=(0.45, -0.14))
ax.set_title("Simulated records by operational category and type", loc="left", fontsize=11, fontweight="bold", color=NAVY)
save(fig, "records_by_category.png", fy=-0.15)
print("charts:", sorted(os.listdir(OUT)))
