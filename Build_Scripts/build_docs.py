# -*- coding: utf-8 -*-
"""Executive summary (1 page), technical methodology and data dictionary, reference register (markdown)."""
import os, sys
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE); sys.path.insert(0, HERE)
import config as C, build_data
from dictionary import DICT
from sources import SOURCES, NOT_VERIFIED
from docx_helpers import *

D = build_data.build(); H, A, R, F = D["hazards"], D["actions"], D["records"], D["facts"]
SD = C.STATUS_DATE.strftime("%d %B %Y")
ib, rb, cb = F["init_bands"], F["res_bands"], F["cur_bands"]
n_a = len(A); pct = lambda a, b: f"{100 * a / b:.1f}%"
FOOT = "SIMULATED DATA: independent portfolio case study, not Powergas records"
od_role = sorted(F["overdue_by_role"].items(), key=lambda kv: -kv[1])

# ============================================================ executive summary
d = new_doc(FOOT, margins=1.8)
d.sections[0].top_margin = Cm(1.4); d.sections[0].bottom_margin = Cm(1.4)
para(d, "INDEPENDENT PORTFOLIO CASE STUDY  |  SIMULATED DATA", bold=True, size=8, color=RED, after=0)
para(d, C.PROJECT_TITLE, bold=True, size=17, color=NAVY, after=0)
para(d, "Executive summary for management discussion  |  Status date " + SD, size=9, color=GREY, after=4)
callout(d, "All hazards, records and outcomes below are **simulated**. They show how the system works and do not describe Powergas's actual safety performance or compliance.", size=8)
para(d, "**What this is.** A working safety management tool for a CNG virtual-pipeline distributor: a scored hazard register, an action tracker that cannot close an action by date alone, an inspection checklist, an incident log and a dashboard, all driven by one Excel workbook.", size=9, after=3)
table(d, ["Measure (simulated)", "Value", "Measure (simulated)", "Value"],
      [["Hazard scenarios", len(H), "Corrective actions", n_a],
       ["Extreme / High (initial)", f"{ib['Extreme']} / {ib['High']}", "Closed and verified", f"{F['closed']} ({pct(F['closed'], n_a)})"],
       ["Extreme / High (projected residual)", f"{rb['Extreme']} / {rb['High']}", "Open / overdue", f"{F['open']} / {F['overdue']}"],
       ["Extreme / High (current, control-verified)", f"{cb['Extreme']} / {cb['High']}", "Awaiting verification", F["status_counts"]["Completed (Awaiting Verification)"]],
       ["Simulated records", f"{len(R)} ({F['rec_types']['Near miss']} near misses)", "Inspection checklist items", F["checklist_n"]]],
      [5.6, 3.2, 5.0, 3.6], size=8.5, align_center=(1, 3))
para(d, "What the simulated results say", bold=True, size=10, color=NAVY, after=2, keep=True)
for t in [f"**Risk is concentrated where people meet gas.** Loading and offloading has the most High hazards ({F['by_category']['Loading & Offloading']['hi_ex']} of {F['by_category']['Loading & Offloading']['n']}); maintenance and contractors has the only two Extreme hazards: isolation before gas-line work and hot work without adequate gas testing.",
          f"**Paper controls do not remove risk.** Projected residual risk is {pct(F['sum_init'] - F['sum_res'], F['sum_init'])} lower than initial, but only {pct(F['sum_init'] - F['sum_cur'], F['sum_init'])} lower today, because most recommended controls are not yet verified as closed. {rb['High']} scenarios stay High even after controls.",
          f"**The action backlog is the main weakness.** {F['overdue']} of {F['open']} open actions are overdue (average {F['avg_days_overdue']:.0f} days); the {od_role[0][0]} holds {od_role[0][1]} of them.",
          f"**Near misses signal more than outcomes.** All {F['rec_types']['Incident (actual)']} simulated incidents were minor, yet {F['rec_high_potential']} of {len(R)} records had a potential severity of 4 or 5."]:
    bullet(d, t).paragraph_format.space_after = Pt(2)
para(d, "Priorities for management", bold=True, size=10, color=NAVY, after=2, keep=True)
table(d, ["Priority", "Action", "Owner", "By"],
      [["P1", "Prove isolation and gate hot work on a recorded gas test (R1)", "Maintenance Mgr / HSE Mgr", "Day 30 interim, 90 full"],
       ["P1", "Weekly action review; clear or re-baseline overdue items; verify completed work within 14 days (R2)", "HSE Manager", "Day 60"],
       ["P2", "Transfer-area barriers, verified connections and hose register (R3); road-transport fatigue and rollover controls (R4); integrity register (R5)", "Station Supervisor, Fleet Mgr, Asset Integrity Eng.", "Day 90"],
       ["P3", "CNG scenario drills (R6); contractor induction gate (R7); close the loop on near misses (R8)", "Emergency Coord., Contracts Mgr, HSE Mgr", "Day 60 to 90"]],
      [1.4, 9.2, 4.2, 2.6], size=8, align_center=(0,))
para(d, "**Limits.** Scores rank scenarios and are not probabilities. The system supports, and does not replace, competent engineering review of pressure equipment, relief devices, hoses and trailers. "
        "No regulatory compliance or ISO 45001 conformity is claimed. Real use needs real data, a calibrated risk team and legal confirmation of applicable Nigerian requirements.", size=8.5, after=0)
d.save(os.path.join(ROOT, "Reports", "Executive_Summary.docx"))

# ============================================================ methodology + dictionary
d = new_doc(FOOT)
para(d, "TECHNICAL METHODOLOGY AND DATA DICTIONARY", bold=True, size=18, color=NAVY, after=2)
para(d, C.PROJECT_TITLE + "  |  Version 1.0  |  Status date " + SD, size=9.5, color=GREY)
callout(d, "All data is simulated for an independent portfolio case study. Scores and statuses are calculated by formulas in the workbook; this document explains the logic so that any KPI can be traced.")
h(d, "1  Scoring logic")
kv_table(d, [("Risk score", "Likelihood (1 to 5) x Severity (1 to 5). Range 1 to 25."),
             ("Categories", "; ".join(f"{n} {lo}-{hi}" for n, lo, hi, _ in C.BANDS) + ". Thresholds are stored once (Risk Assessment Matrix sheet) and read by every formula through named cells."),
             ("Initial risk", "Rated with existing, assumed controls in place."),
             ("Residual risk (projected)", "Rated assuming every recommended control is implemented, working and verified. Limited by the credit table below; the workbook shows OK or REVIEW for each hazard."),
             ("Current risk", "Residual score if every linked action is Closed (Verified); otherwise the initial score. This stops an unimplemented control from reducing risk."),
             ("Priority", "From the initial category: Extreme P1, High P2, Moderate P3, Low P4."),
             ("Aggregates", "Sums and averages of scores are orientation tools only and are not probabilities or loss estimates.")], widths=(3.8, 13.2))
table(d, ["Strongest control", "Max likelihood reduction", "Max severity reduction"],
      [[k, v[0], v[1]] for k, v in C.HIER_CREDIT.items()], [5, 6, 6], size=9, align_center=(1, 2))
h(d, "2  Action status logic (evaluated at the status date)")
table(d, ["Condition (tested in this order)", "Calculated status"],
      [["Completion date AND evidence reference AND verification date present", "Closed (Verified)"],
       ["Completion date present, verification incomplete", "Completed (Awaiting Verification)"],
       ["No completion date and target date earlier than the status date", "Overdue (days overdue = status date minus target date)"],
       ["No completion date, recorded status In Progress", "In Progress"],
       ["Otherwise", "Not Started"]], [11, 6], size=9)
para(d, "Open actions are all actions that are not Closed (Verified). Percentage closed = Closed (Verified) / all actions. A passed target date changes status to Overdue only; it never closes an action. A Status Check column flags records whose recorded status disagrees with the calculated one.")
h(d, "3  Dashboard filters")
para(d, "Five filters (operational category, initial risk level, responsible role, action status, date range) drive hidden filter flags on each data sheet; every card, chart and heat map counts only rows whose flag is 1. The action-status filter applies to action figures only. Action and record rows inherit category and risk level from their linked hazard. A heat map selector switches between initial and residual ratings.")
h(d, "4  Reproducing the project")
for t in ["Run `python3 -I build_data.py` to regenerate the tables and facts (fixed seed).",
          "Run `python3 -I build_workbook.py` to rebuild the workbook; formulas are recalculated in LibreOffice and the results cached.",
          "Run `build_charts.py`, `build_powerbi.py`, `build_report.py` and `build_docs.py` for visuals, the Power BI pack and documents.",
          "To use real data: replace the simulated rows in the input (blue) columns; never edit the grey formula columns; set the status date on the Settings sheet."]:
    bullet(d, t)
h(d, "5  Data dictionary")
para(d, "Input = typed by the user. Calculated = formula; do not overwrite.", size=9)
cur = None
rows = []
for sheet, field, kind, desc, rule in DICT:
    rows.append([sheet, field, kind, desc, rule])
table(d, ["Sheet", "Field", "Type", "Description", "Rule or allowed values"], rows, [3.0, 3.0, 1.6, 5.4, 4.0], size=7, zebra=True)
h(d, "6  Source verification")
para(d, "The reference sources, their working links, dates and verification notes are in Research_and_References/Reference_Register.md and the workbook sheet Reference Sources.")
d.save(os.path.join(ROOT, "Reports", "Technical_Methodology_and_Data_Dictionary.docx"))

# ============================================================ reference register
L = ["# Reference register", "", "Independent portfolio case study (simulated data). Each source was opened between 3 and 4 November 2026. "
     "'Mandatory' means a Nigerian instrument whose applicability to a given licence still needs legal confirmation. Everything else is voluntary guidance or context.", ""]
for s in SOURCES:
    L += [f"## {s['id']}  {s['name']}", f"- **Issuer:** {s['issuer']}", f"- **Status:** {s['status']}", f"- **Date:** {s['date']}", f"- **Link:** {s['url']}",
          f"- **Why relevant:** {s['relevance']}", f"- **What was verified:** {s['verified']}", ""]
L += ["## Not verified (and therefore not asserted)", ""] + [f"- {t}" for t in NOT_VERIFIED]
open(os.path.join(ROOT, "Research_and_References", "Reference_Register.md"), "w", encoding="utf-8").write("\n".join(L) + "\n")
print("docs built")
