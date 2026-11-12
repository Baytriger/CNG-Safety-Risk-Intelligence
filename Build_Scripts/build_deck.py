# -*- coding: utf-8 -*-
"""10-slide management presentation (python-pptx). Numbers from build_data."""
import os, sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE); sys.path.insert(0, HERE)
import config as C, build_data

D = build_data.build(); H, A, R, F = D["hazards"], D["actions"], D["records"], D["facts"]
VIS = os.path.join(ROOT, "Dashboard", "Visuals")
NAVY, GREY, RED, WHITE = RGBColor(0x1F, 0x3A, 0x4D), RGBColor(0x55, 0x55, 0x55), RGBColor(0x9B, 0x2C, 0x2C), RGBColor(255, 255, 255)
BANDRGB = {"Low": "7FB77E", "Moderate": "F2CB5C", "High": "F0924A", "Extreme": "C8342B"}
pct = lambda a, b: f"{100 * a / b:.1f}%"
ib, rb, cb = F["init_bands"], F["res_bands"], F["cur_bands"]
n_a = len(A); od_role = sorted(F["overdue_by_role"].items(), key=lambda kv: -kv[1])
bc = F["by_category"]

prs = Presentation(); prs.core_properties.author = 'Moshood'; prs.core_properties.last_modified_by = 'Moshood'; prs.core_properties.title = 'CNG Safety Risk Intelligence'; prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
blank = prs.slide_layouts[6]
FONT = "Arial"


def tb(s, x, y, w, h, text, size=14, bold=False, color=None, align=None, anchor=None):
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = box.text_frame; tf.word_wrap = True
    if anchor: tf.vertical_anchor = anchor
    lines = text if isinstance(text, list) else [text]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run(); r.text = ln; r.font.size = Pt(size); r.font.bold = bold; r.font.name = FONT
        r.font.color.rgb = color or RGBColor(0x22, 0x22, 0x22)
        if align: p.alignment = align
        p.space_after = Pt(4)
    return box


def bullets(s, x, y, w, h, items, size=14):
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = box.text_frame; tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        head, _, rest = it.partition("|")
        r = p.add_run(); r.text = "▪ " + head; r.font.bold = bool(rest); r.font.size = Pt(size); r.font.name = FONT; r.font.color.rgb = NAVY if rest else RGBColor(0x22, 0x22, 0x22)
        if rest:
            r2 = p.add_run(); r2.text = rest; r2.font.size = Pt(size); r2.font.name = FONT; r2.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
        p.space_after = Pt(8)


def slide(title, subtitle=None):
    s = prs.slides.add_slide(blank)
    bar = s.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(0.12)); bar.fill.solid(); bar.fill.fore_color.rgb = NAVY; bar.line.fill.background()
    tb(s, 0.6, 0.35, 12.1, 0.8, title, 26, True, NAVY)
    if subtitle: tb(s, 0.6, 1.05, 12.1, 0.5, subtitle, 13, False, GREY)
    tb(s, 0.6, 7.05, 10, 0.3, "SIMULATED DATA: independent portfolio case study, not Powergas records", 9, False, RED)
    tb(s, 12.0, 7.05, 0.8, 0.3, str(len(prs.slides)), 9, False, GREY, PP_ALIGN.RIGHT)
    return s


def pic(s, name, x, y, w=None, h=None):
    kw = {}
    if w: kw["width"] = Inches(w)
    if h: kw["height"] = Inches(h)
    s.shapes.add_picture(os.path.join(VIS, name), Inches(x), Inches(y), **kw)


def table(s, x, y, w, header, rows, colw, size=11, rowh=0.38):
    gt = s.shapes.add_table(len(rows) + 1, len(header), Inches(x), Inches(y), Inches(w), Inches(rowh * (len(rows) + 1))).table
    for i, wd in enumerate(colw): gt.columns[i].width = Inches(wd)
    for ci, t in enumerate(header):
        c = gt.cell(0, ci); c.text = t; c.fill.solid(); c.fill.fore_color.rgb = NAVY
        for p in c.text_frame.paragraphs:
            for r in p.runs: r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT
    for ri, row in enumerate(rows, 1):
        for ci, t in enumerate(row):
            c = gt.cell(ri, ci); c.text = str(t); c.fill.solid()
            c.fill.fore_color.rgb = RGBColor.from_string(BANDRGB[t]) if t in BANDRGB else (RGBColor(0xF3, 0xF6, 0xF8) if ri % 2 == 0 else WHITE)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in c.text_frame.paragraphs:
                for r in p.runs: r.font.size = Pt(size); r.font.name = FONT; r.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    return gt


# 1 title
s = prs.slides.add_slide(blank)
bg = s.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height); bg.fill.solid(); bg.fill.fore_color.rgb = NAVY; bg.line.fill.background()
tb(s, 0.9, 1.5, 11.5, 0.4, "INDEPENDENT PORTFOLIO CASE STUDY", 13, True, RGBColor(0xF0, 0x92, 0x4A))
tb(s, 0.9, 2.0, 11.5, 1.8, C.PROJECT_TITLE, 40, True, WHITE)
tb(s, 0.9, 3.9, 11.5, 0.9, "A scored hazard register, action tracker and dashboard for CNG virtual-pipeline distribution", 18, False, RGBColor(0xD0, 0xDA, 0xE3))
tb(s, 0.9, 5.4, 11.5, 1.2, ["Framed around Powergas Global Investment Nigeria Limited and Powergas Ebedei Limited",
                            "All data in this presentation is simulated. No Powergas records were used. Status date " + C.STATUS_DATE.strftime("%d %B %Y") + "."], 12, False, RGBColor(0xD0, 0xDA, 0xE3))

# 2 question and system
s = slide("Can hazards be found, ranked and closed before an incident?", "A working system rather than a report: every figure comes from one workbook")
table(s, 0.6, 1.8, 12.1, ["Component", "What it does", "Simulated size"],
      [["Hazard and risk register", "Scores each scenario on a 5 x 5 matrix; initial, projected residual and current risk", f"{len(H)} hazards, 7 categories"],
       ["Corrective action tracker", "Calculates status; a passed date makes an action Overdue and never Closed", f"{n_a} actions"],
       ["Inspection checklist", "Observation-only checks; technical items flagged for a competent person", f"{F['checklist_n']} items, 8 areas"],
       ["Incident and near-miss log", "Keeps near misses, unsafe acts and conditions, findings and incidents separate", f"{len(R)} records"],
       ["Dashboard", "Five filters; KPI cards, heat map and charts derive from the data", "Excel (live); Power BI import pack"]],
      [3.2, 6.6, 2.3], size=12, rowh=0.62)
tb(s, 0.6, 5.7, 12.1, 0.9, "Not claimed: regulatory compliance, ISO 45001 conformity, accident prediction or financial savings.", 13, True, RED)

# 3 method
s = slide("Method: one scale, and residual risk must be earned", "Risk score = likelihood x severity; boundaries stored once and read by every sheet")
table(s, 0.6, 1.8, 6.0, ["Category", "Score", "Response"],
      [["Low", "1-4", "Routine management"], ["Moderate", "5-9", "Planned, 90 days"], ["High", "10-16", "Senior attention, 60 days"], ["Extreme", "17-25", "Interim controls now"]],
      [1.6, 1.0, 3.4], size=12, rowh=0.5)
bullets(s, 7.0, 1.7, 5.8, 4.8, ["Initial|: rated with existing, assumed controls.",
                                "Residual|: projection if every recommended control is in place and verified.",
                                "Current|: residual only when all linked actions are Closed (Verified); otherwise initial.",
                                "Credit limits|: administrative controls may cut likelihood by 1 level and severity by 0; engineering by 2 and 1.",
                                "Specialist checks|: 13 hazards need competent technical verification."], 14)
tb(s, 0.6, 4.4, 6.0, 1.6, ["Hierarchy of controls: elimination, substitution, engineering, administrative, PPE (source R10)."], 12, False, GREY)

# 4 heat maps
s = slide("Initial versus projected residual risk", f"Initial: {ib['Extreme']} Extreme, {ib['High']} High.  Projected residual: {rb['Extreme']} Extreme, {rb['High']} High.  Current (verified): {cb['Extreme']} Extreme, {cb['High']} High")
pic(s, "heatmap_initial.png", 0.5, 1.7, h=4.9); pic(s, "heatmap_residual.png", 6.9, 1.7, h=4.9)

# 5 concentration
s = slide("Risk concentrates where people, vehicles and gas meet",
          f"Loading and offloading: {bc['Loading & Offloading']['hi_ex']} of {bc['Loading & Offloading']['n']} High; maintenance and contractors holds both Extreme hazards")
pic(s, "hazards_by_category.png", 0.5, 1.7, w=6.4); pic(s, "stage_concentration.png", 6.9, 1.7, w=6.0)
tb(s, 0.6, 5.9, 12.1, 0.9, "Counts depend on how scenarios were split. Sums of scores show where scenarios cluster; they are not probabilities or loss estimates.", 12, False, GREY)

# 6 top hazards
s = slide("Highest scores: isolation and hot work stand apart", "Both are likely events with a catastrophic worst case in a gas facility, and both rely on people following permits")
pic(s, "top10_hazards.png", 0.5, 1.7, w=7.6)
bullets(s, 8.4, 1.9, 4.5, 4.8, ["H-025 and H-026|: score 20 (Extreme); stay High after controls (10 and 15).",
                                "Next three|: exclusion zone, driver fatigue, laden-trailer collision (16, 16, 15).",
                                f"{F['specialist_count']} of {len(H)}|: need a competent person's verification."], 13)

# 7 actions
s = slide("Action backlog is the main weakness",
          f"{F['closed']} of {n_a} closed and verified ({pct(F['closed'], n_a)}); {F['overdue']} of {F['open']} open actions overdue; {F['status_counts']['Completed (Awaiting Verification)']} awaiting verification")
pic(s, "action_status.png", 0.5, 1.8, w=6.2); pic(s, "overdue_by_role.png", 6.9, 1.7, w=6.0)
tb(s, 0.6, 5.5, 12.1, 1.3, [f"Average overdue action: {F['avg_days_overdue']:.0f} days late. {od_role[0][0]} holds {od_role[0][1]} overdue actions.",
                            f"{F['completed_late']} of {F['completed_all']} completed actions finished after target."], 13)

# 8 records
s = slide("Near misses and conditions say more than incident outcomes", f"{len(R)} simulated records; all {F['rec_types']['Incident (actual)']} incidents minor, but {F['rec_high_potential']} records had potential severity 4 or 5")
pic(s, "record_trend.png", 0.5, 1.7, w=7.0)
bullets(s, 7.8, 1.9, 5.1, 4.8, [f"Near misses|: {F['rec_types']['Near miss']}, mostly caught during transfer.",
                                "Unsafe conditions|: gas-detector upkeep, damaged hoses, blocked access.",
                                f"Follow-up|: only {F['rec_followups']} records linked to an action.",
                                "Limit|: no trend or prediction is claimed at this sample size."], 13)

# 9 recommendations
s = slide("Priorities for management", "Every recommendation has an owner, a success indicator and a verification method in the report")
table(s, 0.6, 1.7, 12.1, ["Priority", "Recommendation", "Hierarchy", "Owner", "Window"],
      [["P1", "R1 Prove isolation; gate hot work on a recorded gas test", "Substitution, engineering", "Maintenance Mgr / HSE Mgr", "30 / 90 days"],
       ["P1", "R2 Weekly action review; verify completed work in 14 days", "Administrative", "HSE Manager", "60 days"],
       ["P2", "R3 Transfer-area barriers, verified connections, hose register", "Engineering", "Station Supervisor", "90 days"],
       ["P2", "R4 Fatigue, journey plans, rollover and pre-trip controls", "Engineering, admin.", "Fleet Manager", "90 days"],
       ["P2", "R5 Integrity register; time-limited trip bypass control", "Engineering, admin.", "Asset Integrity Eng.", "60 days"],
       ["P3", "R6 CNG scenario drills; R7 contractor gate; R8 near-miss loop", "Administrative, PPE", "Emergency Coord., Contracts, HSE", "60-90 days"]],
      [0.9, 5.7, 2.2, 2.4, 0.9 + 0.0], size=11, rowh=0.62)

# 10 roadmap and limits
s = slide("90-day roadmap and limits", f"Projected residual risk is {pct(F['sum_init'] - F['sum_res'], F['sum_init'])} below initial; only {pct(F['sum_init'] - F['sum_cur'], F['sum_init'])} is achieved today")
table(s, 0.6, 1.7, 7.4, ["Window", "Focus"],
      [["Days 1-30", "Action review starts; permit audit; hose register; hours limits; integrity register; contractor gate"],
       ["Days 31-60", "Backlog cleared or re-baselined; isolation design review; trip bypass control; first drill"],
       ["Days 61-90", "Isolation and permit changes live; barriers; digital pre-trip checks; responder liaison; first KPI review"]],
      [1.6, 5.8], size=11, rowh=0.9)
bullets(s, 8.4, 1.7, 4.5, 5.0, ["Needed before real use|: real data, a calibrated risk team, competent engineering review and legal confirmation of Nigerian requirements.",
                                "Assumption|: projection holds only if every control is delivered and verified.",
                                "Files|: workbook, report, dashboard, Power BI pack and README are in the project folder."], 12)

out = os.path.join(ROOT, "Presentation", "CNG_Safety_Management_Presentation.pptx")
prs.save(out); print("saved", out, len(prs.slides))
