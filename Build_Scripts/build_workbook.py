# -*- coding: utf-8 -*-
"""Build the Excel risk-assessment workbook with XlsxWriter.
Two passes: pass 1 writes formulas; LibreOffice calculates a scratch copy; pass 2 rewrites the workbook with the
calculated values stored beside the live formulas (so any viewer shows numbers; Excel still recalculates).
Run: python3 -I build_workbook.py"""
import os, subprocess, sys, shutil
from datetime import date, datetime
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import config as C
import build_data
from dictionary import DICT
from sources import SOURCES, NOT_VERIFIED

OUT = os.path.join(ROOT, "Data_and_Excel", "CNG_Safety_Risk_Assessment_Workbook.xlsx")
import tempfile
SCRATCH = tempfile.mkdtemp(prefix="cng_build_")

NAVY, SLATE, GREY, LIGHT, BORDER = "#1F3A4D", "#5B6B7A", "#F2F2F2", "#E8EEF3", "#BFBFBF"
BLUE = "#0000FF"
HRN, TRN, LGN, CLN = "'Hazard Register'!", "'Corrective Action Tracker'!", "'Incident and Near-Miss Log'!", "'Safety Inspection Checklist'!"
MXN, KPN = "'Risk Assessment Matrix'!", "'KPI Dashboard Data'!"
LASTROW = 500


def hr(col):  return f"{HRN}${col}$6:${col}${LASTROW}"
def tr(col):  return f"{TRN}${col}$6:${col}${LASTROW}"
def lg(col):  return f"{LGN}${col}$6:${col}${LASTROW}"


def build_workbook(path, D, cache):
    wb = xlsxwriter.Workbook(path, {"default_format_properties": {"font_name": "Arial", "font_size": 10}})
    wb.set_properties({"author": "Moshood", "company": "Moshood", "comments": "Simulated data for an independent portfolio case study.", "title": C.PROJECT_TITLE, "subject": "Independent portfolio case study: hypothetical data",
                       "author": "Moshood", "comments": C.DISCLOSURE})
    _f = {}

    def fm(**kw):
        key = tuple(sorted(kw.items()))
        if key not in _f:
            base = {"font_name": "Arial", "font_size": 10, "valign": "top"}
            base.update(kw)
            _f[key] = wb.add_format(base)
        return _f[key]

    title = fm(bold=True, font_size=15, font_color=NAVY)
    sub = fm(italic=True, font_size=9, font_color="#9C0006", bg_color="#FDECEA", text_wrap=True, valign="vcenter")
    h_in = fm(bold=True, font_color="#FFFFFF", bg_color=NAVY, border=1, border_color=BORDER, text_wrap=True, valign="vcenter", align="center")
    h_calc = fm(bold=True, font_color="#FFFFFF", bg_color=SLATE, border=1, border_color=BORDER, text_wrap=True, valign="vcenter", align="center")
    sect = fm(bold=True, font_size=11, font_color="#FFFFFF", bg_color=NAVY)
    t_wrap = fm(text_wrap=True, border=1, border_color=BORDER)
    t_ctr = fm(text_wrap=True, border=1, border_color=BORDER, align="center")
    n_in = fm(border=1, border_color=BORDER, align="center", font_color=BLUE)
    d_in = fm(border=1, border_color=BORDER, align="center", font_color=BLUE, num_format="dd-mmm-yyyy")
    s_in = fm(border=1, border_color=BORDER, align="center", font_color=BLUE, text_wrap=True)
    c_num = fm(border=1, border_color=BORDER, align="center", bg_color=GREY)
    c_txt = fm(border=1, border_color=BORDER, align="center", bg_color=GREY, text_wrap=True)
    c_left = fm(border=1, border_color=BORDER, bg_color=GREY, text_wrap=True)
    note = fm(font_size=9, italic=True, font_color="#404040", text_wrap=True)
    bold = fm(bold=True)
    link = fm(font_color="#0563C1", underline=1, text_wrap=True, border=1, border_color=BORDER)

    def cf(bg, fg="#000000", b=False):
        return wb.add_format({"bg_color": bg, "font_color": fg, "bold": b})
    CF = {"Extreme": cf("#C00000", "#FFFFFF", True), "High": cf("#F4B183"), "Moderate": cf("#FFE699"), "Low": cf("#C6E0B4")}

    names = ["Read Me", "Dashboard", "Hazard Register", "Risk Assessment Matrix", "Corrective Action Tracker",
             "Safety Inspection Checklist", "Incident and Near-Miss Log", "KPI Dashboard Data", "Data Dictionary",
             "Reference Sources", "Settings"]
    W = {n: wb.add_worksheet(n) for n in names}
    tabs = {"Read Me": "#7F7F7F", "Dashboard": "#C00000", "Hazard Register": NAVY, "Risk Assessment Matrix": NAVY,
            "Corrective Action Tracker": NAVY, "Safety Inspection Checklist": NAVY, "Incident and Near-Miss Log": NAVY,
            "KPI Dashboard Data": SLATE, "Data Dictionary": "#7F7F7F", "Reference Sources": "#7F7F7F", "Settings": "#7F7F7F"}
    for n, w in W.items():
        w.set_tab_color(tabs[n]); w.set_paper(9); w.set_landscape(); w.fit_to_pages(1, 0)

    def wf(ws, cell, formula, fmt=None):
        v = cache.get((ws.get_name(), cell), 0)
        ws.write_formula(cell, formula, fmt, "" if v is None else v)

    def banner(ws, last_col="L", text=None):
        ws.write("A1", text or C.PROJECT_TITLE, title)
        ws.merge_range(f"A2:{last_col}2", C.DISCLOSURE, sub)
        ws.set_row(1, 28)

    # ============================================================ SETTINGS
    ws = W["Settings"]
    ws.write("B2", "Settings and lists", title)
    ws.write("B4", "Status date (simulated)", bold)
    ws.write_datetime("C4", datetime.combine(C.STATUS_DATE, datetime.min.time()), d_in)
    ws.merge_range("D4:L4", "Overdue logic is evaluated against this date so that published figures are reproducible. For live use, replace with =TODAY().", note)
    wb.define_name("StatusDate", "=Settings!$C$4")
    lists = [("B", "Operational category", ["All"] + C.CAT_NAMES), ("D", "Risk level", ["All"] + C.BAND_ORDER),
             ("F", "Responsible role", ["All"] + C.ROLES), ("H", "Action status", ["All"] + C.COMPUTED_STATUSES),
             ("J", "Control hierarchy", C.HIER), ("L", "Recorded status", C.ACTION_STATUSES),
             ("N", "Record type", C.RECORD_TYPES), ("P", "Checklist response", ["Compliant", "Non-compliant", "Not Applicable"])]
    for col, head, items in lists:
        ws.write(f"{col}7", head, h_in)
        for i, it in enumerate(items):
            ws.write(f"{col}{8 + i}", it, t_wrap)
        ws.set_column(f"{col}:{col}", 30)
    ws.set_column("A:A", 3); ws.set_column("C:C", 16); ws.set_column("E:E", 3); ws.set_column("G:G", 3); ws.set_column("I:I", 3)
    ws.set_column("K:K", 3); ws.set_column("M:M", 3); ws.set_column("O:O", 3)
    SL = {  # list ranges
        "cat_all": "=Settings!$B$8:$B$15", "cat": "=Settings!$B$9:$B$15", "lvl_all": "=Settings!$D$8:$D$12",
        "role_all": "=Settings!$F$8:$F$18", "role": "=Settings!$F$9:$F$18", "stat_all": "=Settings!$H$8:$H$13",
        "hier": "=Settings!$J$8:$J$12", "rec": "=Settings!$L$8:$L$11", "rtype": "=Settings!$N$8:$N$12", "resp": "=Settings!$P$8:$P$10"}

    # ============================================================ RISK ASSESSMENT MATRIX
    ws = W["Risk Assessment Matrix"]
    banner(ws, "G")
    ws.write("A3", "Risk Assessment Matrix and scoring rules. Risk Score = Likelihood x Severity. Blue cells are the only scoring parameters; every other sheet reads them.", note)
    ws.set_column("A:A", 20); ws.set_column("B:B", 16); ws.set_column("C:G", 30)
    # 1 likelihood
    ws.merge_range("A4:G4", "1. Likelihood scale (qualitative; indicative only, not statistical frequencies)", sect)
    for c, h in enumerate(["Level", "Label", "Description"]):
        ws.write(4, c, h, h_in)
    ws.merge_range("C5:G5", "Description", h_in)
    for i in range(1, 6):
        ws.write(4 + i, 0, i, t_ctr); ws.write(4 + i, 1, C.LIKELIHOOD[i][0], t_ctr)
        ws.merge_range(4 + i, 2, 4 + i, 6, C.LIKELIHOOD[i][1], t_wrap)
    # 2 severity
    ws.merge_range("A12:G12", "2. Severity scale (credible worst outcome; the highest applicable column sets the score)", sect)
    for c, h in enumerate(["Level", "Label", "People", "Process safety / loss of containment", "Asset and operations", "Environment and community"]):
        ws.write(12, c, h, h_in)
    for i in range(1, 6):
        s = C.SEVERITY[i]
        ws.write(12 + i, 0, i, t_ctr)
        for c in range(1, 6):
            ws.write(12 + i, c, s[c - 1], t_wrap if c > 1 else t_ctr)
    # 3 matrix
    ws.merge_range("A20:G20", "3. Risk matrix (score = Likelihood x Severity), coloured by the category thresholds in section 4", sect)
    ws.write("A21", "Likelihood \\ Severity", h_in); ws.write("B21", "Level", h_in)
    for j in range(5):
        ws.write_number(20, 2 + j, j + 1, h_in)
        ws.write(21, 2 + j, C.SEVERITY[j + 1][0], h_calc)
    ws.write("A22", "", h_calc); ws.write("B22", "", h_calc)
    for k, lvl in enumerate([5, 4, 3, 2, 1]):
        r = 23 + k
        ws.write(f"A{r}", C.LIKELIHOOD[lvl][0], t_ctr); ws.write_number(f"B{r}", lvl, t_ctr)
        for j in range(5):
            col = "CDEFG"[j]
            wf(ws, f"{col}{r}", f"=$B{r}*{col}$21", fm(border=1, border_color=BORDER, align="center", bold=True))
    # 4 bands
    ws.merge_range("A29:G29", "4. Risk categories (thresholds are editable inputs) and required management response", sect)
    for c, h in enumerate(["Category", "Minimum score", "Maximum score"]):
        ws.write(29, c, h, h_in)
    ws.merge_range("D30:G30", "Required response", h_in)
    mins = {"Low": 1, "Moderate": 5, "High": 10, "Extreme": 17}
    for i, (name, lo, hi, resp) in enumerate(C.BANDS):
        r = 31 + i
        ws.write(f"A{r}", name, t_ctr)
        ws.write_number(f"B{r}", lo, n_in if i else c_num)
        if i < 3:
            wf(ws, f"C{r}", f"=B{r + 1}-1", c_num)
        else:
            ws.write_number(f"C{r}", 25, c_num)
        ws.merge_range(f"D{r}:G{r}", resp, t_wrap)
    wb.define_name("Moderate_Min", f"={MXN}$B$32"); wb.define_name("High_Min", f"={MXN}$B$33"); wb.define_name("Extreme_Min", f"={MXN}$B$34")
    ws.merge_range("A35:G35", "Possible scores are 1,2,3,4,5,6,8,9,10,12,15,16,20,25. Boundaries are placed so that no possible score falls between two bands.", note)
    for name in C.BAND_ORDER:
        pass
    # 5 priority
    ws.merge_range("A36:G36", "5. Action priority and target days (set by the INITIAL risk category of the related hazard)", sect)
    for c, h in enumerate(["Category", "Priority", "Target days", "Extra days for Engineering, Substitution or Elimination actions"]):
        ws.write(36, c, h, h_in)
    ws.merge_range("D37:G37", "Extra days for Engineering, Substitution or Elimination actions (procurement and design lead time)", h_in)
    for i, name in enumerate(C.BAND_ORDER):
        r = 38 + i
        ws.write(f"A{r}", name, t_ctr); ws.write(f"B{r}", C.PRIORITY[name][0], s_in); ws.write_number(f"C{r}", C.PRIORITY[name][1], n_in)
        ws.merge_range(f"D{r}:G{r}", f"+{C.ENGINEERING_LEAD_TIME_DAYS} days (applied when the action was set; stored in the tracker as the Target Completion Date)", t_wrap)
    # 6 hierarchy
    ws.merge_range("A43:G43", "6. Hierarchy of controls and the maximum credit allowed in a RESIDUAL rating", sect)
    for c, h in enumerate(["Control type", "Max likelihood levels", "Max severity levels"]):
        ws.write(43, c, h, h_in)
    ws.merge_range("D44:G44", "Why", h_in)
    why = {"Elimination": "Removes the hazard; can reduce both likelihood and consequence.",
           "Substitution": "Replaces the hazard with a lower one; credit for both, short of elimination.",
           "Engineering": "Reduces the chance of an event and can limit consequence (for example isolation or detection); consequence credit is limited to one level.",
           "Administrative": "Relies on people following procedures; likelihood credit only, one level, and never a reduction in severity.",
           "PPE": "Last line of defence; limits harm to the person but does not reduce the chance of the event."}
    for i, h in enumerate(C.HIER):
        r = 45 + i
        ws.write(f"A{r}", h, t_ctr); ws.write_number(f"B{r}", C.HIER_CREDIT[h][0], n_in); ws.write_number(f"C{r}", C.HIER_CREDIT[h][1], n_in)
        ws.merge_range(f"D{r}:G{r}", why[h], t_wrap)
    # 7/8 heat maps
    def heat(r0, heading, kcol, scol):
        ws.merge_range(f"A{r0}:G{r0}", heading, sect)
        ws.write(f"A{r0 + 1}", "Likelihood \\ Severity", h_in); ws.write(f"B{r0 + 1}", "Level", h_in)
        for j in range(5):
            ws.write_number(r0, 2 + j, j + 1, h_in)
        for k, lvl in enumerate([5, 4, 3, 2, 1]):
            r = r0 + 2 + k
            ws.write(f"A{r}", C.LIKELIHOOD[lvl][0], t_ctr); ws.write_number(f"B{r}", lvl, t_ctr)
            for j in range(5):
                col = "CDEFG"[j]
                wf(ws, f"{col}{r}", f"=COUNTIFS({hr(kcol)},$B{r},{hr(scol)},{col}${r0 + 1})", fm(border=1, border_color=BORDER, align="center", bold=True, num_format="0;-0;;@"))
        rng = f"C{r0 + 2}:G{r0 + 6}"
        for nm, expr in (("Extreme", f"=($B{r0 + 2}*C${r0 + 1})>=Extreme_Min"), ("High", f"=($B{r0 + 2}*C${r0 + 1})>=High_Min"),
                         ("Moderate", f"=($B{r0 + 2}*C${r0 + 1})>=Moderate_Min"), ("Low", f"=($B{r0 + 2}*C${r0 + 1})>=1")):
            ws.conditional_format(rng, {"type": "formula", "criteria": expr, "format": CF[nm], "stop_if_true": True})
    heat(51, "7. Heat map: number of hazards by INITIAL likelihood and severity (all hazards, unfiltered)", "K", "L")
    heat(59, "8. Heat map: number of hazards by PROJECTED RESIDUAL likelihood and severity (valid only once the recommended controls are implemented and verified)", "Q", "R")
    ws.merge_range("A67:G69", "Residual ratings are projections. A control that has been recommended but not implemented and verified does not reduce risk. The register therefore also carries a Current Risk Score: it equals the residual score only when every linked action is Closed (Verified), and otherwise stays at the initial score. A Current heat map is not drawn because the current score is not stored as a separate likelihood and severity pair.", note)
    wb.define_name("BandNames", f"={MXN}$A$31:$A$34")
    # matrix colouring
    for nm, expr in (("Extreme", "=C23>=Extreme_Min"), ("High", "=C23>=High_Min"), ("Moderate", "=C23>=Moderate_Min"), ("Low", "=C23>=1")):
        ws.conditional_format("C23:G27", {"type": "formula", "criteria": expr, "format": CF[nm], "stop_if_true": True})
    for i, nm in enumerate(C.BAND_ORDER):
        ws.conditional_format(f"A{31 + i}", {"type": "no_errors", "format": CF[nm]})

    # generic band text colouring
    def band_cf(ws_, rng):
        for nm in C.BAND_ORDER:
            ws_.conditional_format(rng, {"type": "cell", "criteria": "==", "value": f'"{nm}"', "format": CF[nm]})

    def score_cf(ws_, rng, first):
        for nm, expr in (("Extreme", f"={first}>=Extreme_Min"), ("High", f"={first}>=High_Min"), ("Moderate", f"={first}>=Moderate_Min"), ("Low", f"={first}>=1")):
            ws_.conditional_format(rng, {"type": "formula", "criteria": expr, "format": CF[nm], "stop_if_true": True})

    # ============================================================ HAZARD REGISTER
    ws = W["Hazard Register"]
    banner(ws, "L")
    ws.merge_range("A3:L3", "Hazard and risk register of 40 HYPOTHETICAL scenarios for a CNG distribution operation. Existing controls are assumed for illustration. Blue text = input; grey = formula. Residual ratings are projections that depend on verified implementation.", note)
    cols = [  # header, width, kind
        ("Hazard ID", 9, "in"), ("Operational Category", 20, "in"), ("Operational Stage", 20, "in"), ("Hazard Title", 30, "in"),
        ("Operational Activity", 24, "in"), ("Hazard Description", 44, "in"), ("Potential Incident", 32, "in"), ("Potential Consequences", 36, "in"),
        ("Persons or Assets Affected", 24, "in"), ("Existing Assumed Controls (illustrative)", 34, "in"), ("Initial Likelihood (1-5)", 11, "in"),
        ("Initial Severity (1-5)", 11, "in"), ("Initial Risk Score", 10, "calc"), ("Initial Risk Category", 12, "calc"),
        ("Recommended Additional Controls", 44, "in"), ("Strongest Control Type", 15, "in"), ("Residual Likelihood (1-5)", 11, "in"),
        ("Residual Severity (1-5)", 11, "in"), ("Residual Risk Score", 10, "calc"), ("Residual Risk Category", 12, "calc"),
        ("Residual Rating Basis", 40, "in"), ("Credit Check", 10, "calc"), ("Specialist Assessment Required", 22, "in"),
        ("Recommended Priority", 14, "calc"), ("Responsible Role", 22, "in"), ("Monitoring Frequency", 14, "in"),
        ("Assessment Date (simulated)", 13, "in"), ("Linked Actions", 9, "calc"), ("Closed-Verified Actions", 11, "calc"),
        ("Current Risk Score", 10, "calc"), ("Current Risk Category", 12, "calc"), ("Data Status", 11, "in")]
    for c, (h, w, k) in enumerate(cols):
        ws.write(4, c, h, h_in if k == "in" else h_calc); ws.set_column(c, c, w)
    ws.set_row(4, 42)
    HZ = D["hazards"]
    n = len(HZ)
    for i, h in enumerate(HZ):
        r = 6 + i
        vals = [h["hazard_id"], h["category"], h["stage"], h["title"], h["activity"], h["hazard"], h["incident"], h["consequences"], h["affected"], h["existing_controls"]]
        for c, v in enumerate(vals):
            ws.write(r - 1, c, v, t_ctr if c == 0 else t_wrap)
        ws.write_number(f"K{r}", h["L"], n_in); ws.write_number(f"L{r}", h["S"], n_in)
        wf(ws, f"M{r}", f"=K{r}*L{r}", c_num)
        wf(ws, f"N{r}", f'=IF(M{r}>=Extreme_Min,"Extreme",IF(M{r}>=High_Min,"High",IF(M{r}>=Moderate_Min,"Moderate","Low")))', c_txt)
        ws.write(f"O{r}", h["rec_controls"], t_wrap); ws.write(f"P{r}", h["hierarchy"], s_in)
        ws.write_number(f"Q{r}", h["rL"], n_in); ws.write_number(f"R{r}", h["rS"], n_in)
        wf(ws, f"S{r}", f"=Q{r}*R{r}", c_num)
        wf(ws, f"T{r}", f'=IF(S{r}>=Extreme_Min,"Extreme",IF(S{r}>=High_Min,"High",IF(S{r}>=Moderate_Min,"Moderate","Low")))', c_txt)
        ws.write(f"U{r}", h["rationale"], t_wrap)
        wf(ws, f"V{r}", f'=IF(AND(Q{r}>=1,R{r}>=1,Q{r}<=K{r},R{r}<=L{r},K{r}-Q{r}<=INDEX({MXN}$B$45:$B$49,MATCH(P{r},{MXN}$A$45:$A$49,0)),L{r}-R{r}<=INDEX({MXN}$C$45:$C$49,MATCH(P{r},{MXN}$A$45:$A$49,0))),"OK","REVIEW")', c_txt)
        ws.write(f"W{r}", h["specialist"], t_wrap)
        wf(ws, f"X{r}", f'=IFERROR(INDEX({MXN}$B$38:$B$41,MATCH(N{r},BandNames,0)),"")', c_txt)
        ws.write(f"Y{r}", h["role"], s_in); ws.write(f"Z{r}", h["frequency"], s_in)
        ws.write_datetime(f"AA{r}", datetime.combine(h["assess_date"], datetime.min.time()), d_in)
        wf(ws, f"AB{r}", f"=COUNTIFS({tr('C')},$A{r})", c_num)
        wf(ws, f"AC{r}", f'=COUNTIFS({tr("C")},$A{r},{tr("U")},"Closed (Verified)")', c_num)
        wf(ws, f"AD{r}", f"=IF(AND(AB{r}>0,AC{r}=AB{r}),S{r},M{r})", c_num)
        wf(ws, f"AE{r}", f'=IF(AD{r}>=Extreme_Min,"Extreme",IF(AD{r}>=High_Min,"High",IF(AD{r}>=Moderate_Min,"Moderate","Low")))', c_txt)
        ws.write(f"AF{r}", h["data_status"], t_ctr)
    last = 5 + n
    for rng in (f"N6:N{last}", f"T6:T{last}", f"AE6:AE{last}"):
        band_cf(ws, rng)
    for rng, first in ((f"M6:M{last}", "M6"), (f"S6:S{last}", "S6"), (f"AD6:AD{last}", "AD6")):
        score_cf(ws, rng, first)
    ws.conditional_format(f"V6:V{last}", {"type": "cell", "criteria": "==", "value": '"REVIEW"', "format": CF["Extreme"]})
    ws.data_validation(f"K6:L{LASTROW}", {"validate": "integer", "criteria": "between", "minimum": 1, "maximum": 5, "error_message": "Enter a whole number from 1 to 5."})
    ws.data_validation(f"Q6:R{LASTROW}", {"validate": "integer", "criteria": "between", "minimum": 1, "maximum": 5, "error_message": "Enter a whole number from 1 to 5."})
    ws.data_validation(f"B6:B{LASTROW}", {"validate": "list", "source": SL["cat"]})
    ws.data_validation(f"P6:P{LASTROW}", {"validate": "list", "source": SL["hier"]})
    ws.data_validation(f"Y6:Y{LASTROW}", {"validate": "list", "source": SL["role"]})
    ws.freeze_panes(5, 4)
    ws.autofilter(4, 0, last - 1, len(cols) - 1)
    ws.write(f"A{last + 2}", "To add a hazard: insert a row inside the table, copy the grey formula cells from the row above, and enter the blue inputs. Formulas on other sheets read rows 6 to 500.", note)

    # ============================================================ CORRECTIVE ACTION TRACKER
    ws = W["Corrective Action Tracker"]
    banner(ws, "L")
    ws.merge_range("A3:L3", "Corrective and preventive action tracker (54 SIMULATED actions). Action Status is calculated: Closed needs a completion date, evidence AND a verification date. A passed target date marks an action Overdue; it never marks it complete.", note)
    tcols = [("Action ID", 9, "in"), ("Source", 16, "in"), ("Related Hazard ID", 10, "in"), ("Related Record ID", 11, "in"), ("Identified Issue", 32, "in"),
             ("Required Corrective Action", 46, "in"), ("Control Hierarchy", 14, "in"), ("Responsible Role", 22, "in"), ("Date Identified", 12, "in"),
             ("Target Completion Date", 12, "in"), ("Completion Date", 12, "in"), ("Recorded Status", 20, "in"), ("Verification Requirement", 36, "in"),
             ("Verified By", 14, "in"), ("Verification Date", 12, "in"), ("Evidence of Closure", 38, "in"), ("Hazard Category", 20, "calc"),
             ("Hazard Stage", 20, "calc"), ("Initial Risk Category", 12, "calc"), ("Action Priority", 14, "calc"), ("Action Status", 22, "calc"),
             ("Days Overdue", 9, "calc"), ("Open Flag", 7, "calc"), ("Completed On Time", 10, "calc"), ("Status Check", 28, "calc"), ("Data Status", 11, "in")]
    for c, (h, w, k) in enumerate(tcols):
        ws.write(4, c, h, h_in if k == "in" else h_calc); ws.set_column(c, c, w)
    ws.set_row(4, 42)
    AC = D["actions"]
    def dtw(ws_, cell, v, fmt):
        if v:
            ws_.write_datetime(cell, datetime.combine(v, datetime.min.time()), fmt)
        else:
            ws_.write_blank(cell, None, fmt)
    for i, a in enumerate(AC):
        r = 6 + i
        ws.write(f"A{r}", a["action_id"], t_ctr); ws.write(f"B{r}", a["source"], s_in); ws.write(f"C{r}", a["hazard_id"], s_in)
        ws.write(f"D{r}", a["record_id"], s_in); ws.write(f"E{r}", a["issue"], t_wrap); ws.write(f"F{r}", a["action"], t_wrap)
        ws.write(f"G{r}", a["hierarchy"], s_in); ws.write(f"H{r}", a["role"], s_in)
        dtw(ws, f"I{r}", a["identified"], d_in); dtw(ws, f"J{r}", a["target"], d_in); dtw(ws, f"K{r}", a["completion"], d_in)
        ws.write(f"L{r}", a["recorded_status"], s_in); ws.write(f"M{r}", a["verification"], t_wrap); ws.write(f"N{r}", a["verifier"], s_in)
        dtw(ws, f"O{r}", a["verified_on"], d_in); ws.write(f"P{r}", a["evidence"], t_wrap)
        wf(ws, f"Q{r}", f'=IFERROR(INDEX({hr("B")},MATCH($C{r},{hr("A")},0)),"CHECK ID")', c_txt)
        wf(ws, f"R{r}", f'=IFERROR(INDEX({hr("C")},MATCH($C{r},{hr("A")},0)),"CHECK ID")', c_txt)
        wf(ws, f"S{r}", f'=IFERROR(INDEX({hr("N")},MATCH($C{r},{hr("A")},0)),"CHECK ID")', c_txt)
        wf(ws, f"T{r}", f'=IFERROR(INDEX({MXN}$B$38:$B$41,MATCH($S{r},BandNames,0)),"")', c_txt)
        wf(ws, f"U{r}", f'=IF(AND($K{r}<>"",$P{r}<>"",$O{r}<>""),"Closed (Verified)",IF($K{r}<>"","Completed (Awaiting Verification)",IF($J{r}<StatusDate,"Overdue",IF($L{r}="In Progress","In Progress","Not Started"))))', c_txt)
        wf(ws, f"V{r}", f'=IF($U{r}="Overdue",StatusDate-$J{r},0)', c_num)
        wf(ws, f"W{r}", f'=IF($U{r}="Closed (Verified)",0,1)', c_num)
        wf(ws, f"X{r}", f'=IF($K{r}="","",IF($K{r}<=$J{r},"Yes","No"))', c_txt)
        wf(ws, f"Y{r}", f'=IF(AND($L{r}="Closed (Verified)",$U{r}<>"Closed (Verified)"),"CHECK: marked closed without completion date, evidence and verification date",IF(AND($L{r}<>"Closed (Verified)",$U{r}="Closed (Verified)"),"CHECK: evidence complete but recorded status not updated",IF(AND($L{r}="Not Started",$K{r}<>""),"CHECK: completion date entered but status Not Started","OK")))', c_txt)
        ws.write(f"Z{r}", a["data_status"], t_ctr)
    tl = 5 + len(AC)
    band_cf(ws, f"S6:S{tl}")
    for txt, bg, fg in (("Overdue", "#C00000", "#FFFFFF"), ("Closed (Verified)", "#C6E0B4", "#000000"), ("Completed (Awaiting Verification)", "#BDD7EE", "#000000"),
                        ("In Progress", "#FFE699", "#000000"), ("Not Started", "#D9D9D9", "#000000")):
        ws.conditional_format(f"U6:U{tl}", {"type": "cell", "criteria": "==", "value": f'"{txt}"', "format": cf(bg, fg, txt == "Overdue")})
    ws.conditional_format(f"V6:V{tl}", {"type": "cell", "criteria": ">", "value": 0, "format": wb.add_format({"font_color": "#C00000", "bold": True})})
    ws.conditional_format(f"Y6:Y{tl}", {"type": "text", "criteria": "begins with", "value": "CHECK", "format": CF["Extreme"]})
    ws.conditional_format(f"X6:X{tl}", {"type": "cell", "criteria": "==", "value": '"No"', "format": CF["High"]})
    ws.data_validation(f"G6:G{LASTROW}", {"validate": "list", "source": SL["hier"]})
    ws.data_validation(f"H6:H{LASTROW}", {"validate": "list", "source": SL["role"]})
    ws.data_validation(f"L6:L{LASTROW}", {"validate": "list", "source": SL["rec"]})
    ws.data_validation(f"I6:K{LASTROW}", {"validate": "date", "criteria": ">", "value": datetime(2020, 1, 1)})
    ws.freeze_panes(5, 3)
    ws.autofilter(4, 0, tl - 1, len(tcols) - 1)
    ws.write(f"A{tl + 2}", "To close an action: enter the Completion Date, the Evidence of Closure reference and the Verification Date (verified by a different role). Do not edit the grey columns.", note)

    # ============================================================ SAFETY INSPECTION CHECKLIST
    ws = W["Safety Inspection Checklist"]
    banner(ws, "L")
    ws.merge_range("A3:B3", "Observation-only checklist.", bold)
    ws.write("C3", "Compliant", bold); ws.write("C4", "Non-compliant", bold); ws.write("C5", "Not Applicable", bold); ws.write("C6", "Compliance % (excluding Not Applicable)", bold)
    wf(ws, "D3", '=COUNTIF($F$9:$F$200,"Compliant")', fm(align="left", bold=True))
    wf(ws, "D4", '=COUNTIF($F$9:$F$200,"Non-compliant")', fm(align="left", bold=True))
    wf(ws, "D5", '=COUNTIF($F$9:$F$200,"Not Applicable")', fm(align="left", bold=True))
    wf(ws, "D6", '=IFERROR(D3/(D3+D4),0)', fm(align="left", bold=True, num_format="0.0%"))
    ws.merge_range("E3:L6", "Responses and observations in this sheet are a SIMULATED sample inspection so the formulas can be seen working; clear column F and G to use it. "
                   "No item asks anyone to operate, test, adjust or repair high-pressure gas equipment. Items marked 'Qualified technical inspection' are checked by confirming the competent person's current "
                   "certificate or report; the technical inspection itself is carried out only by a competent, authorised person.", note)
    ccols = [("Item ID", 9), ("Operational Area", 22), ("Inspection Question", 48), ("Expected Safety Condition", 42), ("Check Type", 24), ("Response", 15),
             ("Observation", 34), ("Risk Priority", 10), ("Recommended Corrective Action (if non-compliant)", 40), ("Responsible Role", 22), ("Related Hazard ID", 10), ("Follow-up", 26)]
    for c, (h, w) in enumerate(ccols):
        ws.write(7, c, h, h_calc if h == "Follow-up" else h_in); ws.set_column(c, c, w)
    ws.set_row(7, 36)
    CK = D["checklist"]
    for i, c_ in enumerate(CK):
        r = 9 + i
        ws.write(f"A{r}", c_["item_id"], t_ctr); ws.write(f"B{r}", c_["area"], t_wrap); ws.write(f"C{r}", c_["question"], t_wrap)
        ws.write(f"D{r}", c_["expected"], t_wrap); ws.write(f"E{r}", c_["check_type"], t_wrap); ws.write(f"F{r}", c_["response"], s_in)
        ws.write(f"G{r}", c_["observation"], t_wrap); ws.write(f"H{r}", c_["priority"], t_ctr); ws.write(f"I{r}", c_["corrective"], t_wrap)
        ws.write(f"J{r}", c_["role"], t_wrap); ws.write(f"K{r}", c_["hazard_id"], t_ctr)
        wf(ws, f"L{r}", f'=IF($F{r}="Non-compliant",IF(LEFT($E{r},9)="Qualified","Refer to competent person",IF($H{r}="High","Raise action, high priority","Raise action")),"")', c_txt)
    cl = 8 + len(CK)
    ws.conditional_format(f"F9:F{cl}", {"type": "cell", "criteria": "==", "value": '"Compliant"', "format": CF["Low"]})
    ws.conditional_format(f"F9:F{cl}", {"type": "cell", "criteria": "==", "value": '"Non-compliant"', "format": CF["Extreme"]})
    ws.conditional_format(f"F9:F{cl}", {"type": "cell", "criteria": "==", "value": '"Not Applicable"', "format": cf("#D9D9D9")})
    ws.conditional_format(f"L9:L{cl}", {"type": "text", "criteria": "begins with", "value": "Refer", "format": CF["High"]})
    ws.conditional_format(f"L9:L{cl}", {"type": "text", "criteria": "begins with", "value": "Raise", "format": CF["Moderate"]})
    ws.conditional_format(f"E9:E{cl}", {"type": "text", "criteria": "begins with", "value": "Qualified", "format": cf("#DDEBF7")})
    ws.data_validation(f"F9:F200", {"validate": "list", "source": SL["resp"]})
    ws.freeze_panes(8, 3)
    ws.autofilter(7, 0, cl - 1, len(ccols) - 1)

    # ============================================================ INCIDENT AND NEAR-MISS LOG
    ws = W["Incident and Near-Miss Log"]
    banner(ws, "L")
    ws.merge_range("A3:L3", "60 SIMULATED records: near misses, minor actual incidents, unsafe conditions, unsafe acts and inspection findings. Fictional sites, no names. Record types are kept separate; outcomes are deliberately modest.", note)
    lcols = [("Record ID", 9, "in"), ("Date (simulated)", 12, "in"), ("Month", 9, "calc"), ("Record Type", 17, "in"), ("Operational Category", 22, "calc"),
             ("Operational Stage", 22, "calc"), ("Linked Hazard ID", 10, "in"), ("Site (simulated)", 24, "in"), ("Description", 52, "in"),
             ("Immediate Action Taken", 40, "in"), ("Potential Severity (1-5)", 10, "in"), ("Actual Severity (0=no harm)", 11, "in"), ("Actual Outcome", 28, "in"),
             ("Reported By (role)", 12, "in"), ("Contributing Factor", 20, "in"), ("Linked Action ID", 10, "in"), ("Data Status", 11, "in"),
             ("Linked Hazard Risk Category", 12, "calc"), ("Linked Hazard Responsible Role", 22, "calc")]
    for c, (h, w, k) in enumerate(lcols):
        ws.write(4, c, h, h_in if k == "in" else h_calc); ws.set_column(c, c, w)
    ws.set_row(4, 42)
    RC = D["records"]
    for i, x in enumerate(RC):
        r = 6 + i
        ws.write(f"A{r}", x["record_id"], t_ctr); ws.write_datetime(f"B{r}", datetime.combine(x["date"], datetime.min.time()), d_in)
        wf(ws, f"C{r}", f'=TEXT(B{r},"yyyy-mm")', c_txt)
        ws.write(f"D{r}", x["record_type"], s_in)
        wf(ws, f"E{r}", f'=IFERROR(INDEX({hr("B")},MATCH($G{r},{hr("A")},0)),"CHECK ID")', c_txt)
        wf(ws, f"F{r}", f'=IFERROR(INDEX({hr("C")},MATCH($G{r},{hr("A")},0)),"CHECK ID")', c_txt)
        ws.write(f"G{r}", x["hazard_id"], s_in); ws.write(f"H{r}", x["site"], t_wrap); ws.write(f"I{r}", x["description"], t_wrap)
        ws.write(f"J{r}", x["immediate_action"], t_wrap); ws.write_number(f"K{r}", x["potential_sev"], n_in); ws.write_number(f"L{r}", x["actual_sev"], n_in)
        ws.write(f"M{r}", x["outcome"], t_wrap); ws.write(f"N{r}", x["reporter_role"], t_ctr); ws.write(f"O{r}", x["factor"], t_wrap)
        ws.write(f"P{r}", x["action_id"], s_in); ws.write(f"Q{r}", x["data_status"], t_ctr)
        wf(ws, f"R{r}", f'=IFERROR(INDEX({hr("N")},MATCH($G{r},{hr("A")},0)),"CHECK ID")', c_txt)
        wf(ws, f"S{r}", f'=IFERROR(INDEX({hr("Y")},MATCH($G{r},{hr("A")},0)),"CHECK ID")', c_txt)
    ll = 5 + len(RC)
    band_cf(ws, f"R6:R{ll}")
    for txt, bg in (("Incident (actual)", "#F4B183"), ("Near miss", "#FFE699"), ("Unsafe condition", "#DDEBF7"), ("Unsafe act", "#E4DFEC"), ("Inspection finding", "#EDEDED")):
        ws.conditional_format(f"D6:D{ll}", {"type": "cell", "criteria": "==", "value": f'"{txt}"', "format": cf(bg)})
    ws.data_validation(f"D6:D{LASTROW}", {"validate": "list", "source": SL["rtype"]})
    ws.data_validation(f"K6:L{LASTROW}", {"validate": "integer", "criteria": "between", "minimum": 0, "maximum": 5})
    ws.freeze_panes(5, 2)
    ws.autofilter(4, 0, ll - 1, len(lcols) - 1)

    # ============================================================ KPI DASHBOARD DATA
    ws = W["KPI Dashboard Data"]
    ws.write("A1", "KPI Dashboard Data (calculation sheet: all formulas, driven by the Dashboard filters)", title)
    ws.merge_range("A2:I2", C.DISCLOSURE, sub); ws.set_row(1, 28)
    ws.set_column("A:A", 20); ws.set_column("B:B", 40); ws.set_column("C:H", 17); ws.set_column("I:I", 12)
    ws.write("B3", "Filter criteria (read from the Dashboard)", h_in); ws.write("C3", "Criterion", h_in)
    crit = [("Operational category", '=IF(Dashboard!$B$5="All","*",Dashboard!$B$5)'), ("Risk level (initial)", '=IF(Dashboard!$D$5="All","*",Dashboard!$D$5)'),
            ("Responsible role", '=IF(Dashboard!$F$5="All","*",Dashboard!$F$5)'), ("Action status", '=IF(Dashboard!$H$5="All","*",Dashboard!$H$5)'),
            ("Date from", "=Dashboard!$J$5"), ("Date to", "=Dashboard!$L$5")]
    for i, (lab, fml) in enumerate(crit):
        ws.write(f"B{4 + i}", lab, t_wrap)
        wf(ws, f"C{4 + i}", fml, fm(border=1, border_color=BORDER, align="center", bg_color=GREY, num_format="dd-mmm-yyyy" if i >= 4 else "General"))

    def pairs(items):
        return ",".join(f"{a},{b}" for a, b in items)
    def hz(extra=(), score=None):
        base = [(hr("B"), "$C$4"), (hr("N"), "$C$5"), (hr("Y"), "$C$6"), (hr("AA"), '">="&$C$8'), (hr("AA"), '"<="&$C$9')] + list(extra)
        return f"COUNTIFS({pairs(base)})" if score is None else f"SUMIFS({score},{pairs(base)})"
    def act(extra=()):
        base = [(tr("Q"), "$C$4"), (tr("S"), "$C$5"), (tr("H"), "$C$6"), (tr("U"), "$C$7"), (tr("I"), '">="&$C$8'), (tr("I"), '"<="&$C$9')] + list(extra)
        return f"COUNTIFS({pairs(base)})"
    def rec(extra=()):
        base = [(lg("E"), "$C$4"), (lg("R"), "$C$5"), (lg("S"), "$C$6"), (lg("B"), '">="&$C$8'), (lg("B"), '"<="&$C$9')] + list(extra)
        return f"COUNTIFS({pairs(base)})"

    ws.merge_range("B11:E11", "Headline KPIs", sect)
    for c, h in enumerate(["KPI", "Value", "Definition", "Calculation"]):
        ws.write(11, 1 + c, h, h_calc)
    kpis = [
        ("Total hazards identified", "=" + hz(), "Hazards matching the filters.", "COUNTIFS on Hazard Register"),
        ("Extreme-risk hazards (initial)", "=" + hz([(hr("N"), '"Extreme"')]), "Initial score in the Extreme band.", "COUNTIFS on Initial Risk Category"),
        ("High-risk hazards (initial)", "=" + hz([(hr("N"), '"High"')]), "Initial score in the High band.", "COUNTIFS on Initial Risk Category"),
        ("Total corrective actions", "=" + act(), "Actions matching the filters.", "COUNTIFS on tracker"),
        ("Open corrective actions", "=" + act([(tr("W"), "1")]), "Any action not Closed (Verified).", "COUNTIFS on Open Flag"),
        ("Overdue corrective actions", "=" + act([(tr("U"), '"Overdue"')]), "No completion date and target date before the status date.", "COUNTIFS on Action Status"),
        ("Completed, awaiting verification", "=" + act([(tr("U"), '"Completed (Awaiting Verification)"')]), "Work done; evidence or verification missing.", "COUNTIFS on Action Status"),
        ("Closed and verified (completed actions)", "=" + act([(tr("U"), '"Closed (Verified)"')]), "Completion date, evidence and verification date all present.", "COUNTIFS on Action Status"),
        ("Percentage of actions closed", "=IFERROR(C20/C16,0)", "Closed and verified divided by total actions.", "C20 / C16"),
        ("Simulated near-miss records", "=" + rec([(lg("D"), '"Near miss"')]), "Records of type Near miss.", "COUNTIFS on Record Type"),
        ("All simulated records", "=" + rec(), "All five record types.", "COUNTIFS on log"),
        ("High + Extreme hazards, projected residual", "=" + hz([(hr("T"), '"High"')]) + "+" + hz([(hr("T"), '"Extreme"')]), "If every recommended control is implemented and verified.", "COUNTIFS on Residual Risk Category"),
        ("High + Extreme hazards, current (control-verified)", "=" + hz([(hr("AE"), '"High"')]) + "+" + hz([(hr("AE"), '"Extreme"')]), "Residual credit counted only where all linked actions are Closed (Verified).", "COUNTIFS on Current Risk Category"),
        ("High + Extreme hazards, initial", "=C14+C15", "Initial High plus Extreme.", "C14 + C15"),
    ]
    for i, (lab, fml, dfn, calc) in enumerate(kpis):
        r = 13 + i
        ws.write(f"B{r}", lab, t_wrap)
        wf(ws, f"C{r}", fml, fm(border=1, border_color=BORDER, align="center", bold=True, bg_color=GREY, num_format="0.0%" if "Percentage" in lab else "0"))
        ws.write(f"D{r}", dfn, t_wrap); ws.write(f"E{r}", calc, t_wrap)
    ws.set_column("D:E", 34)

    # Table 1: bands
    ws.merge_range("B29:E29", "Table 1. Hazards by risk band: initial, current (control-verified) and projected residual", sect)
    for c, h in enumerate(["Risk band", "Initial", "Current (control-verified)", "Projected residual"]):
        ws.write(29, 1 + c, h, h_calc)
    for i, nm in enumerate(C.BAND_ORDER):
        r = 31 + i
        ws.write(f"B{r}", nm, t_ctr)
        wf(ws, f"C{r}", "=" + hz([(hr("N"), f"$B{r}")]), c_num)
        wf(ws, f"D{r}", "=" + hz([(hr("AE"), f"$B{r}")]), c_num)
        wf(ws, f"E{r}", "=" + hz([(hr("T"), f"$B{r}")]), c_num)
    ws.write("B35", "Total", bold)
    for col in "CDE":
        wf(ws, f"{col}35", f"=SUM({col}31:{col}34)", fm(bold=True, align="center", border=1, border_color=BORDER))
    # Table 2: category x band
    ws.merge_range("B38:I38", "Table 2. Hazards by operational category and initial risk band", sect)
    for c, h in enumerate(["Operational category"] + C.BAND_ORDER + ["Total", "High + Extreme"]):
        ws.write(38, 1 + c, h, h_calc)
    for i, cn in enumerate(C.CAT_NAMES):
        r = 40 + i
        ws.write(f"B{r}", cn, t_wrap)
        for j, nm in enumerate(C.BAND_ORDER):
            col = "CDEF"[j]
            wf(ws, f"{col}{r}", "=" + hz([(hr("B"), f"$B{r}"), (hr("N"), f'"{nm}"')]), c_num)
        wf(ws, f"G{r}", f"=SUM(C{r}:F{r})", c_num)
        wf(ws, f"H{r}", f"=E{r}+F{r}", c_num)
    ws.write("B47", "Total", bold)
    for col in "CDEFGH":
        wf(ws, f"{col}47", f"=SUM({col}40:{col}46)", fm(bold=True, align="center", border=1, border_color=BORDER))
    # Table 3: status
    ws.merge_range("B50:C50", "Table 3. Corrective actions by status", sect)
    ws.write("B51", "Action status", h_calc); ws.write("C51", "Actions", h_calc)
    for i, st in enumerate(C.COMPUTED_STATUSES):
        r = 52 + i
        ws.write(f"B{r}", st, t_wrap)
        wf(ws, f"C{r}", "=" + act([(tr("U"), f"$B{r}")]), c_num)
    ws.write("B57", "Total", bold)
    wf(ws, "C57", "=SUM(C52:C56)", fm(bold=True, align="center", border=1, border_color=BORDER))
    # Table 4: role
    ws.merge_range("B60:D60", "Table 4. Overdue and open actions by responsible role", sect)
    for c, h in enumerate(["Responsible role", "Overdue", "Open (not closed-verified)"]):
        ws.write(60, 1 + c, h, h_calc)
    for i, ro in enumerate(C.ROLES):
        r = 62 + i
        ws.write(f"B{r}", ro, t_wrap)
        wf(ws, f"C{r}", "=" + act([(tr("H"), f"$B{r}"), (tr("U"), '"Overdue"')]), c_num)
        wf(ws, f"D{r}", "=" + act([(tr("H"), f"$B{r}"), (tr("W"), "1")]), c_num)
    ws.write("B72", "Total", bold)
    for col in "CD":
        wf(ws, f"{col}72", f"=SUM({col}62:{col}71)", fm(bold=True, align="center", border=1, border_color=BORDER))
    # Table 5: monthly
    ws.merge_range("B75:I75", "Table 5. Records by month (month start dates are inputs; add rows for later months)", sect)
    for c, h in enumerate(["Month start", "Month"] + ["Near miss", "Unsafe condition", "Unsafe act", "Incident (actual)", "Inspection finding", "Total"]):
        ws.write(75, 1 + c, h, h_calc)
    months = [date(2026, m, 1) for m in range(3, 11)]
    typecols = ["D", "E", "F", "G", "H"]
    for i, ms in enumerate(months):
        r = 77 + i
        ws.write_datetime(f"B{r}", datetime.combine(ms, datetime.min.time()), fm(border=1, border_color=BORDER, align="center", font_color=BLUE, num_format="dd-mmm-yyyy"))
        wf(ws, f"C{r}", f'=TEXT(B{r},"mmm yyyy")', c_txt)
        for j, tn in enumerate(["Near miss", "Unsafe condition", "Unsafe act", "Incident (actual)", "Inspection finding"]):
            wf(ws, f"{typecols[j]}{r}", "=" + rec([(lg("D"), f'"{tn}"'), (lg("B"), f'">="&$B{r}'), (lg("B"), f'"<"&EDATE($B{r},1)')]), c_num)
        wf(ws, f"I{r}", f"=SUM(D{r}:H{r})", c_num)
    ws.write("B85", "Total", bold)
    for col in "DEFGHI":
        wf(ws, f"{col}85", f"=SUM({col}77:{col}84)", fm(bold=True, align="center", border=1, border_color=BORDER))
    # Table 6: stage
    ws.merge_range("B88:F88", "Table 6. Risk concentration by operational stage", sect)
    for c, h in enumerate(["Operational stage", "Hazards", "Sum of initial scores", "Sum of projected residual scores", "High + Extreme hazards (initial)"]):
        ws.write(88, 1 + c, h, h_calc)
    for i, sg in enumerate(C.STAGES.values()):
        r = 90 + i
        ws.write(f"B{r}", sg, t_wrap)
        wf(ws, f"C{r}", "=" + hz([(hr("C"), f"$B{r}")]), c_num)
        wf(ws, f"D{r}", "=" + hz([(hr("C"), f"$B{r}")], score=hr("M")), c_num)
        wf(ws, f"E{r}", "=" + hz([(hr("C"), f"$B{r}")], score=hr("S")), c_num)
        wf(ws, f"F{r}", "=" + hz([(hr("C"), f"$B{r}"), (hr("N"), '"High"')]) + "+" + hz([(hr("C"), f"$B{r}"), (hr("N"), '"Extreme"')]), c_num)
    ws.write("B97", "Total", bold)
    for col in "CDEF":
        wf(ws, f"{col}97", f"=SUM({col}90:{col}96)", fm(bold=True, align="center", border=1, border_color=BORDER))
    ws.set_row(88, 30); ws.set_row(60, 30); ws.set_row(29, 30)
    # Table 7: heat map filtered
    ws.merge_range("A100:G100", "Table 7. Heat map of hazards by initial likelihood and severity (filtered)", sect)
    ws.write("A101", "Likelihood", h_calc); ws.write("B101", "Level", h_calc)
    for j in range(5):
        ws.write_number(100, 2 + j, j + 1, h_calc)
    for k, lvl in enumerate([5, 4, 3, 2, 1]):
        r = 102 + k
        ws.write(f"A{r}", C.LIKELIHOOD[lvl][0], t_ctr); ws.write_number(f"B{r}", lvl, t_ctr)
        for j in range(5):
            col = "CDEFG"[j]
            wf(ws, f"{col}{r}", "=" + hz([(hr("K"), f"$B{r}"), (hr("L"), f"{col}$101")]), fm(border=1, border_color=BORDER, align="center", bold=True, num_format="0;-0;;@"))
    for nm, expr in (("Extreme", "=($B102*C$101)>=Extreme_Min"), ("High", "=($B102*C$101)>=High_Min"), ("Moderate", "=($B102*C$101)>=Moderate_Min"), ("Low", "=($B102*C$101)>=1")):
        ws.conditional_format("C102:G106", {"type": "formula", "criteria": expr, "format": CF[nm], "stop_if_true": True})
    ws.freeze_panes(3, 0)

    # ============================================================ DASHBOARD
    ws = W["Dashboard"]
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2); ws.set_column("B:M", 12.5)
    ws.write("B1", "CNG Safety Intelligence Dashboard", title)
    ws.merge_range("B2:M2", C.DISCLOSURE, sub); ws.set_row(1, 30)
    wf(ws, "B3", '="Status date: "&TEXT(StatusDate,"dd mmm yyyy")&"  |  All figures are formulas over the Hazard Register, Corrective Action Tracker and Incident Log. Change a filter below and every card and chart responds."', fm(italic=True, font_size=9, font_color="#404040"))
    lab = fm(bold=True, font_size=9, font_color="#FFFFFF", bg_color=SLATE, align="center", valign="vcenter")
    inp = fm(bold=True, font_color=BLUE, bg_color="#FFF2CC", border=1, border_color="#BF9000", align="center", valign="vcenter")
    inp_d = fm(bold=True, font_color=BLUE, bg_color="#FFF2CC", border=1, border_color="#BF9000", align="center", valign="vcenter", num_format="dd-mmm-yyyy")
    for (a, b, text) in (("B", "C", "Operational category"), ("D", "E", "Risk level (initial)"), ("F", "G", "Responsible role"), ("H", "I", "Action status"),
                         ("J", "K", "Date from"), ("L", "M", "Date to")):
        ws.merge_range(f"{a}4:{b}4", text, lab)
    ws.merge_range("B5:C5", "All", inp); ws.merge_range("D5:E5", "All", inp); ws.merge_range("F5:G5", "All", inp); ws.merge_range("H5:I5", "All", inp)
    ws.merge_range("J5:K5", datetime(2026, 1, 1), inp_d); ws.merge_range("L5:M5", datetime(2026, 12, 31), inp_d)
    ws.set_row(4, 22)
    for cell, src in (("B5", "cat_all"), ("D5", "lvl_all"), ("F5", "role_all"), ("H5", "stat_all")):
        ws.data_validation(cell, {"validate": "list", "source": SL[src], "input_title": "Filter", "input_message": "Choose a value or All."})
    ws.data_validation("J5", {"validate": "date", "criteria": ">", "value": datetime(2020, 1, 1)})
    ws.data_validation("L5", {"validate": "date", "criteria": ">", "value": datetime(2020, 1, 1)})
    ws.merge_range("B6:M6", "Filters apply to: hazards (category, risk level, role, assessment date), actions (category, risk level of the related hazard, role, status, date identified) and records (category, risk level and role of the linked hazard, record date). Action status applies to actions only.", note)
    ws.set_row(5, 26)
    card_l = fm(bold=True, font_size=9, font_color="#FFFFFF", bg_color=NAVY, align="center", valign="vcenter", text_wrap=True)
    card_v = fm(bold=True, font_size=24, font_color=NAVY, bg_color=LIGHT, align="center", valign="vcenter")
    card_p = fm(bold=True, font_size=24, font_color=NAVY, bg_color=LIGHT, align="center", valign="vcenter", num_format="0.0%")
    card_s = fm(font_size=8, font_color="#404040", bg_color=LIGHT, align="center", valign="vcenter", text_wrap=True)
    cards = [("Total hazards identified", "C13", False, '="Hypothetical scenarios matching filters"'),
             ("Extreme-risk hazards", "C14", False, '="Initial score "&Extreme_Min&"-25"'),
             ("High-risk hazards", "C15", False, '="Initial score "&High_Min&"-"&(Extreme_Min-1)'),
             ("Open corrective actions", "C17", False, '="of "&' + KPN + '$C$16&" actions; not closed-verified"'),
             ("Overdue corrective actions", "C18", False, '="No completion; target before status date"'),
             ("Completed corrective actions", "C20", False, '="Closed-verified; plus "&' + KPN + '$C$19&" awaiting verification"'),
             ("Simulated near-miss records", "C22", False, '="of "&' + KPN + '$C$23&" simulated records"'),
             ("Percentage of actions closed", "C21", True, '="Closed-verified / total actions"')]
    for i, (lbl, ref, pct, subf) in enumerate(cards):
        row0 = 7 if i < 4 else 11
        c0 = 1 + 3 * (i % 4)
        ws.merge_range(row0 - 1, c0, row0 - 1, c0 + 2, lbl, card_l)
        ws.merge_range(row0, c0, row0, c0 + 2, "", card_p if pct else card_v)
        wf(ws, xl_rowcol_to_cell(row0, c0), f"={KPN}${ref[0]}${ref[1:]}", card_p if pct else card_v)
        ws.merge_range(row0 + 1, c0, row0 + 1, c0 + 2, "", card_s)
        wf(ws, xl_rowcol_to_cell(row0 + 1, c0), subf, card_s)
        ws.set_row(row0, 38)
    ws.set_row(6, 18); ws.set_row(10, 18)
    # charts
    def style(ch, ttl, w=540, h=285):
        ch.set_title({"name": ttl, "name_font": {"name": "Arial", "size": 11, "bold": True, "color": NAVY}})
        ch.set_size({"width": w, "height": h}); ch.set_chartarea({"border": {"color": "#D9D9D9"}})
        ch.set_legend({"position": "bottom", "font": {"name": "Arial", "size": 9}})
    AX = {"num_font": {"name": "Arial", "size": 9}}
    BC = {"Low": "#A9D18E", "Moderate": "#FFD966", "High": "#F4B183", "Extreme": "#C00000"}
    # 1 risk distribution by band
    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "Hazards (initial)", "categories": f"={KPN}$B$31:$B$34", "values": f"={KPN}$C$31:$C$34",
                   "points": [{"fill": {"color": BC[b]}} for b in C.BAND_ORDER], "data_labels": {"value": True, "font": {"name": "Arial", "size": 9}}, "gap": 80})
    style(ch, "Risk distribution by risk category (initial)"); ch.set_legend({"none": True}); ch.set_y_axis({"min": 0, "major_gridlines": {"visible": True, "line": {"color": "#E6E6E6"}}, **AX}); ch.set_x_axis(AX)
    ws.insert_chart("B15", ch)
    # 2 initial vs current vs residual
    ch = wb.add_chart({"type": "column"})
    for col, nm, colr in (("C", "Initial", NAVY), ("D", "Current (control-verified)", "#5B84A8"), ("E", "Projected residual", "#A9C3DB")):
        ch.add_series({"name": nm, "categories": f"={KPN}$B$31:$B$34", "values": f"={KPN}${col}$31:${col}$34", "fill": {"color": colr}, "data_labels": {"value": True, "font": {"name": "Arial", "size": 8}}})
    style(ch, "Initial vs current vs projected residual risk"); ch.set_y_axis({"min": 0, "major_gridlines": {"visible": True, "line": {"color": "#E6E6E6"}}, **AX}); ch.set_x_axis(AX)
    ws.insert_chart("H15", ch)
    # 3 hazards by category stacked
    ch = wb.add_chart({"type": "bar", "subtype": "stacked"})
    for j, nm in enumerate(C.BAND_ORDER):
        col = "CDEF"[j]
        ch.add_series({"name": nm, "categories": f"={KPN}$B$40:$B$46", "values": f"={KPN}${col}$40:${col}$46", "fill": {"color": BC[nm]}, "gap": 60})
    style(ch, "Hazards by operational category and initial band"); ch.set_y_axis({"reverse": True, **AX}); ch.set_x_axis({"min": 0, "major_gridlines": {"visible": True, "line": {"color": "#E6E6E6"}}, **AX})
    ws.insert_chart("B30", ch)
    # 4 actions by status
    ch = wb.add_chart({"type": "bar"})
    sc = {"Not Started": "#BFBFBF", "In Progress": "#FFD966", "Overdue": "#C00000", "Completed (Awaiting Verification)": "#9DC3E6", "Closed (Verified)": "#70AD47"}
    ch.add_series({"name": "Actions", "categories": f"={KPN}$B$52:$B$56", "values": f"={KPN}$C$52:$C$56", "points": [{"fill": {"color": sc[s]}} for s in C.COMPUTED_STATUSES],
                   "data_labels": {"value": True, "font": {"name": "Arial", "size": 9}}, "gap": 60})
    style(ch, "Corrective actions by status"); ch.set_legend({"none": True}); ch.set_y_axis({"reverse": True, **AX}); ch.set_x_axis({"min": 0, "major_gridlines": {"visible": True, "line": {"color": "#E6E6E6"}}, **AX})
    ws.insert_chart("H30", ch)
    # 5 overdue by role
    ch = wb.add_chart({"type": "bar"})
    ch.add_series({"name": "Overdue", "categories": f"={KPN}$B$62:$B$71", "values": f"={KPN}$C$62:$C$71", "fill": {"color": "#C00000"}, "data_labels": {"value": True, "font": {"name": "Arial", "size": 9}}, "gap": 50})
    style(ch, "Overdue actions by responsible role"); ch.set_legend({"none": True}); ch.set_y_axis({"reverse": True, **AX}); ch.set_x_axis({"min": 0, "major_gridlines": {"visible": True, "line": {"color": "#E6E6E6"}}, **AX})
    ws.insert_chart("B45", ch)
    # 6 trend
    ch = wb.add_chart({"type": "line"})
    for col, nm, colr in (("D", "Near miss", "#C55A11"), ("E", "Unsafe condition", NAVY), ("F", "Unsafe act", "#7F7F7F")):
        ch.add_series({"name": nm, "categories": f"={KPN}$C$77:$C$84", "values": f"={KPN}${col}$77:${col}$84", "line": {"color": colr, "width": 2.25}, "marker": {"type": "circle", "size": 5, "fill": {"color": colr}, "border": {"color": colr}}})
    style(ch, "Near-miss, unsafe-condition and unsafe-act records by month (simulated)"); ch.set_y_axis({"min": 0, "major_gridlines": {"visible": True, "line": {"color": "#E6E6E6"}}, **AX}); ch.set_x_axis(AX)
    ws.insert_chart("H45", ch)
    # 7 stage concentration
    ch = wb.add_chart({"type": "bar"})
    ch.add_series({"name": "Sum of initial scores", "categories": f"={KPN}$B$90:$B$96", "values": f"={KPN}$D$90:$D$96", "fill": {"color": NAVY}, "gap": 50, "data_labels": {"value": True, "font": {"name": "Arial", "size": 8}}})
    ch.add_series({"name": "Sum of projected residual scores", "categories": f"={KPN}$B$90:$B$96", "values": f"={KPN}$E$90:$E$96", "fill": {"color": "#A9C3DB"}, "data_labels": {"value": True, "font": {"name": "Arial", "size": 8}}})
    style(ch, "Risk concentration by operational stage (sum of scores)"); ch.set_y_axis({"reverse": True, **AX}); ch.set_x_axis({"min": 0, "major_gridlines": {"visible": True, "line": {"color": "#E6E6E6"}}, **AX})
    ws.insert_chart("B60", ch)
    # 8 heat map (cells)
    ws.merge_range("H60:M60", "5 x 5 risk heat map: hazards by initial likelihood and severity (filtered)", fm(bold=True, font_size=11, font_color=NAVY))
    ws.write("H61", "Likelihood \\ Severity", fm(bold=True, font_size=8, font_color="#FFFFFF", bg_color=SLATE, align="center", valign="vcenter", shrink=True))
    for j in range(5):
        ws.write_number(60, 8 + j, j + 1, fm(bold=True, font_size=9, font_color="#FFFFFF", bg_color=SLATE, align="center", valign="vcenter", num_format='"Severity "0'))
    for k, lvl in enumerate([5, 4, 3, 2, 1]):
        r = 62 + k
        ws.write_number(f"H{r}", lvl, fm(bold=True, font_size=9, bg_color="#D9E1E8", align="left", valign="vcenter", shrink=True, num_format=f'0" {C.LIKELIHOOD[lvl][0]}"', border=1, border_color=BORDER))
        for j in range(5):
            col = "IJKLM"[j]; src = "CDEFG"[j]
            wf(ws, f"{col}{r}", f"={KPN}{src}${102 + k}", fm(bold=True, font_size=14, align="center", valign="vcenter", border=1, border_color="#FFFFFF", num_format="0;-0;;@"))
        ws.set_row(r - 1, 34)
    for nm, expr in (("Extreme", "=($H62*I$61)>=Extreme_Min"), ("High", "=($H62*I$61)>=High_Min"), ("Moderate", "=($H62*I$61)>=Moderate_Min"), ("Low", "=($H62*I$61)>=1")):
        ws.conditional_format("I62:M66", {"type": "formula", "criteria": expr, "format": CF[nm], "stop_if_true": True})
    ws.merge_range("H67:M69", "Cell colour shows the risk category of that likelihood and severity combination; the number is how many hazards sit there. Green Low, yellow Moderate, orange High, red Extreme (thresholds on the Risk Assessment Matrix sheet).", note)
    ws.set_row(60, 26)
    ws.activate()

    # ============================================================ DATA DICTIONARY
    ws = W["Data Dictionary"]
    banner(ws, "E")
    for c, (h, w) in enumerate([("Sheet", 26), ("Field", 34), ("Kind", 11), ("Description", 70), ("Rule or allowed values", 70)]):
        ws.write(3, c, h, h_in); ws.set_column(c, c, w)
    for i, row in enumerate(DICT):
        for c, v in enumerate(row):
            ws.write_string(4 + i, c, v, t_wrap if c != 2 else t_ctr)
    ws.autofilter(3, 0, 3 + len(DICT), 4); ws.freeze_panes(4, 0)
    ws.write(5 + len(DICT), 0, "Dashboard, KPI Dashboard Data, Risk Assessment Matrix and Settings contain only formulas, labels and parameters; their calculations are described on those sheets.", note)

    # ============================================================ REFERENCE SOURCES
    ws = W["Reference Sources"]
    banner(ws, "H")
    ws.merge_range("A3:H3", "Each source below was opened and read between 3 and 4 November 2026. 'What was verified' states exactly what was and was not confirmed. Mandatory Nigerian requirements are kept separate from voluntary international guidance.", note)
    for c, (h, w) in enumerate([("Ref", 6), ("Source", 40), ("Issuer", 30), ("Status", 28), ("Date", 26), ("Link", 44), ("Why it is relevant", 56), ("What was verified", 56)]):
        ws.write(4, c, h, h_in); ws.set_column(c, c, w)
    for i, s in enumerate(SOURCES):
        r = 5 + i
        for c, k in enumerate(["id", "name", "issuer", "status", "date"]):
            ws.write(r, c, s[k], t_wrap)
        ws.write_url(r, 5, s["url"], link, s["url"]); ws.write(r, 6, s["relevance"], t_wrap); ws.write(r, 7, s["verified"], t_wrap)
    r0 = 6 + len(SOURCES)
    ws.merge_range(r0, 0, r0, 7, "Not verified, therefore not claimed", sect)
    for i, t in enumerate(NOT_VERIFIED):
        ws.merge_range(r0 + 1 + i, 0, r0 + 1 + i, 7, f"{i + 1}. {t}", t_wrap)
    ws.freeze_panes(5, 2)

    # ============================================================ READ ME
    ws = W["Read Me"]
    ws.hide_gridlines(2)
    ws.set_column("A:A", 3); ws.set_column("B:B", 30); ws.set_column("C:C", 100)
    ws.write("B1", C.PROJECT_TITLE, title)
    ws.merge_range("B2:C2", C.DISCLOSURE, sub); ws.set_row(1, 32)
    ws.write("B3", "Prepared for discussion with: " + C.CASE_STUDY_FOR + " (independent case study; no engagement or access implied).", note)
    ws.write("B5", "Sheet", h_in); ws.write("C5", "Purpose", h_in)
    guide = [("Dashboard", "Management view. Five filters, eight KPI cards, eight visuals including the 5 x 5 heat map. Every figure is a formula."),
             ("Hazard Register", "40 hypothetical hazard scenarios with initial and residual scoring, control hierarchy, credit check, priority and current (control-verified) risk."),
             ("Risk Assessment Matrix", "Scales, matrix, category thresholds, priority rules, hierarchy-of-controls credit limits and unfiltered heat maps. Thresholds here drive every sheet."),
             ("Corrective Action Tracker", "54 simulated actions with calculated status, priority and overdue identification. Closure needs evidence and verification."),
             ("Safety Inspection Checklist", "48 observation-only items across eight areas; technical items are verified by a competent person. Includes a simulated sample inspection."),
             ("Incident and Near-Miss Log", "60 simulated records separated into near miss, actual incident, unsafe condition, unsafe act and inspection finding."),
             ("KPI Dashboard Data", "Calculation sheet behind the dashboard: criteria, KPIs, chart tables."),
             ("Data Dictionary", "Every field: kind (input or calculated), meaning and rule."),
             ("Reference Sources", "Verified sources with dates, links, mandatory versus voluntary status and what was verified."),
             ("Settings", "Status date (simulated) and dropdown lists.")]
    for i, (n_, p) in enumerate(guide):
        ws.write_url(5 + i, 1, f"internal:'{n_}'!A1", fm(font_color="#0563C1", underline=1, border=1, border_color=BORDER), n_)
        ws.write(5 + i, 2, p, t_wrap)
    r = 17
    ws.write(f"B{r}", "How to use", h_in); ws.write(f"C{r}", "", h_in)
    steps = ["1. Replace the simulated hazards with your own: edit the blue inputs (likelihood, severity, control type, residual scores). Scores, categories, priority and checks update from the grey formulas.",
             "2. Record actions in the tracker. Enter the completion date only when the work is done. An action closes only when completion date, evidence reference and verification date are all present.",
             "3. Set the status date on the Settings sheet (or use =TODAY()). Overdue is calculated against it.",
             "4. Use the Dashboard filters. Hazards, actions and records are filtered together as described on the dashboard.",
             "5. Residual ratings are projections. The Current Risk Score recognises a reduction only when every linked action is closed and verified.",
             "6. Rows 6 to 500 of the register, tracker and log are read by the formulas. Insert rows inside the table and copy the grey formulas down."]
    for i, s in enumerate(steps):
        ws.write(r + i, 1, "", None); ws.merge_range(r + i, 1, r + i, 2, s, t_wrap)
    r = 25
    ws.write(f"B{r}", "Colour legend", h_in); ws.write(f"C{r}", "", h_in)
    ws.write(f"B{r + 1}", "Blue text", fm(font_color=BLUE, border=1, border_color=BORDER)); ws.write(f"C{r + 1}", "Input or parameter you may edit.", t_wrap)
    ws.write(f"B{r + 2}", "Grey cell", c_num); ws.write(f"C{r + 2}", "Formula. Do not overwrite.", t_wrap)
    ws.write(f"B{r + 3}", "Risk colours", t_wrap); ws.write(f"C{r + 3}", "Low green, Moderate yellow, High orange, Extreme red. Boundaries: Low 1-4, Moderate 5-9, High 10-16, Extreme 17-25.", t_wrap)
    ws.merge_range("B30:C32", "Limits: this workbook supports hazard identification, risk ranking and action tracking. It is not a quantitative risk assessment, a substitute for competent engineering review of pressure equipment, or evidence of regulatory compliance or ISO 45001 conformity. Every figure in it comes from simulated data.", note)

    wb.close()


def recalc_values(path):
    """Recalculate a scratch copy in LibreOffice (headless); returns {(sheet, cell): value} and a small status dict."""
    from openpyxl import load_workbook
    os.makedirs(SCRATCH, exist_ok=True)
    src = os.path.join(SCRATCH, "recalc_in.xlsx"); out = os.path.join(SCRATCH, "recalc_out")
    shutil.copy(path, src)
    from lo_recalc import recalc_convert
    recalc_convert(src, out)
    wbv = load_workbook(os.path.join(out, "recalc_in.xlsx"), data_only=True)
    wbf = load_workbook(path)
    cache, errors, formulas = {}, 0, 0
    for sh in wbv.worksheets:
        for row in sh.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                v = c.value
                if isinstance(v, str) and v.startswith("#") and v[1:4].isupper():
                    errors += 1
                if isinstance(v, (datetime, date)):
                    v = (v - datetime(1899, 12, 30)).days if isinstance(v, datetime) else (datetime.combine(v, datetime.min.time()) - datetime(1899, 12, 30)).days
                cache[(sh.title, c.coordinate)] = v
    for sh in wbf.worksheets:
        for row in sh.iter_rows():
            formulas += sum(1 for c in row if isinstance(c.value, str) and c.value.startswith("="))
    return cache, {"status": "success" if not errors else "errors", "total_formulas": formulas, "total_errors": errors}


if __name__ == "__main__":
    D = build_data.build()
    build_data.export(D)
    os.makedirs(SCRATCH, exist_ok=True)
    p1 = os.path.join(SCRATCH, "pass1.xlsx")
    build_workbook(p1, D, {})
    cache, info = recalc_values(p1)
    print("pass-1 recalc:", {k: info.get(k) for k in ("status", "total_formulas", "total_errors", "error")})
    if info.get("total_errors"):
        print(info.get("error_summary"))
    build_workbook(OUT, D, cache)
    print("written", OUT)
