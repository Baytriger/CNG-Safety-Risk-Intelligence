# -*- coding: utf-8 -*-
"""Builds the main HSE project report (DOCX). All numbers come from build_data (same source as the workbook)."""
import os, sys
from collections import Counter
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE); sys.path.insert(0, HERE)
import config as C, build_data
from sources import SOURCES, NOT_VERIFIED
from docx_helpers import *

D = build_data.build(); H, A, R, CL, F = D["hazards"], D["actions"], D["records"], D["checklist"], D["facts"]
VIS = os.path.join(ROOT, "Dashboard", "Visuals")
SD = C.STATUS_DATE.strftime("%d %B %Y")
n_h, n_a, n_r = len(H), len(A), len(R)
ib, rb, cb = F["init_bands"], F["res_bands"], F["cur_bands"]
st = F["status_counts"]; closed, openn, overdue = F["closed"], F["open"], F["overdue"]
awaiting = st["Completed (Awaiting Verification)"]
pct = lambda a, b: f"{100 * a / b:.1f}%"
red_proj = pct(F["sum_init"] - F["sum_res"], F["sum_init"]); red_cur = pct(F["sum_init"] - F["sum_cur"], F["sum_init"])
top = sorted(H, key=lambda h: (-h["init_score"], h["hazard_id"]))
rt = F["rec_types"]
nm_by_cat = Counter(r["category"] for r in R if r["record_type"] == "Near miss")
uc_by_cat = Counter(r["category"] for r in R if r["record_type"] == "Unsafe condition")
rec_by_cat = F["rec_by_category"]
FOOT = "SIMULATED DATA: independent portfolio case study, not Powergas records"

d = new_doc(FOOT)

# ------------------------------------------------------------------ cover
for _ in range(5): para(d, "")
para(d, "INDEPENDENT PORTFOLIO CASE STUDY", bold=True, size=10, color=RED, after=2)
para(d, C.PROJECT_TITLE, bold=True, size=24, color=NAVY, after=6)
para(d, "Hazard and risk register, corrective action tracking and management dashboard for a CNG virtual-pipeline distributor", size=12, color=GREY, after=18)
kv_table(d, [("Case study framed around", C.CASE_STUDY_FOR + " (CNG production and virtual-pipeline distribution)"),
             ("Prepared by", "Moshood (independent)"),
             ("Reporting status date", SD + " (simulated; all overdue logic is evaluated at this date)"),
             ("Report date", "4 November 2026"),
             ("Version", "1.0 for discussion"),
             ("Data basis", "Entirely hypothetical and simulated. No Powergas records, personnel, facilities, inspections or incident reports were accessed or used.")], widths=(4.2, 12.8))
callout(d, "**Read this first.** Every hazard, observation, incident, near miss, action, date, facility and outcome in this report is **simulated** to demonstrate the method. "
           "Nothing here describes the actual safety performance of Powergas or any other company, and nothing here shows regulatory compliance or ISO 45001 conformity. "
           "Findings describe the simulated dataset only.")
page_break(d)

# ------------------------------------------------------------------ contents
h(d, "Contents", 1)
for t in ["1  Purpose and scope", "2  Operating context", "3  Regulatory and technical basis", "4  Method", "5  Hazard identification and risk results",
          "6  Corrective action performance", "7  Simulated incident, near-miss and inspection records", "8  Analytical findings",
          "9  Recommendations", "10  Implementation roadmap: 30, 60 and 90 days", "11  Limitations and items needing manual verification",
          "Appendix A  Hazard register summary", "Appendix B  Project files and how they connect"]:
    para(d, t, after=2)

# ------------------------------------------------------------------ 1
h(d, "1  Purpose and scope")
para(d, "A CNG distribution company needs a way to answer one practical question: how do we find operational hazards systematically, rank them consistently, check that preventive controls "
        "exist and work, and make sure corrective actions are actually closed before an incident occurs? This project builds that capability as a working system and demonstrates it on a simulated dataset.")
para(d, "The system has four connected parts: a hazard and risk register scored on a 5 x 5 matrix; a corrective action tracker that cannot mark an action complete just because a target date has passed; "
        "a CNG inspection checklist and incident log that feed the register; and a dashboard that derives every figure from those tables.")
para(d, f"The simulated dataset contains **{n_h} hazard scenarios** across seven operating categories, **{n_a} corrective actions**, **{n_r} incident, near-miss, unsafe-condition, unsafe-act and inspection records**, and a **{F['checklist_n']}-item inspection checklist**.")
h(d, "What the project does not do", 2)
bullet(d, "It is not a quantitative risk assessment. Scores rank scenarios; they are not probabilities of loss.")
bullet(d, "It does not replace competent engineering review of pressure equipment, relief devices, hoses, trailers or electrical area classification.")
bullet(d, "It does not show that any company complies with Nigerian law or with ISO 45001. It borrows ISO 45001 and ILO-OSH principles only.")
bullet(d, "It does not predict accidents. The incident log is too small and too artificial to support prediction, and the report makes no such claim.")

# ------------------------------------------------------------------ 2
h(d, "2  Operating context")
para(d, "Powergas is publicly described as a producer and distributor of compressed natural gas that supplies customers beyond the pipeline network by road, using tube trailers (a virtual pipeline), "
        "with gas at Ebedei that would otherwise be flared (source R12; company description only). The hazard set is built around that chain rather than around generic workplace safety:")
table(d, ["Stage", "What happens", "Why the hazards differ here"],
      [["1 Compression and processing", "Gas is dried, compressed to storage and transfer pressure and held under automatic protection.", "High-pressure rotating equipment and flammable gas in a shelter; protective devices must work when called on."],
       ["2 Storage", "Cascade banks of pressure vessels hold compressed gas.", "Stored energy is large; failures are low-frequency and high-consequence, and depend on inspection discipline."],
       ["3 Transfer (loading and offloading)", "Hoses and couplings connect trailers to the compressor station and to customer pressure-reduction skids.", "Connections are made and broken repeatedly by people; small errors release gas close to workers and vehicles."],
       ["4 Road transport", "Laden tube trailers travel to customer sites.", "Rollover, fatigue and vehicle defects put a pressurised load on public roads (source R11, R14)."],
       ["5 Customer delivery site", "Arrival, positioning, connection and offloading at a site the company does not control.", "Unfamiliar layouts, ignition sources and third parties."],
       ["6 Maintenance and contractor work", "Isolation, hot work, inspection and repair.", "Gas lines must be proved safe to work on; contractors add unfamiliarity."],
       ["7 Site support", "PPE, training, reporting, emergency arrangements.", "These decide whether a small event stays small."]],
      [4.2, 6.3, 6.5], size=8.5)

# ------------------------------------------------------------------ 3
h(d, "3  Regulatory and technical basis")
para(d, "Sources were opened and read during the project (3 and 4 November 2026). The table separates mandatory Nigerian instruments from voluntary international guidance. "
        "A source appears here only if it was read; where only a catalogue page could be read, the table says so.")
mand = [s for s in SOURCES if s["status"].startswith("Mandatory")]
vol = [s for s in SOURCES if not s["status"].startswith("Mandatory")]
h(d, "3.1  Mandatory Nigerian instruments", 2)
table(d, ["Ref", "Source and date", "Used for", "What was verified"],
      [[s["id"], f"**{s['name']}**. {s['date']}", s["relevance"], s["verified"]] for s in mand], [1.0, 5.0, 6.0, 5.0], size=7.5)
h(d, "3.2  Voluntary international guidance and context", 2)
table(d, ["Ref", "Source, status and date", "Used for", "What was verified"],
      [[s["id"], f"**{s['name']}**. {s['status']}. {s['date']}", s["relevance"], s["verified"]] for s in vol], [1.0, 5.0, 6.0, 5.0], size=7.5)
h(d, "3.3  What could not be confirmed", 2)
para(d, "The following were not verified and are therefore not asserted as requirements anywhere in this project:")
for t in NOT_VERIFIED: bullet(d, t)
para(d, "Working links for every source are in the workbook sheet Reference Sources and in Research_and_References/Reference_Register.md.", italic=True, size=9)

# ------------------------------------------------------------------ 4
h(d, "4  Method")
h(d, "4.1  Risk matrix and categories", 2)
para(d, "Risk Score = Likelihood x Severity, each on a 1 to 5 scale. The same boundaries are used in the register, matrix, heat maps, tracker, dashboard and this report; in the workbook they are held in one place and every sheet reads them.")
table(d, ["Level", "Likelihood", "Meaning used in this project"], [[k, v[0], v[1]] for k, v in C.LIKELIHOOD.items()], [1.4, 3.0, 12.6], size=8.5, align_center=(0,))
table(d, ["Level", "Severity", "People", "Process safety", "Asset and operations"], [[k, v[0], v[1], v[2], v[3]] for k, v in C.SEVERITY.items()], [1.2, 2.6, 4.0, 5.2, 4.0], size=8, align_center=(0,))
table(d, ["Category", "Score range", "Required response (illustrative company policy, not a legal requirement)"],
      [[n, f"{lo} to {hi}", r] for n, lo, hi, r in C.BANDS], [2.4, 2.4, 12.2], size=8.5, band_col=0)
para(d, "The Extreme band starts at 17, so only a likely catastrophic event, an almost certain major event, or a combination above those reaches it. A score of 15 or 16 is High. Organisations that want a stricter boundary can change one cell in the workbook and every figure updates.")
h(d, "4.2  Initial, residual and current risk", 2)
para(d, "**Initial risk** is rated with the existing, assumed controls in place. **Residual risk** is a projection: the rating expected once every recommended control is implemented, working and verified. "
        "Because a recommended control reduces nothing until it exists, the register also holds a **current risk score**. It equals the residual score only when every linked action is Closed (Verified), and otherwise stays at the initial score.")
para(d, "To stop residual ratings from being optimistic, the strongest recommended control type limits how far a rating may fall. The workbook flags any hazard that exceeds these limits.")
table(d, ["Strongest control", "Maximum likelihood reduction", "Maximum severity reduction", "Reasoning"],
      [["Elimination", "4 levels", "4 levels", "The hazard is removed."],
       ["Substitution", "3", "3", "A less hazardous method or material replaces it."],
       ["Engineering", "2", "1", "Physical design reduces the chance, and may limit inventory or exposure."],
       ["Administrative", "1", "0", "Procedures and training depend on people and do not change what a release can do."],
       ["PPE", "0", "1", "Protects the wearer; does not prevent the event."]], [3.2, 4.3, 4.0, 5.5], size=8.5)
h(d, "4.3  Action tracking rules", 2)
bullet(d, "**Priority** follows the initial risk category: Extreme P1 (target 30 days from identification), High P2 (60), Moderate P3 (90), Low P4 (180). Engineering, substitution and elimination actions receive 30 extra days for design and procurement. These are illustrative policy values.")
bullet(d, "**Status** is calculated. Closed (Verified) needs a completion date, a reference to closure evidence and a verification date. A completion date without verification is Completed (Awaiting Verification). With no completion date, an action is Overdue once the status date passes its target; otherwise In Progress or Not Started.")
bullet(d, "A passed target date can never close an action. It can only make it Overdue.")
h(d, "4.4  The simulated dataset", 2)
para(d, "Hazard scenarios, controls, actions and records were written for this project and are labelled SIMULATED in every table. Dates are generated with a fixed seed so the results are reproducible. "
        "Sites are fictional (for example Simulated Compression Station A). The incident log separates near misses, actual incidents, unsafe conditions, unsafe acts and inspection findings, and every actual incident is minor.")

# ------------------------------------------------------------------ 5
h(d, "5  Hazard identification and risk results")
para(d, f"The register holds {n_h} simulated hazard scenarios. On initial ratings, **{ib['Extreme']} are Extreme, {ib['High']} High, {ib['Moderate']} Moderate and {ib['Low']} Low**. "
        f"On projected residual ratings, **{rb['Extreme']} are Extreme, {rb['High']} High, {rb['Moderate']} Moderate and {rb['Low']} Low**. "
        f"Counting only controls that have been verified as closed, the current picture is {cb['Extreme']} Extreme, {cb['High']} High, {cb['Moderate']} Moderate and {cb['Low']} Low.")
figure(d, os.path.join(VIS, "risk_distribution.png"), 14.5, "Figure 1. Hazards by risk category: initial, current (control-verified) and projected residual. Simulated data.")
figure(d, os.path.join(VIS, "heatmap_initial.png"), 11.5, "Figure 2. 5 x 5 heat map of initial risk. Small figure = cell score; large figure = number of hazards.")
figure(d, os.path.join(VIS, "heatmap_residual.png"), 11.5, "Figure 3. Heat map of projected residual risk. A projection that holds only if every recommended control is implemented and verified.")
h(d, "5.1  Where the high ratings sit", 2)
bc = F["by_category"]
table(d, ["Operational category", "Hazards", "Extreme", "High", "Avg initial", "Avg residual", "High/Extreme after controls", "Open actions", "Overdue"],
      [[c, v["n"], v["extreme"], v["high"], f"{v['avg_init']:.1f}", f"{v['avg_res']:.1f}", v["res_hi_ex"], v["open_actions"], v["overdue"]] for c, v in bc.items()],
      [4.4, 1.4, 1.4, 1.2, 1.6, 1.7, 2.2, 1.5, 1.6], size=8, align_center=(1, 2, 3, 4, 5, 6, 7, 8))
mx = max(bc.items(), key=lambda kv: kv[1]["hi_ex"])
para(d, f"Loading and offloading holds the most High or Extreme hazards ({bc['Loading & Offloading']['hi_ex']} of {bc['Loading & Offloading']['n']}), followed by compression and processing, transportation and fleet, and maintenance and contractors ({bc['Gas Compression & Processing']['hi_ex']}, {bc['Transportation & Fleet']['hi_ex']} and {bc['Maintenance & Contractors']['hi_ex']} each). "
        f"Maintenance and contractors is the only category with Extreme hazards and has the highest average initial score ({bc['Maintenance & Contractors']['avg_init']:.1f}). "
        "Counts depend on how scenarios were split; compare averages and the stage view below rather than relying on counts alone.")
figure(d, os.path.join(VIS, "hazards_by_category.png"), 14.5, "Figure 4. Hazards by operational category and initial risk category.")
figure(d, os.path.join(VIS, "stage_concentration.png"), 14.5, "Figure 5. Sum of risk scores by operational stage. Site support (occupational and emergency) and maintenance carry the most summed risk because they hold the most scenarios.")
h(d, "5.2  Highest initial scores", 2)
table(d, ["ID", "Hazard (simulated)", "Initial", "Category", "Residual", "Category", "Strongest control", "Specialist check"],
      [[x["hazard_id"], x["title"], x["init_score"], x["init_cat"], x["res_score"], x["res_cat"], x["hierarchy"], "Yes" if x["specialist"].lower().startswith("yes") else "No"] for x in top[:10]],
      [1.2, 6.0, 1.3, 1.7, 1.4, 1.7, 2.3, 1.4], size=8, band_col=(3, 5), align_center=(0, 2, 4, 6, 7))
para(d, "The two Extreme scenarios are isolation before work on gas lines (H-025) and hot work without adequate gas testing (H-026). Both are likely events with a catastrophic worst case in a gas facility, and both rely heavily on people following permits. "
        "The next three (H-015, H-019 and H-022) are all in transfer and road transport, where people, vehicles and flammable gas meet.")
h(d, "5.3  Scenarios that stay High after controls", 2)
stay = [x for x in H if x["res_cat"] in ("High", "Extreme")]
table(d, ["ID", "Hazard (simulated)", "Initial", "Residual", "Why it remains High"],
      [[x["hazard_id"], x["title"], x["init_score"], x["res_score"], x["rationale"]] for x in stay], [1.2, 4.5, 1.2, 1.4, 8.7], size=7.5, align_center=(2, 3))
para(d, f"{F['specialist_count']} of the {n_h} hazards are flagged for specialist assessment or competent technical verification because their controls depend on engineering judgement, for example pressure-equipment integrity, relief devices, protective trips, hoses, tube trailers and isolation design. "
        "The workbook marks them in the register and the checklist separates them from routine visual checks.")

# ------------------------------------------------------------------ 6
h(d, "6  Corrective action performance")
para(d, f"At {SD} the tracker holds {n_a} simulated actions: **{closed} Closed (Verified) ({pct(closed, n_a)})**, {awaiting} Completed (Awaiting Verification), {st['In Progress']} In Progress, {st['Not Started']} Not Started and **{overdue} Overdue**. "
        f"That leaves {openn} actions open, of which {overdue} ({pct(overdue, openn)}) are past their target date. The average overdue action is {F['avg_days_overdue']:.0f} days late.")
figure(d, os.path.join(VIS, "action_status.png"), 12.5, "Figure 6. Corrective actions by calculated status.")
od_role = sorted(F["overdue_by_role"].items(), key=lambda kv: -kv[1])
para(d, f"Overdue actions are concentrated: the {od_role[0][0]} holds {od_role[0][1]} of the {overdue}, followed by the {od_role[1][0]} ({od_role[1][1]}). "
        f"By initial risk, {F['overdue_by_init_band']['High']} overdue actions belong to High hazards, {F['overdue_by_init_band']['Moderate']} to Moderate, {F['overdue_by_init_band']['Low']} to Low and none to Extreme hazards. "
        f"Among the {F['completed_all']} actions that have a completion date, {F['completed_late']} were finished after their target.")
figure(d, os.path.join(VIS, "overdue_by_role.png"), 14.0, "Figure 7. Open and overdue actions by responsible role.")
para(d, f"The {awaiting} actions awaiting verification matter as much as the overdue ones. Until a second person confirms the work, the related hazard keeps its initial score in the current-risk view.")

# ------------------------------------------------------------------ 7
h(d, "7  Simulated incident, near-miss and inspection records")
para(d, f"The log holds {n_r} simulated records: {rt['Near miss']} near misses, {rt['Unsafe condition']} unsafe conditions, {rt['Unsafe act']} unsafe acts, {rt['Inspection finding']} inspection findings and {rt['Incident (actual)']} actual incidents. "
        "The types are kept separate because they answer different questions. A near miss is an event that could have caused harm but did not. An unsafe condition or act is something seen before an event occurs. An inspection finding comes from a planned check. An incident is an event that did cause an outcome; here all eight were minor (first aid, minor equipment damage or a small contained release).")
figure(d, os.path.join(VIS, "record_trend.png"), 14.5, "Figure 8. Simulated records by month and type.")
figure(d, os.path.join(VIS, "records_by_category.png"), 14.5, "Figure 9. Simulated records by operational category and type.")
para(d, f"What the simulated records show, and what they cannot show:")
bullet(d, f"Loading and offloading ({rec_by_cat['Loading & Offloading']} records), transportation ({rec_by_cat['Transportation & Fleet']}) and compression ({rec_by_cat['Gas Compression & Processing']}) account for {pct(rec_by_cat['Loading & Offloading'] + rec_by_cat['Transportation & Fleet'] + rec_by_cat['Gas Compression & Processing'], n_r)} of records. These are also the stages with the most High hazards, which is consistent with the register but does not prove it, because the records were written against the register.")
bullet(d, f"{nm_by_cat.most_common(1)[0][1]} of the {rt['Near miss']} near misses sit in {nm_by_cat.most_common(1)[0][0]}, where transfer errors are caught before gas is released.")
bullet(d, f"{F['rec_high_potential']} of {n_r} records were rated with a potential severity of 4 or 5, against an actual outcome of minor at worst. That gap is the reason to investigate near misses and conditions with the same care as injuries.")
bullet(d, f"By contributing factor (preliminary, not root cause): equipment or condition {F['rec_by_factor']['Equipment or condition']}, human performance {F['rec_by_factor']['Human performance']}, procedure or planning {F['rec_by_factor']['Procedure or planning']}, management system {F['rec_by_factor']['Management system']}.")
bullet(d, f"{F['rec_followups']} records are linked to a corrective action. Records without an action should carry a reason.")
bullet(d, "Month-to-month changes are noise at this sample size. No trend, correlation or prediction is claimed.")
h(d, "7.1  Inspection checklist", 2)
para(d, f"The checklist has {F['checklist_n']} items in eight areas. Each states the expected condition, the response options (Compliant, Non-compliant, Not Applicable), a priority, the corrective action and the responsible role. "
        f"{F['checklist_tech']} items are marked 'Qualified technical inspection': the inspector confirms that a competent person's current record exists and does not open, test or adjust equipment. "
        f"A simulated sample inspection shows {F['checklist']['Compliant']} compliant, {F['checklist']['Non-compliant']} non-compliant and {F['checklist']['Not Applicable']} not applicable responses ({F['checklist_compliance_pct']}% compliance, excluding not-applicable items).")

# ------------------------------------------------------------------ 8
h(d, "8  Analytical findings")
h(d, "8.1  Answers to the analysis questions", 2)
rh = F["reduction_by_hierarchy"]
br = F["biggest_reductions"]
table(d, ["Question", "Finding (simulated data)"],
      [["Which categories hold the most high-risk hazards?", f"Loading and offloading ({bc['Loading & Offloading']['hi_ex']} High or Extreme of {bc['Loading & Offloading']['n']}), then compression ({bc['Gas Compression & Processing']['hi_ex']}), transportation ({bc['Transportation & Fleet']['hi_ex']}) and maintenance ({bc['Maintenance & Contractors']['hi_ex']}). Maintenance and contractors has the only Extreme hazards."],
       ["Which hazards have the highest initial scores?", f"{top[0]['hazard_id']} ({top[0]['title']}) and {top[1]['hazard_id']} ({top[1]['title']}) at {top[0]['init_score']}; then {top[2]['hazard_id']} and {top[3]['hazard_id']} at {top[2]['init_score']}; {top[4]['hazard_id']} at {top[4]['init_score']}."],
       ["Which controls give the largest assumed reduction?", f"Engineering controls account for the largest total reduction ({rh['Engineering']['total']} points across {rh['Engineering']['n']} hazards, average {rh['Engineering']['avg']}), ahead of administrative controls ({rh['Administrative']['total']} across {rh['Administrative']['n']}, average {rh['Administrative']['avg']}). The largest single reductions are {br[0]['id']} (-{br[0]['reduction']}, {br[0]['hier'].lower()}), {br[1]['id']} (-{br[1]['reduction']}) and {br[2]['id']}, {br[3]['id']}, {br[4]['id']}, {br[5]['id']} (-{br[2]['reduction']} each). These are outputs of the rating rules, not measured effects."],
       ["What proportion of actions remain outstanding?", f"{openn} of {n_a} ({pct(openn, n_a)}) are not Closed (Verified); {overdue} of those are overdue."],
       ["Which roles hold the most overdue actions?", f"{od_role[0][0]} ({od_role[0][1]}), {od_role[1][0]} ({od_role[1][1]}), then {od_role[2][0]}, {od_role[3][0]} and {od_role[4][0]} ({od_role[2][1]} each)."],
       ["What patterns appear in the near-miss records?", "Across all record types, transfer hoses, exclusion zones, vehicle condition, gas-detector upkeep and slips recur (three records each for H-012, H-015, H-020, H-002 and H-032). See section 7 for limits."],
       ["Which interventions should management prioritise?", "Isolation and hot-work controls; clearing the overdue and unverified action backlog; transfer-area controls; road-transport fatigue and rollover controls. See section 9."],
       ["What limits apply with real data?", "See section 11."]], [5.0, 12.0], size=8.5)
h(d, "8.2  Projected improvement, with assumptions", 2)
para(d, f"The sum of initial scores is {F['sum_init']}. The projected residual sum is {F['sum_res']} ({red_proj} lower); the current control-verified sum is {F['sum_cur']} ({red_cur} lower). "
        "The projection assumes every recommended control is delivered as described, performs as designed, and is kept in place; it assumes the likelihood and severity credit limits in section 4.2; and it treats each hazard independently. "
        "It is a planning figure for comparing options, not a forecast of incident reduction, and no financial benefit is claimed.")
h(d, "8.3  Four different activities", 2)
table(d, ["Activity", "Question it answers", "Typical output", "Role here"],
      [["Risk assessment", "What could go wrong, how likely, how bad, and what controls are needed?", "Hazard register with ratings and actions", "The register and matrix"],
       ["Safety inspection", "Is what is required actually in place today?", "Checklist responses and findings", "The checklist; findings feed the register and tracker"],
       ["Incident investigation", "Why did this specific event happen, and what must change?", "Root-cause analysis and actions", "Not performed; the log only records preliminary factors"],
       ["Predictive safety analytics", "Which leading signals suggest rising risk?", "Models validated on large real datasets", "Not performed; sample is too small and artificial"]],
      [3.2, 5.0, 4.4, 4.4], size=8.5)

# ------------------------------------------------------------------ 9
h(d, "9  Recommendations")
para(d, "Recommendations are listed by priority. Each addresses a result in this simulated dataset; for a real operator, each should be re-tested against real data. Resources are described by type, not cost, because no cost data exists.")
RECS = [
 ("P1", "R1  Prove isolation and control hot work before gas-line work", [
   ("Issue", "H-025 and H-026 are the only Extreme hazards (score 20 each). They remain High after controls (10 and 15), and one of H-025's linked actions is still awaiting verification."),
   ("Why necessary", "Both scenarios rely on people following permits near flammable gas. The worst case is an ignition or an uncontrolled release during maintenance."),
   ("Recommended control", "Positive isolation with proving and a lockable isolation design for gas lines, assessed by a competent engineer; a permit-to-work that cannot be issued for hot work without a recorded, in-date gas test; cold-work methods where practicable."),
   ("Control hierarchy", "Substitution (cold work) and engineering (isolation design); administrative (permit gate)."),
   ("Responsible role", "Maintenance Manager (isolation); HSE Manager (hot work)."),
   ("Timeframe", "Interim controls and independent permit audit within 30 days; design review within 60; implementation within 90."),
   ("Resources", "Competent engineer for isolation design review; permit-system change; toolbox time for maintenance crews."),
   ("Success indicator", "Zero permits issued without a gas-test record in the monthly audit; verified closure of all actions linked to H-025 and H-026."),
   ("Verification", "HSE Manager audits a sample of permits each month; competent person signs off the isolation design. Residual scores stay at initial until verified.")]),
 ("P1", "R2  Clear the overdue and unverified action backlog", [
   ("Issue", f"{overdue} of {openn} open actions are overdue (average {F['avg_days_overdue']:.0f} days) and {awaiting} completed actions await verification."),
   ("Why necessary", "A register that is not closed out gives a false sense of control. Unverified actions leave current risk at the initial score."),
   ("Recommended control", "A weekly 30-minute action review led by the HSE Manager; re-baseline or escalate each overdue action with a reason; second-person verification within 14 days of completion; extend dashboard review to the monthly management meeting."),
   ("Control hierarchy", "Administrative."),
   ("Responsible role", "HSE Manager; line owners for their actions."),
   ("Timeframe", "Start in week 1; backlog cleared or formally re-baselined by day 60."),
   ("Resources", "About 30 minutes weekly from each role holder; the existing tracker and dashboard."),
   ("Success indicator", "Overdue actions below 10% of open actions; no action awaiting verification for more than 14 days."),
   ("Verification", "Tracker and dashboard figures reviewed monthly; status logic unchanged (no closure without date, evidence and verification date).")]),
 ("P2", "R3  Strengthen transfer-area controls", [
   ("Issue", "Exclusion zone not enforced (H-015, score 16) and coupling errors and hose condition (H-013, H-012, score 12 each) lead the loading and offloading category."),
   ("Why necessary", "Transfer is where people, vehicles, couplings and gas meet. Near misses here are frequent in the simulated log."),
   ("Recommended control", "Fixed physical barriers and ignition-source controls at the bay; connection verification before pressurisation designed with a competent engineer; hose inspection, tagging and retirement register; second-person check at customer sites."),
   ("Control hierarchy", "Engineering (barriers, verified connection), administrative (register, second-person check)."),
   ("Responsible role", "Station Supervisor; Asset Integrity Engineer for hoses."),
   ("Timeframe", "Register and checks within 30 days; barriers and connection verification within 90."),
   ("Resources", "Barrier materials; engineering review; hose inspection tags."),
   ("Success indicator", "All hoses in service have an in-date inspection tag; zero exclusion-zone breaches recorded in weekly checks."),
   ("Verification", "Checklist items for loading and offloading; competent-person sign-off on the connection scheme.")]),
 ("P2", "R4  Reduce road-transport exposure", [
   ("Issue", "Driver fatigue (H-019, score 16), a laden trailer collision (H-022, 15) and vehicle defects (H-020, 12). H-019 and H-022 stay High after controls."),
   ("Why necessary", "A rollover or collision puts a pressurised load on a public road. Source R11 reports that CNG trailers with a high centre of gravity roll over more than most trailers and that driver error featured in the crashes reviewed."),
   ("Recommended control", "Driving-hours limits and journey plans with rest points; speed and roll-stability monitoring; digital pre-trip inspection that blocks departure with a critical defect; route assessment for restricted roads and weather."),
   ("Control hierarchy", "Engineering (monitoring, blocking), administrative (hours, plans)."),
   ("Responsible role", "Fleet Manager; Logistics Coordinator for journey plans."),
   ("Timeframe", "Hours and plans within 30 days; monitoring and digital checks within 90."),
   ("Resources", "Telematics or equivalent; driver time for training."),
   ("Success indicator", "100% of trips have a journey plan; zero departures with an open critical defect."),
   ("Verification", "Trip-record sampling by the HSE Manager; competent review of trailer and monitoring arrangements.")]),
 ("P2", "R5  Make pressure-equipment and protective-system integrity visible", [
   ("Issue", "Relief devices past test interval (H-004), inspection dates not tracked (H-007) and protective trips inhibited (H-003) each score 10. H-004 has an overdue action."),
   ("Why necessary", "Failure of stored-energy equipment is rare but catastrophic, and depends on dates being kept."),
   ("Recommended control", "A single integrity register with due dates, owner and competent-person certificate reference; bypass control for trips and alarms with time limits and sign-off; overdue items reported to the HSE Manager."),
   ("Control hierarchy", "Engineering (trip and alarm bypass control), administrative (register)."),
   ("Responsible role", "Asset Integrity Engineer; Operations Manager for trips."),
   ("Timeframe", "Register within 30 days; bypass control within 60."),
   ("Resources", "Competent person for inspections; register maintained in the workbook or CMMS."),
   ("Success indicator", "No integrity item past due without a documented extension; every bypass has an approver and an end time."),
   ("Verification", "Checklist items marked 'Qualified technical inspection'; certificates sighted, not re-tested by the inspector.")]),
 ("P3", "R6  Test emergency arrangements against CNG scenarios", [
   ("Issue", "Drills do not cover key scenarios (H-039, score 12), leak escalation is slow (H-036), communication can fail (H-038) and there is no liaison with external responders for road incidents (H-040). Several linked actions are overdue."),
   ("Why necessary", "Tube trailer releases can continue for an hour or longer and need a wide isolation distance (source R14, a US responder guide; not a Nigerian figure)."),
   ("Recommended control", "Scenario drills for a coupling release, a trailer leak at a customer site and a road incident; a 24-hour contact and escalation chart; briefings for external responders along key routes."),
   ("Control hierarchy", "Administrative; engineering for alarm and radio coverage."),
   ("Responsible role", "Emergency Response Coordinator."),
   ("Timeframe", "First drill within 60 days; liaison within 90."),
   ("Resources", "Drill time; contact chart; external responder time."),
   ("Success indicator", "One scenario drill per quarter with recorded lessons; every route has a named external contact."),
   ("Verification", "Drill reports reviewed by the HSE Manager; actions entered in the tracker.")]),
 ("P3", "R7  Control contractor entry, training and PPE", [
   ("Issue", "Contractor unfamiliarity (H-028), training refreshers (H-034) and PPE matching (H-031) carry overdue actions."),
   ("Why necessary", "Contractors and newcomers work on or near gas lines with less site knowledge."),
   ("Recommended control", "Contractor competence check and site induction before permit issue; training matrix with refresher dates; PPE matched to tasks, including flame-resistant clothing where the risk assessment calls for it."),
   ("Control hierarchy", "Administrative and PPE."),
   ("Responsible role", "Contracts Manager; Training Officer; HSE Manager."),
   ("Timeframe", "Induction gate within 30 days; training matrix within 60."),
   ("Resources", "Induction material; training time."),
   ("Success indicator", "No contractor works under a permit without a recorded induction."),
   ("Verification", "Permit audit; training matrix review.")]),
 ("P3", "R8  Close the loop on near misses and conditions", [
   ("Issue", f"{F['rec_high_potential']} of {n_r} simulated records had a potential severity of 4 or 5; only {F['rec_followups']} are linked to an action."),
   ("Why necessary", "Reporting only improves safety when reports lead to decisions."),
   ("Recommended control", "Every record with potential severity 4 or 5 gets an owner, a preliminary cause and either an action or a documented reason; feedback to reporters within 7 days; recognise reporting rather than blame."),
   ("Control hierarchy", "Administrative."),
   ("Responsible role", "HSE Manager."),
   ("Timeframe", "From week 2."),
   ("Resources", "Time for a short weekly triage."),
   ("Success indicator", "All high-potential records have an action or a reason within 7 days."),
   ("Verification", "Monthly sample of the log against the tracker.")]),
]
for prio, title, rows in RECS:
    h(d, f"{title}  [{prio}]", 3)
    kv_table(d, rows, widths=(3.4, 13.6), size=8.5)

# ------------------------------------------------------------------ 10
h(d, "10  Implementation roadmap: 30, 60 and 90 days")
table(d, ["Window", "Actions", "Owner", "Evidence of completion"],
      [["Days 1 to 30", "Weekly action review starts (R2). Interim controls and permit audit for isolation and hot work (R1). Hose register and second-person checks (R3). Hours limits and journey plans (R4). Integrity register (R5). Contractor induction gate (R7). Near-miss triage starts (R8).", "HSE Manager with line owners", "Review minutes; audit sample; registers in place"],
       ["Days 31 to 60", "Backlog cleared or re-baselined (R2). Isolation design review (R1). Bypass control for trips and alarms (R5). First scenario drill (R6). Training matrix (R7).", "Maintenance Manager, Asset Integrity Engineer, Emergency Response Coordinator", "Closed (Verified) entries; drill report"],
       ["Days 61 to 90", "Isolation and permit changes implemented (R1). Barriers and verified connection scheme (R3). Monitoring and digital pre-trip checks (R4). External responder liaison (R6). First management review of KPI trends.", "Station Supervisor, Fleet Manager, HSE Manager", "Competent-person sign-offs; dashboard at the status date"]],
      [2.3, 8.2, 3.5, 3.0], size=8.5)
para(d, "Suggested KPIs to track on the dashboard are all calculated from the workbook: extreme and high hazards, open and overdue actions, percentage closed and verified, near misses and unsafe conditions per month, and checklist compliance. Management trigger levels should be set by the company, not taken from this simulation.")

# ------------------------------------------------------------------ 11
h(d, "11  Limitations and items needing manual verification")
h(d, "If this system were used with real data", 2)
bullet(d, "Real hazards, controls and ratings must come from workshops with operators, drivers, maintenance staff and a competent process-safety engineer. The simulated scores here only demonstrate the method.")
bullet(d, "Ratings are judgements. Two teams can score the same scenario differently; calibrate with a worked example before using the matrix.")
bullet(d, "Real incident data will be uneven, delayed and under-reported. Rates need exposure data (trips, transfers, hours) to mean anything, and trends need enough records to be stable.")
bullet(d, "Aggregate scores (sums and averages) are for orientation. Never use them to trade a low-risk improvement against a high-risk one.")
bullet(d, "Overdue logic depends on a correct status date. In live use, replace the Settings date with the review date or with TODAY().")
h(d, "Items requiring manual verification before any real-world use", 2)
for t in ["Applicability and clause-level obligations of the NMDPRA safety and operations regulations to the company's licence (primary texts and legal advice).",
          "Whether the Factories Act applies to each site, and the notification duties under it and under the Employee's Compensation Act.",
          "The edition of NFPA 52, ISO 45001 or any CNG standard that the company has adopted or is contractually bound to.",
          "Exclusion distances, relief and venting design, hose rating and trailer test intervals, which must come from the equipment designer, the applicable standard and a competent person.",
          "Power BI: no .pbix was produced; the import pack must be assembled in Power BI Desktop and its figures reconciled with the table in PowerBI_Build_Guide.md.",
          "Open the Excel workbook in desktop Excel once and confirm that charts, dropdown filters and conditional formatting display correctly. Formulas were recalculated and checked in LibreOffice, not in Excel."]:
    bullet(d, t)

# ------------------------------------------------------------------ appendix A
page_break(d)
h(d, "Appendix A  Hazard register summary (simulated)")
para(d, "Likelihood x severity for initial and projected residual ratings. Full text, controls, assumptions and actions are in the workbook sheet Hazard Register.", size=9)
table(d, ["ID", "Category", "Hazard", "Init", "Initial", "Res", "Residual", "Strongest control", "Priority", "Role"],
      [[x["hazard_id"], x["category"].split(" &")[0], x["title"], f"{x['L']}x{x['S']}={x['init_score']}", x["init_cat"], f"{x['rL']}x{x['rS']}={x['res_score']}", x["res_cat"], x["hierarchy"], x["priority"].split(" ")[0], x["role"]] for x in H],
      [1.1, 1.9, 4.6, 1.5, 1.3, 1.5, 1.3, 1.6, 0.9, 1.9], size=6.5, band_col=(4, 6), align_center=(0, 3, 5, 8))

# ------------------------------------------------------------------ appendix B
h(d, "Appendix B  Project files and how they connect")
table(d, ["Folder or file", "Contents"],
      [["Final_HSE_Report.docx / .pdf", "This report."],
       ["Data_and_Excel/CNG_Safety_Risk_Assessment_Workbook.xlsx", "Working tool: Read Me, Dashboard, Hazard Register, Risk Assessment Matrix, Corrective Action Tracker, Safety Inspection Checklist, Incident and Near-Miss Log, KPI Dashboard Data, Data Dictionary, Reference Sources, Settings."],
       ["Data_and_Excel/csv/", "The same tables as CSV for other tools."],
       ["Dashboard/", "Visuals (PNG) and the Power BI import pack (data, Power Query, DAX, theme, build guide)."],
       ["Presentation/", "Management deck (10 slides)."],
       ["Reports/", "One-page executive summary; technical methodology and data dictionary; QA audit."],
       ["Research_and_References/", "Reference register with links and verification notes."],
       ["Build_Scripts/", "Python scripts that regenerate everything from one set of simulated inputs."]], [6.5, 10.5], size=8.5)

out = os.path.join(ROOT, "Final_HSE_Report.docx")
d.save(out); print("saved", out)
