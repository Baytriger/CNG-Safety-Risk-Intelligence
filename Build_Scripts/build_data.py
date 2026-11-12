# -*- coding: utf-8 -*-
"""Assemble all SIMULATED tables, apply the scoring rules in config.py, export CSVs and facts.json.
Run:  python3 -I build_data.py      (deterministic: fixed seed)"""
import importlib.util, json, os, random, sys
from collections import Counter, defaultdict
from datetime import date, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import config as C
from hazards import H as HAZ_RAW
from incidents import R as REC_RAW, FOLLOWUPS
from checklist import AREAS
from sources import SOURCES, NOT_VERIFIED

SEED = 20261031
rng = random.Random(SEED)
SD = C.STATUS_DATE


def d(s):
    return date.fromisoformat(s) if isinstance(s, str) else s


# ---------------------------------------------------------------- hazards
ASSESS_DATES = {"A": date(2026, 2, 3), "B": date(2026, 2, 5), "C": date(2026, 2, 10), "D": date(2026, 2, 12),
                "E": date(2026, 2, 17), "F": date(2026, 2, 19), "G": date(2026, 2, 24)}


def build_hazards():
    rows = []
    for i, x in enumerate(HAZ_RAW, 1):
        (cat, stage, title, activity, hazard, incident, conseq, affected, ctrls, L, S, rec, hier, rL, rS,
         rationale, specialist, role, freq, actions) = x
        rows.append(dict(
            hazard_id=f"H-{i:03d}", category=C.CATS[cat], stage=C.STAGES[stage], title=title, activity=activity,
            hazard=hazard, incident=incident, consequences=conseq, affected=affected, existing_controls=ctrls,
            L=L, S=S, init_score=L * S, init_cat=C.band(L * S),
            rec_controls=rec, hierarchy=hier, rL=rL, rS=rS, res_score=rL * rS, res_cat=C.band(rL * rS),
            rationale=rationale, specialist=specialist, role=role, frequency=freq,
            priority=C.PRIORITY[C.band(L * S)][0], assess_date=ASSESS_DATES[cat],
            _actions=actions, data_status=C.SIM_TAG))
    return rows


# ---------------------------------------------------------------- incident log
def build_records(hazards):
    hmap = {h["hazard_id"]: h for h in hazards}
    start, end = date(2026, 3, 2), date(2026, 10, 28)
    span = (end - start).days
    tmp = []
    for k, rec in enumerate(REC_RAW):
        (rtype, hno, site, desc, imm, pot, act, outcome, reporter, factor, fu) = rec
        day = start + timedelta(days=rng.randint(0, span))
        while day.weekday() == 6:  # no Sunday dates, purely cosmetic
            day += timedelta(days=1)
        tmp.append([day, k, rtype, hno, site, desc, imm, pot, act, outcome, reporter, factor, fu])
    tmp.sort(key=lambda t: (t[0], t[1]))
    rows = []
    for n, t in enumerate(tmp, 1):
        day, _, rtype, hno, site, desc, imm, pot, act, outcome, reporter, factor, fu = t
        hz = hmap[f"H-{hno:03d}"]
        rows.append(dict(record_id=f"REC-{n:03d}", date=day, month=day.strftime("%Y-%m"), record_type=rtype,
                         category=hz["category"], stage=hz["stage"], hazard_id=hz["hazard_id"],
                         site_code=site, site=C.SITES[site], description=desc, immediate_action=imm,
                         potential_sev=pot, actual_sev=act, outcome=outcome, reporter_role=reporter,
                         factor=factor, followup_key=fu, action_id="", data_status=C.SIM_TAG))
    return rows


# ---------------------------------------------------------------- actions
EVIDENCE = {
    "Engineering": "Close-out report and installation photographs; competent-person sign-off where applicable",
    "Administrative": "Issued procedure or register, training or briefing record, and audit sample",
    "Substitution": "Change-management record and verified replacement of the previous method",
    "PPE": "Task-based PPE assessment and issue record; spot-check observation",
    "Elimination": "Verification that the hazard source has been removed",
}


def draw_status(days_past_target):
    """Seeded outcome mix. Weights are simulation settings, not findings."""
    if days_past_target > 120:
        w = [("closed", 60), ("awaiting", 12), ("progress", 18), ("notstarted", 10)]
    elif days_past_target > 0:
        w = [("closed", 40), ("awaiting", 15), ("progress", 30), ("notstarted", 15)]
    else:
        w = [("closed", 12), ("awaiting", 18), ("progress", 40), ("notstarted", 30)]
    tot = sum(x[1] for x in w)
    r = rng.uniform(0, tot)
    acc = 0
    for name, wt in w:
        acc += wt
        if r <= acc:
            return name
    return w[-1][0]


def make_action(aid, source, hazard, rec_id, issue, action, hier, role, identified, verification):
    init_band = hazard["init_cat"]
    label, sla = C.PRIORITY[init_band]
    lead = C.ENGINEERING_LEAD_TIME_DAYS if hier in ("Engineering", "Substitution", "Elimination") else 0
    target = identified + timedelta(days=sla + lead)
    outcome = draw_status((SD - target).days)
    completion = verified_on = None
    evidence = ""
    recorded = "Not Started"
    verifier = ""
    if outcome in ("closed", "awaiting"):
        done_by = min(SD - timedelta(days=1), target + timedelta(days=25))
        earliest = identified + timedelta(days=7)
        if done_by <= earliest:
            done_by = earliest + timedelta(days=1)
        completion = earliest + timedelta(days=rng.randint(0, max(1, (done_by - earliest).days)))
        if completion >= SD:
            completion = SD - timedelta(days=2)
        recorded = "Completed (Awaiting Verification)"
        if outcome == "closed":
            v = completion + timedelta(days=rng.randint(4, 20))
            if v >= SD:
                outcome = "awaiting"
            else:
                verified_on = v
                recorded = "Closed (Verified)"
                evidence = f"[SIMULATED] EV-{aid[3:]}: {EVIDENCE[hier]}."
                verifier = "HSE Manager" if role != "HSE Manager" else "Operations Manager"
    elif outcome == "progress":
        recorded = "In Progress"
    return dict(action_id=aid, source=source, hazard_id=hazard["hazard_id"], record_id=rec_id, issue=issue,
                action=action, hierarchy=hier, role=role, identified=identified, target=target,
                completion=completion, recorded_status=recorded, verification=verification, verifier=verifier,
                verified_on=verified_on, evidence=evidence, data_status=C.SIM_TAG)


def build_actions(hazards, records):
    hmap = {h["hazard_id"]: h for h in hazards}
    out = []
    n = 0
    for h in hazards:
        for (text, hier, verification) in h["_actions"]:
            n += 1
            out.append(make_action(f"CA-{n:03d}", "Hazard assessment", h, "", h["title"], text, hier, h["role"],
                                   h["assess_date"], verification))
    # follow-ups raised from records, in record-date order
    for rec in records:
        key = rec["followup_key"]
        if not key:
            continue
        hno, source, issue, text, hier, role, verification = FOLLOWUPS[key]
        n += 1
        ident = rec["date"] + timedelta(days=rng.randint(2, 6))
        a = make_action(f"CA-{n:03d}", source, hmap[f"H-{hno:03d}"], rec["record_id"], issue, text, hier, role, ident, verification)
        rec["action_id"] = a["action_id"]
        out.append(a)
    return out


# ---------------------------------------------------------------- derived (Python mirror of the workbook formulas)
def computed_status(a):
    if a["completion"] and a["evidence"] and a["verified_on"]:
        return "Closed (Verified)"
    if a["completion"]:
        return "Completed (Awaiting Verification)"
    if a["target"] < SD:
        return "Overdue"
    return "In Progress" if a["recorded_status"] == "In Progress" else "Not Started"


def derive(hazards, actions, records):
    hmap = {h["hazard_id"]: h for h in hazards}
    by_h = defaultdict(list)
    for a in actions:
        h = hmap[a["hazard_id"]]
        a["category"], a["stage"], a["init_cat"] = h["category"], h["stage"], h["init_cat"]
        a["priority"] = C.PRIORITY[h["init_cat"]][0]
        a["status"] = computed_status(a)
        a["days_overdue"] = (SD - a["target"]).days if a["status"] == "Overdue" else 0
        a["is_open"] = 0 if a["status"] == "Closed (Verified)" else 1
        a["on_time"] = ("Yes" if a["completion"] <= a["target"] else "No") if a["completion"] else ""
        by_h[a["hazard_id"]].append(a)
    for h in hazards:
        acts = by_h[h["hazard_id"]]
        h["n_actions"] = len(acts)
        h["n_closed"] = sum(1 for a in acts if a["status"] == "Closed (Verified)")
        verified = h["n_actions"] > 0 and h["n_closed"] == h["n_actions"]
        h["cur_score"] = h["res_score"] if verified else h["init_score"]
        h["cur_cat"] = C.band(h["cur_score"])
        h["reduction"] = h["init_score"] - h["res_score"]
        cL, cS = C.HIER_CREDIT[h["hierarchy"]]
        h["credit_ok"] = (h["L"] - h["rL"] <= cL) and (h["S"] - h["rS"] <= cS) and h["rL"] >= 1 and h["rS"] >= 1
    return hmap


def build_checklist(hazards):
    rows, n = [], 0
    hmap = {h["hazard_id"]: h for h in hazards}
    for area, items in AREAS:
        for (q, exp, ctype, prio, corr, role, hno) in items:
            n += 1
            rows.append(dict(item_id=f"CL-{n:03d}", area=area, question=q, expected=exp, check_type=ctype,
                             priority=prio, corrective=corr, role=role, hazard_id=f"H-{hno:03d}",
                             response="", observation=""))
    # simulated sample inspection: seeded responses (clearly labelled)
    nc_obs = ["[SIMULATED] Gap seen during the sample walk-through; refer to the tracker.",
              "[SIMULATED] Record incomplete in the sample reviewed.",
              "[SIMULATED] Condition not as expected during the sample inspection."]
    for r_ in rows:
        x = rng.random()
        if r_["check_type"].startswith("Qualified") and x < 0.25:
            r_["response"] = "Non-compliant"
        elif x < 0.20:
            r_["response"] = "Non-compliant"
        elif x < 0.26:
            r_["response"] = "Not Applicable"
        else:
            r_["response"] = "Compliant"
        r_["observation"] = rng.choice(nc_obs) if r_["response"] == "Non-compliant" else (
            "[SIMULATED] Not applicable to the sample site." if r_["response"] == "Not Applicable" else "[SIMULATED] Observed as expected.")
    return rows


def facts(hazards, actions, records, checklist):
    F = {}
    F["status_date"] = SD.isoformat()
    F["n_hazards"] = len(hazards)
    F["n_actions"] = len(actions)
    F["n_records"] = len(records)
    for tag, key in (("init", "init_cat"), ("res", "res_cat"), ("cur", "cur_cat")):
        c = Counter(h[key] for h in hazards)
        F[f"{tag}_bands"] = {b: c.get(b, 0) for b in C.BAND_ORDER}
    st = Counter(a["status"] for a in actions)
    F["status_counts"] = {s: st.get(s, 0) for s in C.COMPUTED_STATUSES}
    F["closed"] = st["Closed (Verified)"]
    F["completed_all"] = st["Closed (Verified)"] + st["Completed (Awaiting Verification)"]
    F["open"] = sum(1 for a in actions if a["is_open"])
    F["overdue"] = st["Overdue"]
    F["pct_closed"] = round(100 * F["closed"] / len(actions), 1)
    F["pct_overdue_of_open"] = round(100 * F["overdue"] / F["open"], 1)
    F["pct_open"] = round(100 * F["open"] / len(actions), 1)
    F["completed_late"] = sum(1 for a in actions if a["on_time"] == "No")
    F["completed_on_time"] = sum(1 for a in actions if a["on_time"] == "Yes")
    od = Counter(a["role"] for a in actions if a["status"] == "Overdue")
    F["overdue_by_role"] = dict(od.most_common())
    F["open_by_role"] = dict(Counter(a["role"] for a in actions if a["is_open"]).most_common())
    F["avg_days_overdue"] = round(sum(a["days_overdue"] for a in actions if a["status"] == "Overdue") / max(1, F["overdue"]), 1)
    F["overdue_by_init_band"] = {b: sum(1 for a in actions if a["status"] == "Overdue" and a["init_cat"] == b) for b in C.BAND_ORDER}
    F["by_category"] = {}
    for cat in C.CAT_NAMES:
        hs = [h for h in hazards if h["category"] == cat]
        F["by_category"][cat] = dict(
            n=len(hs), extreme=sum(1 for h in hs if h["init_cat"] == "Extreme"), high=sum(1 for h in hs if h["init_cat"] == "High"),
            hi_ex=sum(1 for h in hs if h["init_cat"] in ("High", "Extreme")),
            avg_init=round(sum(h["init_score"] for h in hs) / len(hs), 1), avg_res=round(sum(h["res_score"] for h in hs) / len(hs), 1),
            sum_init=sum(h["init_score"] for h in hs), res_hi_ex=sum(1 for h in hs if h["res_cat"] in ("High", "Extreme")),
            open_actions=sum(1 for a in actions if a["category"] == cat and a["is_open"]),
            overdue=sum(1 for a in actions if a["category"] == cat and a["status"] == "Overdue"))
    F["by_stage"] = {}
    for sg in C.STAGES.values():
        hs = [h for h in hazards if h["stage"] == sg]
        F["by_stage"][sg] = dict(n=len(hs), sum_init=sum(h["init_score"] for h in hs), hi_ex=sum(1 for h in hs if h["init_cat"] in ("High", "Extreme")),
                                 sum_res=sum(h["res_score"] for h in hs))
    top = sorted(hazards, key=lambda h: (-h["init_score"], h["hazard_id"]))
    F["top_hazards"] = [dict(id=h["hazard_id"], title=h["title"], score=h["init_score"], cat=h["init_cat"], res=h["res_score"],
                             res_cat=h["res_cat"], role=h["role"]) for h in top[:10]]
    F["sum_init"] = sum(h["init_score"] for h in hazards)
    F["sum_res"] = sum(h["res_score"] for h in hazards)
    F["sum_cur"] = sum(h["cur_score"] for h in hazards)
    F["hi_ex_init"] = F["init_bands"]["High"] + F["init_bands"]["Extreme"]
    F["hi_ex_res"] = F["res_bands"]["High"] + F["res_bands"]["Extreme"]
    F["hi_ex_cur"] = F["cur_bands"]["High"] + F["cur_bands"]["Extreme"]
    # reduction by hierarchy
    rh = defaultdict(lambda: [0, 0, 0])
    for h in hazards:
        rh[h["hierarchy"]][0] += 1
        rh[h["hierarchy"]][1] += h["reduction"]
    F["reduction_by_hierarchy"] = {k: dict(n=v[0], total=v[1], avg=round(v[1] / v[0], 1)) for k, v in rh.items()}
    F["biggest_reductions"] = [dict(id=h["hazard_id"], title=h["title"], reduction=h["reduction"], hier=h["hierarchy"],
                                    init=h["init_score"], res=h["res_score"])
                               for h in sorted(hazards, key=lambda h: (-h["reduction"], h["hazard_id"]))[:8]]
    F["specialist_count"] = sum(1 for h in hazards if h["specialist"].startswith("Yes"))
    F["hier_actions"] = dict(Counter(a["hierarchy"] for a in actions))
    # records
    F["rec_types"] = dict(Counter(r["record_type"] for r in records))
    F["near_miss"] = F["rec_types"].get("Near miss", 0)
    F["rec_by_month"] = {}
    months = sorted({r["month"] for r in records})
    for m in months:
        F["rec_by_month"][m] = {t: sum(1 for r in records if r["month"] == m and r["record_type"] == t) for t in C.RECORD_TYPES}
    F["rec_by_category"] = dict(Counter(r["category"] for r in records).most_common())
    F["rec_by_type_cat"] = {t: dict(Counter(r["category"] for r in records if r["record_type"] == t)) for t in C.RECORD_TYPES}
    F["rec_by_factor"] = dict(Counter(r["factor"] for r in records).most_common())
    F["rec_by_site"] = dict(Counter(r["site"] for r in records).most_common())
    F["rec_high_potential"] = sum(1 for r in records if r["potential_sev"] >= 4)
    F["rec_high_potential_nm"] = sum(1 for r in records if r["potential_sev"] >= 4 and r["record_type"] == "Near miss")
    F["rec_linked_to_hi_ex"] = sum(1 for r in records if next(h for h in hazards if h["hazard_id"] == r["hazard_id"])["init_cat"] in ("High", "Extreme"))
    F["actual_incident_outcomes"] = dict(Counter(r["actual_sev"] for r in records if r["record_type"] == "Incident (actual)"))
    hz_counts = Counter(r["hazard_id"] for r in records)
    F["rec_top_hazards"] = [dict(id=k, n=v, title=next(h["title"] for h in hazards if h["hazard_id"] == k)) for k, v in hz_counts.most_common(6)]
    F["rec_followups"] = sum(1 for r in records if r["action_id"])
    F["checklist"] = dict(Counter(r["response"] for r in checklist))
    ap = F["checklist"].get("Compliant", 0) + F["checklist"].get("Non-compliant", 0)
    F["checklist_compliance_pct"] = round(100 * F["checklist"].get("Compliant", 0) / ap, 1) if ap else 0
    F["checklist_n"] = len(checklist)
    F["checklist_tech"] = sum(1 for r in checklist if r["check_type"].startswith("Qualified"))
    F["checklist_nc_tech"] = sum(1 for r in checklist if r["check_type"].startswith("Qualified") and r["response"] == "Non-compliant")
    return F


def build():
    hazards = build_hazards()
    records = build_records(hazards)
    actions = build_actions(hazards, records)
    hmap = derive(hazards, actions, records)
    checklist = build_checklist(hazards)
    F = facts(hazards, actions, records, checklist)
    return dict(hazards=hazards, records=records, actions=actions, checklist=checklist, facts=F)


# ---------------------------------------------------------------- CSV export
def export(D):
    import csv
    outdir = os.path.join(ROOT, "Data_and_Excel", "csv")
    os.makedirs(outdir, exist_ok=True)

    def w(name, header, rows):
        with open(os.path.join(outdir, name), "w", newline="", encoding="utf-8-sig") as f:
            cw = csv.writer(f)
            cw.writerow(header)
            cw.writerows(rows)

    def dt(x):
        return x.isoformat() if x else ""

    hz = D["hazards"]
    w("hazard_register.csv",
      ["Hazard ID", "Operational Category", "Operational Stage", "Hazard Title", "Operational Activity", "Hazard Description", "Potential Incident",
       "Potential Consequences", "Persons or Assets Affected", "Existing Assumed Controls (illustrative)", "Initial Likelihood", "Initial Severity",
       "Initial Risk Score (calc)", "Initial Risk Category (calc)", "Recommended Additional Controls", "Strongest Control Type", "Residual Likelihood",
       "Residual Severity", "Residual Risk Score (calc)", "Residual Risk Category (calc)", "Residual Rating Basis", "Specialist Assessment Required",
       "Recommended Priority (calc)", "Responsible Role", "Monitoring Frequency", "Assessment Date (simulated)", "Linked Actions (calc)",
       "Closed-Verified Actions (calc)", "Current Risk Score (calc)", "Current Risk Category (calc)", "Data Status"],
      [[h["hazard_id"], h["category"], h["stage"], h["title"], h["activity"], h["hazard"], h["incident"], h["consequences"], h["affected"],
        h["existing_controls"], h["L"], h["S"], h["init_score"], h["init_cat"], h["rec_controls"], h["hierarchy"], h["rL"], h["rS"], h["res_score"],
        h["res_cat"], h["rationale"], h["specialist"], h["priority"], h["role"], h["frequency"], dt(h["assess_date"]), h["n_actions"], h["n_closed"],
        h["cur_score"], h["cur_cat"], h["data_status"]] for h in hz])
    w("corrective_action_tracker.csv",
      ["Action ID", "Source", "Related Hazard ID", "Related Record ID", "Identified Issue", "Required Corrective Action", "Control Hierarchy", "Responsible Role",
       "Date Identified", "Target Completion Date", "Completion Date", "Recorded Status", "Verification Requirement", "Verified By", "Verification Date",
       "Evidence of Closure", "Hazard Category (calc)", "Hazard Stage (calc)", "Initial Risk Category (calc)", "Action Priority (calc)", "Action Status (calc)",
       "Days Overdue (calc)", "Open Flag (calc)", "Completed On Time (calc)", "Data Status"],
      [[a["action_id"], a["source"], a["hazard_id"], a["record_id"], a["issue"], a["action"], a["hierarchy"], a["role"], dt(a["identified"]), dt(a["target"]),
        dt(a["completion"]), a["recorded_status"], a["verification"], a["verifier"], dt(a["verified_on"]), a["evidence"], a["category"], a["stage"],
        a["init_cat"], a["priority"], a["status"], a["days_overdue"], a["is_open"], a["on_time"], a["data_status"]] for a in D["actions"]])
    w("incident_near_miss_log.csv",
      ["Record ID", "Date (simulated)", "Month", "Record Type", "Operational Category", "Operational Stage", "Linked Hazard ID", "Site (simulated)", "Description",
       "Immediate Action Taken", "Potential Severity (1-5)", "Actual Severity (0=no harm)", "Actual Outcome", "Reported By (role)", "Contributing Factor",
       "Linked Action ID", "Data Status"],
      [[r["record_id"], dt(r["date"]), r["month"], r["record_type"], r["category"], r["stage"], r["hazard_id"], r["site"], r["description"], r["immediate_action"],
        r["potential_sev"], r["actual_sev"], r["outcome"], r["reporter_role"], r["factor"], r["action_id"], r["data_status"]] for r in D["records"]])
    w("inspection_checklist.csv",
      ["Item ID", "Operational Area", "Inspection Question", "Expected Safety Condition", "Check Type", "Risk Priority", "Recommended Corrective Action",
       "Responsible Role", "Related Hazard ID", "Response (SIMULATED sample)", "Observation (SIMULATED)"],
      [[c["item_id"], c["area"], c["question"], c["expected"], c["check_type"], c["priority"], c["corrective"], c["role"], c["hazard_id"], c["response"], c["observation"]]
       for c in D["checklist"]])
    w("reference_sources.csv", ["Ref", "Source", "Issuer", "Status", "Date", "URL", "Relevance", "What was verified"],
      [[s["id"], s["name"], s["issuer"], s["status"], s["date"], s["url"], s["relevance"], s["verified"]] for s in SOURCES])
    with open(os.path.join(ROOT, "Data_and_Excel", "facts.json"), "w", encoding="utf-8") as f:
        json.dump(D["facts"], f, indent=2, default=str)
    return outdir


if __name__ == "__main__":
    D = build()
    out = export(D)
    F = D["facts"]
    print("exported to", out)
    for k in ("n_hazards", "n_actions", "n_records", "init_bands", "res_bands", "cur_bands", "status_counts", "pct_closed", "overdue",
              "overdue_by_role", "rec_types", "checklist", "checklist_compliance_pct", "specialist_count", "sum_init", "sum_res"):
        print(k, "=", F[k])
