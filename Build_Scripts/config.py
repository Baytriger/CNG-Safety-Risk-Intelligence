# -*- coding: utf-8 -*-
"""Single source of truth for scoring rules, labels and project constants.
Every script (workbook, charts, report, deck, QA) imports from here so numbers cannot drift."""
from datetime import date

PROJECT_TITLE = "CNG Safety Risk Intelligence and Preventive Action System"
SHORT_TITLE = "CNG Safety Risk Intelligence System"
CASE_STUDY_FOR = "Powergas Global Investment Nigeria Limited and Powergas Ebedei Limited"
DISCLOSURE = ("Independent portfolio case study. All hazards, observations, incidents, near misses, "
              "facilities, personnel, dates and outcomes are HYPOTHETICAL and SIMULATED. "
              "No Powergas records, employees, facilities or findings were used or accessed.")
SIM_TAG = "SIMULATED"

# Simulated reporting ("status") date. All overdue logic is evaluated against this date so that
# published numbers are reproducible. In the workbook it is a single cell on the Settings sheet and can be
# replaced with =TODAY() for live use.
STATUS_DATE = date(2026, 10, 31)
SIM_START = date(2026, 1, 5)

# ---- Risk bands: Risk Score = Likelihood x Severity -------------------------------------------
BANDS = [  # name, min score, max score, required response
    ("Low", 1, 4, "Manage by routine procedures; review at the next scheduled assessment."),
    ("Moderate", 5, 9, "Management attention required; actions planned and tracked within 90 days."),
    ("High", 10, 16, "Senior management attention; actions started promptly and completed within 60 days; competent-person review where flagged."),
    ("Extreme", 17, 25, "Unacceptable without immediate additional control; consider stopping or restricting the activity; actions within 30 days."),
]
BAND_ORDER = [b[0] for b in BANDS]
BAND_COLOURS = {"Low": "C6E0B4", "Moderate": "FFE699", "High": "F4B183", "Extreme": "E06666"}


def band(score: int) -> str:
    for name, lo, hi, _ in BANDS:
        if lo <= score <= hi:
            return name
    raise ValueError(score)


# ---- Action priority by INITIAL risk band -----------------------------------------------------
PRIORITY = {  # band -> (label, target days from date identified)
    "Extreme": ("P1 Immediate", 30),
    "High": ("P2 Urgent", 60),
    "Moderate": ("P3 Planned", 90),
    "Low": ("P4 Routine", 180),
}
ENGINEERING_LEAD_TIME_DAYS = 30  # added to target for Engineering/Substitution/Elimination actions (procurement, design)

LIKELIHOOD = {
    1: ("Rare", "Not expected during the life of the operation; would need several independent control failures."),
    2: ("Unlikely", "Has happened in the industry but is not expected at this operation within several years."),
    3: ("Possible", "Could occur within a few years; relies partly on consistent human performance or unverified controls."),
    4: ("Likely", "Expected to occur within a year; known weaknesses exist in the controls."),
    5: ("Almost certain", "Expected to occur repeatedly; the conditions for the event exist routinely."),
}
SEVERITY = {  # level: (label, people, process safety, asset/operations, environment/community)
    1: ("Insignificant", "No injury.", "Negligible weep; no flammable accumulation.", "No downtime; no damage.", "No off-site effect."),
    2: ("Minor", "First aid case.", "Small, quickly isolated release; no ignition potential realised.", "Less than one shift lost; minor repair.", "Minor on-site spill contained."),
    3: ("Moderate", "Medical treatment, restricted work or lost-time injury.", "Release needing isolation and area evacuation.", "Equipment damage; outage of days.", "Reportable on-site release; local nuisance."),
    4: ("Major", "Serious injury or permanent disability.", "Significant release with credible ignition; fire or flash-fire potential.", "Major equipment damage; extended outage.", "Alert or evacuation beyond the fence line."),
    5: ("Catastrophic", "Fatality or multiple serious injuries.", "Jet fire, explosion or rupture of high-pressure equipment.", "Loss of a facility or trailer; prolonged shutdown.", "Public harm or major community impact."),
}

# Maximum reduction (in levels) the strongest recommended control may claim when rating RESIDUAL risk.
# (likelihood levels, severity levels)
HIER_CREDIT = {"Elimination": (4, 4), "Substitution": (3, 3), "Engineering": (2, 1), "Administrative": (1, 0), "PPE": (0, 1)}
HIER = ["Elimination", "Substitution", "Engineering", "Administrative", "PPE"]

ROLES = ["HSE Manager", "Operations Manager", "Station Supervisor", "Maintenance Manager",
         "Asset Integrity Engineer", "Fleet Manager", "Logistics Coordinator", "Training Officer",
         "Emergency Response Coordinator", "Contracts Manager"]

CATS = {"A": "Gas Compression & Processing", "B": "CNG Storage", "C": "Loading & Offloading",
        "D": "Transportation & Fleet", "E": "Maintenance & Contractors", "F": "Occupational Safety",
        "G": "Emergency Preparedness"}
CAT_NAMES = list(CATS.values())
STAGES = {1: "1 Compression & processing", 2: "2 Storage", 3: "3 Transfer (loading/offloading)",
          4: "4 Road transport", 5: "5 Customer delivery site", 6: "6 Maintenance & contractor work",
          7: "7 Site support (occupational & emergency)"}

ACTION_STATUSES = ["Not Started", "In Progress", "Completed (Awaiting Verification)", "Closed (Verified)"]
COMPUTED_STATUSES = ["Not Started", "In Progress", "Overdue", "Completed (Awaiting Verification)", "Closed (Verified)"]
RECORD_TYPES = ["Near miss", "Incident (actual)", "Unsafe condition", "Unsafe act", "Inspection finding"]

SITES = {  # all fictional
    "CS-A": "Simulated Compression Station A",
    "LB-A": "Simulated Loading Bay A",
    "DP-B": "Simulated Fleet Depot B",
    "CU-1": "Simulated Customer Site 1",
    "CU-2": "Simulated Customer Site 2",
    "WS-A": "Simulated Workshop A",
    "RD":   "Simulated Road Route",
}
