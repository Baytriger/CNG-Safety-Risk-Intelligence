# -*- coding: utf-8 -*-
"""SIMULATED incident, near-miss, unsafe-condition, unsafe-act and inspection-finding records (60).
Every record is fictional. Nothing here describes Powergas. Severity of outcomes is deliberately modest."""

# Follow-up corrective actions raised from records: key -> (hazard_no, source, issue, action, hierarchy, role, verification)
FOLLOWUPS = {
    "F01": (12, "Incident follow-up", "Abraded transfer hose found at pre-use check.", "Introduce hose retirement tags with a fixed service life and a quarantine rack for suspect hoses.", "Administrative", "Asset Integrity Engineer", "Hose register reconciled with tags by HSE Manager; quarantine rack inspected."),
    "F02": (15, "Incident follow-up", "Third-party vehicle entered the transfer exclusion area.", "Install fixed barriers and refresh floor and sign markings at the loading bay.", "Engineering", "Station Supervisor", "Photo and walk-through verification; barrier positions checked against the layout drawing."),
    "F03": (25, "Incident follow-up", "Isolation valve found not fully closed during a zero-energy check.", "Add a documented second-person verification and lock-and-tag record to the isolation step of the permit.", "Administrative", "Maintenance Manager", "Sample of ten permits audited for the second-person signature."),
    "F04": (26, "Incident follow-up", "Hot work began before a gas-test result was recorded.", "Make a recorded gas test a hold point on the hot-work permit, with the permit issuer physically present.", "Administrative", "HSE Manager", "Permit audit; supervisor observation of three hot-work starts."),
    "F05": (3, "Incident follow-up", "Alarm inhibit not restored at shift handover.", "Add active inhibits to the shift handover record and display inhibited alarms on the control screen.", "Engineering", "Operations Manager", "Handover records audited; screen display witnessed."),
    "F06": (14, "Incident follow-up", "Trailer earthing clamp not attached before transfer start.", "Add earthing verification to the transfer checklist and evaluate a transfer permissive that depends on earthing proof.", "Engineering", "Station Supervisor", "Checklist audit; engineering evaluation report reviewed by Asset Integrity Engineer."),
    "F07": (18, "Incident follow-up", "Delivery aborted: customer site access unsuitable for the trailer.", "Introduce a pre-delivery site survey template, completed before the first delivery to any customer.", "Administrative", "Logistics Coordinator", "Completed surveys sampled for all new customer sites in the quarter."),
    "F08": (36, "Incident follow-up", "Gas odour reported late by an operator.", "Run a short scenario briefing and drill on immediate reporting of suspected gas leaks, with a reporting card at each station.", "Administrative", "Emergency Response Coordinator", "Drill record; random check that staff can state the reporting route."),
    "F09": (16, "Incident follow-up", "Radio failure during a transfer.", "Provide a back-up communication method for each transfer point and test it at shift start.", "Engineering", "Station Supervisor", "Shift-start test log; witnessed test by Operations Manager."),
    "F10": (29, "Incident follow-up", "Contractor began a task before the permit was signed.", "Retrain permit issuers and contractor supervisors on the permit hold point; add a gate check that no work begins without a signed permit.", "Administrative", "Contracts Manager", "Training record; gate-check sample audit."),
    "F11": (8, "Incident follow-up", "Reversing vehicle came close to the storage protection line.", "Install additional impact protection and a marked reversing route away from the storage area.", "Engineering", "Station Supervisor", "Inspection of installed protection against drawing; traffic route walk-through."),
    "F12": (20, "Incident follow-up", "Trailer air-brake leak found at a roadside stop.", "Add braking-system checks and a recorded fault-reporting step to the pre-trip routine.", "Administrative", "Fleet Manager", "Pre-trip sheets sampled weekly for a month by Fleet Manager."),
    "F13": (32, "Incident follow-up", "Operator slipped on an oily bay surface.", "Degrease and apply a non-slip surface treatment to the bay, with a spill-clean-up standard.", "Engineering", "Station Supervisor", "Surface inspection after treatment; spill clean-up checked during rounds."),
    "F14": (34, "Inspection finding", "Refresher training overdue for several operators.", "Build a training matrix with due dates by role and schedule overdue refreshers.", "Administrative", "Training Officer", "Matrix reconciled with attendance records by HSE Manager."),
}

# (type, hazard_no, site, description, immediate_action, potential_severity, actual_severity (0=no harm), outcome, reporter, factor, followup)
R = []
def r(*a): R.append(a)

# ---- Incidents (actual, minor) ----
r("Incident (actual)",12,"CU-1","Minor gas weep detected at a hose end fitting during the post-connection leak check at a customer site.","Transfer stopped, line isolated, fitting replaced by competent person before resuming.",3,1,"Small release, no injury, delivery delayed about 40 minutes.","Driver","Equipment or condition",None)
r("Incident (actual)",32,"LB-A","Operator slipped on an oily patch beside the loading bay and sprained a wrist.","First aid; area cleaned; operator returned to light duties.",2,2,"First aid case; one shift of restricted work.","Operator","Equipment or condition","F13")
r("Incident (actual)",21,"DP-B","Trailer mirror struck a gatepost while manoeuvring in the depot yard.","Vehicle checked, damage recorded, no gas equipment affected.",2,1,"Minor property damage.","Driver","Procedure or planning",None)
r("Incident (actual)",5,"CS-A","Small lubricating-oil leak from a compressor skid reached the yard drain before it was contained.","Leak stopped, spill kit used, drain cleaned.",2,1,"Minor on-site spill contained.","Operator","Equipment or condition",None)
r("Incident (actual)",20,"RD","Laden trailer broke down on a road shoulder with a brake-air fault; no gas release.","Hazard warning set out, dispatch notified, recovery arranged.",4,1,"Delay and temporary lane obstruction; no release or injury.","Driver","Equipment or condition","F12")
r("Incident (actual)",2,"CS-A","Precautionary compressor shutdown after a gas-detector fault alarm; fault traced to a failed sensor head.","Unit shut down as designed, sensor replaced by vendor.",2,1,"About four hours of lost production; no gas release.","Operator","Equipment or condition",None)
r("Incident (actual)",30,"WS-A","Contractor technician received a minor burn from an uninsulated hot surface during routine work.","First aid given; area marked.",2,2,"First aid case.","Contractor","Human performance",None)
r("Incident (actual)",24,"CU-2","Vehicle reversed into a customer-site bollard because no banksman was available.","Delivery completed with a second person guiding; damage noted.",3,1,"Minor property damage.","Driver","Procedure or planning",None)

# ---- Near misses ----
r("Near miss",13,"LB-A","A coupling was found only part-engaged by a second person before pressurisation began.","Transfer paused and coupling reseated; second-person check confirmed.",4,0,"No release.","Operator","Human performance",None)
r("Near miss",12,"LB-A","A hose with visible outer abrasion was noticed at the pre-transfer check and had been issued from the rack.","Hose tagged out and replaced.",4,0,"No release.","Operator","Equipment or condition","F01")
r("Near miss",15,"LB-A","A third-party delivery vehicle drove into the marked exclusion area during a transfer.","Supervisor stopped the vehicle and paused the transfer until it left.",4,0,"No ignition or injury.","Supervisor","Procedure or planning","F02")
r("Near miss",19,"RD","A driver reported briefly losing concentration late in a long return leg and stopped at the next rest point.","Driver rested; dispatch revised the next day's plan.",5,0,"No collision.","Driver","Management system",None)
r("Near miss",22,"RD","A car cut across the path of a laden trailer; the driver braked and avoided contact.","Driver reported to dispatch; no vehicle damage.",5,0,"No collision.","Driver","Human performance",None)
r("Near miss",25,"WS-A","An isolation valve was found not fully closed during the zero-energy check before a flange was opened.","Work stopped; valve closed and re-verified; permit annotated.",5,0,"No release.","Contractor","Procedure or planning","F03")
r("Near miss",26,"WS-A","Hot work was about to start before the gas-test result was recorded on the permit.","Supervisor intervened; gas test performed and recorded.",5,0,"No ignition.","Supervisor","Procedure or planning","F04")
r("Near miss",3,"CS-A","A high-vibration alarm inhibited during troubleshooting was not restored at shift handover; found on rounds.","Inhibit removed and logged.",5,0,"No equipment damage.","Operator","Procedure or planning","F05")
r("Near miss",14,"LB-A","The trailer earthing clamp was not attached; caught when the transfer checklist was read back.","Transfer held until earthing was connected.",4,0,"No release.","Operator","Human performance","F06")
r("Near miss",18,"CU-2","The access road at a new customer site was too narrow for a safe turn; delivery aborted before the trailer neared live equipment.","Delivery rescheduled; site referred to Logistics.",4,0,"No harm.","Driver","Procedure or planning","F07")
r("Near miss",36,"CS-A","An operator noticed a gas odour and waited about ten minutes before reporting it; a minor seep was later confirmed and isolated.","Area checked with fixed detection and by the supervisor; fitting tightened by a competent person.",4,0,"No injury.","Operator","Human performance","F08")
r("Near miss",16,"LB-A","The handheld radio failed during a transfer and the operator could not reach the control room.","Transfer stopped as the procedure requires until communication was restored.",3,0,"No release.","Operator","Equipment or condition","F09")
r("Near miss",29,"WS-A","A contractor began a task before the work permit had been signed.","Work stopped by the supervisor and permit completed.",4,0,"No harm.","Supervisor","Procedure or planning","F10")
r("Near miss",8,"CS-A","A reversing truck stopped short of the storage protection line after the driver saw the bollards.","Driver briefed; reversing route reviewed.",4,0,"No contact.","Driver","Human performance","F11")
r("Near miss",20,"DP-B","A brake-air leak on a trailer was found at the end of a trip; the trailer returned to the depot without incident.","Trailer removed from service for repair.",4,0,"No loss of control.","Driver","Equipment or condition",None)
r("Near miss",6,"CS-A","A low suction-pressure alarm was acknowledged but not investigated until the next round, where a trend was spotted.","Operator checked the supply line; trend reviewed by the supervisor.",3,0,"No equipment damage.","Operator","Human performance",None)

# ---- Unsafe conditions ----
r("Unsafe condition",32,"LB-A","Oil film on the loading-bay floor near the hose rack.","Cleaned up; cleaning frequency increased.",2,0,"No harm.","Supervisor","Equipment or condition",None)
r("Unsafe condition",12,"LB-A","Transfer hoses left coiled on the ground rather than on the rack.","Hoses returned to rack.",3,0,"No harm.","Operator","Procedure or planning",None)
r("Unsafe condition",8,"CS-A","One impact bollard near the storage area was missing after earlier damage.","Temporary barrier placed; replacement requested.",4,0,"No harm.","Supervisor","Equipment or condition",None)
r("Unsafe condition",37,"WS-A","A workshop emergency exit partly blocked by stored materials.","Cleared; storage area marked.",4,0,"No harm.","Supervisor","Procedure or planning",None)
r("Unsafe condition",2,"CS-A","One fixed gas detector displayed a fault state for two shifts before being reported.","Reported; sensor replaced.",4,0,"No release.","Operator","Equipment or condition",None)
r("Unsafe condition",4,"CS-A","A relief-valve certification tag showed a test date past its interval.","Referred to Asset Integrity Engineer for competent-person review.",5,0,"No overpressure event.","Operator","Management system",None)
r("Unsafe condition",15,"LB-A","Exclusion-zone floor markings faded and partly invisible.","Re-marked; signs checked.",4,0,"No harm.","Supervisor","Equipment or condition",None)
r("Unsafe condition",30,"WS-A","Loose cables and spare fittings in a workshop walkway.","Cleared; housekeeping round added.",2,0,"No harm.","Supervisor","Procedure or planning",None)
r("Unsafe condition",1,"CS-A","A small-bore fitting on discharge piping showed a slight weep in the leak check.","Isolated and referred to a competent person.",4,0,"No injury.","Operator","Equipment or condition",None)
r("Unsafe condition",9,"CS-A","The padlock on the storage-area gate was found broken.","Gate secured with a temporary lock; replacement fitted.",4,0,"No unauthorised entry recorded.","Supervisor","Equipment or condition",None)
r("Unsafe condition",20,"DP-B","A trailer marker light was found inoperative at the pre-trip check.","Repaired before departure.",3,0,"No harm.","Driver","Equipment or condition",None)
r("Unsafe condition",32,"DP-B","Worn tread on a trailer access step.","Step replaced.",2,0,"No harm.","Driver","Equipment or condition",None)
r("Unsafe condition",10,"CS-A","An emergency shutdown push-button partly hidden behind stored equipment.","Equipment moved; button area marked.",4,0,"No harm.","Operator","Procedure or planning",None)
r("Unsafe condition",14,"LB-A","An earthing cable end was found corroded at the loading bay.","Cable replaced.",4,0,"No release.","Operator","Equipment or condition",None)
r("Unsafe condition",25,"WS-A","Valves on a maintenance line were unlabelled, making positive identification harder during isolation.","Valves labelled.",5,0,"No release.","Contractor","Equipment or condition",None)
r("Unsafe condition",5,"CS-A","Oil drips under a compressor skid close to a warm surface.","Cleaned; leak referred to maintenance.",3,0,"No ignition.","Operator","Equipment or condition",None)

# ---- Unsafe acts ----
r("Unsafe act",31,"WS-A","A technician handled hoses without gloves.","Corrected on the spot; briefed at toolbox talk.",2,0,"No harm.","Operator","Human performance",None)
r("Unsafe act",11,"CS-A","A visitor used a mobile phone close to the hazardous-area boundary.","Phone put away; visitor briefing reinforced.",4,0,"No ignition.","Supervisor","Human performance",None)
r("Unsafe act",24,"CU-1","A driver reversed at a customer site without a guide.","Stopped by the receiving customer's staff; driver coached.",3,0,"No contact.","Driver","Human performance",None)
r("Unsafe act",17,"CU-2","A single driver made the connection at a customer site where the procedure calls for a second person.","Reviewed with the driver; second person assigned.",4,0,"No release.","Driver","Procedure or planning",None)
r("Unsafe act",23,"DP-B","A pre-trip checklist was ticked without all items being physically checked.","Driver re-did the checks; reviewed with Fleet Manager.",4,0,"No defect missed on this occasion.","Supervisor","Human performance",None)
r("Unsafe act",3,"CS-A","A bypass key was left in a panel after a test.","Key returned; key control reinforced.",4,0,"No loss of protection.","Operator","Human performance",None)
r("Unsafe act",15,"LB-A","A person took a shortcut through the marked exclusion zone while a transfer was active.","Stopped by the supervisor.",4,0,"No harm.","Supervisor","Human performance",None)
r("Unsafe act",33,"CS-A","A hose coil was lifted by one person with a poor posture.","Technique corrected; lifting aid provided.",2,0,"No harm.","Operator","Human performance",None)
r("Unsafe act",19,"RD","A driver ran ahead of the planned driving hours to complete one extra delivery.","Driver stopped at the next rest point; dispatch informed.",5,0,"No collision.","Driver","Management system",None)
r("Unsafe act",11,"CS-A","Smoking was seen just outside the marked hazardous-area boundary near the yard.","Person redirected to the designated area; signs reviewed.",4,0,"No ignition.","Supervisor","Human performance",None)

# ---- Inspection findings ----
r("Inspection finding",2,"CS-A","Detector register has no calibration due dates recorded.","Register creation requested.",4,0,"Administrative gap.","Inspector","Management system",None)
r("Inspection finding",39,"CS-A","Annual drill programme has no tube-trailer release scenario.","Scenario added to the programme for next quarter.",3,0,"Planning gap.","Inspector","Management system",None)
r("Inspection finding",34,"CS-A","Refresher training records missing for several operators.","Training matrix requested.",4,0,"Administrative gap.","Inspector","Management system","F14")
r("Inspection finding",31,"WS-A","PPE issued is generic and not specified by task.","Task-based PPE assessment requested.",3,0,"Planning gap.","Inspector","Management system",None)
r("Inspection finding",28,"WS-A","Contractor induction records incomplete for two visiting technicians.","Induction completed before next visit.",4,0,"Administrative gap.","Inspector","Management system",None)
r("Inspection finding",40,"DP-B","Emergency contact list for route incidents is out of date.","List updated and posted.",3,0,"Administrative gap.","Inspector","Management system",None)
r("Inspection finding",37,"CS-A","Muster point headcount did not reconcile with the visitor log in an unannounced drill.","Visitor log process reviewed.",4,0,"Drill shortfall.","Inspector","Procedure or planning",None)
r("Inspection finding",29,"WS-A","Several completed work permits had blank verification fields.","Permit audit scheduled.",5,0,"Administrative gap.","Inspector","Management system",None)
r("Inspection finding",23,"DP-B","Trailer pre-trip sheets completed inconsistently across drivers.","Standard sheet issued.",4,0,"Administrative gap.","Inspector","Management system",None)
r("Inspection finding",7,"CS-A","Storage inspection dates are held on separate lists and not consolidated.","Consolidated register requested.",5,0,"Administrative gap.","Inspector","Management system",None)

assert len(R) == 60, len(R)
