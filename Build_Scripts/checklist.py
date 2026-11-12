# -*- coding: utf-8 -*-
"""CNG safety inspection checklist items. Observation-only: no item asks anyone to operate, test, adjust or repair
high-pressure equipment. Items marked 'Competent person' are verified by checking the competent person's current
certificate or report, not by performing the technical task.
(area, question, expected, check_type, priority, corrective_action, role, hazard_no)"""
ROUTINE = "Routine visual check"
ADMIN = "Administrative record check"
TECH = "Qualified technical inspection (competent person)"

AREAS = [
 ("1. Compression and processing", [
  ("Is a current vibration and pipe-support review on file for compressor discharge piping?", "Review report dated within the interval set by the Asset Integrity Engineer; actions tracked.", TECH, "High", "Commission the review through the Asset Integrity Engineer; track actions in the tracker.", "Maintenance Manager", 1),
  ("Does the fixed gas-detector register show calibration and function-check dates that are in date?", "Every detector has a recorded, in-date calibration and function-check certificate.", ADMIN, "High", "Update register; arrange vendor service for any detector out of date.", "Maintenance Manager", 2),
  ("Is there a register of safety-system overrides and inhibits, with approvals and time limits?", "No active override without an approval, reason and expiry; handover record lists active inhibits.", ADMIN, "High", "Start the register; remove unapproved inhibits via the Operations Manager.", "Operations Manager", 3),
  ("Are relief-device test and certification dates in date?", "Each device has a certificate within its interval, issued by a competent person.", TECH, "High", "Refer overdue devices to the Asset Integrity Engineer for competent-person action.", "Asset Integrity Engineer", 4),
  ("Is the compressor area free of oil leaks and drips near hot surfaces?", "No visible oil on skids, floor or hot surfaces; drip trays in place.", ROUTINE, "Medium", "Clean; raise a maintenance request for the source of the leak.", "Station Supervisor", 5),
  ("Are operator round sheets complete and abnormal readings escalated?", "Round sheets completed each shift; abnormal readings show a recorded escalation.", ADMIN, "Medium", "Reinforce escalation expectations; review sheets weekly.", "Operations Manager", 6)]),
 ("2. Gas storage", [
  ("Is a consolidated inspection register available for storage vessels, cylinders and piping?", "One register showing last and next inspection by a competent person for each item.", ADMIN, "High", "Consolidate registers; schedule overdue inspections.", "Asset Integrity Engineer", 7),
  ("Is impact protection (bollards or barriers) intact around storage and piping?", "No missing or damaged protection; vehicle routes keep clear of the storage area.", ROUTINE, "High", "Replace damaged protection; review vehicle routes.", "Station Supervisor", 8),
  ("Is the storage area secured against unauthorised access?", "Gate and fence intact, lock working, warning signs legible, visitor log used.", ROUTINE, "Medium", "Repair lock or fence; refresh signs.", "Station Supervisor", 9),
  ("Is the emergency isolation location known, accessible and unobstructed?", "Emergency shutdown buttons and isolation points are visible, labelled and clear.", ROUTINE, "High", "Clear obstructions; label; escalate if remote isolation is not available.", "Operations Manager", 10),
  ("Are ignition sources controlled in and around hazardous areas?", "Hazardous-area boundaries marked; no smoking, phones or unapproved equipment inside.", ROUTINE, "High", "Correct on the spot; review signage and visitor briefing.", "Station Supervisor", 11),
  ("Is electrical equipment in hazardous areas covered by a current competent-person inspection record?", "Inspection record in date; no unapproved modifications noted.", TECH, "High", "Arrange a competent-person electrical inspection.", "Maintenance Manager", 11)]),
 ("3. Loading and offloading", [
  ("Is each transfer hose within its service life and free of visible damage?", "Hose tag in date; no cuts, abrasion, kinks or damaged end fittings.", ROUTINE, "High", "Remove damaged hoses from service; replace.", "Asset Integrity Engineer", 12),
  ("Is the transfer checklist completed with a second-person read-back for the coupling?", "Checklist signed by two people before pressurisation.", ADMIN, "High", "Retrain on the checklist; audit sheets weekly.", "Station Supervisor", 13),
  ("Is trailer earthing and restraint confirmed before transfer starts?", "Earthing connected and chocks in place, recorded on the checklist.", ROUTINE, "High", "Hold the transfer until confirmed; review the checklist.", "Station Supervisor", 14),
  ("Is the exclusion zone marked, barriered and clear of unauthorised people and vehicles?", "Markings visible, barriers in place, only authorised people inside.", ROUTINE, "High", "Re-mark and barrier; brief the loading team.", "Station Supervisor", 15),
  ("Is a working communication method confirmed at the start of each shift?", "Primary and back-up communication tested and logged.", ADMIN, "Medium", "Repair or replace radios; add a back-up method.", "Station Supervisor", 16),
  ("Is a supervisor present or reachable for each transfer, as defined by the procedure?", "Named supervisor recorded on the transfer record.", ADMIN, "Medium", "Revise rostering; reinforce the supervision rule.", "Operations Manager", 13)]),
 ("4. Vehicle and tube trailer inspection", [
  ("Is the tube trailer's periodic inspection certificate in date?", "Certificate issued by a competent person is in date and held with the trailer record.", TECH, "High", "Take the trailer out of service until the certificate is renewed.", "Fleet Manager", 23),
  ("Was the pre-trip checklist completed in full, including brakes, lights, tyres and valve guards?", "All checklist items completed and signed by the driver.", ROUTINE, "High", "Reject the trip until complete; coach the driver.", "Fleet Manager", 23),
  ("Are valve protection, relief outlets and pipework free of visible damage or leakage signs?", "No visible damage, corrosion or leak indication; guards in place.", ROUTINE, "High", "Do not move the trailer; refer to a competent person.", "Fleet Manager", 23),
  ("Are trailer access steps and walkways free of damage and slip hazards?", "Treads intact; no oil or mud build-up.", ROUTINE, "Medium", "Repair or clean before use.", "Fleet Manager", 32),
  ("Is the vehicle's braking, steering and stability equipment covered by a current maintenance record?", "Maintenance record in date; defects tracked to closure.", ADMIN, "High", "Schedule maintenance; remove the vehicle from service if defects are open.", "Fleet Manager", 20),
  ("Are driver and vehicle documents (licence, vehicle papers, emergency information) on board?", "All documents on board and in date.", ADMIN, "Medium", "Hold the trip until provided.", "Logistics Coordinator", 21)]),
 ("5. Transportation readiness", [
  ("Is the planned route recorded, with known risk points and rest stops?", "Route and rest points on the dispatch record.", ADMIN, "Medium", "Complete route planning before release.", "Logistics Coordinator", 21),
  ("Is each driver within the planned driving-hour and rest limits?", "Driver hours recorded and inside company limits.", ADMIN, "High", "Stand the driver down; adjust the schedule.", "Fleet Manager", 19),
  ("Is the customer site assessed and recorded as suitable before the first delivery?", "Site survey on file for each customer.", ADMIN, "High", "Complete a survey before further deliveries.", "Logistics Coordinator", 18),
  ("Are two trained people assigned to each customer connection?", "Names of the connection team recorded.", ADMIN, "High", "Reassign staff; reinforce the two-person rule.", "Logistics Coordinator", 17),
  ("Do drivers carry emergency information cards and the dispatch emergency number?", "Card present and legible.", ROUTINE, "Medium", "Issue cards.", "Emergency Response Coordinator", 40),
  ("Has driver training on trailer stability and rollover risk been recorded?", "Training record in date for each driver.", ADMIN, "Medium", "Schedule training.", "Training Officer", 22)]),
 ("6. Maintenance", [
  ("Is a valid work permit in place before any work on gas-containing equipment?", "Signed permit with isolation points and gas-test result recorded.", ADMIN, "High", "Stop work until the permit is complete.", "Maintenance Manager", 29),
  ("Is isolation verified by a second person and tagged before a line is opened?", "Lock-and-tag records and second-person verification on the permit.", ADMIN, "High", "Stop work; re-isolate and verify.", "Maintenance Manager", 25),
  ("For hot work, is a gas test recorded before the work starts?", "Gas test result and time recorded on the permit by the issuer.", ADMIN, "High", "Stop hot work until tested.", "HSE Manager", 26),
  ("Are contractor inductions and competence records complete for everyone on site?", "Induction record exists for each contractor person.", ADMIN, "High", "Do not allow work to begin until complete.", "Contracts Manager", 28),
  ("Is maintenance on gas equipment carried out only by named, competent people?", "Authorised-person list in use and consistent with who is working.", ADMIN, "High", "Stop unauthorised work; update the authorised list.", "Maintenance Manager", 27),
  ("Are the workshop and yard free of trip hazards and stored materials in walkways?", "Walkways clear; materials stored in marked places.", ROUTINE, "Low", "Clear and correct on the spot.", "Station Supervisor", 30)]),
 ("7. PPE and occupational safety", [
  ("Is PPE specified for each task and available in the right size?", "Task-based PPE list available; stock matches the list.", ADMIN, "Medium", "Complete a task-based PPE assessment.", "HSE Manager", 31),
  ("Is the PPE in use in good condition and worn as specified?", "No damaged or missing PPE observed.", ROUTINE, "Medium", "Replace and coach on the spot.", "Station Supervisor", 31),
  ("Are floors, bay surfaces and steps free of oil, water and damage?", "Dry, clean surfaces; spills cleaned promptly.", ROUTINE, "Medium", "Clean and treat surface; repair damage.", "Station Supervisor", 32),
  ("Are lifting aids provided and used for hoses and heavy fittings?", "Lifting aids available; manual handling guidance posted.", ROUTINE, "Low", "Provide aids; brief the team.", "HSE Manager", 33),
  ("Does the training matrix show everyone's required courses and refreshers in date?", "Matrix complete and in date for each role.", ADMIN, "High", "Schedule overdue training.", "Training Officer", 34),
  ("Are hazard and near-miss reports being submitted and answered in the last month?", "Reports logged with a recorded response.", ADMIN, "Medium", "Reinforce reporting; give feedback to reporters.", "HSE Manager", 35)]),
 ("8. Emergency preparedness", [
  ("Do staff know how to report a suspected gas leak, and is the route posted?", "Reporting card posted; staff can state the first action.", ROUTINE, "High", "Brief staff; post the card.", "Emergency Response Coordinator", 36),
  ("Is a muster point defined, and does the headcount reconcile with the site and visitor logs?", "Last drill reconciled people on site with people mustered.", ADMIN, "High", "Fix the visitor log process; repeat the drill.", "Emergency Response Coordinator", 37),
  ("Is a back-up emergency communication method available and tested?", "Test recorded within the last quarter.", ADMIN, "Medium", "Provide and test a back-up method.", "Emergency Response Coordinator", 38),
  ("Does the drill programme include a tube-trailer release and fire at the loading bay?", "Scenario scheduled or completed with a recorded debrief.", ADMIN, "Medium", "Add the scenario and run it.", "Emergency Response Coordinator", 39),
  ("Are fire-fighting and gas-detection equipment inspected by a competent person within their intervals?", "Inspection certificates in date.", TECH, "High", "Arrange competent-person inspection.", "Maintenance Manager", 37),
  ("Are external responder contacts and route-incident arrangements current?", "Contact list updated within the last six months.", ADMIN, "Medium", "Update the list; confirm liaison.", "Emergency Response Coordinator", 40)]),
]
