# QA audit

Independent portfolio case study. All data simulated. Audit run 4 November 2026 on the delivered workbook.

## Method
1. The workbook was recalculated from scratch in LibreOffice (separate from the build process) and every sheet was scanned for formula errors: **0 errors**; 1591 formulas (count at build).
2. Row-level results (scores, categories, residual checks, action status, days overdue) read from the recalculated workbook were compared with an independent Python calculation: **0 mismatches**.
3. The Excel dashboard filters were set to seven different combinations (category, risk level, role, status, date range and combined). The eleven headline KPIs were compared with a Python calculation each time: **all seven combinations match**. Script: `Build_Scripts/qa/qa_filter_test.py`.
4. Data and logic checks (`Build_Scripts/qa/qa_audit.py`):

| Result | Check | Detail |
|---|---|---|
| PASS | Hazard IDs unique | 40 hazards |
| PASS | Action IDs unique | 54 actions |
| PASS | Record IDs unique | 60 records |
| PASS | Checklist IDs unique | 48 items |
| PASS | Scores = L x S |  |
| PASS | Band boundaries consistent |  |
| PASS | Residual <= initial (L and S) |  |
| PASS | Residual within hierarchy credit limits |  |
| PASS | Action->hazard references valid |  |
| PASS | Record->hazard references valid |  |
| PASS | Record->action references valid |  |
| PASS | Every hazard has >=1 action |  |
| PASS | Action status = independent recomputation |  |
| PASS | No overdue action has a completion date |  |
| PASS | No action closed without date+evidence+verification |  |
| PASS | Past-target open actions are never Closed |  |
| PASS | No future-dated completions/records |  |
| PASS | Record types limited to 5 defined types |  |
| PASS | >=35 hazards, >=50 records |  |
| PASS | All rows tagged SIMULATED |  |
| PASS | No record/hazard text mentions Powergas as source |  |
| PASS | Workbook Hazard Register formulas = Python (scores, categories, residual check) | 0 mismatches |
| PASS | Workbook tracker status/overdue formulas = Python; Status Check all OK | 0 mismatches |
| PASS | Every reference has a URL and status | 15 sources |

## Reconciliation of headline figures (all files use these)
Hazards 40; Extreme/High initial 2/14; projected residual High or Extreme 5; current High or Extreme 14; actions 54; closed-verified 23 (42.6%); open 31; overdue 19; near misses 16; records 60; checklist compliance 86.7%. The workbook dashboard, report, executive summary, deck and Power BI guide read these from the same source tables.

## Not verified automatically (manual verification required)
- **Desktop Excel:** formulas were recalculated and tested in LibreOffice. Open the workbook once in Microsoft Excel to confirm that charts, dropdown filters and conditional formatting render as intended.
- **Power BI:** no .pbix was created (Power BI Desktop unavailable). The import pack must be assembled and its cards reconciled with the table in `PowerBI_Build_Guide.md`.
- **Reference links:** each source was opened with the fetch tool while writing the register. The build shell cannot reach external sites, so links were not re-tested by script; click each link before submission. The ISO catalogue page could not be retrieved, so ISO 45001 details rest on a secondary catalogue.
- **Regulatory content:** clause-level obligations in the NMDPRA regulations, Factories Act and Employee's Compensation Act were not extracted. Obtain primary texts and legal advice before relying on any of them.
- **Technical judgement:** simulated scores, controls and assumptions need review by a competent process-safety engineer before any real use.
