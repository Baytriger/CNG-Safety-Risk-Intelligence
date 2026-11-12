# Power BI import pack: how to build the report

**Status: import-ready, not a .pbix.** Power BI Desktop is not available in the environment that produced this project, so no .pbix file was created and none is claimed. The pack below lets an analyst with Power BI Desktop build the report in about 30 minutes. The fully working dashboard is the **Dashboard** sheet in `Data_and_Excel/CNG_Safety_Risk_Assessment_Workbook.xlsx`.

All data is **SIMULATED** (independent portfolio case study; no Powergas records were used).

## Contents
| File | Purpose |
|---|---|
| `data/Hazards.csv`, `Actions.csv`, `Records.csv`, `Checklist.csv` | Fact tables exported from the workbook (same record IDs). |
| `data/DimRiskBand.csv`, `DimLikelihood.csv`, `DimSeverity.csv`, `Settings.csv` | Scoring scales, band thresholds and the reporting status date (31 October 2026). |
| `PowerQuery_Load.m` | Eight Power Query queries (create a `DataFolder` parameter first). |
| `DAX_Measures.dax` | Date table, calculated columns and measures. Scores, categories and action status are recalculated from inputs. |
| `Theme_CNG_Safety.json` | Muted industrial theme (View > Themes > Browse for themes). |

## Steps
1. Open Power BI Desktop. Create a Text parameter `DataFolder` with the path to the `data` folder, ending in a backslash.
2. Add the eight queries from `PowerQuery_Load.m`; name them `Hazards`, `Actions`, `Records`, `Checklist`, `DimRiskBand`, `DimLikelihood`, `DimSeverity`, `Settings`. Close and apply.
3. Create `DimDate`, then the calculated columns and measures from `DAX_Measures.dax`.
4. Relationships (all single direction, one-to-many): `Hazards[Hazard ID]` to `Actions[Related Hazard ID]`; `Hazards[Hazard ID]` to `Records[Linked Hazard ID]`; `DimDate[Date]` to `Records[Date (simulated)]`; `DimDate[Date]` to `Actions[Date Identified]` (make this one inactive and activate it with USERELATIONSHIP if needed). `DimLikelihood`, `DimSeverity`, `DimRiskBand` and `Settings` stay disconnected.
5. Import the theme and build the pages below.

## Page 1: Management overview
- **Cards (8):** Total Hazards, Extreme Hazards, High Hazards, Open Actions, Overdue Actions, Closed Actions, Near Miss Records, Pct Actions Closed.
- **Clustered column:** axis `Hazards[Initial Risk Category]` and `Hazards[Residual Risk Category]` counts (initial versus residual); sort by `Initial Band Order`.
- **Matrix heat map:** rows `DimLikelihood[Level]` (descending), columns `DimSeverity[Level]`, values `[Hazards In Cell]`, background from `[Cell Colour]`.
- **Stacked bar:** hazards by `Operational Category`, legend `Initial Risk Category`.
- **Bar:** `Sum Initial Score` and `Sum Residual Score` by `Operational Stage` (risk concentration).
- **Slicers:** Operational Category, Initial Risk Category (Hazards); Responsible Role and Action Status (Actions); date range on `DimDate[Date]`.

## Page 2: Corrective actions
- Bar: Actions by `Action Status`. Bar: Overdue Actions by `Responsible Role`. Table: Action ID, issue, role, target date, Days Overdue, status, sorted by Days Overdue.
- Keep the Overdue visual honest: an action is Overdue only when it has no completion date and its target date is before the status date.

## Page 3: Near misses and conditions
- Column chart by `DimDate[Month]`, legend `Records[Record Type]`. Bar: records by `Operational Category`. Table of records linked to Extreme or High hazards.

## Reconciliation: figures to expect with this dataset
Check these values after loading. If a card differs, a relationship, filter or status date is wrong.

| Measure | Expected |
|---|---|
| Total Hazards | 40 |
| Extreme Hazards (initial 17-25) | 2 |
| High Hazards (initial 10-16) | 14 |
| Total Actions | 54 |
| Closed Actions (verified) | 23 |
| Open Actions | 31 |
| Overdue Actions | 19 |
| Awaiting Verification | 7 |
| Pct Actions Closed | 42.6% |
| Near Miss Records | 16 |
| Total records | 60 |
| Residual High or Extreme | 5 |
| Checklist Compliance % | 86.7% |

## Using real data later
Replace the CSVs with exports from your register, keep the column headers, and set `Settings[StatusDate]` to the review date. Do not mix simulated and real records in one file.
