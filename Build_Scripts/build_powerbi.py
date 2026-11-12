# -*- coding: utf-8 -*-
"""Builds the import-ready Power BI pack: CSV tables, Power Query (M), DAX measures, theme and page specification.
No .pbix file is produced: Power BI Desktop is not available in this environment.
Run: python3 -I build_powerbi.py"""
import csv, json, os, shutil, sys
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import config as C, build_data

PK = os.path.join(ROOT, "Dashboard", "PowerBI_Import_Pack"); DATA = os.path.join(PK, "data")
os.makedirs(DATA, exist_ok=True)
D = build_data.build(); build_data.export(D); F = D["facts"]
SRC = os.path.join(ROOT, "Data_and_Excel", "csv")
for a, b in (("hazard_register.csv", "Hazards.csv"), ("corrective_action_tracker.csv", "Actions.csv"),
             ("incident_near_miss_log.csv", "Records.csv"), ("inspection_checklist.csv", "Checklist.csv")):
    shutil.copy(os.path.join(SRC, a), os.path.join(DATA, b))


def w(name, header, rows):
    with open(os.path.join(DATA, name), "w", newline="", encoding="utf-8-sig") as f:
        cw = csv.writer(f); cw.writerow(header); cw.writerows(rows)


w("DimRiskBand.csv", ["BandName", "MinScore", "MaxScore", "SortOrder", "RequiredResponse"],
  [[n, lo, hi, i + 1, r] for i, (n, lo, hi, r) in enumerate(C.BANDS)])
w("DimLikelihood.csv", ["Level", "Label", "Definition"], [[k, v[0], v[1]] for k, v in C.LIKELIHOOD.items()])
w("DimSeverity.csv", ["Level", "Label", "People", "ProcessSafety", "AssetOperations", "EnvironmentCommunity"], [[k, *v] for k, v in C.SEVERITY.items()])
w("Settings.csv", ["Name", "StatusDate", "Note"], [["Reporting status date", C.STATUS_DATE.isoformat(),
  "Simulated reporting date used for overdue logic. Replace with the real review date, or use TODAY() in a measure for live use."]])

M = r'''// Power Query (M): paste each query into Power BI Desktop > Transform data > New Source > Blank Query > Advanced Editor.
// First create a Text parameter named DataFolder whose value is the full path of the "data" folder (for example C:\CNG_Safety_Risk_Intelligence\Dashboard\PowerBI_Import_Pack\data\).
// All records are SIMULATED. Columns ending "(calc)" are dropped on load; the DAX file recomputes every score, category and status from the inputs.

// ---------- Hazards
let
    Source = Csv.Document(File.Contents(DataFolder & "Hazards.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Initial Likelihood", Int64.Type}, {"Initial Severity", Int64.Type}, {"Residual Likelihood", Int64.Type}, {"Residual Severity", Int64.Type}, {"Assessment Date (simulated)", type date}}),
    Dropped = Table.RemoveColumns(Typed, List.Select(Table.ColumnNames(Typed), each Text.EndsWith(_, "(calc)")), MissingField.Ignore)
in
    Dropped

// ---------- Actions
let
    Source = Csv.Document(File.Contents(DataFolder & "Actions.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Date Identified", type date}, {"Target Completion Date", type date}, {"Completion Date", type date}, {"Verification Date", type date}}),
    Blanks = Table.ReplaceValue(Typed, "", null, Replacer.ReplaceValue, {"Evidence of Closure", "Verified By", "Related Record ID"}),
    Dropped = Table.RemoveColumns(Blanks, List.Select(Table.ColumnNames(Blanks), each Text.EndsWith(_, "(calc)")), MissingField.Ignore)
in
    Dropped

// ---------- Records (incident and near-miss log)
let
    Source = Csv.Document(File.Contents(DataFolder & "Records.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Date (simulated)", type date}, {"Potential Severity (1-5)", Int64.Type}, {"Actual Severity (0=no harm)", Int64.Type}})
in
    Typed

// ---------- Checklist
let
    Source = Csv.Document(File.Contents(DataFolder & "Checklist.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
in
    Promoted

// ---------- DimRiskBand
let
    Source = Csv.Document(File.Contents(DataFolder & "DimRiskBand.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"MinScore", Int64.Type}, {"MaxScore", Int64.Type}, {"SortOrder", Int64.Type}})
in
    Typed

// ---------- DimLikelihood
let
    Source = Csv.Document(File.Contents(DataFolder & "DimLikelihood.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Level", Int64.Type}})
in
    Typed

// ---------- DimSeverity
let
    Source = Csv.Document(File.Contents(DataFolder & "DimSeverity.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Level", Int64.Type}})
in
    Typed

// ---------- Settings
let
    Source = Csv.Document(File.Contents(DataFolder & "Settings.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"StatusDate", type date}})
in
    Typed
'''
open(os.path.join(PK, "PowerQuery_Load.m"), "w", encoding="utf-8").write(M)

DAX = r'''// DAX for the CNG Safety Risk Intelligence model. Create these in Power BI Desktop (Modeling > New table / New column / New measure).
// Nothing here is a typed-in KPI value: every figure is calculated from the loaded tables, and risk bands come from DimRiskBand.
// All data is SIMULATED.

// ===== Date table (new table) =====
DimDate =
ADDCOLUMNS (
    CALENDAR ( DATE ( 2026, 1, 1 ), DATE ( 2026, 12, 31 ) ),
    "Month", FORMAT ( [Date], "yyyy-MM" ),
    "MonthStart", DATE ( YEAR ( [Date] ), MONTH ( [Date] ), 1 )
)

// ===== Hazards: calculated columns =====
Initial Score = Hazards[Initial Likelihood] * Hazards[Initial Severity]
Residual Score = Hazards[Residual Likelihood] * Hazards[Residual Severity]
Initial Risk Category =
VAR s = Hazards[Initial Score]
RETURN MAXX ( TOPN ( 1, FILTER ( DimRiskBand, DimRiskBand[MinScore] <= s ), DimRiskBand[MinScore], DESC ), DimRiskBand[BandName] )
Residual Risk Category =
VAR s = Hazards[Residual Score]
RETURN MAXX ( TOPN ( 1, FILTER ( DimRiskBand, DimRiskBand[MinScore] <= s ), DimRiskBand[MinScore], DESC ), DimRiskBand[BandName] )
Initial Band Order = MAXX ( FILTER ( DimRiskBand, DimRiskBand[BandName] = Hazards[Initial Risk Category] ), DimRiskBand[SortOrder] )
Residual Band Order = MAXX ( FILTER ( DimRiskBand, DimRiskBand[BandName] = Hazards[Residual Risk Category] ), DimRiskBand[SortOrder] )

// ===== Actions: calculated columns (same logic as the Excel tracker) =====
Action Status =
IF (
    NOT ISBLANK ( Actions[Completion Date] ) && NOT ISBLANK ( Actions[Evidence of Closure] ) && NOT ISBLANK ( Actions[Verification Date] ),
    "Closed (Verified)",
    IF (
        NOT ISBLANK ( Actions[Completion Date] ),
        "Completed (Awaiting Verification)",
        IF (
            Actions[Target Completion Date] < MAX ( Settings[StatusDate] ),
            "Overdue",
            IF ( Actions[Recorded Status] = "In Progress", "In Progress", "Not Started" )
        )
    )
)
Days Overdue = IF ( Actions[Action Status] = "Overdue", DATEDIFF ( Actions[Target Completion Date], MAX ( Settings[StatusDate] ), DAY ), 0 )
Hazard Category = LOOKUPVALUE ( Hazards[Operational Category], Hazards[Hazard ID], Actions[Related Hazard ID] )
Hazard Initial Risk Category = LOOKUPVALUE ( Hazards[Initial Risk Category], Hazards[Hazard ID], Actions[Related Hazard ID] )

// ===== KPI measures =====
Total Hazards = DISTINCTCOUNT ( Hazards[Hazard ID] )
Extreme Hazards = CALCULATE ( [Total Hazards], Hazards[Initial Risk Category] = "Extreme" )
High Hazards = CALCULATE ( [Total Hazards], Hazards[Initial Risk Category] = "High" )
Total Actions = COUNTROWS ( Actions )
Closed Actions = CALCULATE ( COUNTROWS ( Actions ), Actions[Action Status] = "Closed (Verified)" )
Open Actions = [Total Actions] - [Closed Actions]
Overdue Actions = CALCULATE ( COUNTROWS ( Actions ), Actions[Action Status] = "Overdue" )
Awaiting Verification = CALCULATE ( COUNTROWS ( Actions ), Actions[Action Status] = "Completed (Awaiting Verification)" )
Pct Actions Closed = DIVIDE ( [Closed Actions], [Total Actions] )
Pct Open Actions Overdue = DIVIDE ( [Overdue Actions], [Open Actions] )
Near Miss Records = CALCULATE ( COUNTROWS ( Records ), Records[Record Type] = "Near miss" )
Unsafe Condition Records = CALCULATE ( COUNTROWS ( Records ), Records[Record Type] = "Unsafe condition" )
Actual Incident Records = CALCULATE ( COUNTROWS ( Records ), Records[Record Type] = "Incident (actual)" )
Sum Initial Score = SUM ( Hazards[Initial Score] )
Sum Residual Score = SUM ( Hazards[Residual Score] )
Residual High or Extreme = CALCULATE ( [Total Hazards], Hazards[Residual Risk Category] IN { "High", "Extreme" } )
Checklist Compliance % =
VAR c = CALCULATE ( COUNTROWS ( Checklist ), Checklist[Response (SIMULATED sample)] = "Compliant" )
VAR n = CALCULATE ( COUNTROWS ( Checklist ), Checklist[Response (SIMULATED sample)] = "Non-compliant" )
RETURN DIVIDE ( c, c + n )

// ===== 5 x 5 heat map (matrix visual: rows DimLikelihood[Level], columns DimSeverity[Level], values [Hazards In Cell]) =====
Hazards In Cell =
VAR l = SELECTEDVALUE ( DimLikelihood[Level] )
VAR s = SELECTEDVALUE ( DimSeverity[Level] )
RETURN CALCULATE ( COUNTROWS ( Hazards ), Hazards[Initial Likelihood] = l, Hazards[Initial Severity] = s )
Cell Score = SELECTEDVALUE ( DimLikelihood[Level] ) * SELECTEDVALUE ( DimSeverity[Level] )
Cell Colour =
VAR s = [Cell Score]
RETURN SWITCH ( TRUE (),
    s >= 17, "#C8342B",   // Extreme  (keep in step with DimRiskBand)
    s >= 10, "#F0924A",   // High
    s >= 5,  "#F2CB5C",   // Moderate
    "#7FB77E" )           // Low
// Apply: matrix > Cell elements > Background colour > Format by "Field value" > [Cell Colour]. Show zero as blank by hiding values of 0 with a conditional format on font colour.
'''
open(os.path.join(PK, "DAX_Measures.dax"), "w", encoding="utf-8").write(DAX)

THEME = {"name": "CNG Safety Risk Intelligence (simulated data)", "dataColors": ["#1F3A4D", "#5B7A99", "#A9C3DB", "#D97B29", "#8A949C", "#6FA3C7", "#4E9A5B", "#C8342B"],
         "background": "#FFFFFF", "foreground": "#222222", "tableAccent": "#1F3A4D",
         "textClasses": {"title": {"fontFace": "Segoe UI Semibold", "fontSize": 14, "color": "#1F3A4D"}, "label": {"fontFace": "Segoe UI", "fontSize": 10, "color": "#222222"}}}
json.dump(THEME, open(os.path.join(PK, "Theme_CNG_Safety.json"), "w"), indent=2)

exp = [("Total Hazards", F["n_hazards"]), ("Extreme Hazards (initial 17-25)", F["init_bands"]["Extreme"]), ("High Hazards (initial 10-16)", F["init_bands"]["High"]),
       ("Total Actions", F["n_actions"]), ("Closed Actions (verified)", F["closed"]), ("Open Actions", F["open"]), ("Overdue Actions", F["overdue"]),
       ("Awaiting Verification", F["status_counts"]["Completed (Awaiting Verification)"]), ("Pct Actions Closed", f'{F["pct_closed"]}%'),
       ("Near Miss Records", F["near_miss"]), ("Total records", F["n_records"]), ("Residual High or Extreme", F["hi_ex_res"]),
       ("Checklist Compliance %", f'{F["checklist_compliance_pct"]}%')]
SPEC = f"""# Power BI import pack: how to build the report

**Status: import-ready, not a .pbix.** Power BI Desktop is not available in the environment that produced this project, so no .pbix file was created and none is claimed. The pack below lets an analyst with Power BI Desktop build the report in about 30 minutes. The fully working dashboard is the **Dashboard** sheet in `Data_and_Excel/CNG_Safety_Risk_Assessment_Workbook.xlsx`.

All data is **SIMULATED** (independent portfolio case study; no Powergas records were used).

## Contents
| File | Purpose |
|---|---|
| `data/Hazards.csv`, `Actions.csv`, `Records.csv`, `Checklist.csv` | Fact tables exported from the workbook (same record IDs). |
| `data/DimRiskBand.csv`, `DimLikelihood.csv`, `DimSeverity.csv`, `Settings.csv` | Scoring scales, band thresholds and the reporting status date ({C.STATUS_DATE:%d %B %Y}). |
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
""" + "\n".join(f"| {k} | {v} |" for k, v in exp) + """

## Using real data later
Replace the CSVs with exports from your register, keep the column headers, and set `Settings[StatusDate]` to the review date. Do not mix simulated and real records in one file.
"""
open(os.path.join(PK, "PowerBI_Build_Guide.md"), "w", encoding="utf-8").write(SPEC)
print("power bi pack:", sorted(os.listdir(PK)), sorted(os.listdir(DATA)))
