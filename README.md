# CNG Safety Risk Intelligence and Preventive Action System

A safety risk management tool for a CNG virtual-pipeline distributor, built as an independent portfolio case study around Powergas operations.

It ranks operational hazards on a 5 x 5 matrix, tracks corrective actions with calculated status and overdue flags, holds an inspection checklist and an incident and near-miss log, and shows the results on a filterable Excel dashboard.

**All data is simulated.** No Powergas records, people, sites or findings were used, and nothing here describes Powergas's real safety performance or claims regulatory compliance.

## Start here
- `Data_and_Excel/CNG_Safety_Risk_Assessment_Workbook.xlsx`: the working tool (open the Dashboard sheet).
- `Final_HSE_Report.pdf`: method, findings and recommendations.
- `Reports/Executive_Summary.pdf`: one page.
- `Presentation/`: 10-slide deck.
- `Dashboard/PowerBI_Import_Pack/`: data, Power Query and DAX for building the dashboard in Power BI (no .pbix included).
- `Research_and_References/Reference_Register.md`: sources and what was verified.
- `Build_Scripts/`: Python that regenerates everything (run `build_data.py`, then the other `build_*.py` scripts).
