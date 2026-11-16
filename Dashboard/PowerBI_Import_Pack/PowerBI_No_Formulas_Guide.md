# Power BI without writing any formulas

All data is SIMULATED. `PowerBI_Ready.xlsx` already contains every score, risk category and action status, plus 1/0 flag columns, so cards need no DAX.

## Load
Get data > Excel workbook > `PowerBI_Ready.xlsx` > tick Hazards, Actions, Records, Checklist > Load. No relationships are needed. Each sheet already holds the fields it needs.

## Cards (drag the field into a Card, set it to Sum)
| Card | Table and field | Expected |
|---|---|---|
| Total Hazards | Hazards, Hazard ID (Count distinct) | 40 |
| Extreme Hazards | Hazards, Is Extreme (Sum) | 2 |
| High Hazards | Hazards, Is High (Sum) | 14 |
| Total Actions | Actions, Action ID (Count distinct) | 54 |
| Open Actions | Actions, Open Flag (Sum) | 31 |
| Overdue Actions | Actions, Is Overdue (Sum) | 19 |
| Closed Actions | Actions, Is Closed (Sum) | 23 |
| Near Miss Records | Records, Is Near Miss (Sum) | 16 |

## Charts
- Initial versus residual risk: clustered column, axis `Initial Risk Category`, value count of `Hazard ID`. Sort by `Initial Risk Order`.
- Hazards by category: stacked bar, axis `Operational Category`, legend `Initial Risk Category`, value count of `Hazard ID`.
- Heat map: Matrix, rows `Initial Likelihood`, columns `Initial Severity`, values count of `Hazard ID`.
- Actions by status: bar, axis `Action Status`, value count of `Action ID`.
- Overdue by role: bar, axis `Responsible Role`, value sum of `Is Overdue`.
- Slicers: `Operational Category`, `Initial Risk Category`, `Action Status`.
