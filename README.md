# Case Study 2 — SIOP Forecast Reporting (Power BI)

Internship program assessment. The solution transforms SIOP forecast **snapshots** into a
side-by-side **lag comparison** (Resultant / Lag 0 / Lag 1), merges them with **actuals**, and
surfaces **Forecast Accuracy** and **Forecast Bias** KPIs in an executive Power BI dashboard.

## Highlights
- **Scalable lag mapping** — snapshots are ranked by `AsOfPeriod`, so new snapshots are absorbed with no formula changes.
- **Data-quality fix** — the Actuals source was ~88% exact-duplicate rows (442k of 500k); these are removed before aggregation, which is what makes the KPIs trustworthy.
- **KPIs** — Forecast Accuracy % = 1 − ABS(Lag − Actual)/Actual; Forecast Bias % = (Lag − Actual)/Actual, both zero-guarded.
- **Finding** — forecasts systematically under-shoot actuals, and accuracy decays as the forecast horizon lengthens.

## Contents
| Path | Description |
|------|-------------|
| `CS2_SIOP_Reporting.pbix` | Power BI dashboard (Executive Overview / Lag Comparison / Accuracy & Bias) |
| `output/CS2_Report.pdf` / `.docx` | Approach & findings report |
| `output/PowerQuery_Forecast.m`, `output/PowerQuery_Actual_and_Merge.m` | Power Query (M) transformations |
| `output/DAX_Measures.txt` | DAX measures |
| `output/SIOP_Theme.json` | Power BI theme |
| `output/CS2_GroundTruth_Reference.xlsx` | Independent validation of the KPI figures |
| `data/Forecast.csv`, `data/Actual.csv` | Source data (enables dashboard refresh) |

## How to open
Open `CS2_SIOP_Reporting.pbix` in Power BI Desktop. To refresh, keep the `data/` CSVs at the
path referenced in Power Query (or repoint the source).
