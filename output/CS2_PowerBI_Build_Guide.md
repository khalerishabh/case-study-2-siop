# Case Study 2 — SIOP Reporting: Power BI Build Guide (click-by-click)

> Build the `.pbix` **entirely by hand** in Power BI Desktop. **Do not use Copilot or any
> "auto-generate report/visual" AI feature** — those are what the brief rejects. Typing the
> provided M/DAX yourself is normal Power BI development and is fine.

---

## Phase 0 — Setup
1. Install **Power BI Desktop** (free, Microsoft Store) and open it.
2. `File > Options and settings > Options > Preview features` → ensure **Copilot / AI report creation is OFF**.
3. Keep these files open next to you:
   - `PowerQuery_Forecast.m`, `PowerQuery_Actual_and_Merge.m`, `DAX_Measures.txt`
   - `CS2_GroundTruth_Reference.xlsx` (for verifying numbers at the end).

## Phase 1 — Get Data
1. `Home > Get Data > More… > Excel workbook` → choose the `.xlsb`
   (if it's hidden, set the file dialog filter to **All files**).
2. In the Navigator, tick **Forecast** and **Actual** → click **Transform Data** (NOT Load).

## Phase 2 — Build `Forecast_Long`
1. In Power Query, click the **Forecast** query → rename it to **`Forecast_Long`**.
2. `Home > Advanced Editor` → delete everything → paste the contents of **`PowerQuery_Forecast.m`**.
3. **Fix the file path** in the `Source` line to your local `.xlsb`.
4. Confirm the output columns are: `Item, Country, Period, Lag, AsOfPeriod, PeriodOffset, Set_vs_FG, Segmentation_* (5), MeasureType, ForecastQty`.
   - Sanity check the **Lag** column shows `Resultant / Lag0 / Lag1`.
   - `PeriodOffset` (0–14) = forecast horizon in months → used on Page 3. `Set_vs_FG` and the
     `Segmentation_*` codes are extra drill/slicer dimensions. (Lifecycle is ~95% blank, so it's
     intentionally excluded — add it back in the M if you want it.)

## Phase 3 — Build `Actual_Long`
1. Click the **Actual** query → rename to **`Actual_Long`**.
2. `Advanced Editor` → paste **Query 2** from `PowerQuery_Actual_and_Merge.m` (the `Actual_Long` block) and fix the path.
3. ⚠️ **Do not skip the `Table.Distinct` step** in that query. The Actual sheet is ~88%
   exact-duplicate rows (442k of 500k). Removing them is what makes the KPIs correct; the
   `Remove Duplicates` happens **before** the group-sum.
4. Output columns: `Item, Country, Period, MeasureType, ActualQty`.

## Phase 4 — Merge into `Fact_ForecastActual`
1. Right-click **`Forecast_Long`** → **Reference**. Rename the new query to **`Fact_ForecastActual`**.
2. `Home > Merge Queries`:
   - Top table = `Fact_ForecastActual`; bottom = `Actual_Long`.
   - **Ctrl-click in the same order** to select the match keys on both: `Item`, `Country`, `Period`, `MeasureType`.
   - **Join Kind = Left Outer** → OK.
3. Click the **expand** ⧉ icon on the new `Actual_Long` column → tick **ActualQty** only, **untick** "Use original column name as prefix".
4. Select `ActualQty` → `Transform > Replace Values` → replace **null** with **0**.
5. `Add Column > Custom Column` → name **`ActualRealized`**, formula `= [ActualQty] <> 0`.
   *(Or just paste the Query 3 block from the file, which does steps 2–5 at once.)*
6. **Reduce model size:** right-click `Forecast_Long` and `Actual_Long` → **untick "Enable load"** (they're only staging).
7. `Home > Close & Apply`.

## Phase 5 — Supporting table for the Lag filter
1. `Home > Enter Data`. Create table **`Lag Selection`** with:

   | Lag | SortOrder |
   |-----|-----------|
   | Resultant | 0 |
   | Lag0 | 1 |
   | Lag1 | 2 |
2. Load it. In **Data view**, select column `Lag` → `Column tools > Sort by column > SortOrder`.
3. Leave it **disconnected** (no relationship) — the SWITCH measure reads it.

## Phase 6 — Model
1. `Model view`: confirm `Lag Selection` has **no relationship** to the fact.
2. (Optional but tidy) create an empty **`_Measures`** table (`Enter Data`, load empty) to hold all measures.

## Phase 7 — Create measures
1. Open `DAX_Measures.txt`. For each measure: select `_Measures` (or the fact) → `New Measure` → paste → Enter.
2. Format **Accuracy %** and **Bias %** as **Percentage** (Measure tools > Format > %).

## Phase 8 — Page 1: "Executive Forecast Overview"
- **Three clustered column charts** (one each Revenue / Non-Revenue / Capex):
  - X-axis = `Country`, Values = `Selected Lag Forecast`.
  - Visual-level filter: `MeasureType = Revenue` (then NonRevenue, then Capex). Title each chart.
- **KPI cards**: `Selected Lag Forecast`, `Forecast Accuracy %`, `Forecast Bias %`.
- **Slicers**: `Lag Selection[Lag]` (set to **Single select** in Format), `Period`, `Country`, `MeasureType`.
- Test: changing the Lag slicer to **Lag0** should make the cards/charts show Lag0 sums.

## Phase 9 — Page 2: "Lag Comparison (side-by-side)"
- **Matrix** visual → Rows: `Country` › `Item` › `Period`; Columns: `Lag`; Values: `Forecast Qty`.
  - This produces the literal **Resultant / Lag0 / Lag1** columns the brief asks for.
- Add `Actual Qty` as a value, and a variance measure if you like.
- Confirm column order is Resultant → Lag0 → Lag1 (driven by the sort key).

## Phase 10 — Page 3: "Forecast Accuracy & Bias" (your proposed analysis)
- Use the **(Realized)** measures here so empty future periods don't distort %.
- **Line chart (the headline insight)**: `Accuracy % (Realized)` on Y, **`PeriodOffset` (0–14)** on X
  → shows accuracy decaying as the forecast horizon lengthens (≈27% near-term → ≈16% mid-horizon).
  Filter to `MeasureType = Revenue` (the only measure with meaningful actuals).
- **Column chart**: `Accuracy % (Realized)` by `Lag`, legend = `MeasureType`.
- **Bar chart**: `Forecast Bias %` by `Country` → over- vs under-forecasting (bias is ~−75% throughout).
- Optional: `Set_vs_FG` as a slicer; scatter of Bias vs forecast volume by Item.

## Phase 11 — Interactivity & polish
1. `View > Sync slicers` → sync `Period` and `Country` across all pages.
2. Apply a clean theme; add page titles; align visuals; concise tooltips.

## Phase 12 — Validate (do not skip)
1. Open `CS2_GroundTruth_Reference.xlsx`.
2. On Page 2 matrix, filter to a Country and compare **Forecast Qty by Lag** to the
   `Forecast_by_Country_Lag` sheet. They should match.
3. Spot-check Accuracy/Bias for a Country×Lag against `KPI_Country_Lag`.

## Phase 13 — Save & submit
- Save the `.pbix`. Submit it with the written report.

---

### Common gotchas
- **xlsb not visible** in Get Data → set dialog filter to *All files* or use the Excel connector.
- **Item key mismatches** — Items are mixed numeric/text in the source; the M forces `Item` to **text** on both sides so the merge keys line up. Don't change that.
- **Accuracy shows blank** for Capex — expected: Capex actuals are 0 in this dataset (`DIVIDE` guards it).
- **% too low when "all periods" selected** — restrict KPI visuals to `ActualRealized = TRUE` (Page 3 uses the Realized measures for this reason).
