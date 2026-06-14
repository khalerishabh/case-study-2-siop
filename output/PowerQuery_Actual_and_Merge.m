// ============================================================
// QUERY 2:  Actual_Long
// Unpivots the 3 actual measures so they align with Forecast_Long.
// ============================================================
let
    Source = Excel.Workbook(File.Contents("C:\Users\khale\Downloads\Case Study 2  SIOP Reporting.xlsb"), null, true),
    Actual = Source{[Item="Actual",Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Actual, [PromoteAllScalars=true]),
    Renamed = Table.RenameColumns(Promoted,{
        {"Actuals___Revenue_Qty","Revenue"},
        {"Actuals___NonRevenue_Qty","NonRevenue"},
        {"Actuals___Capex_Qty","Capex"}}),
    Typed = Table.TransformColumnTypes(Renamed,{
        {"Item", type text}, {"Country", type text}, {"Period", type text},
        {"Revenue", type number}, {"NonRevenue", type number}, {"Capex", type number}}),

    // *** CRITICAL DATA-QUALITY STEP ***
    // The Actual sheet contains ~442k EXACT duplicate rows (88% of 500k). Summing them
    // blindly inflates actuals ~9x and destroys the KPIs. Remove exact duplicates first.
    Distinct = Table.Distinct(Typed),

    Unpiv = Table.UnpivotOtherColumns(Distinct, {"Item","Country","Period"}, "MeasureType", "ActualQty"),
    // After de-duplication, ~7% of keys still legitimately have multiple distinct rows,
    // so we sum within key to get one actual per Item/Country/Period/MeasureType.
    Grouped = Table.Group(Unpiv, {"Item","Country","Period","MeasureType"},
                 {{"ActualQty", each List.Sum([ActualQty]), type number}})
in
    Grouped

// ============================================================
// QUERY 3:  Fact_ForecastActual   (the merge the brief asks for)
// Merge Actuals onto Forecast on Period + Item + Country (+ MeasureType).
// Build this by: reference Forecast_Long > Merge Queries > pick Actual_Long
//   match on Item, Country, Period, MeasureType > Left Outer > expand ActualQty.
// Equivalent M:
// ============================================================
let
    Source = Forecast_Long,
    Merged = Table.NestedJoin(Source, {"Item","Country","Period","MeasureType"},
                Actual_Long, {"Item","Country","Period","MeasureType"}, "act", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Merged, "act", {"ActualQty"}, {"ActualQty"}),
    FillZero = Table.ReplaceValue(Expanded, null, 0, Replacer.ReplaceValue, {"ActualQty"}),
    // Flag rows where an actual is realized -> use this to scope KPI visuals fairly
    AddRealized = Table.AddColumn(FillZero, "ActualRealized", each [ActualQty] <> 0, type logical)
in
    AddRealized
