// ============================================================
// QUERY 1:  Forecast_Long
// Paste into:  Power Query > New Source > Blank Query > Advanced Editor
// Replace the file path in Source with your local .xlsb path.
// What it does: ranks snapshots (scalable), maps to Resultant/Lag0/Lag1,
//               then unpivots the 3 forecast measures into a long fact.
// ============================================================
let
    Source = Excel.Workbook(File.Contents("C:\Users\khale\Downloads\Case Study 2  SIOP Reporting.xlsb"), null, true),
    Forecast = Source{[Item="Forecast",Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Forecast, [PromoteAllScalars=true]),

    // --- standardize types & fix source typos ('Foreacast', spaces) ---
    Renamed = Table.RenameColumns(Promoted,{
        {"Non Revenue Forecast","NonRevenue"},
        {"CAPEX Foreacast","Capex"},
        {"Revenue Forecast","Revenue"}}),
    Typed = Table.TransformColumnTypes(Renamed,{
        {"Item", type text}, {"Country", type text},
        {"AsOfPeriod", type text}, {"Period", type text},
        {"NonRevenue", type number}, {"Capex", type number}, {"Revenue", type number}}),

    // --- SCALABLE lag mapping: rank AsOfPeriod descending (1 = latest) ---
    // 'YYYY Pnn' text sorts chronologically, so we just sort the distinct list desc.
    Snapshots = List.Sort(List.Distinct(Typed[AsOfPeriod]), Order.Descending),
    AddRank = Table.AddColumn(Typed, "SnapshotRank",
                 each List.PositionOf(Snapshots, [AsOfPeriod]) + 1, Int64.Type),
    AddLag = Table.AddColumn(AddRank, "Lag",
                 each if [SnapshotRank] = 1 then "Resultant"
                      else "Lag" & Text.From([SnapshotRank] - 2), type text),
    // (optional) keep only the 3 most recent snapshots:
    // AddLag = Table.SelectRows(AddLag, each [SnapshotRank] <= 3),

    // --- keep analytically useful attributes (don't drop the rest of the dataset) ---
    //   AsOfPeriod   : traceability to the source snapshot
    //   PeriodOffset : true forecast horizon in months (0..14) -> accuracy-by-horizon analysis
    //   Set_vs_FG    : fully-populated FG/Set breakdown
    //   Segmentation_* : optional drill hierarchy (anonymised codes). Drop any you don't want.
    // (Lifecycle is only ~5% populated, so it is intentionally left out — add it back if needed.)
    Keep = Table.SelectColumns(AddLag, {
        "Item","Country","Period","Lag","AsOfPeriod","PeriodOffset","Set_vs_FG",
        "Segmentation_Region","Segmentation_Cluster","Segmentation_CountryGroup",
        "Segmentation_Country","Segmentation_Global",
        "Revenue","NonRevenue","Capex"}),

    // --- unpivot ONLY the 3 measures into MeasureType / ForecastQty (everything else is an id) ---
    Unpiv = Table.UnpivotOtherColumns(Keep,
        {"Item","Country","Period","Lag","AsOfPeriod","PeriodOffset","Set_vs_FG",
         "Segmentation_Region","Segmentation_Cluster","Segmentation_CountryGroup",
         "Segmentation_Country","Segmentation_Global"},
        "MeasureType", "ForecastQty"),
    Result = Table.TransformColumnTypes(Unpiv,{{"ForecastQty", type number},{"PeriodOffset", Int64.Type}})
in
    Result
