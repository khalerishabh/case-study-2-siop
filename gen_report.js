const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType, ShadingType,
  TableOfContents, PageNumber, Header, Footer, PageBreak, ImageRun, VerticalAlign } = require("docx");

function placeholder(title, instruction, heightDxa = 2700) {
  const d = { style: BorderStyle.DASHED, size: 6, color: "9AA5C0" };
  return new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
    rows: [new TableRow({ height: { value: heightDxa, rule: "atLeast" }, children: [
      new TableCell({
        width: { size: 9360, type: WidthType.DXA },
        borders: { top: d, bottom: d, left: d, right: d },
        shading: { fill: "F0F2F8", type: ShadingType.CLEAR },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
            children: [new TextRun({ text: title, bold: true, color: "1E2761", size: 22 })] }),
          new Paragraph({ alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: instruction, italics: true, color: GREY, size: 18 })] }),
        ],
      }),
    ] })],
  });
}
const phCaption = (t) => new Paragraph({ spacing: { before: 20, after: 200 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: t, italics: true, size: 17, color: GREY })] });

function fig(path, caption, w = 460) {
  const h = Math.round(w * 510 / 900);
  return [
    new Paragraph({ spacing: { before: 100, after: 20 }, alignment: AlignmentType.CENTER,
      children: [new ImageRun({ type: "png", data: fs.readFileSync(path),
        transformation: { width: w, height: h },
        altText: { title: caption, description: caption, name: caption } })] }),
    new Paragraph({ spacing: { after: 160 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: caption, italics: true, size: 17, color: GREY })] }),
  ];
}

const NAVY = "1E2761", ICE = "CADCFC", GREY = "59607A", LINE = "BBBBBB";
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const P = (t, opts = {}) => new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: t, ...opts })] });
const bullet = (t, bold) => new Paragraph({ numbering: { reference: "b", level: 0 }, spacing: { after: 60 },
  children: Array.isArray(t) ? t : [new TextRun({ text: t, bold: !!bold })] });
const num = (t) => new Paragraph({ numbering: { reference: "n", level: 0 }, spacing: { after: 60 }, children: [new TextRun(t)] });

function table(headers, rows, widths) {
  const tw = widths.reduce((a, b) => a + b, 0);
  const hdr = new TableRow({ tableHeader: true, children: headers.map((h, i) =>
    new TableCell({ borders, width: { size: widths[i], type: WidthType.DXA },
      shading: { fill: NAVY, type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, color: "FFFFFF", size: 19 })] })] })) });
  const body = rows.map((r, ri) => new TableRow({ children: r.map((c, i) =>
    new TableCell({ borders, width: { size: widths[i], type: WidthType.DXA },
      shading: { fill: ri % 2 ? "F2F5FB" : "FFFFFF", type: ShadingType.CLEAR },
      margins: { top: 50, bottom: 50, left: 100, right: 100 },
      children: [new Paragraph({ children: [new TextRun({ text: String(c), size: 19 })] })] })) }));
  return new Table({ width: { size: tw, type: WidthType.DXA }, columnWidths: widths, rows: [hdr, ...body] });
}

const doc = new Document({
  creator: "Candidate",
  styles: {
    default: { document: { run: { font: "Calibri", size: 22, color: "222222" } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: "Calibri", color: NAVY },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 25, bold: true, font: "Calibri", color: "2E3A6B" },
        paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [
    { reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 280 } } } }] },
    { reference: "n", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 280 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Case Study 2 – SIOP Reporting   |   Page ", size: 16, color: GREY }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: GREY })] })] }) },
    children: [
      // ---- Title block ----
      new Paragraph({ spacing: { before: 1400, after: 0 }, children: [new TextRun({ text: "Case Study 2", size: 30, color: GREY })] }),
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "SIOP Forecast Reporting", bold: true, size: 60, color: NAVY })] }),
      new Paragraph({ spacing: { after: 40 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 6 } },
        children: [new TextRun({ text: "Data Transformation, KPI Modelling & Dashboard Approach", size: 26, color: "2E3A6B" })] }),
      new Paragraph({ spacing: { before: 200 }, children: [new TextRun({ text: "Internship Program Assessment – June 2026", size: 22, color: GREY })] }),
      new Paragraph({ spacing: { after: 400 }, children: [new TextRun({ text: "Deliverables: Power BI report (.pbix) + this approach document", size: 20, italics: true, color: GREY })] }),

      H1("Table of Contents"),
      new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
      new Paragraph({ children: [new PageBreak()] }),

      // ---- Executive Summary ----
      H1("Executive Summary"),
      P("This document describes my understanding of the SIOP (Sales, Inventory & Operations Planning) forecasting case study and the approach taken to deliver a robust, scalable Power BI solution. SIOP forecasts are stored as monthly snapshots; the same future month is forecast repeatedly from different vantage points in time. The core task is to restructure these snapshots so that successive forecast vintages can be compared side by side, merge them with realised actuals, and surface forecast-quality KPIs (Accuracy and Bias) in an executive dashboard."),
      P("Two findings drove the design. First, the snapshot-to-lag mapping was implemented with a parameter-free ranking rule so the model scales automatically as new snapshots arrive. Second, a significant data-quality issue was discovered in the Actuals data – roughly 88% of its rows are exact duplicates – which, if summed naively, inflates actual quantities by ~9x and renders every KPI meaningless. Handling this correctly is central to a trustworthy solution."),

      // ---- 1. Problem understanding ----
      H1("1. Problem Understanding"),
      P("In SIOP, a forecast is regenerated every period and saved as a snapshot identified by its “As-Of Period” (the month the forecast was created). Each snapshot contains forecasts for many future periods. Because the same target month is forecast from multiple snapshots, we can study how the forecast for a given month evolved as it was revised over time – this is forecast-vintage (lag) analysis."),
      P("The brief defines the labelling relative to the latest snapshot:", { bold: false }),
      bullet([new TextRun({ text: "Resultant ", bold: true }), new TextRun("– the latest snapshot (here 2023 P12).")]),
      bullet([new TextRun({ text: "Lag 0 ", bold: true }), new TextRun("– the previous snapshot (2023 P11).")]),
      bullet([new TextRun({ text: "Lag 1 ", bold: true }), new TextRun("– the snapshot before that (2023 P10).")]),
      P("For any fixed target Period, Resultant / Lag 0 / Lag 1 are therefore forecasts of the same month made at progressively earlier times (i.e. at longer horizons). Aligning them on the target Period and comparing against actuals tells us whether forecasts improve as the target month approaches.", { }),

      // ---- 2. Dataset overview ----
      H1("2. Dataset Overview"),
      P("The source workbook contains two sheets:"),
      table(["Sheet", "Rows", "Grain", "Key columns"],
        [["Forecast", "246,477", "Item × Country × AsOfPeriod × Period",
          "AsOfPeriod (snapshot), Period (target), Revenue / Non-Revenue / CAPEX Forecast"],
         ["Actual", "500,000", "Item × Country × Period (heavily duplicated)",
          "Actuals Revenue / Non-Revenue / Capex Qty"]],
        [1400, 1100, 3260, 3600]),
      P("There are three snapshots (2023 P10, P11, P12); target periods span 2023 P10 to 2024 P12; four countries (A–D). The forecast measure column “CAPEX Foreacast” is mis-spelled at source and is corrected during transformation.", { }),

      // ---- 3. Approach ----
      H1("3. Solution Approach & Architecture"),
      P("A key design decision was how to represent the lags. Two options were considered:"),
      table(["Option", "Description", "Trade-off"],
        [["A – Wide pivot", "Hard-pivot to 9 columns (Resultant/Lag0/Lag1 × 3 measures).",
          "Matches the brief wording literally, but rigid and makes the Lag slicer awkward."],
         ["B – Normalised (chosen)", "Keep one Lag column; produce side-by-side columns via a Matrix visual.",
          "Single clean fact, native slicing, single-definition KPIs, scales to N snapshots."]],
        [1900, 3760, 3700]),
      P("I chose Option B as the engine and reproduce the literal side-by-side columns through a Matrix visual (Lag on columns). This satisfies the stated requirement while keeping the model scalable and the DAX simple."),

      // ---- 4. Transformation ----
      H1("4. Data Transformation (Power Query)"),
      P("All transformation is done in Power Query (M) so it is repeatable and refreshable. Steps on the Forecast data:"),
      num("Promote headers, set types, and rename the three measures (fixing the source typo)."),
      num("Derive a scalable lag label: rank distinct AsOfPeriods in descending order (1 = latest). Rank 1 → “Resultant”; rank n>1 → “Lag” & (n−2). This yields Lag0, Lag1, Lag2… automatically with no hard-coded snapshot names."),
      num("Unpivot the three measure columns into MeasureType + ForecastQty, producing a long fact (739,431 rows)."),
      P("Because the lag label is derived from a ranking rather than fixed values, dropping in a future snapshot (e.g. 2024 P01) automatically re-labels it as Resultant and shifts the others – no formula changes required.", { italics: true }),

      // ---- 5. Merge ----
      H1("5. Merging Actuals (the critical data-quality step)"),
      P("Actuals are unpivoted to MeasureType + ActualQty and merged onto the forecast fact on Item + Country + Period + MeasureType (a left join, keeping all forecast rows)."),
      P("Investigating the Actuals revealed a serious quality issue:", { bold: true }),
      table(["Check", "Result"],
        [["Total rows", "500,000"],
         ["Exact-duplicate rows", "442,054 (88.4%)"],
         ["Distinct rows", "57,946"],
         ["Keys with genuinely differing values", "3,919 (7.4%)"]],
        [5800, 3560]),
      P("Summing the raw rows inflates actuals by roughly 9x and produces nonsensical KPIs (e.g. forecast appearing to be ~5% of actual). The robust rule applied is: remove exact duplicate rows first (Table.Distinct), then sum within each key to respect the small share of keys that legitimately carry multiple distinct values. This correction alone moved Revenue accuracy from ~5% to ~20–30%."),

      // ---- 6. Model ----
      H1("6. Data Model"),
      bullet("Fact_ForecastActual – the merged long fact (Item, Country, Period, Lag, MeasureType, ForecastQty, ActualQty, ActualRealized)."),
      bullet("Lag Selection – a small disconnected table (Resultant / Lag0 / Lag1, with a sort key) that drives the interactive lag filter via a SWITCH-style measure."),
      bullet("Staging queries (Forecast_Long, Actual_Long) have load disabled to keep the model lean."),
      placeholder("[ Screenshot 1 — Data Model view ]", "In Power BI: Model view → capture the tables and relationships. Paste here.", 2200),
      phCaption("Screenshot 1 — Star schema: Fact_ForecastActual with the disconnected Lag Selection table."),

      // ---- 7. KPIs ----
      H1("7. KPI Definitions & Assumptions"),
      P("KPIs follow the brief, implemented with DIVIDE to guard against division by zero:"),
      bullet([new TextRun({ text: "Forecast Accuracy % ", bold: true }), new TextRun("= 1 − ABS(Lag Qty − Actual Qty) / Actual Qty")]),
      bullet([new TextRun({ text: "Forecast Bias % ", bold: true }), new TextRun("= (Lag Qty − Actual Qty) / Actual Qty")]),
      P("Assumptions documented for the assessor:"),
      bullet("“Sales Actuals” in the Bias formula is interpreted as the actual quantity of the same measure type as the forecast (Revenue↔Revenue, etc.)."),
      bullet("Where actuals are zero or absent (e.g. CAPEX actuals are entirely zero in this sample, and future 2024 periods have no actual yet), the KPI is left blank rather than forced – DIVIDE handles this."),
      bullet("Accuracy/Bias visuals are scoped to rows where an actual is realised (ActualRealized = TRUE) so unrealised future forecasts do not distort the percentages."),
      P("The Lag Selection filter is implemented so that choosing “Lag 0” returns the sum of the Lag 0 columns across Revenue, Non-Revenue and CAPEX, as required."),

      // ---- 8. Dashboard ----
      H1("8. Dashboard Design"),
      P("Three pages, all built manually in Power BI Desktop:"),
      bullet([new TextRun({ text: "Executive Forecast Overview ", bold: true }), new TextRun("– Revenue, Non-Revenue and CAPEX forecast by Country; KPI cards; a Lag Selection slicer plus Period / Country / Measure slicers that recalculate on selection.")]),
      bullet([new TextRun({ text: "Lag Comparison ", bold: true }), new TextRun("– a Matrix with Lag on columns, giving the literal Resultant / Lag 0 / Lag 1 side-by-side view against actuals.")]),
      bullet([new TextRun({ text: "Forecast Accuracy & Bias ", bold: true }), new TextRun("– accuracy plotted against the true forecast horizon (PeriodOffset, 0–14 months) and bias by country (proposed analysis), using the realised-only measures.")]),
      P("A variance measure (Forecast vs Actual = Forecast Qty − Actual Qty) was added to the Lag Comparison matrix so each vintage can be read against the realised actual in the same view.", { italics: true }),
      placeholder("[ Screenshot 2 — Page 1: Executive Forecast Overview ]", "Capture the full Page 1 (Lag slicer, KPI cards, the three by-Country charts). Paste here."),
      phCaption("Screenshot 2 — Executive Forecast Overview: Revenue / Non-Revenue / CAPEX by Country with the Lag-selection KPIs."),
      placeholder("[ Screenshot 3 — Page 2: Lag Comparison matrix ]", "Set the MeasureType slicer to Revenue, then capture the full Page 2 matrix. Paste here."),
      phCaption("Screenshot 3 — Lag Comparison matrix: Resultant / Lag 0 / Lag 1 side-by-side with actuals and variance."),
      placeholder("[ Screenshot 4 — Page 3: Forecast Accuracy & Bias ]", "Capture the full Page 3 (horizon line, accuracy-by-lag, bias-by-country, KPI cards). Paste here."),
      phCaption("Screenshot 4 — Accuracy & Bias page: accuracy vs horizon and bias by country."),

      // ---- 9. Findings ----
      H1("9. Key Findings & Insights"),
      P("Computed on the corrected data (exact duplicates removed, realised periods only):"),
      table(["Lag", "Measure", "Forecast Qty", "Actual Qty", "Accuracy %", "Bias %"],
        [["Resultant", "Revenue", "1,276,145", "6,289,571", "20.3%", "-79.7%"],
         ["Lag 0", "Revenue", "1,801,029", "6,102,108", "29.5%", "-70.5%"],
         ["Lag 1", "Revenue", "1,203,338", "6,176,877", "19.5%", "-80.5%"]],
        [1500, 1700, 1900, 1900, 1180, 1180]),
      bullet("Forecasts systematically under-shoot actuals (large negative bias) at every horizon, pointing to a persistent under-forecasting tendency that the business should investigate."),
      bullet("Accuracy decays with horizon – near-term Revenue forecasts (~27% at 0 months out) outperform mid-horizon ones (~16% around 8–9 months out) – which is exactly why PeriodOffset was retained."),
      bullet("Non-Revenue and CAPEX actuals are near-zero / zero in this sample, so Revenue is the only measure with a meaningful accuracy signal – flagged rather than hidden."),
      bullet("These figures are sensitive to the duplicate-handling rule; the headline message of the case is that data hygiene must precede KPI calculation."),
      ...fig("output/fig_horizon.png", "Figure 1 — Accuracy generally weakens as the forecast horizon lengthens (Revenue, realised periods)."),
      ...fig("output/fig_bias.png", "Figure 2 — Forecast bias by Country: A and C under-forecast, B and D over-forecast (Revenue, realised)."),
      P("Note: Figures 1–2 are charted directly from the model data to illustrate the findings; Screenshots 1–4 above show the live Power BI dashboard.", { italics: true }),

      // ---- 10. Scalability ----
      H1("10. Scalability & Robustness"),
      bullet("Lag labelling is rank-derived, so new snapshots are absorbed without edits."),
      bullet("All logic lives in Power Query (M) and DAX measures – fully refreshable, no manual restructuring."),
      bullet("Exact-duplicate removal and DIVIDE-guarded KPIs make the model resilient to the data issues present in the source."),
      bullet("Staging queries disabled from load; long-and-narrow fact keeps the model compact despite ~0.7M rows."),

      // ---- 11. Assumptions ----
      H1("11. Assumptions & Limitations"),
      bullet("Period text (“YYYY Pnn”) sorts chronologically; a dedicated calendar dimension can be added for richer time intelligence."),
      bullet("PeriodOffset (horizon) and Set_vs_FG are retained and used; the five anonymised Segmentation_ codes are kept as optional drill dimensions but Country remains the main analysis axis per the brief. Lifecycle (~95% blank) is excluded."),
      bullet("The Bias “Sales Actuals” interpretation (same-measure actual) is stated above and can be re-pointed to Revenue actuals if the business intends otherwise."),
      bullet("The dashboard was authored manually in Power BI Desktop; no generative-AI report-building feature was used, per the assessment rules."),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync("output/CS2_Report.docx", buf); console.log("wrote output/CS2_Report.docx"); });
