import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "/Users/benjohnson/dev/science/qrahss";
const outputPath = `${root}/outputs/qr-funding-series/qr_funding_by_discipline_2010-11_to_2024-25.xlsx`;

function parseCsv(text) {
  const rows = [];
  let row = [], field = "", quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (ch === '"') quoted = false;
      else field += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === ',') { row.push(field); field = ""; }
    else if (ch === '\n') { row.push(field.replace(/\r$/, "")); rows.push(row); row = []; field = ""; }
    else field += ch;
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  const headers = rows.shift();
  return rows.filter(r => r.some(v => v !== "")).map(r => Object.fromEntries(headers.map((h, i) => [h, r[i] ?? ""])));
}

const components = parseCsv(await fs.readFile(`${root}/tmp/qr_series/qr_components.csv`, "utf8"));
const series = parseCsv(await fs.readFile(`${root}/tmp/qr_series/qr_share_series.csv`, "utf8"));
const wb = Workbook.create();
const font = "Arial";
const navy = "#17324D", blue = "#246B8E", teal = "#2A9D8F", gold = "#E9C46A";
const lightBlue = "#EAF2F7", lightGold = "#FFF5D6", grey = "#5B6573", pale = "#F5F7F9";

function title(sheet, range, text) {
  sheet.getRange(range).values = [[text]];
  sheet.getRange(range).format.font = { name: font, size: 16, bold: true, color: navy };
}
function section(sheet, range, text) {
  sheet.getRange(range.split(":")[0]).values = [[text]];
  sheet.getRange(range).format = { fill: lightBlue, font: { name: font, size: 10, bold: true, color: navy },
    borders: { preset: "outside", style: "thin", color: "#B8C7D1" } };
}
function header(sheet, range) {
  sheet.getRange(range).format = { fill: navy, font: { name: font, size: 10, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center", verticalAlignment: "center",
    borders: { preset: "inside", style: "thin", color: "#FFFFFF" } };
}
function body(sheet, range) {
  sheet.getRange(range).format.font = { name: font, size: 10, color: "#1F2933" };
  sheet.getRange(range).format.verticalAlignment = "center";
}

// Components: source-like inputs with transparent formula totals.
const comp = wb.worksheets.add("Components");
comp.showGridlines = false;
title(comp, "A2", "QR and identifiable research funding components");
comp.getRange("A3:N3").values = [["Academic year", "Broad discipline", "Mainstream QR", "London QR", "RDP QR",
  "Charity QR", "Business QR", "Total QR", "Research Councils", "Charities", "Central government", "Industry", "Overseas", "Identifiable funding"]];
header(comp, "A3:N3");
const compValues = components.map(r => [r.academic_year, r.discipline,
  +r.mainstream_qr, +r.london_qr, +r.rdp_qr, +r.charity_qr, +r.business_qr, null,
  +r.research_councils, +r.charities, +r.central_government, +r.industry, +r.overseas, null]);
comp.getRange(`A4:N${3 + compValues.length}`).values = compValues;
for (let i = 0; i < compValues.length; i++) {
  const row = 4 + i;
  comp.getRange(`H${row}`).formulas = [[`=SUM(C${row}:G${row})`]];
  comp.getRange(`N${row}`).formulas = [[`=SUM(H${row}:M${row})`]];
}
body(comp, `A4:N${3 + compValues.length}`);
comp.getRange(`C4:N${3 + compValues.length}`).setNumberFormat("#,##0");
comp.getRange("A3:N63").format.rowHeight = 18;
comp.getRange("A24:N39").format.fill = lightGold;
comp.getRange("A60:N63").format.fill = lightGold;
comp.getRange("A3:N63").format.autofitColumns();
comp.getRange("A:A").format.columnWidth = 13;
comp.getRange("B:B").format.columnWidth = 22;
comp.getRange("C:N").format.columnWidth = 16;

// Annual series: calculations linked to Components.
const annual = wb.worksheets.add("Annual series");
annual.showGridlines = false;
title(annual, "A2", "QR share of identifiable research funding by broad discipline");
annual.getRange("A3").values = [["All values are England totals. Percentages are nominal funding shares; the shaded years use reconstructed QR subject allocations."]];
annual.getRange("A3:J3").format.font = { name: font, size: 10, italic: true, color: grey };
annual.getRange("A4:J4").values = [["Academic year", "Broad discipline", "QR data status", "Total QR (£)", "Identifiable funding (£)",
  "QR share", "Mainstream QR share", "Other QR share", "Sensitivity low", "Sensitivity high"]];
header(annual, "A4:J4");
for (let i = 0; i < series.length; i++) {
  const row = 5 + i, crow = 4 + i, s = series[i];
  annual.getRange(`A${row}:C${row}`).values = [[s.academic_year, s.discipline, s.qr_data_status]];
  annual.getRange(`D${row}:H${row}`).formulas = [[
    `=Components!H${crow}`, `=Components!N${crow}`, `=D${row}/E${row}`,
    `=Components!C${crow}/E${row}`, `=SUM(Components!D${crow}:G${crow})/E${row}`
  ]];
  annual.getRange(`I${row}:J${row}`).values = [[+s.qr_share_sensitivity_low, +s.qr_share_sensitivity_high]];
  if (s.academic_year >= "2015-16" && s.academic_year <= "2018-19") annual.getRange(`A${row}:J${row}`).format.fill = lightGold;
}
body(annual, "A5:J64");
annual.getRange("D5:E64").setNumberFormat("#,##0");
annual.getRange("F5:J64").setNumberFormat("0.0%");
annual.getRange("A4:J64").format.rowHeight = 18;
annual.getRange("A4:J64").format.autofitColumns();
annual.getRange("A:A").format.columnWidth = 13;
annual.getRange("B:B").format.columnWidth = 22;
annual.getRange("C:C").format.columnWidth = 48;
annual.getRange("D:E").format.columnWidth = 18;
annual.getRange("F:J").format.columnWidth = 18;

// Summary: headline reproduction, latest comparison, and full trend.
const summary = wb.worksheets.add("Summary");
summary.showGridlines = false;
body(summary, "A1:P32");
title(summary, "A2", "The role of QR funding across disciplines");
summary.getRange("A3").values = [["Reconstruction of HEFCE's 2010/11 result and extension to the latest common QR/HESA year (2024/25)"]];
summary.getRange("A3:H3").format.font = { name: font, size: 10, italic: true, color: grey };
section(summary, "A5:C5", "2010/11 Arts and Humanities reproduction");
summary.getRange("A6:C6").values = [["Mainstream QR", "Other QR", "Total QR share"]];
header(summary, "A6:C6");
summary.getRange("A7:C7").formulas = [["='Annual series'!G5", "='Annual series'!H5", "='Annual series'!F5"]];
summary.getRange("A7:C7").setNumberFormat("0.00%");
summary.getRange("A7:C7").format.font = { name: font, size: 14, bold: true, color: navy };
section(summary, "E5:H5", "2024/25 QR share");
summary.getRange("E6:H6").values = [["A&H", "Social Sciences", "STEM", "Health"]];
header(summary, "E6:H6");
summary.getRange("E7:H7").formulas = [["='Annual series'!F61", "='Annual series'!F62", "='Annual series'!F63", "='Annual series'!F64"]];
summary.getRange("E7:H7").setNumberFormat("0.0%");
summary.getRange("E7:H7").format.font = { name: font, size: 14, bold: true, color: blue };
summary.getRange("A9").values = [["The report's ‘around 70%’ is 68.20% on the reconstructed data. QR remains the largest identifiable stream for A&H in 2024/25."]];
summary.getRange("A9:H9").format.font = { name: font, size: 10, bold: true, color: "#334155" };

summary.getRange("A12:E12").values = [["Academic year", "Arts and Humanities", "Social Sciences", "STEM", "Health"]];
header(summary, "A12:E12");
const years = [...new Set(series.map(r => r.academic_year))];
for (let yi = 0; yi < years.length; yi++) {
  const row = 13 + yi, annualStart = 5 + yi * 4;
  summary.getRange(`A${row}`).values = [[years[yi]]];
  summary.getRange(`B${row}:E${row}`).formulas = [[0,1,2,3].map(offset => `='Annual series'!F${annualStart + offset}`)];
}
body(summary, "A13:E27");
summary.getRange("B13:E27").setNumberFormat("0.0%");
summary.getRange("A18:E21").format.fill = lightGold;
const chart = summary.charts.add("line", summary.getRange("A12:E27"));
chart.title = "QR share of identifiable research funding";
chart.titleTextStyle.fontSize = 12;
chart.titleTextStyle.typeface = font;
chart.legend = { position: "top", textStyle: { typeface: font, fontSize: 10 } };
chart.xAxis = { axisType: "textAxis", textStyle: { typeface: font, fontSize: 9 } };
chart.yAxis = { numberFormatCode: "0%", numberFormatSourceLinked: false, minimumScale: 0, maximumScale: 0.75,
  textStyle: { typeface: font, fontSize: 9 } };
chart.setPosition("G12", "P28");
const palette = [navy, blue, teal, gold];
chart.series.items.forEach((s, i) => { s.line = { fill: palette[i], style: "solid", width: i === 0 ? 3 : 2 }; });
summary.getRange("A30:B30").values = [[
  "Reading the bridge", "2015/16–2018/19 use the published national pots and reconstruct the missing subject distribution. Sensitivity columns show the effect of using the 2014/15 rather than 2019/20 RDP discipline mix."
]];
summary.getRange("A30").format.font = { name: font, size: 10, bold: true, color: navy };
summary.getRange("B30").format.font = { name: font, size: 10, color: grey };
summary.getRange("B30").format.wrapText = true;
summary.getRange("A:A").format.columnWidth = 15;
summary.getRange("B:B").format.columnWidth = 34;
summary.getRange("C:H").format.columnWidth = 18;
summary.getRange("A9:H9").format.rowHeight = 24;
summary.getRange("A30:B30").format.rowHeight = 48;

// Method and source register.
const method = wb.worksheets.add("Method");
method.showGridlines = false;
title(method, "A2", "Method, definitions and source register");
method.getRange("A4:D4").values = [["Topic", "Method used", "Implication", "Primary source"]];
header(method, "A4:D4");
const methods = [
  ["Scope", "England; academic years 2010/11–2024/25; four broad disciplines.", "2024/25 is the latest year with both QR allocations and a HESA denominator.", "HEFCE/Research England; HESA Finance Table 5/5b"],
  ["QR numerator", "Mainstream QR + London weighting + RDP supervision + charity support + business research element.", "National Research Libraries and temporary GCRF QR are excluded for consistency with the 2014 figure.", "HEFCE/Research England detailed allocation tables"],
  ["Denominator", "QR numerator + Research Councils + UK charities + central government + UK industry + overseas research income.", "HESA UK other sources and the source-table total are excluded, matching Figure 9.3.", "HESA Finance Plus Table 5b / Table 5 CSV"],
  ["2010/11 result", "Provider support-fund grants are split by the provider's qualifying HESA income across cost centres, then aggregated.", "A&H: 54.66% mainstream QR + 13.53% other QR = 68.20% total QR.", "resdata1011.xls; business1011.xls; HE_Finance_Plus_2010-11.zip"],
  ["2015/16–2018/19 mainstream", "Use the 2019/20 REF2014 UOA distribution, reverse UOA 4's 1.60 cost weight to 1.42, and scale to the published annual mainstream-plus-London pot.", "Reconstructs the stable formula base while excluding separately identified GCRF additions.", "HEFCE funding guides; RE 2019/20 data-source note"],
  ["2015/16–2018/19 RDP", "Use the adjusted 2019/20 UOA distribution and scale to each annual RDP pot.", "Exact values require unpublished student-level derived inputs; low/high columns use 2014/15 and 2019/20 discipline mixes as endpoint sensitivity.", "RE RDP 2019/20 workbook; HEFCE/RE funding guides"],
  ["Charity/business QR", "Split published provider grants by qualifying charity/industry income across that provider's HESA cost centres. Unmatched residuals use the national qualifying-income mix.", "The grant totals are observed; their discipline allocation is constructed consistently for every year.", "HEFCE/RE provider grants; HESA Finance Table 5"],
  ["2024/25 RDP", "Carry forward each provider's 2023/24 UOA shares and scale to its published 2024/25 RDP total.", "Matches Research England's stated carry-forward approach; new providers use national shares.", "Research England provider grant tables 2023/24 and 2024/25"],
  ["Broad disciplines", "Legacy HESA: Health 01,02,04–08; STEM 03,10–25 selected; Social 26–29,34,38,41; A&H 30,31,33,35,37. New HESA: Health 101–107; STEM 109–123; Social 108,124,127–136; A&H 125–126,137–145.", "Keeps the report's disciplinary logic across the 2012/13 cost-centre change.", "HESA cost-centre classifications"],
  ["REF UOA mapping", "REF2014: Health 1–4; STEM 5–16; Social 17–26; A&H 27–36. REF2021: Health 1–4; STEM 5–13; Social 14 and 16–24; A&H 15 and 25–34.", "Architecture stays in STEM; archaeology/area studies stay in A&H.", "REF UOA definitions and HESA mapping workbook"],
  ["Prices", "Nominal pounds; shares are calculated within each year.", "No deflator is required for the percentage series.", "All source tables"],
  ["Interpretation", "The metric measures QR as a share of identifiable income streams, not total research resources or full economic cost.", "Changes may reflect both QR allocations and movements in external income.", "HEFCE 2014 report, Figure 9.3"]
];
method.getRange(`A5:D${4 + methods.length}`).values = methods;
body(method, `A5:D${4 + methods.length}`);
method.getRange(`A5:D${4 + methods.length}`).format.wrapText = true;
method.getRange(`A5:D${4 + methods.length}`).format.verticalAlignment = "top";
method.getRange("A:A").format.columnWidth = 24;
method.getRange("B:B").format.columnWidth = 72;
method.getRange("C:C").format.columnWidth = 62;
method.getRange("D:D").format.columnWidth = 48;
method.getRange(`A5:D${4 + methods.length}`).format.rowHeight = 52;
section(method, "A19:D19", "Source links");
const sources = [
  ["2014 report", `${root}/2014_qrreview.pdf`, "Attached source report"],
  ["RE 2019/20 detailed QR data", "https://www.ukri.org/publications/funding-allocations-2019-2020-quality-related-research-funding-data/", "Mainstream, London and RDP UOA workbooks"],
  ["RE 2023/24 provider tables", "https://www.ukri.org/publications/higher-education-provider-grant-data-tables-2023-to-2024/", "Provider and UOA allocations"],
  ["RE 2024/25 provider tables", "https://www.ukri.org/publications/higher-education-provider-grant-data-tables-2024-to-2025/", "Latest common-year QR allocations"],
  ["Local source inventory", `${root}/data_sources/UKRI_QR_download_manifest.md`, "Downloaded workbooks and provenance"]
];
method.getRange("A20:C20").values = [["Source", "Location", "Use"]];
header(method, "A20:C20");
method.getRange(`A21:C${20 + sources.length}`).values = sources;
body(method, `A21:C${20 + sources.length}`);
method.getRange(`A21:C${20 + sources.length}`).format.wrapText = true;
method.getRange(`A21:C${20 + sources.length}`).format.rowHeight = 34;

// Checks: compare calculated QR totals with independently published control totals.
const checks = wb.worksheets.add("Checks");
checks.showGridlines = false;
title(checks, "A2", "QR control-total checks");
checks.getRange("A3").values = [["Calculated broad-discipline totals reconcile to the published QR component totals; tolerance is £2 for source rounding."]];
checks.getRange("A3:E3").format.font = { name: font, size: 10, italic: true, color: grey };
checks.getRange("A5:E5").values = [["Academic year", "Published/control QR (£)", "Calculated QR (£)", "Difference (£)", "Status"]];
header(checks, "A5:E5");
const controls = [1596545711,1551545721,1551545692,1551545692,1551498547,1552000000,1552000000,1552000000,1578000000,1622387305,1622387117,1622387148,1967226170,1974225843,1979999992];
for (let i = 0; i < years.length; i++) {
  const row = 6 + i, cstart = 4 + i * 4, cend = cstart + 3;
  checks.getRange(`A${row}:B${row}`).values = [[years[i], controls[i]]];
  checks.getRange(`C${row}:E${row}`).formulas = [[`=SUM(Components!H${cstart}:H${cend})`, `=C${row}-B${row}`, `=IF(ABS(D${row})<=2,"OK","Review")`]];
  if (years[i] >= "2015-16" && years[i] <= "2018-19") checks.getRange(`A${row}:E${row}`).format.fill = lightGold;
}
body(checks, "A6:E20");
checks.getRange("B6:D20").setNumberFormat("#,##0");
checks.getRange("A:E").format.autofitColumns();
checks.getRange("A:A").format.columnWidth = 14;
checks.getRange("B:D").format.columnWidth = 24;
checks.getRange("E:E").format.columnWidth = 12;
checks.getRange("A23:C23").values = [["Additional check", "Result", "Meaning"]];
header(checks, "A23:C23");
checks.getRange("A24:C27").values = [
  ["2010/11 A&H QR share", null, "Should round to the report's ‘around 70%’"],
  ["2010/11 component identity", null, "Mainstream share + other QR share = total QR share"],
  ["Reconstructed years", 4, "2015/16–2018/19"],
  ["Latest common year", "2024-25", "Later QR years lack a matching HESA denominator"]
];
checks.getRange("B24").formulas = [["='Annual series'!F5"]];
checks.getRange("B25").formulas = [["='Annual series'!G5+'Annual series'!H5-'Annual series'!F5"]];
checks.getRange("B24:B25").setNumberFormat("0.00%");
body(checks, "A24:C27");
checks.getRange("A:A").format.columnWidth = 32;
checks.getRange("B:B").format.columnWidth = 20;
checks.getRange("C:C").format.columnWidth = 58;

// Navigation order: outputs, calculations, inputs, documentation/checks.
wb.worksheets.getItem("Summary").position = 0;
wb.worksheets.getItem("Annual series").position = 1;
wb.worksheets.getItem("Components").position = 2;
wb.worksheets.getItem("Method").position = 3;
wb.worksheets.getItem("Checks").position = 4;

wb.recalculate();
const inspect = await wb.inspect({ kind: "table", range: "Summary!A5:H9", include: "values,formulas", tableMaxRows: 10, tableMaxCols: 10 });
console.log(inspect.ndjson);
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 300 }, summary: "final formula error scan" });
console.log(errors.ndjson);
for (const name of ["Summary", "Annual series", "Components", "Method", "Checks"]) {
  const preview = await wb.render({ sheetName: name, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${root}/tmp/artifact_builder/${name.replace(/ /g, "_")}.png`, new Uint8Array(await preview.arrayBuffer()));
}
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);
console.log(JSON.stringify({ outputPath, sheets: ["Summary", "Annual series", "Components", "Method", "Checks"] }));
