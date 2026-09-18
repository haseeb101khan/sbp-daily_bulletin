import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const rootDir = path.resolve(".");
const outputDir = path.resolve("../../outputs/daily_bulletin_vercel");
const dataDir = path.join(rootDir, "data");
await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(dataDir, { recursive: true });

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Sheet3");
sheet.showGridLines = true;

const headers = ["Domain", "Heading", "Paper", "Link"];
sheet.getRange("A1:D1").values = [headers];
sheet.getRange("A1:D1").format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF", size: 11 },
  borders: { preset: "all", style: "thin", color: "#B7C9D8" },
  verticalAlignment: "center",
};

const rows = Array.from({ length: 60 }, () => ["", "", "", ""]);
sheet.getRange("A2:D61").values = rows;
sheet.getRange("A2:D61").format = {
  fill: "#FFFFFF",
  font: { color: "#111827", size: 10 },
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#D9E2EA" },
  verticalAlignment: "top",
};

const widths = [18, 72, 18, 104];
for (let col = 0; col < widths.length; col += 1) {
  sheet.getRangeByIndexes(0, col, 61, 1).format.columnWidth = widths[col];
}

sheet.getRange("A1:D1").format.rowHeight = 24;
sheet.getRange("A2:D61").format.rowHeight = 42;
sheet.freezePanes.freezeRows(1);

sheet.getRange("A2:A61").dataValidation = {
  rule: {
    type: "list",
    values: [
      "SBP Related News",
      "Domestic News",
      "Editorials/Articles",
      "Foreign News",
      "SBP Related News / Press Release",
      "Editorials / Opinion / Analysis",
      "Foreign Media Updates",
    ],
  },
};

const check = await workbook.inspect({
  kind: "table",
  range: "Sheet3!A1:D12",
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 4,
});
console.log(check.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const preview = await workbook.render({ sheetName: "Sheet3", range: "A1:D18", scale: 1, format: "png" });
await fs.writeFile(path.join(outputDir, "ecd_news_links_template_preview.png"), new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
const outputPath = path.join(outputDir, "Daily_Bulletin_Empty_Template.xlsx");
await output.save(outputPath);
await fs.copyFile(outputPath, path.join(dataDir, "Daily_Bulletin_Empty_Template.xlsx"));

console.log(`Saved ${outputPath}`);
