import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workDir = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(workDir, "../../..");
const inputs = [
  path.join(rootDir, "DNTT", "Mau.xlsx"),
  path.join(rootDir, "DNTT", "Phiếu chi 2023 - ĐNTT.xlsx"),
];

const results = [];

for (const inputPath of inputs) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
  const sheetOverview = await workbook.inspect({
    kind: "sheet",
    include: "id,name",
    maxChars: 12000,
  });
  const summary = await workbook.inspect({
    kind: "workbook,sheet,table,definedName,drawing",
    maxChars: 30000,
    tableMaxRows: 15,
    tableMaxCols: 16,
    tableMaxCellChars: 160,
  });

  const sheets = [];
  for (let index = 0; ; index += 1) {
    let sheet;
    try {
      sheet = workbook.worksheets.getItemAt(index);
    } catch {
      break;
    }
    if (!sheet) break;

    const used = sheet.getUsedRange();
    let address = null;
    let values = [];
    let formulas = [];
    if (used) {
      address = used.address;
      values = used.values;
      formulas = used.formulas;
    }

    const previewPath = path.join(
      workDir,
      `${path.basename(inputPath, ".xlsx").replaceAll(/[^a-zA-Z0-9_-]+/g, "_")}-${index + 1}.png`,
    );
    let previewError = null;
    try {
      const preview = await workbook.render({
        sheetName: sheet.name,
        autoCrop: "all",
        scale: 1.25,
        format: "png",
      });
      await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
    } catch (error) {
      previewError = String(error?.message || error);
    }

    sheets.push({
      index,
      name: sheet.name,
      usedRange: address,
      rowCount: values.length,
      columnCount: values.reduce((max, row) => Math.max(max, row.length), 0),
      values,
      formulas,
      previewPath,
      previewError,
    });
  }

  results.push({
    file: inputPath,
    sheetOverview: sheetOverview.ndjson,
    summary: summary.ndjson,
    sheets,
  });
}

await fs.writeFile(
  path.join(workDir, "inspection.json"),
  JSON.stringify(results, null, 2),
  "utf8",
);

for (const result of results) {
  console.log(`FILE: ${result.file}`);
  for (const sheet of result.sheets) {
    console.log(
      `  SHEET: ${sheet.name} | used=${sheet.usedRange} | rows=${sheet.rowCount} | cols=${sheet.columnCount}`,
    );
    if (sheet.previewError) console.log(`  PREVIEW ERROR: ${sheet.previewError}`);
  }
}
