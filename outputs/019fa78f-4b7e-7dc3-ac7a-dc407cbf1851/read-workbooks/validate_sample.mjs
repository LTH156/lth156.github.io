import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const projectDir = path.resolve("../../..");
const samplePath = path.join(projectDir, "DNTT", "ket_qua_test", "PC2023.012.xlsx");
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(samplePath));

const check = await workbook.inspect({
  kind: "table",
  range: "ĐNTT!A1:H25",
  include: "values,formulas",
  tableMaxRows: 25,
  tableMaxCols: 8,
  maxChars: 14000,
});
console.log(check.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "sample formula error scan",
  maxChars: 6000,
});
console.log(errors.ndjson);

const preview = await workbook.render({
  sheetName: "ĐNTT",
  autoCrop: "all",
  scale: 1.5,
  format: "png",
});
await fs.writeFile(
  path.join(projectDir, "DNTT", "ket_qua_test", "PC2023.012-preview.png"),
  new Uint8Array(await preview.arrayBuffer()),
);
