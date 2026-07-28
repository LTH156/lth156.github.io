import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const projectDir = path.resolve("../../..");
const sourcePath = path.join(projectDir, "DNTT", "Mau.xlsx");
const outputPath = path.join(projectDir, "DNTT", "Mau_template.xlsx");

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
const sheet = workbook.worksheets.getItem("ĐNTT");

sheet.getRange("D6").values = [["%NGAYLAP%"]];
sheet.getRange("D6").format.numberFormat = "dd/mm/yyyy";
sheet.getRange("B16").values = [["%NOIDUNG%"]];
sheet.getRange("F16").values = [["%DONGIA%"]];
sheet.getRange("G16").formulas = [["=E16*F16"]];
sheet.getRange("H16").values = [[null]];
sheet.getRange("G22").formulas = [["=SUM(G16:G21)"]];
sheet.getRange("C23").values = [["%SOTIENBANGCHU%"]];
for (const headerRange of ["A7:H7", "A10:H10", "A15:H15"]) {
  sheet.getRange(headerRange).format = {
    fill: "#FFFFFF",
    font: {
      bold: true,
      color: "#000000",
    },
  };
}

const preview = await workbook.render({
  sheetName: "ĐNTT",
  autoCrop: "all",
  scale: 1.5,
  format: "png",
});
await fs.writeFile(
  path.join(projectDir, "outputs", "019fa78f-4b7e-7dc3-ac7a-dc407cbf1851", "read-workbooks", "template-preview.png"),
  new Uint8Array(await preview.arrayBuffer()),
);

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);

const check = await workbook.inspect({
  kind: "table",
  range: "ĐNTT!A1:H25",
  include: "values,formulas",
  tableMaxRows: 25,
  tableMaxCols: 8,
  maxChars: 12000,
});
console.log(check.ndjson);
