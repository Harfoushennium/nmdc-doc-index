import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbookPath = path.resolve(process.argv[2] || "excel/NMDC_Document_Index_Base.xlsx");
const previewDir = path.resolve(process.argv[3] || "/tmp/nmdc-workbook-preview");
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(workbookPath));

const requiredSheets = [
  "Home", "Master Documents", "Revisions", "Transactions", "Pending Update",
  "Review Flags", "User Decisions", "Configuration", "Rules & Mappings",
  "Update History", "Error Log", "Help", "System Data",
];
const actualSheets = workbook.worksheets.items.map((sheet) => sheet.name);
for (const name of requiredSheets) {
  if (!actualSheets.includes(name)) throw new Error(`Required worksheet missing: ${name}`);
}

const summary = await workbook.inspect({
  kind: "sheet,table",
  maxChars: 12000,
  tableMaxRows: 3,
  tableMaxCols: 8,
  tableMaxCellChars: 80,
});
await fs.mkdir(previewDir, { recursive: true });
await fs.writeFile(path.join(previewDir, "workbook-inspection.ndjson"), summary.ndjson, "utf8");

for (const sheetName of ["Home", "Configuration", "Rules & Mappings", "Help"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(
    path.join(previewDir, `${sheetName.replaceAll(" ", "_")}.png`),
    new Uint8Array(await preview.arrayBuffer()),
  );
}

console.log(JSON.stringify({ workbookPath, sheets: actualSheets, previewDir }));
