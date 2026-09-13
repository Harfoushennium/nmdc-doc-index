import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";
import JSZip from "jszip";

const args = process.argv.slice(2);
function arg(name, fallback = "") {
  const index = args.indexOf(name);
  return index >= 0 && index + 1 < args.length ? args[index + 1] : fallback;
}

const root = path.resolve(arg("--root", process.cwd()));
const exchangeDir = path.resolve(arg("--exchange", path.join(root, "outputs", "nmdc_doc_index_real_review", "exchange")));
const outputPath = path.resolve(arg("--output", path.join(root, "outputs", "nmdc_doc_index_real_review", "NMDC_Document_Index_Real_Data_Review.xlsx")));
const packageMode = arg("--mode", "review").toLowerCase();
const productionMode = packageMode === "production";
const eventReviewLimit = Math.max(500, Number(arg("--event-limit", "5000")) || 5000);
const detailReviewLimit = Math.max(500, Number(arg("--detail-limit", "1500")) || 1500);
const nativeLinks = [];

const NAVY = "#0B1F33";
const BLUE = "#146C94";
const TEAL = "#118AB2";
const GREEN = "#2E7D32";
const AMBER = "#F4B400";
const RED = "#C62828";
const INK = "#1F2937";
const MUTED = "#5B6470";
const PALE_BLUE = "#EAF4FB";
const PALE_GREEN = "#EAF6EC";
const PALE_AMBER = "#FFF7DB";
const PALE_RED = "#FDECEC";
const GRID = "#D8E1E8";

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          quoted = false;
        }
      } else {
        field += ch;
      }
    } else if (ch === '"') {
      quoted = true;
    } else if (ch === ",") {
      row.push(field);
      field = "";
    } else if (ch === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (ch !== "\r") {
      field += ch;
    }
  }
  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }
  if (rows.length && rows[0].length) rows[0][0] = rows[0][0].replace(/^\uFEFF/, "");
  return rows;
}

async function csvFile(name) {
  return parseCsv(await fs.readFile(path.join(exchangeDir, name), "utf8"));
}

function colName(number) {
  let n = number;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function address(row, col) {
  return `${colName(col)}${row}`;
}

function safeSheetName(name) {
  return name.replace(/'/g, "''");
}

function linkFormula(sheetName, cell = "A1", label = "Open") {
  return `=HYPERLINK("#'${safeSheetName(sheetName)}'!${cell}","${String(label).replace(/"/g, '""')}")`;
}

function registerLink(sheetName, ref, targetSheet, targetCell, label) {
  nativeLinks.push({ sheetName, ref, targetSheet, targetCell, label });
}

function setFont(range, config) {
  range.format.font = { ...(range.format.font || {}), ...config };
}

function baseSheet(sheet, tabColor = BLUE) {
  sheet.showGridLines = false;
  sheet.tabColor = tabColor;
}

function titleBlock(sheet, title, subtitle, width = 10) {
  const end = colName(width);
  const titleRange = sheet.getRange(`A1:${end}1`);
  titleRange.merge();
  titleRange.values = [[title]];
  titleRange.format = {
    fill: NAVY,
    font: { bold: true, color: "#FFFFFF", size: 16 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
  titleRange.format.rowHeight = 30;
  const subRange = sheet.getRange(`A2:${end}2`);
  subRange.merge();
  subRange.values = [[subtitle]];
  subRange.format = {
    fill: PALE_BLUE,
    font: { color: MUTED, italic: true, size: 10 },
    wrapText: true,
    verticalAlignment: "center",
  };
  subRange.format.rowHeight = 30;
}

function homeLink(sheet, target = "Home") {
  const range = sheet.getRange("A3:B3");
  range.merge();
  range.values = [["← Home"]];
  registerLink(sheet.name, "A3:B3", target, "A1", "← Home");
  range.format = {
    fill: PALE_BLUE,
    font: { bold: true, color: BLUE, underline: true },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  range.format.rowHeight = 22;
}

function setWidths(sheet, widths) {
  for (const [index, width] of widths.entries()) {
    sheet.getRange(`${colName(index + 1)}:${colName(index + 1)}`).format.columnWidth = width;
  }
}

function headerStyle(range) {
  range.format = {
    fill: BLUE,
    font: { bold: true, color: "#FFFFFF", size: 10 },
    wrapText: true,
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: GRID },
  };
  range.format.rowHeight = 32;
}

function bodyStyle(range) {
  range.format = {
    font: { color: INK, size: 9 },
    verticalAlignment: "top",
    wrapText: false,
    borders: { preset: "inside", style: "thin", color: GRID },
  };
  range.format.numberFormat = "@";
}

function addTable(sheet, startRow, rows, tableName, widths = []) {
  if (!rows.length) return { rowCount: 0, colCount: 0 };
  const colCount = rows[0].length;
  const rowCount = rows.length;
  const range = sheet.getRangeByIndexes(startRow - 1, 0, rowCount, colCount);
  range.values = rows;
  headerStyle(sheet.getRangeByIndexes(startRow - 1, 0, 1, colCount));
  if (rowCount > 1) bodyStyle(sheet.getRangeByIndexes(startRow, 0, rowCount - 1, colCount));
  // Excel tables require at least one data row in this artifact runtime.  A
  // header-only view is still useful (for example Pending Update before the
  // first staged run), so leave it as a styled range instead of creating an
  // invalid/empty table.
  if (rowCount > 1) {
    const table = sheet.tables.add(range, true, tableName);
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
  }
  if (widths.length) setWidths(sheet, widths);
  sheet.freezePanes.freezeRows(startRow);
  return { rowCount, colCount };
}

function addConditionalLevelFormatting(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  range.conditionalFormats.add("containsText", { text: "REVIEW", format: { fill: PALE_AMBER, font: { color: "#7A5A00", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "CONFLICT", format: { fill: PALE_RED, font: { color: RED, bold: true } } });
  range.conditionalFormats.add("containsText", { text: "OK", format: { fill: PALE_GREEN, font: { color: GREEN, bold: true } } });
}

function rowsToObjects(matrix) {
  if (!matrix.length) return { headers: [], rows: [] };
  const headers = matrix[0];
  return { headers, rows: matrix.slice(1) };
}

function lookup(headers, row, field) {
  const index = headers.indexOf(field);
  return index >= 0 ? row[index] || "" : "";
}

function addDetailSheet(workbook, name, title, subtitle, matrix, tableName, widths, tabColor = BLUE) {
  const sheet = workbook.worksheets.add(name);
  baseSheet(sheet, tabColor);
  const { headers, rows } = rowsToObjects(matrix);
  titleBlock(sheet, title, subtitle, Math.max(headers.length, 8));
  homeLink(sheet);
  if (headers.length) {
    addTable(sheet, 5, [headers, ...rows], tableName, widths);
    if (headers[0] === "Flag Level") addConditionalLevelFormatting(sheet, `A6:A${Math.max(6, rows.length + 5)}`);
  }
  return sheet;
}

function addTextSheet(workbook, name, title, subtitle, lines, tabColor = BLUE) {
  const sheet = workbook.worksheets.add(name);
  baseSheet(sheet, tabColor);
  titleBlock(sheet, title, subtitle, 9);
  homeLink(sheet);
  let row = 5;
  for (const item of lines) {
    const range = sheet.getRange(`A${row}:I${row}`);
    range.merge();
    range.values = [[item]];
    range.format = { font: { color: INK, size: 10 }, wrapText: true, verticalAlignment: "top" };
    range.format.rowHeight = item.length > 130 ? 46 : 28;
    row += 1;
  }
  setWidths(sheet, [28, 16, 16, 16, 16, 16, 16, 16, 16]);
  return sheet;
}

async function main() {
  const dashboardMatrix = await csvFile("dashboard.csv");
  const docsMatrix = await csvFile("master_documents.csv");
  const revisionsMatrix = await csvFile("revisions.csv");
  const eventsMatrix = await csvFile("events.csv");
  const pendingMatrix = await csvFile("pending_update.csv");
  const flagsMatrix = await csvFile("flags.csv");
  const historyMatrix = await csvFile("history.csv");
  const errorsMatrix = await csvFile("errors.csv");
  const inventoryMatrix = await csvFile("source_inventory.csv");
  const rulesMatrix = await csvFile("classification_rules.csv");
  const baseline = JSON.parse(await fs.readFile(path.join(exchangeDir, "baseline_summary.json"), "utf8"));
  const dashboard = rowsToObjects(dashboardMatrix);
  const dashboardRow = dashboard.rows[0] || [];
  const value = (field) => lookup(dashboard.headers, dashboardRow, field);
  const eventRows = eventsMatrix.slice(1);
  // Authoring uses a bounded event slice so the artifact authoring runtime
  // stays within its memory budget. After export, the OOXML patch below
  // restores the complete event table into the delivered workbook.
  const eventReviewRows = eventRows.length <= eventReviewLimit
    ? eventRows
    : eventRows.slice(0, Math.ceil(eventReviewLimit * 0.8)).concat(eventRows.slice(-Math.floor(eventReviewLimit * 0.2)));
  const eventsReviewMatrix = [eventsMatrix[0], ...eventReviewRows];
  const detailRows = (matrix) => {
    const rows = matrix.slice(1);
    if (rows.length <= detailReviewLimit) return matrix;
    const head = Math.ceil(detailReviewLimit * 0.8);
    return [matrix[0], ...rows.slice(0, head).concat(rows.slice(-Math.floor(detailReviewLimit * 0.2)))];
  };
  const docsReviewMatrix = detailRows(docsMatrix);
  const revisionsReviewMatrix = detailRows(revisionsMatrix);

  const workbook = Workbook.create();

  // Home is the first sheet and contains real, clickable internal links.  The
  // action labels intentionally route to the relevant review page in this
  // snapshot; execution actions will be wired to the macro-enabled package.
  const home = workbook.worksheets.add("Home");
  baseSheet(home, TEAL);
  titleBlock(home, "NMDC Document Index", productionMode ? "PRODUCTION PACKAGE • Excel control panel" : "REAL-DATA REVIEW PACKAGE • click a control to move to the relevant review page", productionMode ? 12 : 10);
  const status = home.getRange(productionMode ? "A4:L4" : "A4:J4");
  status.merge();
  status.values = [[productionMode ? "STATUS: PRODUCTION SETUP • Select the DATA folder, stage an update, then review before approval" : "STATUS: REAL-DATA BASELINE — REVIEW ONLY • No approved master data was changed"]];
  status.format = { fill: PALE_AMBER, font: { bold: true, color: "#7A5A00", size: 11 }, horizontalAlignment: "center", verticalAlignment: "center" };
  status.format.rowHeight = 26;

  if (productionMode) {
    const blocks = [
      ["A5:C5", "APPROVED MASTER", [["A6", "Status", "B6:C6", "NOT YET APPROVED"], ["A7", "Run", "B7:C7", ""], ["A8", "Data folder", "B8:C8", "Select from Home"], ["A9", "Last update", "B9:C9", "Never"]]],
      ["D5:F5", "APPROVED COUNTS", [["D6", "Documents", "E6:F6", "0"], ["D7", "Revisions", "E7:F7", "0"], ["D8", "Transactions", "E8:F8", "0"], ["D9", "Baseline", "E9:F9", `${value("Approved Documents")} documents for first scan`]]],
      ["G5:I5", "PENDING UPDATE", [["G6", "Status", "H6:I6", "NONE"], ["G7", "Run", "H7:I7", ""], ["G8", "Changes", "H8:I8", "Stage an update first"], ["G9", "Safety", "H9:I9", "Approval required"]]],
      ["J5:L5", "REVIEW CONTROL", [["J6", "Review flags", "K6:L6", "0"], ["J7", "Conflicts", "K7:L7", "0"], ["J8", "Real baseline", "K8:L8", String(value("Review Flags"))], ["J9", "Sources", "K9:L9", String(baseline.counts.selected_workbooks)]]],
    ];
    for (const [headerRange, headerText, rows] of blocks) {
      const header = home.getRange(headerRange);
      header.merge();
      header.values = [[headerText]];
      header.format = { fill: NAVY, font: { bold: true, color: "#FFFFFF", size: 9 }, horizontalAlignment: "center", verticalAlignment: "center" };
      for (const [labelCell, labelText, valueRange, initialValue] of rows) {
        const label = home.getRange(labelCell);
        label.values = [[labelText]];
        label.format = { fill: PALE_BLUE, font: { bold: true, color: INK, size: 8 }, verticalAlignment: "center", borders: { preset: "all", style: "thin", color: GRID } };
        const output = home.getRange(valueRange);
        output.merge();
        output.values = [[initialValue]];
        output.format = { fill: "#FFFFFF", font: { bold: true, color: TEAL, size: 9 }, verticalAlignment: "center", wrapText: true, borders: { preset: "all", style: "thin", color: GRID } };
      }
    }
  } else {
    const cards = [
      ["Documents", value("Approved Documents"), "Master Documents", "A1"],
      ["Revisions", value("Approved Revisions"), "Revisions", "A1"],
      ["Transactions", value("Approved Transactions"), "Transactions", "A1"],
      ["Review flags", value("Review Flags"), "Review Flags", "A1"],
      ["Source workbooks", baseline.counts.selected_workbooks, "System Data", "A1"],
    ];
    for (let i = 0; i < cards.length; i += 1) {
      const col = i * 2 + 1;
      const top = home.getRange(`${colName(col)}6:${colName(col + 1)}6`);
      top.merge();
      top.values = [[cards[i][0]]];
      top.format = { fill: NAVY, font: { bold: true, color: "#FFFFFF", size: 10 }, horizontalAlignment: "center", verticalAlignment: "center" };
      const number = home.getRange(`${colName(col)}7:${colName(col + 1)}8`);
      number.merge();
      number.values = [[String(cards[i][1])]];
      number.format = { fill: "#FFFFFF", font: { bold: true, color: TEAL, size: 18 }, horizontalAlignment: "center", verticalAlignment: "center", borders: { preset: "all", style: "thin", color: GRID } };
      const target = home.getRange(`${colName(col)}9:${colName(col + 1)}9`);
      target.merge();
      target.values = [[`Open ${cards[i][2]}`]];
      registerLink(home.name, `${colName(col)}9:${colName(col + 1)}9`, cards[i][2], cards[i][3], `Open ${cards[i][2]}`);
      target.format = { fill: PALE_BLUE, font: { color: BLUE, underline: true, size: 9 }, horizontalAlignment: "center", verticalAlignment: "center" };
    }
  }

  const buttonSpecs = [
    ["🔄 Update Changed Files", "Help", "A22", PALE_BLUE],
    ["♻️ Full Rescan / Rebuild All", "Help", "A22", PALE_BLUE],
    ["👀 Review Pending Update", "Pending Update", "A1", PALE_AMBER],
    ["✅ Approve Update", "User Decisions", "A1", PALE_GREEN],
    ["⏸️ Hold Update", "User Decisions", "A1", PALE_AMBER],
    ["❌ Reject Update", "User Decisions", "A1", PALE_RED],
    ["📂 Select Data Folder", "Configuration", "A1", PALE_BLUE],
    ["🚩 Review Flags", "Review Flags", "A1", PALE_AMBER],
    ["🚩 Flag Wrong Data", "Review Flags", "A1", PALE_AMBER],
    ["⚙️ Configuration", "Configuration", "A1", PALE_BLUE],
    ["📝 Report Requirement / Problem", "Help", "A34", PALE_BLUE],
    ["📋 View Log", "Error Log", "A1", PALE_RED],
    ["🔃 Refresh Dashboard", "Home", "A1", PALE_GREEN],
    ["Rules & Mappings", "Rules & Mappings", "A1", PALE_BLUE],
    ["Help", "Help", "A1", PALE_BLUE],
  ];
  let bRow = 12;
  for (let i = 0; i < buttonSpecs.length; i += 1) {
    const [label, targetSheet, targetCell, fill] = buttonSpecs[i];
    const col = productionMode ? (i % 3 === 0 ? "A" : i % 3 === 1 ? "E" : "I") : (i % 3 === 0 ? "A" : i % 3 === 1 ? "D" : "G");
    const end = productionMode ? (col === "A" ? "D" : col === "E" ? "H" : "L") : (col === "A" ? "C" : col === "D" ? "F" : "I");
    const row = bRow + Math.floor(i / 3) * 3;
    const range = home.getRange(`${col}${row}:${end}${row + 1}`);
    range.merge();
    range.values = [[label]];
    registerLink(home.name, `${col}${row}:${end}${row + 1}`, targetSheet, targetCell, label);
    range.format = { fill, font: { bold: true, color: BLUE, underline: true, size: 10 }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true, borders: { preset: "all", style: "thin", color: GRID } };
    range.format.rowHeight = 28;
  }
  const note = home.getRange(productionMode ? "A29:L33" : "A29:J33");
  note.merge();
  note.values = [[productionMode ? "Production workflow: select the DATA folder, press Update Changed Files or Full Rescan, review Pending Update and Review Flags, then Approve, Hold, or Reject. Approved data changes only after the explicit Approve action. Document links become clickable after the engine resolves them against the selected DATA folder." : "What is functional in this file: every colored control is a real Excel hyperlink to the relevant review page, the tables contain the Cycle 3 real-data baseline, and document/source identifiers are stored as text. The file is deliberately review-only. Running the engine, selecting an external folder, and changing the approved master still require the macro-enabled production package; this workbook does not pretend those actions are available here."]];
  note.format = { fill: PALE_BLUE, font: { color: INK, size: 10 }, wrapText: true, verticalAlignment: "center", borders: { preset: "all", style: "thin", color: GRID } };
  note.format.rowHeight = 78;
  setWidths(home, productionMode ? [15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15] : [18, 18, 4, 18, 18, 4, 18, 18, 4, 18]);

  addDetailSheet(workbook, "Master Documents", "Master Documents", `Real-data document view • ${value("Approved Documents")} documents • identifiers preserved as text`, docsReviewMatrix, "MasterDocuments", [10, 12, 18, 22, 18, 24, 28, 42, 24, 12, 14, 22, 34, 42, 26, 10, 10, 30]);
  addDetailSheet(workbook, "Revisions", "Revisions", `Real-data revision view • ${value("Approved Revisions")} revisions • identifiers preserved as text`, revisionsReviewMatrix, "RevisionRegister", [10, 12, 30, 42, 12, 16, 32, 42, 26, 10, 10]);
  addDetailSheet(workbook, "Transactions", "Transactions / Events", `Lossless event-level view • ${value("Approved Transactions")} event records • source locations preserved`, eventsReviewMatrix, "EventRegister", [10, 12, 30, 42, 12, 34, 14, 28, 24, 45, 16, 34, 42, 26, 10, 10, 32]);

  const pending = addDetailSheet(workbook, "Pending Update", "Pending Update", "No staged update is included in this review-only baseline. Use the macro-enabled package to create a fresh staged comparison.", pendingMatrix, "PendingUpdate", [16, 32, 12, 30, 12, 30, 42, 52, 16], PALE_AMBER);
  if (!pendingMatrix.slice(1).length) {
    const empty = pending.getRange("A5:I7");
    empty.merge();
    empty.values = [["NO STAGED UPDATE\n\nThis workbook contains the extracted real-data baseline for review. It has not run an update and it has not changed approved master data."]];
    empty.format = { fill: PALE_AMBER, font: { bold: true, color: "#7A5A00", size: 11 }, wrapText: true, horizontalAlignment: "center", verticalAlignment: "center", borders: { preset: "all", style: "thin", color: GRID } };
    empty.format.rowHeight = 70;
  }
  addDetailSheet(workbook, "Review Flags", "Review Flags", `${value("Review Flags")} worksheets require safe layout review • no conflicts were created`, flagsMatrix, "ReviewFlags", [12, 24, 48, 55, 12, 24, 12, 46, 28, 12, 12, 18, 26, 18, 32], PALE_AMBER);
  const decisions = addTextSheet(workbook, "User Decisions", "User Decisions / Overrides", productionMode ? "Auditable owner decisions and configuration changes" : "Owner decision area • this review package records no approval, hold, reject, or override", [
    productionMode ? "Decisions are recorded against the exact staged run displayed in Excel." : "Review package status: REVIEW ONLY.",
    productionMode ? "Approve changes the approved master pointer only after confirmation. Hold and Reject keep the approved master unchanged." : "There is no staged run ID to approve, hold, or reject in this snapshot.",
    productionMode ? "Rules & Mappings changes are validated and recorded before a staged update is created." : "When the macro-enabled workflow is available, decisions will be tied to the exact displayed pending run ID and logged before any approved master pointer can change.",
    "Owner override rule: an explicit owner identity decision takes priority, but contradictions raise APPROVED_OVERRIDE_CONTRADICTED for review.",
  ], PALE_GREEN);
  const decisionHeader = ["Decision ID", "Date/Time", "User", "Decision Type", "Source File", "Project / Document / Revision", "Previous Value", "Approved Value", "User Note", "Status", "Configuration Version"];
  addTable(decisions, 11, [decisionHeader, new Array(decisionHeader.length).fill("")], "UserDecisionLog", [30, 22, 24, 24, 42, 36, 28, 28, 55, 24, 26]);
  decisions.getRange("D12").dataValidation = { rule: { type: "list", values: ["APPROVE", "HOLD", "REJECT", "OVERRIDE", "CONFIGURATION CHANGE", ""] } };
  addDetailSheet(workbook, "Configuration", "Configuration", productionMode ? "Editable production settings" : "Editable control surface for the future macro-enabled engine integration", [
    ["Setting", "Current value", "Owner note"],
    ["Package mode", productionMode ? "PRODUCTION" : "REAL-DATA REVIEW ONLY", productionMode ? "Operational buttons are attached during Windows Excel packaging." : "This workbook is a review snapshot; engine execution is not attached."],
    ["Source set", "Cycle 3 selected DATA workbooks", "46 selected workbooks were extracted without changing DATA/."],
    ["Data Folder", productionMode ? "" : "DATA", productionMode ? "Use Select Data Folder on Home." : "The production workbook will let the owner choose an external data folder."],
    ["Runtime Folder", "", "Leave blank to use the runtime folder beside the workbook."],
    ["Engine Executable Path", "", "Leave blank to use engine\\nmdc_index_engine.exe beside the workbook."],
    ["Classification Rules File", "config\\classification_rules.csv", "Rules & Mappings exports to this controlled file."],
    ["Project Identity Overrides File", "config\\project_identity_overrides.csv", "Owner-approved project identities."],
    ["Parser Version", "cycle3-extractor-v1", "Displayed for traceability."],
    ["Configuration Version", "INITIAL", "Changes automatically when Rules & Mappings is saved."],
    ["Approval safety", "Explicit approval only", "Staged data must never silently replace approved data."],
  ], "Configuration", [24, 34, 80], PALE_BLUE);
  const rulesSheet = addDetailSheet(workbook, "Rules & Mappings", "Rules & Mappings", productionMode ? "Edit business classification rules here; changes are validated before staging" : "Business classification rules loaded from the real project configuration • editability is shown for review", rulesMatrix, "ClassificationRules", [12, 10, 14, 18, 34, 34, 28, 28, 28, 28, 28, 12, 12, 14, 14, 14, 42], PALE_BLUE);
  const ruleNote = rulesSheet.getRange("A4:Q4");
  ruleNote.merge();
  ruleNote.values = [[productionMode ? "Rule changes are validated and versioned when you start an update. They trigger reprocessing and remain staged until you explicitly approve the resulting master update." : "Owner note: rule changes must be validated, versioned, reprocessed, staged, and explicitly approved in the production workflow. This review file does not write rule changes back to the engine."]];
  ruleNote.format = { fill: PALE_AMBER, font: { color: "#7A5A00", bold: true, size: 9 }, wrapText: true, verticalAlignment: "center" };
  ruleNote.format.rowHeight = 32;
  // Move the table down visually by leaving the existing table readable below
  // the note; the note is intentionally secondary and does not alter values.
  addDetailSheet(workbook, "Update History", "Update History", "Audit trail for this package", historyMatrix, "UpdateHistory", [24, 24, 34, 18, 18, 18, 14, 16, 16, 16, 18, 18, 14, 14, 70], PALE_BLUE);
  addDetailSheet(workbook, "Error Log", "Error / Debug Log", "Technical conflicts are shown here; the real-data baseline produced no parser/hash/duplicate-key conflicts", errorsMatrix, "ErrorLog", [24, 14, 24, 55, 55, 30, 30, 42, 28, 18], PALE_RED);
  const help = addTextSheet(workbook, "Help", "Help", productionMode ? "Plain-language production guide" : "Plain-language guide for reviewing this real-data package", [
    productionMode ? "Start on Home. Use Select Data Folder first, then use Update Changed Files for routine updates." : "Start on Home. The coloured controls are clickable Excel links to the relevant page.",
    "Master Documents is the one-row-per-document view. Revisions preserves every detected revision. Transactions / Events preserves one row per extracted event.",
    "Review Flags lists 15 unusual or empty worksheet layouts that were deliberately not guessed. Review them before those worksheets can be included safely.",
    "The green OK level means the row was extracted without a row-level parser warning. Yellow REVIEW means a human decision is needed. Red CONFLICT is reserved for blocking integrity problems.",
    productionMode ? "Every update is staged first. Review Pending Update and Review Flags before choosing Approve, Hold, or Reject." : "The baseline is real extracted data, but it is not an approved master snapshot. No approval action was recorded and no source file was changed.",
    productionMode ? "Use Full Rescan after a parser or rules change, or when you want to rebuild every eligible source from scratch." : "NEXT PRODUCTION STEP (Update Changed Files / Full Rescan): open the macro-enabled package, choose the data folder, run the engine, review the staged tables, then approve, hold, or reject the exact pending run.",
    "A wrong-data report should capture the run, source workbook, worksheet, row/cell, document, revision, current value, expected value, and user note.",
    "Report Requirement / Problem should be plain English; the production engine will capture the active run and configuration context automatically.",
  ], PALE_BLUE);
  // Anchor points used by Home controls.
  const helpAnchor = help.getRange("A22:I24");
  helpAnchor.merge();
  helpAnchor.values = [[productionMode ? "ENGINE ACTIONS\n\nHome buttons run the packaged engine silently. If Windows or corporate policy blocks the engine, Excel records the failure in Error Log and approved data remains unchanged." : "ENGINE ACTIONS IN THIS REVIEW FILE\n\nThe links from Home take you to this explanation. The actual silent engine launch and Excel-to-engine refresh are deliberately not claimed until the macro-enabled package is attached and validated in Microsoft Excel."]];
  helpAnchor.format = { fill: PALE_AMBER, font: { bold: true, color: "#7A5A00", size: 10 }, wrapText: true, verticalAlignment: "center", borders: { preset: "all", style: "thin", color: GRID } };
  helpAnchor.format.rowHeight = 65;
  const reportAnchor = help.getRange("A34:I36");
  reportAnchor.merge();
  reportAnchor.values = [["REPORT REQUIREMENT / PROBLEM\n\nIn the production workbook, enter a plain-English description. The support package will include the active run, source context, parser version, and configuration version."]];
  reportAnchor.format = { fill: PALE_BLUE, font: { bold: true, color: BLUE, size: 10 }, wrapText: true, verticalAlignment: "center", borders: { preset: "all", style: "thin", color: GRID } };
  reportAnchor.format.rowHeight = 65;

  const system = workbook.worksheets.add("System Data");
  baseSheet(system, NAVY);
  titleBlock(system, "System Data", productionMode ? "Protected runtime support area" : "Read-only support view for the baseline package", 9);
  homeLink(system);
  const countRows = [
    ["Measure", "Value"],
    ["Selected workbooks", String(baseline.counts.selected_workbooks)],
    ["Processed workbooks", String(baseline.counts.processed_workbooks)],
    ["Worksheets extracted", String(baseline.counts.worksheets_extracted)],
    ["Worksheets requiring review", String(baseline.counts.worksheets_review_required)],
    ["Documents", String(baseline.counts.documents)],
    ["Revisions", String(baseline.counts.revisions)],
    ["Event records", String(baseline.counts.events)],
    ["Hyperlinks preserved", String(baseline.counts.hyperlinks)],
    ["Review flags", String(baseline.counts.review_flags)],
    ["Conflict flags", String(baseline.counts.conflict_flags)],
    ["Baseline run ID", String(baseline.run_id)],
    ["Source DATA state", "READ-ONLY / UNCHANGED"],
  ];
  addTable(system, 5, countRows, "BaselineCounts", [36, 32]);
  if (inventoryMatrix.length) {
    const invStart = 21;
    const inv = addTable(system, invStart, inventoryMatrix, "SourceInventory", [42, 34, 14, 12, 14, 68, 12, 20, 22, 18, 14, 16, 22, 22, 24, 18, 24, 24, 24, 24, 24, 32, 16, 16]);
    if (inv.rowCount > 1) system.getRange(`A${invStart + 1}:A${invStart + inv.rowCount - 1}`).format.numberFormat = "@";
  }

  // Recalculate once after all edits, then export.
  workbook.recalculate();
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const file = await SpreadsheetFile.exportXlsx(workbook);
  await file.save(outputPath);
  await patchNativeHyperlinks(outputPath, workbook);
  await patchFullDataTables(outputPath, {
    2: { matrix: docsMatrix, table: "table1.xml", lastColumn: "R" },
    3: { matrix: revisionsMatrix, table: "table2.xml", lastColumn: "K" },
    4: { matrix: eventsMatrix, table: "table3.xml", lastColumn: "Q" },
  });
  // Artifact Tool writes a large inspection sidecar for authoring QA.  It is
  // useful during construction but is not part of the user-facing package.
  try { await fs.unlink(`${outputPath}.inspect.ndjson`); } catch (_) { /* optional sidecar */ }
  console.log(JSON.stringify({ outputPath, packageMode, sheets: workbook.worksheets.items.length, documents: value("Approved Documents"), revisions: value("Approved Revisions"), events: value("Approved Transactions"), flags: value("Review Flags") }));
}

function xmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

async function patchNativeHyperlinks(filePath, workbook) {
  // Artifact Tool intentionally does not calculate HYPERLINK formulas.  Use
  // native OOXML internal hyperlinks after export so Excel displays the clean
  // button labels and follows the links without a formula error/cache value.
  const zip = await JSZip.loadAsync(await fs.readFile(filePath));
  const sheetIndex = new Map(workbook.worksheets.items.map((sheet, index) => [sheet.name, index + 1]));
  const grouped = new Map();
  for (const link of nativeLinks) {
    if (!grouped.has(link.sheetName)) grouped.set(link.sheetName, []);
    grouped.get(link.sheetName).push(link);
  }
  for (const [sheetName, links] of grouped.entries()) {
    const index = sheetIndex.get(sheetName);
    if (!index || !links.length) continue;
    const sheetPath = `xl/worksheets/sheet${index}.xml`;
    const existing = zip.file(sheetPath);
    if (!existing) continue;
    let sheetXml = await existing.async("string");
    const hyperlinkXml = links.map((link) => `    <x:hyperlink ref="${xmlEscape(link.ref)}" location="'${xmlEscape(link.targetSheet)}'!${xmlEscape(link.targetCell)}" display="${xmlEscape(link.label)}" />`).join("");
    const block = `<x:hyperlinks>${hyperlinkXml}</x:hyperlinks>`;
    if (sheetXml.includes("<x:pageMargins")) {
      sheetXml = sheetXml.replace("<x:pageMargins", `${block}<x:pageMargins`);
    } else {
      sheetXml = sheetXml.replace("</x:worksheet>", `${block}</x:worksheet>`);
    }
    zip.file(sheetPath, sheetXml);
  }
  await fs.writeFile(filePath, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

function dataCellXml(rowNumber, colNumber, value) {
  const ref = `${colName(colNumber)}${rowNumber}`;
  const text = String(value ?? "");
  if (!text) return `<x:c r="${ref}" t="str" />`;
  return `<x:c r="${ref}" t="str"><x:v>${xmlEscape(text)}</x:v></x:c>`;
}

function dataRowsXml(matrix) {
  return matrix.slice(1).map((row, offset) => {
    const rowNumber = offset + 6;
    const cells = row.map((value, index) => dataCellXml(rowNumber, index + 1, value)).join("");
    return `<x:row r="${rowNumber}">${cells}</x:row>`;
  }).join("");
}

async function patchFullDataTables(filePath, specs) {
  // The artifact authoring runtime keeps cell objects in memory, so the
  // review UI is authored with bounded slices.  Once exported, append the
  // complete real-data tables directly to OOXML.  Excel receives the same
  // styled workbook and the complete rows without the authoring heap limit.
  const zip = await JSZip.loadAsync(await fs.readFile(filePath));
  for (const [sheetIndex, spec] of Object.entries(specs)) {
    const sheetPath = `xl/worksheets/sheet${sheetIndex}.xml`;
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) continue;
    const matrix = spec.matrix;
    const dataCount = Math.max(0, matrix.length - 1);
    let xml = await sheetFile.async("string");
    const sheetDataMatch = xml.match(/<x:sheetData>([\s\S]*?)<\/x:sheetData>/);
    if (!sheetDataMatch) continue;
    const keptRows = [];
    for (const match of sheetDataMatch[1].matchAll(/<x:row r="(\d+)"[\s\S]*?<\/x:row>/g)) {
      if (Number(match[1]) <= 5) keptRows.push(match[0]);
    }
    const replacement = `<x:sheetData>${keptRows.join("")}${dataRowsXml(matrix)}</x:sheetData>`;
    xml = xml.replace(sheetDataMatch[0], replacement);
    const lastRow = dataCount + 5;
    const tableRef = `A5:${spec.lastColumn}${lastRow}`;
    const tablePath = `xl/tables/${spec.table}`;
    const tableFile = zip.file(tablePath);
    if (tableFile) {
      let tableXml = await tableFile.async("string");
      tableXml = tableXml.replace(/ref="A5:[A-Z]+\d+"/g, `ref="${tableRef}"`);
      zip.file(tablePath, tableXml);
    }
    xml = xml.replace(/sqref="A6:A\d+"/g, `sqref="A6:A${lastRow}"`);
    zip.file(sheetPath, xml);
  }
  await fs.writeFile(filePath, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

await main();
