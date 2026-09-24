import fs from "fs";
import path from "path";
import * as XLSX from "xlsx";

/** Read all sheets from an .xlsx file */
export function readExcelFile(
  filePath: string
): Record<string, Record<string, unknown>[]> {
  const abs = path.resolve(filePath);
  if (!fs.existsSync(abs)) throw new Error(`File not found: ${abs}`);
  const workbook = XLSX.readFile(abs);
  const result: Record<string, Record<string, unknown>[]> = {};
  for (const sheetName of workbook.SheetNames) {
    result[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  }
  return result;
}

/** Read a specific range from a sheet (e.g. "A1:D10") */
export function readExcelRange(
  filePath: string,
  sheetName: string,
  range: string
): unknown[][] {
  const abs = path.resolve(filePath);
  if (!fs.existsSync(abs)) throw new Error(`File not found: ${abs}`);
  const workbook = XLSX.readFile(abs);
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found`);
  const subset = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    range,
  });
  return subset;
}

/** Write rows to a sheet in an .xlsx file (creates the file if it doesn't exist) */
export function writeExcelSheet(
  filePath: string,
  sheetName: string,
  rows: Record<string, unknown>[]
): void {
  const abs = path.resolve(filePath);
  let workbook: XLSX.WorkBook;
  if (fs.existsSync(abs)) {
    workbook = XLSX.readFile(abs);
  } else {
    workbook = XLSX.utils.book_new();
  }
  const ws = XLSX.utils.json_to_sheet(rows);
  if (workbook.SheetNames.includes(sheetName)) {
    workbook.Sheets[sheetName] = ws;
  } else {
    XLSX.utils.book_append_sheet(workbook, ws, sheetName);
  }
  XLSX.writeFile(workbook, abs);
}

/** Update a single cell value (e.g. cellAddress = "B2") */
export function updateExcelCell(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  value: unknown
): void {
  const abs = path.resolve(filePath);
  if (!fs.existsSync(abs)) throw new Error(`File not found: ${abs}`);
  const workbook = XLSX.readFile(abs);
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found`);
  if (!sheet[cellAddress]) sheet[cellAddress] = {};
  (sheet[cellAddress] as XLSX.CellObject).v = value as XLSX.CellObject["v"];
  XLSX.writeFile(workbook, abs);
}

/** List all sheet names */
export function listExcelSheets(filePath: string): string[] {
  const abs = path.resolve(filePath);
  if (!fs.existsSync(abs)) throw new Error(`File not found: ${abs}`);
  return XLSX.readFile(abs).SheetNames;
}
