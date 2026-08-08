import { defaultSplitRatio, monthlySheet, monthlySheetEditCols, otherSheetContants, summarySheet } from "./constants";
import { SHEET_IDS } from "./constants/globals";
import { calculateSummary } from "./main";

export function onOpen(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();

  Logger.log(`Workbook ID: ${spreadsheet.getId()}`);
  Logger.log(`Worksheet ID: ${activeSheet.getSheetId()}`);
}

export function onSelectionChange(): void {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Logger.log(`Worksheet ID: ${activeSheet.getSheetId()}`);
}

export function onCalculate(): void {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (activeSheet.getSheetId().toString() !== SHEET_IDS.Z_SUMMARY) return;

  try {
    calculateSummary();
  } catch (err) {
    Logger.log("error in onCalculate is");
    Logger.log(err);
  }
}

export function handleEdit(e?: GoogleAppsScript.Events.SheetsOnEdit): void {
  if (!e?.range) return;

  const sheet = e.range.getSheet();
  if (
    sheet.getSheetId().toString() === SHEET_IDS.Z_SUMMARY ||
    sheet.getName() === summarySheet.name
  ) {
    return;
  }

  const colNumber = e.range.getColumn();
  if (!monthlySheetEditCols.includes(colNumber)) {
    return;
  }

  const rowNumber = e.range.getRow();
  if (rowNumber < otherSheetContants.rowStart) {
    return;
  }

  const entireRow = sheet.getRange(
    rowNumber,
    otherSheetContants.colStart,
    1,
    sheet.getLastColumn(),
  );
  const rowValues = entireRow.getValues()[0] as unknown[];

  const amount = Number(rowValues[monthlySheet.amountCol - 1]);
  if (!amount) return;

  const splitRatioCell = rowValues[monthlySheet.splitRatioCol - 1];
  const splitRatio =
    typeof splitRatioCell === "string" && splitRatioCell.trim() !== ""
      ? splitRatioCell.trim()
      : defaultSplitRatio;

  if (!splitRatio.includes(":")) {
    throw new Error("ValidationError: ratio should include ':' symbol");
  }

  const [x, y] = splitRatio.split(":").map(Number);
  if (isNaN(x) || isNaN(y) || x < 0 || y < 0 || x + y === 0) {
    throw new Error("ValidationError: Invalid ratio format");
  }

  const amitsShare = (x / (x + y)) * amount;
  const shreyasShare = (y / (x + y)) * amount;

  if (splitRatio === defaultSplitRatio) {
    sheet.getRange(rowNumber, monthlySheet.splitRatioCol).setValue(defaultSplitRatio);
  }

  sheet.getRange(rowNumber, monthlySheet.amitsShareCol, 1, 2).setValues([[amitsShare, shreyasShare]]);
}

