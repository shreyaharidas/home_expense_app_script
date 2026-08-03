import { defaultSplitRatio, monthlySheet, monthlySheetEditCols, otherSheetContants, summarySheet } from "./constants";
import { calculateSummary, executeCalculationProcess } from "./main";
import { getAndValidateSpreadSheet } from "./utils/sheet";

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
  const activeSheetName = activeSheet.getName();
  Logger.log(`Active sheet changed to: ${activeSheetName}`);

  if (activeSheetName === summarySheet.name) {
    try {
      Logger.log("inside try");
      executeCalculationProcess(calculateSummary);
    } catch (err) {
      Logger.log("error in onCalculate is");
      Logger.log(err);
    }
  }
}

export function handleEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  const ss = getAndValidateSpreadSheet();
  if (!ss) return;

  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  if (sheetName === "Z-Summary") return;

  const activeRange = sheet.getActiveRange();
  const rowNumber = activeRange.getRow();
  const colNumber = activeRange.getColumn();
  const entireRow = sheet.getRange(
    rowNumber,
    otherSheetContants.colStart,
    1,
    sheet.getLastColumn(),
  );
  const rowValues = entireRow.getValues()[0] as unknown[];

  if (monthlySheetEditCols.includes(colNumber)) {
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

    sheet.getRange(rowNumber, monthlySheet.amitsShareCol).setValue(amitsShare);
    sheet.getRange(rowNumber, monthlySheet.shreyasShareCol).setValue(shreyasShare);
  }

}
