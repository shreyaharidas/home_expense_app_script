import { bgColors, monthlySheet, otherSheetContants, sheetTypes, summarySheet } from "../constants";

export function getAndValidateSpreadSheet(): GoogleAppsScript.Spreadsheet.Spreadsheet | undefined {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (ss.getName() !== "Home Expenses") {
    return undefined;
  }

  return ss;
}

function getAllSheets(): string[] {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = activeSpreadsheet.getSheets();

  Logger.log("activesheet sheets");
  Logger.log(sheets);

  const sheetNames: string[] = [];
  sheets.forEach((sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
    const name = sheet.getName();
    Logger.log("Sheet Name: " + name);
    sheetNames.push(name);
  });

  Logger.log("all sheet names: ");
  Logger.log(sheetNames);
  return sheetNames;
}

export function findCurrentSheetRowIndex(rowData: unknown[], sheetName: string): boolean {
  return rowData[0] === sheetName;
}

export function getRangeData(sheetType: string, editSheetName: string): Array<Array<unknown>> {
  Logger.log("from get range 1");
  Logger.log(sheetType);
  Logger.log("from get range 2");
  Logger.log(editSheetName);

  if (!sheetType || !editSheetName) {
    throw new Error("Invalid arguments: 'sheetType' and 'editSheetName' are required.");
  }

  const totalCols = sheetType === sheetTypes.summary ? summarySheet.totalCols : monthlySheet.totalCols;

  if (!otherSheetContants.rowStart || !otherSheetContants.colStart) {
    throw new Error("Invalid range constants: 'rowStart' and 'colStart' are required.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetToEdit = ss.getSheetByName(editSheetName);

  if (!sheetToEdit) {
    throw new Error(`Sheet not found: ${editSheetName}`);
  }

  Logger.log("to edit");
  Logger.log(sheetToEdit);

  const lastRow = sheetToEdit.getLastRow();
  if (lastRow <= 1) {
    Logger.log("No data beyond the header row.");
    return [];
  }

  const range = sheetToEdit.getRange(otherSheetContants.rowStart, otherSheetContants.colStart, lastRow - 1, totalCols);
  const values = range.getValues().filter((row) => row.some((cell) => cell !== ""));

  return values.length ? values : [];
}

export function getUntalliedSheets(): string[] {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheetRef = ss.getSheetByName("Z-Summary");

  if (!summarySheetRef) {
    return [];
  }

  const dataRange = summarySheetRef.getRange(2, 1, summarySheetRef.getLastRow() - 1, 2);
  const values = dataRange.getValues() as Array<Array<string | number | boolean>>;

  const resultMap: Record<string, string | number | boolean> = {};
  values.forEach((row) => {
    const key = row[0];
    const value = row[1];
    if (typeof key === "string" && key !== "") {
      resultMap[key] = value;
    }
  });

  Logger.log("tally result map");
  Logger.log(resultMap);

  const untalliedSheets: string[] = [];
  for (const sheetName in resultMap) {
    if (!resultMap[sheetName]) {
      untalliedSheets.push(sheetName);
    }
  }
  return untalliedSheets;
}

export function getSheets(sheetType?: string): {
  summarySheet?: GoogleAppsScript.Spreadsheet.Sheet;
  monthlySheets?: GoogleAppsScript.Spreadsheet.Sheet[];
} {
  const aSS = SpreadsheetApp.getActiveSpreadsheet();

  if (sheetType === sheetTypes.summary) {
    return {
      summarySheet: aSS.getSheetByName(summarySheet.name) ?? undefined,
    };
  }

  if (sheetType === sheetTypes.monthlySheet) {
    const sheetNames = getAllSheets();
    const monthlySheets = sheetNames
      .filter((sheetName) => sheetName !== "Z-Summary")
      .map((sheetName) => aSS.getSheetByName(sheetName))
      .filter((sheet): sheet is GoogleAppsScript.Spreadsheet.Sheet => Boolean(sheet));

    return { monthlySheets };
  }

  if (!sheetType) {
    const sheetNames = getAllSheets();
    const monthlySheets = sheetNames
      .filter((sheetName) => sheetName !== "Z-Summary")
      .map((sheetName) => aSS.getSheetByName(sheetName))
      .filter((sheet): sheet is GoogleAppsScript.Spreadsheet.Sheet => Boolean(sheet));

    return {
      monthlySheets,
      summarySheet: aSS.getSheetByName(summarySheet.name) ?? undefined,
    };
  }

  return {};
}

function loader(load: boolean) {
  const sheet = getSheets(sheetTypes.summary).summarySheet;

  if (!sheet) {
    throw new Error("Summary sheet not found.");
  }

  const range = sheet.getRange("K1:K4");

  if (load) {
    range.setBackground(bgColors.toEditColor);
    range.setValue("Calculating...");
  } else {
    range.setBackground("white");
    range.setValue("");
  }

  Logger.log(`Loader set: ${load ? "Loading" : "Cleared"}`);
}
