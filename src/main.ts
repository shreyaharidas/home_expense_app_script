import { currentTime } from "./utils";
import { financialSourceOwners, monthlySheet, sheetTypes, summarySheet } from "./constants";
import { getRangeData, getUntalliedSheets, findCurrentSheetRowIndex } from "./utils/sheet";

export function calculateSummary(): void {
  try {
    Logger.log("inside calculateSummary")
    const summarySheetToEdit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(summarySheet.name);

    if (!summarySheetToEdit) {
      Logger.log("Summary sheet not found.");
      return;
    }

    const sharesOwed: Record<string, { Amit: number; Shreya: number }> = {};

    const untalliedSheets = getUntalliedSheets();

    Logger.log("untalliedSheets")
    Logger.log(untalliedSheets)


    untalliedSheets.forEach((sheet) => {
      Logger.log("inside forEach")
      Logger.log("shhetname")
      Logger.log(sheet)
      sharesOwed[sheet] = { Amit: 0, Shreya: 0 };

      const data = getRangeData(sheetTypes.monthlySheet, sheet);
      Logger.log("Started loop");
      try {
        for (const row of data) {
          Logger.log("row data")
          Logger.log(row)
          const financialSource = row[monthlySheet.financialSourceCol - 1];

          Logger.log("financialSource")
          Logger.log(financialSource)

          if (typeof financialSource !== "string") {
            continue;
          }

          Logger.log("financialSourceOwners[financialSource]")
          Logger.log(financialSourceOwners[financialSource])

          const owner = financialSourceOwners[financialSource];

          if (owner === "Amit") {
            sharesOwed[sheet].Shreya += Number(row[monthlySheet.shreyasShareCol - 1]);
            Logger.log("SFS")
            Logger.log(row[monthlySheet.shreyasShareCol - 1])

          } else if (owner === "Shreya") {
            sharesOwed[sheet].Amit += Number(row[monthlySheet.amitsShareCol - 1]);
            Logger.log("AFS")
            Logger.log(row[monthlySheet.amitsShareCol - 1])
          }
        }
      }
      catch (err) {
        Logger.log(err)
        throw new Error("error is in loop")
      }
      Logger.log("Stopped loop");
      Logger.log(`Final shares for sheet: ${JSON.stringify(sharesOwed)}`);
    });

    const sheetRange = summarySheetToEdit.getDataRange();
    const summaryData = sheetRange.getValues();

    untalliedSheets.forEach((sheetName) => {
      const foundRowIndex = summaryData.findIndex((item) => findCurrentSheetRowIndex(item, sheetName));
      const editRowNumber = foundRowIndex === -1 ? summarySheetToEdit.getLastRow() + 1 : foundRowIndex + 1;

      Logger.log("edit row number");
      Logger.log(editRowNumber);

      const rowData = [
        sheetName,
        "Yes",
        sharesOwed[sheetName].Amit,
        sharesOwed[sheetName].Shreya,
        Math.max(0, sharesOwed[sheetName].Amit - sharesOwed[sheetName].Shreya),
        Math.max(0, sharesOwed[sheetName].Shreya - sharesOwed[sheetName].Amit),
        currentTime(),
      ];

      summarySheetToEdit.getRange(editRowNumber, 1, 1, rowData.length).setValues([rowData]);
    });
  }
  catch (err) {
    Logger.log("error is")
    Logger.log(err)
  }
}
