import { currentTime } from "./utils";
import { financialSourceOwners, monthlySheet, otherSheetContants, sheetTypes, summarySheet } from "./constants";
import { getRangeData, getSheets, getUntalliedSheets, findCurrentSheetRowIndex } from "./utils/sheet";

export function calculateSummary(): void {
  try {
    Logger.log("inside calculateSummary")
    const { monthlySheets, summarySheet: summarySheetToEdit } = getSheets();
    const sharesOwed: Record<string, { Amit: number; Shreya: number }> = {};
    
  const untalliedSheets=getUntalliedSheets();

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
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
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

    [...getUntalliedSheets()].forEach((sheetName) => {
      const foundRowIndex = summaryData.findIndex((item) => findCurrentSheetRowIndex(item, sheetName));
      const editRowNumber = foundRowIndex === -1 ? summarySheetToEdit.getLastRow() + 1 : foundRowIndex + 1;

      if (foundRowIndex === -1) {
        const lastRowIndex = sheetRange.getLastRow();
        summarySheetToEdit.getRange(lastRowIndex + 1, otherSheetContants.colStart).setValue(sheetName);
      }

      Logger.log("edit row number")
      Logger.log(editRowNumber)

      Object.values(summarySheet.summarySheetHeaders).forEach((header, colIndex) => {

        if(![1,2].includes(colIndex+1)){

        const cell = summarySheetToEdit.getRange(editRowNumber, colIndex+1 );
        Logger.log("row data for set")
        Logger.log(cell.getA1Notation())
        if (header === summarySheet.summarySheetHeaders.amitOwes) {
          cell.setValue(sharesOwed[sheetName].Amit);
        } else if (header === summarySheet.summarySheetHeaders.shreyaOwes) {
          cell.setValue(sharesOwed[sheetName].Shreya);
        } else if (header === summarySheet.summarySheetHeaders.S2A) {
          const diff = Number(sharesOwed[sheetName].Shreya - sharesOwed[sheetName].Amit);
          cell.setValue(diff > 0 ? diff : 0);
        } else if (header === summarySheet.summarySheetHeaders.A2S) {
          const diff = Number(sharesOwed[sheetName].Amit - sharesOwed[sheetName].Shreya);
          cell.setValue(diff > 0 ? diff : 0);
        } else if (header === summarySheet.summarySheetHeaders.editLog) {
          cell.setValue(currentTime());
        }

        const tallyTrackCell=summarySheetToEdit.getRange(editRowNumber, 2);
        tallyTrackCell.setValue("Yes")
        }
      });
    });
  }
  catch (err) {
    Logger.log("error is")
    Logger.log(err)
  }
}

export function executeCalculationProcess(calculatorCallback: () => void): void {

  Logger.log("inside execute")

  // const summarySheetToEdit = getSheets(sheetTypes.summary).summarySheet;
  const summarySheetData = getRangeData(sheetTypes.summary,summarySheet.name);

  // function setBackground(editDone) {
  //   summarySheetData.forEach((row) => {
  //     if (editedSheets.has(row[0])) {
  //       const rowRange = summarySheetToEdit.getRange(2, 1, 1, summarySheetToEdit.getLastColumn());
  //       rowRange.setBackground(editDone ? bgColors.editedColor : bgColors.toEditColor);
  //       Logger.log(`Background color for row ${row} set to ${bgColors.toEditColor}`);
  //     }
  //   });
  // }

  try {
    calculatorCallback();
  } catch (err) {
    Logger.log(err);
    throw new Error("Error in calculateSummary");
  }
}