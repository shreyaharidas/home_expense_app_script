

function calculateSummary() {
  try {
    Logger.log ("testing if code is pushed")
    Logger.log("inside calculateSummary")
    const { monthlySheets, summarySheet: summarySheetToEdit } = getSheets();
    const sharesOwed = {};
    
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
          Logger.log("financialSourceOwners[financialSource]")
          Logger.log(financialSourceOwners[financialSource])

          if (financialSourceOwners[financialSource] === "Amit") {
            sharesOwed[sheet].Shreya += Number(row[monthlySheet.shreyasShareCol - 1]);
             Logger.log("SFS")
            Logger.log(row[monthlySheet.shreyasShareCol - 1])
           
          } else if (financialSourceOwners[financialSource] === "Shreya") {
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
      const editRowNumber = summaryData.findIndex((item) => findCurrentSheetRowIndex(item, sheetName)) + 1;

      if (editRowNumber === -1) {
        const lastRowIndex = sheetRange.getLastRow();
        summarySheetToEdit.getRange(lastRowIndex + 1, otherSheetContants.colStart).setValue(sheetName);
      }

      Logger.log("edit row number")
      Logger.log(editRowNumber)

      Object.values(summarySheet.summarySheetHeaders).forEach((header, colIndex) => {

        if(![1,2].includes(colIndex+1)){

        const cell = summarySheetToEdit.getRange(editRowNumber, colIndex+1 );
        Logger.log("row data for set")
        Logger.log(Array.from(cell))
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

function executeCalculationProcess(calculatorCallback) {

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

  // loader(true);
  try {
    calculatorCallback();
  } catch (err) {
    Logger.log(err)
    throw new Error("Error in calculateSummary");
  }
  // loader(false);
}

function onCalculate() {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeSheetName = activeSheet.getName();
  Logger.log(`Active sheet changed to: ${activeSheetName}`);

  if (activeSheetName === summarySheet.name) {
    try {
      Logger.log("inside try")
      executeCalculationProcess(calculateSummary);
    }
    catch (err) {
      Logger.log("error in onCalculate is")
      Logger.log(err)
    }
  }
}


function handleEdit(e) {
  const ss = getAndValidateSpreadSheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  if (sheetName === "Z-Summary") return;

  const activeRange = sheet.getActiveRange();
  const rowNumber = activeRange.getRow();
  const colNumber = activeRange.getColumn();
  const entireRow = sheet.getRange(rowNumber, otherSheetContants.colStart, 1, sheet.getLastColumn());
  const rowValues = entireRow.getValues()[0];

  if (monthlySheetEditCols.includes(colNumber)) {
    const amount = Number(rowValues[monthlySheet.amountCol - 1]);
    if (!amount) return;

    let splitRatio = rowValues[monthlySheet.splitRatioCol - 1] || defaultSplitRatio;
    if (!splitRatio.includes(":")) throw new Error("ValidationError: ratio should include ':' symbol");

    const [x, y] = splitRatio.split(":").map(Number);
    if (isNaN(x) || isNaN(y) || x < 0 || y < 0 || x + y === 0) {
      throw new Error("ValidationError: Invalid ratio format");
    }

    const amitsShare = (x / (x + y)) * amount;
    const shreyasShare = (y / (x + y)) * amount;

if(splitRatio===defaultSplitRatio){
  sheet.getRange(rowNumber, monthlySheet.splitRatioCol).setValue(defaultSplitRatio);
}
    sheet.getRange(rowNumber, monthlySheet.amitsShareCol).setValue(amitsShare);
    sheet.getRange(rowNumber, monthlySheet.shreyasShareCol).setValue(shreyasShare);
  }
}
