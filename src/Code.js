



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
