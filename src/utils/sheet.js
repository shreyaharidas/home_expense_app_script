function getAndValidateSpreadSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check if the active spreadsheet is "Home Expenses"
  if (ss.getName() !== "Home Expenses") return; // Exit if it's not "Home Expenses"

  return ss;
}

function getAllSheets() {
  // Get the active spreadsheet
  const activespreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get all sheets in the spreadsheet
  const sheets = activespreadsheet.getSheets();
Logger.log("activesheet sheets")
  Logger.log(sheets)

  const sheetNames = []
  // Loop through each sheet and log its name
  sheets.forEach(function (sheet) {
    const name = sheet.getName()
    Logger.log("Sheet Name: " + name);
    sheetNames.push(name)
  });

  Logger.log("all sheet names: ")
  Logger.log(sheetNames)
  return sheetNames;
}

function findCurrentSheetRowIndex(rowData, sheetName) {
  return rowData[0] === sheetName
}


function getUntalliedSheets(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("Z-Summary");

  var dataRange = summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 2);
  var values = dataRange.getValues(); // Returns a 2D array

  // Create the map
  var resultMap = {};
  values.forEach(function(row) {
    var key = row[0];
    var value = row[1];
    if (key !== "") {
      resultMap[key] = value;
    }
  });

  Logger.log("tally result map")
  Logger.log(resultMap);
  
  let untalliedSheets=[]
  for(const sheetName in resultMap){
    if(!resultMap[sheetName]){
      untalliedSheets.push(sheetName)
    }
  }
  return untalliedSheets;
}

function getSheets(sheetType) {
  const aSS = SpreadsheetApp.getActiveSpreadsheet();
  if (sheetType === sheetTypes.summary) {
    return {
      summarySheet: aSS.getSheetByName(summarySheet.name)
    };
  }
  else if (sheetType === sheetTypes.monthlySheet) {

    const monthlySheets=[getAllSheets()].map(sheetName => sheetName!=="Z-Summary"?aSS.getSheetByName(sheetName):undefined).filter(sheet=>sheet)
    return{ monthlySheets  }
  
  }
  else if (!sheetType) {
    const monthlySheets=[getAllSheets()].map(sheetName => sheetName!=="Z-Summary"?aSS.getSheetByName(sheetName):undefined).filter(sheet=>sheet)
    return {
      monthlySheets,
      summarySheet: aSS.getSheetByName(summarySheet.name),
    }
  }
}

function loader(load) {
  const sheet = getSheets(sheetTypes.summary).summarySheet

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

function getRangeData(sheetType, editSheetName) {

  Logger.log("from get range 1")
  Logger.log(sheetType)
  Logger.log("from get range 2")
  Logger.log(editSheetName)
  if (!sheetType || !editSheetName) {
    throw new Error("Invalid arguments: 'sheetType' and 'editSheetName' are required.");
  }

  const totalCols = sheetType === sheetTypes.summary ? summarySheet.totalCols : monthlySheet.totalCols;

  if (!otherSheetContants.rowStart || !otherSheetContants.colStart) {
    throw new Error("Invalid range constants: 'rowStart' and 'colStart' are required.");
  }

 const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetToEdit = ss.getSheetByName(editSheetName);
  Logger.log("to edit")
  Logger.log(sheetToEdit)
  const lastRow = sheetToEdit.getLastRow();

  if (lastRow <= 1) {
    Logger.log("No data beyond the header row.");
    return [];
  }

  const range = sheetToEdit.getRange(otherSheetContants.rowStart, otherSheetContants.colStart, lastRow - 1, totalCols);

  const values = range.getValues().filter(row => row.some(cell => cell !== ""));
  return values.length ? values : [];
}