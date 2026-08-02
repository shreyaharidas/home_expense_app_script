const monthlySheet = {
  financialSourceCol: 3,
  amountCol: 5,
  splitRatioCol: 6,
  amitsShareCol: 7,
  shreyasShareCol: 8,
  totalCols: 8,
}

const monthlySheetEditCols = [monthlySheet.amountCol, monthlySheet.splitRatioCol]

const summarySheet = {
  name: "Z-Summary",
  totalCols: 7,
  loaderIndex : 12,
  loaderWidthInPx:113,
  summarySheetHeaders:{
  month: "Month",
  tallied:"Tallied",
  amitOwes: "Amit Owes",
  shreyaOwes: "Shreya Owes",
  A2S: "A to S",
  S2A: "S to A",
  editLog: "lastEdit",
},
}

const sheetTypes = {
  summary: summarySheet.name,
  monthlySheet: "monthly"
}

const otherSheetContants = {
  rowStart: 2,
  colStart: 1,
}