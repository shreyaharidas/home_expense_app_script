export interface SheetFormat {
  [colName: string]: number | string | Record<string, string>;
}

export interface SummarySheetHeaders {
  month: string;
  tallied: string;
  amitOwes: string;
  shreyaOwes: string;
  A2S: string;
  S2A: string;
  editLog: string;
}

export interface monthlySheetFormat extends SheetFormat {
  financialSourceCol: number;
  amountCol: number;
  splitRatioCol: number;
  amitsShareCol: number;
  shreyasShareCol: number;
  totalCols: number;
}

export type summarySheetFormat = SheetFormat & {
  name: string;
  totalCols: number;
  summarySheetHeaders: SummarySheetHeaders;
};

export const monthlySheet: monthlySheetFormat = {
  financialSourceCol: 3,
  amountCol: 5,
  splitRatioCol: 6,
  amitsShareCol: 7,
  shreyasShareCol: 8,
  totalCols: 8,
};

export const monthlySheetEditCols = [monthlySheet.amountCol, monthlySheet.splitRatioCol];

export const summarySheet: summarySheetFormat = {
  name: "Z-Summary",
  totalCols: 7,
  summarySheetHeaders: {
    month: "Month",
    tallied: "Tallied",
    amitOwes: "Amit Owes",
    shreyaOwes: "Shreya Owes",
    A2S: "A to S",
    S2A: "S to A",
    editLog: "lastEdit",
  },
};

export const sheetTypes = {
  summary: summarySheet.name,
  monthlySheet: "monthly",
};

export const otherSheetContants = {
  rowStart: 2,
  colStart: 1,
};