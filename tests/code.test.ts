import { beforeEach, describe, expect, it, vi } from "vitest";

import { handleEdit, onCalculate, onOpen } from "../src/Code";
import { defaultSplitRatio, monthlySheet, otherSheetContants, summarySheet } from "../src/constants";
import { SHEET_IDS, WORKBOOK_ID } from "../src/constants/globals";
import * as mainModule from "../src/main";

type RecordedWrite = { row: number; col: number; value: unknown };

function createSheet({
  name,
  rowValues = [],
  activeRow = 2,
  activeCol = monthlySheet.amountCol,
  lastColumn = monthlySheet.totalCols,
  sheetId = 101,
}: {
  name: string;
  rowValues?: Array<unknown>;
  activeRow?: number;
  activeCol?: number;
  lastColumn?: number;
  sheetId?: number | string;
}) {
  const writes: RecordedWrite[] = [];

  const rangeObj = {
    getRow: () => activeRow,
    getColumn: () => activeCol,
    getSheet: () => sheet,
    setValue: (value: unknown) => {
      writes.push({ row: activeRow, col: activeCol, value });
    },
    getValues: () => [rowValues],
  };

  const sheet = {
    getName: vi.fn(() => name),
    getSheetId: vi.fn(() => sheetId),
    getLastColumn: vi.fn(() => lastColumn),
    getActiveRange: vi.fn(() => rangeObj),
    getRange: vi.fn((row: number, col: number, numRows?: number, numCols?: number) => ({
      getRow: () => row,
      getColumn: () => col,
      getSheet: () => sheet,
      setValue: (value: unknown) => {
        writes.push({ row, col, value });
      },
      setValues: (values: Array<Array<unknown>>) => {
        values.forEach((rowVals, rIdx) => {
          rowVals.forEach((val, cIdx) => {
            writes.push({ row: row + rIdx, col: col + cIdx, value: val });
          });
        });
      },
      getValues: () => [rowValues],
    })),
  } as unknown as GoogleAppsScript.Spreadsheet.Sheet & { writes: RecordedWrite[] };

  (sheet as typeof sheet & { writes: RecordedWrite[] }).writes = writes;
  return sheet;
}

function createSpreadsheet(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  return {
    getId: vi.fn(() => WORKBOOK_ID),
    getName: vi.fn(() => "Home Expenses"),
    getActiveSheet: vi.fn(() => sheet),
  } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
}

function createEditEvent(sheet: GoogleAppsScript.Spreadsheet.Sheet, range?: GoogleAppsScript.Spreadsheet.Range) {
  return {
    range: range ?? sheet.getActiveRange(),
    source: createSpreadsheet(sheet),
  } as unknown as GoogleAppsScript.Events.SheetsOnEdit;
}

describe("Code.ts event handlers", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    vi.stubGlobal("Logger", { log: vi.fn() });
    vi.stubGlobal("console", { log: vi.fn() });
  });

  it("logs the workbook id and current sheet id when the spreadsheet opens", () => {
    const sheet = createSheet({ name: "January", sheetId: 42 });
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => ({
        getId: vi.fn(() => "workbook-123"),
        getActiveSheet: vi.fn(() => sheet),
      })),
    });

    const logSpy = vi.spyOn((globalThis as any).Logger, "log").mockImplementation(() => undefined);

    onOpen();

    expect(logSpy).toHaveBeenCalledWith("Workbook ID: workbook-123");
    expect(logSpy).toHaveBeenCalledWith("Worksheet ID: 42");
  });

  it("does nothing when onCalculate is run on a non-summary sheet", () => {
    const sheet = createSheet({ name: "January", sheetId: 101 });
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => ({
        getActiveSheet: vi.fn(() => sheet),
      })),
    });

    const calculateSpy = vi.spyOn(mainModule, "calculateSummary").mockImplementation(() => undefined);

    onCalculate();

    expect(calculateSpy).not.toHaveBeenCalled();
  });

  it("runs the calculation process when onCalculate is triggered from the summary sheet", () => {
    const sheet = createSheet({ name: summarySheet.name, sheetId: Number(SHEET_IDS.Z_SUMMARY) });
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => ({
        getActiveSheet: vi.fn(() => sheet),
      })),
    });

    const calculateSpy = vi.spyOn(mainModule, "calculateSummary").mockImplementation(() => undefined);

    onCalculate();

    expect(calculateSpy).toHaveBeenCalledTimes(1);
  });

  it("swallows calculation errors in onCalculate without throwing", () => {
    const sheet = createSheet({ name: summarySheet.name, sheetId: Number(SHEET_IDS.Z_SUMMARY) });
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => ({
        getActiveSheet: vi.fn(() => sheet),
      })),
    });

    vi.spyOn(mainModule, "calculateSummary").mockImplementation(() => {
      throw new Error("boom");
    });

    expect(() => onCalculate()).not.toThrow();
  });

  it("returns early when handleEdit is called without a valid event or range", () => {
    expect(() => handleEdit()).not.toThrow();
    expect(() => handleEdit({} as GoogleAppsScript.Events.SheetsOnEdit)).not.toThrow();
  });

  it("returns early when the edited sheet is the summary sheet", () => {
    const sheet = createSheet({ name: summarySheet.name, sheetId: Number(SHEET_IDS.Z_SUMMARY) });

    expect(() => handleEdit(createEditEvent(sheet))).not.toThrow();
    expect((sheet as typeof sheet & { writes: RecordedWrite[] }).writes).toEqual([]);
  });

  it("does not calculate shares when the edited column is not tracked", () => {
    const sheet = createSheet({
      name: "January",
      rowValues: ["", "", "", "", 100, "3:2", "", ""],
      activeCol: 1,
    });

    handleEdit(createEditEvent(sheet));

    expect((sheet as typeof sheet & { writes: RecordedWrite[] }).writes).toEqual([]);
  });

  it("returns early when the edited row is before rowStart", () => {
    const sheet = createSheet({
      name: "January",
      rowValues: ["", "", "", "", 100, "3:2", "", ""],
      activeRow: 1,
      activeCol: monthlySheet.amountCol,
    });

    handleEdit(createEditEvent(sheet));

    expect((sheet as typeof sheet & { writes: RecordedWrite[] }).writes).toEqual([]);
  });

  it("uses the default split ratio when the ratio cell is empty and writes the calculated shares", () => {
    const sheet = createSheet({
      name: "January",
      rowValues: ["", "", "", "", 100, "", "", ""],
      activeCol: monthlySheet.amountCol,
    });

    handleEdit(createEditEvent(sheet));

    expect((sheet as typeof sheet & { writes: RecordedWrite[] }).writes).toEqual(
      expect.arrayContaining([
        { row: 2, col: monthlySheet.splitRatioCol, value: defaultSplitRatio },
        { row: 2, col: monthlySheet.amitsShareCol, value: 60 },
        { row: 2, col: monthlySheet.shreyasShareCol, value: 40 },
      ]),
    );
  });

  it("returns early when the amount value is zero or blank", () => {
    const sheet = createSheet({
      name: "January",
      rowValues: ["", "", "", "", 0, "3:2", "", ""],
      activeCol: monthlySheet.amountCol,
    });

    handleEdit(createEditEvent(sheet));

    expect((sheet as typeof sheet & { writes: RecordedWrite[] }).writes).toEqual([]);
  });

  it.each(["3 2", "a:b", "-1:2", "0:0"])("throws a validation error for invalid split ratio %s", (ratio) => {
    const sheet = createSheet({
      name: "January",
      rowValues: ["", "", "", "", 100, ratio, "", ""],
      activeCol: monthlySheet.splitRatioCol,
    });

    expect(() => handleEdit(createEditEvent(sheet))).toThrow(/ValidationError/);
  });

  it("computes exact shares for a standard ratio and writes them to the correct columns", () => {
    const sheet = createSheet({
      name: "January",
      rowValues: ["", "", "", "", 100, "1:1", "", ""],
      activeCol: monthlySheet.splitRatioCol,
    });

    handleEdit(createEditEvent(sheet));

    expect((sheet as typeof sheet & { writes: RecordedWrite[] }).writes).toEqual(
      expect.arrayContaining([
        { row: 2, col: monthlySheet.amitsShareCol, value: 50 },
        { row: 2, col: monthlySheet.shreyasShareCol, value: 50 },
      ]),
    );
  });
});
