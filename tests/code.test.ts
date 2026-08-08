import { beforeEach, describe, expect, it, vi } from "vitest";

import { handleEdit, onCalculate, onOpen, onSelectionChange } from "../src/Code";
import { defaultSplitRatio, monthlySheet, otherSheetContants, summarySheet } from "../src/constants";
import { SHEET_IDS, WORKBOOK_ID } from "../src/constants/globals";
import * as mainModule from "../src/main";
import { setGlobalIds } from "../src/utils/globalIdUtils";

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

  const sheet = {
    getName: vi.fn(() => name),
    getSheetId: vi.fn(() => sheetId),
    getLastColumn: vi.fn(() => lastColumn),
    getActiveRange: vi.fn(() => ({
      getRow: () => activeRow,
      getColumn: () => activeCol,
    })),
    getRange: vi.fn((row: number, col: number) => ({
      setValue: (value: unknown) => {
        writes.push({ row, col, value });
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

describe("Code.ts event handlers", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    vi.stubGlobal("Logger", { log: vi.fn() });
    vi.stubGlobal("console", { log: vi.fn() });
    setGlobalIds({ type: "workbookId", id: WORKBOOK_ID });
    setGlobalIds({ type: "activeWorkSheetId", id: "101" });
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

  it("logs the active sheet id when the selection changes", () => {
    const sheet = createSheet({ name: "January", sheetId: 77 });
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => ({
        getActiveSheet: vi.fn(() => sheet),
      })),
    });

    const logSpy = vi.spyOn((globalThis as any).Logger, "log").mockImplementation(() => undefined);

    onSelectionChange();

    expect(logSpy).toHaveBeenCalledWith("Worksheet ID: 77");
  });

  it("does nothing when onCalculate is run on a non-summary sheet", () => {
    setGlobalIds({ type: "activeWorkSheetId", id: "101" });
    const sheet = createSheet({ name: "January" });
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => ({
        getActiveSheet: vi.fn(() => sheet),
      })),
    });

    const executeSpy = vi.spyOn(mainModule, "executeCalculationProcess").mockImplementation(() => undefined);

    onCalculate();

    expect(executeSpy).not.toHaveBeenCalled();
  });

  it("runs the calculation process when onCalculate is triggered from the summary sheet", () => {
    setGlobalIds({ type: "activeWorkSheetId", id: SHEET_IDS.Z_SUMMARY });
    const sheet = createSheet({ name: summarySheet.name, sheetId: Number(SHEET_IDS.Z_SUMMARY) });
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => ({
        getActiveSheet: vi.fn(() => sheet),
      })),
    });

    const executeSpy = vi.spyOn(mainModule, "executeCalculationProcess").mockImplementation(() => undefined);

    onCalculate();

    expect(executeSpy).toHaveBeenCalledTimes(1);
    expect(executeSpy).toHaveBeenCalledWith(expect.any(Function));
  });

  it("swallows calculation errors in onCalculate without throwing", () => {
    setGlobalIds({ type: "activeWorkSheetId", id: SHEET_IDS.Z_SUMMARY });
    const sheet = createSheet({ name: summarySheet.name, sheetId: Number(SHEET_IDS.Z_SUMMARY) });
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => ({
        getActiveSheet: vi.fn(() => sheet),
      })),
    });

    vi.spyOn(mainModule, "executeCalculationProcess").mockImplementation(() => {
      throw new Error("boom");
    });

    expect(() => onCalculate()).not.toThrow();
  });

  it("returns early when handleEdit is called for a spreadsheet that is not validated", () => {
    setGlobalIds({ type: "workbookId", id: "other-workbook-id" });

    const sheet = createSheet({ name: "January" });
    const spreadsheet = createSpreadsheet(sheet);
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => spreadsheet),
    });

    expect(() => handleEdit({} as GoogleAppsScript.Events.SheetsOnEdit)).not.toThrow();
    expect((sheet as typeof sheet & { writes: RecordedWrite[] }).writes).toEqual([]);
  });

  it("returns early when the edited sheet is the summary sheet", () => {
    setGlobalIds({ type: "activeWorkSheetId", id: SHEET_IDS.Z_SUMMARY });
    const sheet = createSheet({ name: summarySheet.name, sheetId: Number(SHEET_IDS.Z_SUMMARY) });
    const spreadsheet = createSpreadsheet(sheet);
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => spreadsheet),
    });

    expect(() => handleEdit({} as GoogleAppsScript.Events.SheetsOnEdit)).not.toThrow();
    expect((sheet as typeof sheet & { writes: RecordedWrite[] }).writes).toEqual([]);
  });

  it("does not calculate shares when the edited column is not tracked", () => {
    const sheet = createSheet({
      name: "January",
      rowValues: ["", "", "", "", 100, "3:2", "", ""],
      activeCol: 1,
    });
    const spreadsheet = createSpreadsheet(sheet);
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => spreadsheet),
    });

    handleEdit({} as GoogleAppsScript.Events.SheetsOnEdit);

    expect((sheet as typeof sheet & { writes: RecordedWrite[] }).writes).toEqual([]);
  });

  it("uses the default split ratio when the ratio cell is empty and writes the calculated shares", () => {
    const sheet = createSheet({
      name: "January",
      rowValues: ["", "", "", "", 100, "", "", ""],
      activeCol: monthlySheet.amountCol,
    });
    const spreadsheet = createSpreadsheet(sheet);
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => spreadsheet),
    });

    handleEdit({} as GoogleAppsScript.Events.SheetsOnEdit);

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
    const spreadsheet = createSpreadsheet(sheet);
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => spreadsheet),
    });

    handleEdit({} as GoogleAppsScript.Events.SheetsOnEdit);

    expect((sheet as typeof sheet & { writes: RecordedWrite[] }).writes).toEqual([]);
  });

  it.each(["3 2", "a:b", "-1:2", "0:0"])("throws a validation error for invalid split ratio %s", (ratio) => {
    const sheet = createSheet({
      name: "January",
      rowValues: ["", "", "", "", 100, ratio, "", ""],
      activeCol: monthlySheet.splitRatioCol,
    });
    const spreadsheet = createSpreadsheet(sheet);
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => spreadsheet),
    });

    expect(() => handleEdit({} as GoogleAppsScript.Events.SheetsOnEdit)).toThrow(/ValidationError/);
  });

  it("computes exact shares for a standard ratio and writes them to the correct columns", () => {
    const sheet = createSheet({
      name: "January",
      rowValues: ["", "", "", "", 100, "1:1", "", ""],
      activeCol: monthlySheet.splitRatioCol,
    });
    const spreadsheet = createSpreadsheet(sheet);
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => spreadsheet),
    });

    handleEdit({} as GoogleAppsScript.Events.SheetsOnEdit);

    expect((sheet as typeof sheet & { writes: RecordedWrite[] }).writes).toEqual(
      expect.arrayContaining([
        { row: 2, col: monthlySheet.amitsShareCol, value: 50 },
        { row: 2, col: monthlySheet.shreyasShareCol, value: 50 },
      ]),
    );
  });
});
