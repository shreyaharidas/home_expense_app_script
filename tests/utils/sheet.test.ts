import { beforeEach, describe, expect, it, vi } from "vitest";

import {
  getAndValidateSpreadSheet,
  findCurrentSheetRowIndex,
  getRangeData,
  getUntalliedSheets,
  getSheets,
} from "../../src/utils/sheet";
import { sheetTypes, summarySheet, otherSheetContants } from "../../src/constants";
import { WORKBOOK_ID } from "../../src/constants/globals";
import { setGlobalIds } from "../../src/utils/globalIdUtils";

describe("sheet utilities", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.stubGlobal("Logger", { log: vi.fn() });
  });

  describe("getAndValidateSpreadSheet", () => {
    const getName = vi.fn();
    const spreadsheet = { getName } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;

    beforeEach(() => {
      vi.stubGlobal("SpreadsheetApp", {
        getActiveSpreadsheet: vi.fn(() => spreadsheet),
      });
    });

    it("returns the active spreadsheet when it is the Home Expenses workbook", () => {
      setGlobalIds({ type: "workbookId", id: WORKBOOK_ID });

      expect(getAndValidateSpreadSheet()).toBe(spreadsheet);
    });

    it("returns undefined for a different workbook", () => {
      setGlobalIds({ type: "workbookId", id: "other-workbook-id" });

      expect(getAndValidateSpreadSheet()).toBeUndefined();
    });
  });

  describe("findCurrentSheetRowIndex", () => {
    it("returns true when rowData[0] matches the sheet name", () => {
      expect(findCurrentSheetRowIndex(["Jan", 1, 2], "Jan")).toBe(true);
    });

    it("returns false when rowData[0] does not match the sheet name", () => {
      expect(findCurrentSheetRowIndex(["Feb", 1, 2], "Mar")).toBe(false);
    });
  });

  describe("getRangeData", () => {
    it("throws when required args are missing", () => {
      expect(() => getRangeData("", "")).toThrow();
    });

    it("throws when the sheet to edit is not found", () => {
      const aSS = { getSheetByName: vi.fn(() => undefined) } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
      vi.stubGlobal("SpreadsheetApp", { getActiveSpreadsheet: vi.fn(() => aSS) });

      expect(() => getRangeData(sheetTypes.monthlySheet, "Unknown")).toThrow(/Sheet not found/);
    });

    it("returns empty array when sheet has only header (lastRow <= 1)", () => {
      const sheetToEdit = { getLastRow: vi.fn(() => 1) } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
      const aSS = { getSheetByName: vi.fn(() => sheetToEdit) } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
      vi.stubGlobal("SpreadsheetApp", { getActiveSpreadsheet: vi.fn(() => aSS) });

      expect(getRangeData(sheetTypes.monthlySheet, "Sheet1")).toEqual([]);
    });

    it("throws when range constants are missing or falsy", () => {
      const origRow = otherSheetContants.rowStart;
      const origCol = otherSheetContants.colStart;
      (otherSheetContants as any).rowStart = 0;
      (otherSheetContants as any).colStart = 0;

      const sheetToEdit = { getLastRow: vi.fn(() => 2) } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
      const aSS = { getSheetByName: vi.fn(() => sheetToEdit) } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
      vi.stubGlobal("SpreadsheetApp", { getActiveSpreadsheet: vi.fn(() => aSS) });

      expect(() => getRangeData(sheetTypes.monthlySheet, "Sheet1")).toThrow(/Invalid range constants/);

      (otherSheetContants as any).rowStart = origRow;
      (otherSheetContants as any).colStart = origCol;
    });

    it("returns [] when all rows in the range are empty strings", () => {
      const values = [["", ""], ["", ""]];
      const range = { getValues: vi.fn(() => values) } as unknown as GoogleAppsScript.Spreadsheet.Range;
      const sheetToEdit = { getLastRow: vi.fn(() => 3), getRange: vi.fn(() => range) } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
      const aSS = { getSheetByName: vi.fn(() => sheetToEdit) } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
      vi.stubGlobal("SpreadsheetApp", { getActiveSpreadsheet: vi.fn(() => aSS) });

      expect(getRangeData(sheetTypes.monthlySheet, "Sheet1")).toEqual([]);
    });

    it("includes rows with non-empty non-string values (e.g., 0 or false)", () => {
      const values = [["", ""], [0, false]];
      const range = { getValues: vi.fn(() => values) } as unknown as GoogleAppsScript.Spreadsheet.Range;
      const sheetToEdit = { getLastRow: vi.fn(() => 3), getRange: vi.fn(() => range) } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
      const aSS = { getSheetByName: vi.fn(() => sheetToEdit) } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
      vi.stubGlobal("SpreadsheetApp", { getActiveSpreadsheet: vi.fn(() => aSS) });

      const res = getRangeData(sheetTypes.monthlySheet, "Sheet1");
      expect(res).toEqual([values[1]]);
    });

    it("uses summary sheet totalCols when sheetType is summary", () => {
      const range = { getValues: vi.fn(() => [["X"]]) } as unknown as GoogleAppsScript.Spreadsheet.Range;
      const getRange = vi.fn(() => range);
      const sheetToEdit = { getLastRow: vi.fn(() => 2), getRange } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
      const aSS = { getSheetByName: vi.fn(() => sheetToEdit) } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
      vi.stubGlobal("SpreadsheetApp", { getActiveSpreadsheet: vi.fn(() => aSS) });

      getRangeData(sheetTypes.summary, summarySheet.name);

      expect(getRange).toHaveBeenCalledWith(otherSheetContants.rowStart, otherSheetContants.colStart, 1, summarySheet.totalCols);
    });

    it("returns filtered non-empty rows from the sheet range", () => {
      const values = [
        ["Jan", "", ""],
        ["", "", ""],
        ["Feb", 123, ""],
      ];

      const range = { getValues: vi.fn(() => values) } as unknown as GoogleAppsScript.Spreadsheet.Range;
      const sheetToEdit = { getLastRow: vi.fn(() => 4), getRange: vi.fn(() => range) } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
      const aSS = { getSheetByName: vi.fn(() => sheetToEdit) } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
      vi.stubGlobal("SpreadsheetApp", { getActiveSpreadsheet: vi.fn(() => aSS) });

      const result = getRangeData(sheetTypes.monthlySheet, "Sheet1");
      expect(result).toEqual([values[0], values[2]]);
    });
  });

  describe("getUntalliedSheets", () => {
    it("returns empty array when summary sheet not present", () => {
      const aSS = { getSheetByName: vi.fn(() => undefined) } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
      vi.stubGlobal("SpreadsheetApp", { getActiveSpreadsheet: vi.fn(() => aSS) });

      expect(getUntalliedSheets()).toEqual([]);
    });

    it("returns sheet names that are untallied (falsy value)", () => {
      const values = [["Jan", false], ["Feb", true], ["Mar", ""]];
      const summarySheetRef = {
        getLastRow: vi.fn(() => 4),
        getRange: vi.fn(() => ({ getValues: () => values })),
      } as unknown as GoogleAppsScript.Spreadsheet.Sheet;

      const aSS = { getSheetByName: vi.fn((name: string) => (name === "Z-Summary" ? summarySheetRef : undefined)) } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
      vi.stubGlobal("SpreadsheetApp", { getActiveSpreadsheet: vi.fn(() => aSS) });

      const untallied = getUntalliedSheets();
      expect(untallied).toEqual(["Jan", "Mar"]);
    });

    it("ignores non-string keys and only returns string-named untallied sheets", () => {
      const values = [[1, false], ["Feb", false], [null, false]];
      const summarySheetRef = {
        getLastRow: vi.fn(() => 4),
        getRange: vi.fn(() => ({ getValues: () => values })),
      } as unknown as GoogleAppsScript.Spreadsheet.Sheet;

      const aSS = { getSheetByName: vi.fn((name: string) => (name === "Z-Summary" ? summarySheetRef : undefined)) } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
      vi.stubGlobal("SpreadsheetApp", { getActiveSpreadsheet: vi.fn(() => aSS) });

      const untallied = getUntalliedSheets();
      expect(untallied).toEqual(["Feb"]);
    });
  });

  describe("getSheets", () => {
    it("returns empty object for unknown sheetType", () => {
      const aSS = { getSheets: vi.fn(() => []), getSheetByName: vi.fn(() => undefined) } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
      vi.stubGlobal("SpreadsheetApp", { getActiveSpreadsheet: vi.fn(() => aSS) });

      const res = getSheets("unknown-type");
      expect(res).toEqual({});
    });

    it("returns the summary sheet when requested", () => {
      const summary = {} as GoogleAppsScript.Spreadsheet.Sheet;
      const aSS = { getSheetByName: vi.fn((name: string) => (name === summarySheet.name ? summary : undefined)) } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
      vi.stubGlobal("SpreadsheetApp", { getActiveSpreadsheet: vi.fn(() => aSS) });

      const res = getSheets(sheetTypes.summary);
      expect(res.summarySheet).toBe(summary);
      expect(res.monthlySheets).toBeUndefined();
    });

    it("returns monthly sheets when requested", () => {
      const sheetA = { getName: vi.fn(() => "Jan") } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
      const sheetB = { getName: vi.fn(() => "Z-Summary") } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
      const aSS = {
        getSheets: vi.fn(() => [sheetA, sheetB]),
        getSheetByName: vi.fn((name: string) => (name === "Jan" ? sheetA : name === "Z-Summary" ? sheetB : undefined)),
      } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;

      vi.stubGlobal("SpreadsheetApp", { getActiveSpreadsheet: vi.fn(() => aSS) });

      const res = getSheets(sheetTypes.monthlySheet);
      expect(res.monthlySheets).toBeDefined();
      expect(res.monthlySheets?.length).toBe(1);
      expect(res.monthlySheets?.[0]).toBe(sheetA);
    });

    it("returns both summary and monthly sheets when no type provided", () => {
      const sheetA = { getName: vi.fn(() => "Jan") } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
      const summaryRef = {} as GoogleAppsScript.Spreadsheet.Sheet;
      const aSS = {
        getSheets: vi.fn(() => [sheetA]),
        getSheetByName: vi.fn((name: string) => (name === summarySheet.name ? summaryRef : name === "Jan" ? sheetA : undefined)),
      } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;

      vi.stubGlobal("SpreadsheetApp", { getActiveSpreadsheet: vi.fn(() => aSS) });

      const res = getSheets();
      expect(res.monthlySheets).toBeDefined();
      expect(res.summarySheet).toBe(summaryRef);
    });
  });
});
