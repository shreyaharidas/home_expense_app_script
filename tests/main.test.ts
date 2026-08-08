import { beforeEach, describe, expect, it, vi } from "vitest";

import { calculateSummary } from "../src/main";
import { monthlySheet, otherSheetContants, summarySheet } from "../src/constants";

function createSheet(name: string, rows: Array<Array<unknown>> = []) {
  const writes: Array<{ row: number; col: number; value: unknown }> = [];
  const lastRow = rows.length + 1;

  const sheet = {
    getName: vi.fn(() => name),
    getLastRow: vi.fn(() => lastRow),
    getRange: vi.fn((row: number, col: number) => ({
      setValue: (value: unknown) => {
        writes.push({ row, col, value });
      },
      getValues: () => rows,
      getA1Notation: () => `${String.fromCharCode(64 + col)}${row}`,
    })),
  } as unknown as GoogleAppsScript.Spreadsheet.Sheet & { writes: Array<{ row: number; col: number; value: unknown }> };

  (sheet as typeof sheet & { writes: Array<{ row: number; col: number; value: unknown }> }).writes = writes;

  return sheet;
}

function createSummarySheet({
  dataRows = [],
  untalliedValues = [],
}: {
  dataRows?: Array<Array<unknown>>;
  untalliedValues?: Array<Array<unknown>>;
}) {
  const writes: Array<{ row: number; col: number; value: unknown }> = [];
  let lastRow = Math.max(dataRows.length, 1) + 1;

  const sheet = {
    getName: vi.fn(() => summarySheet.name),
    getLastRow: vi.fn(() => lastRow),
    getDataRange: vi.fn(() => ({
      getValues: () => dataRows,
      getLastRow: () => lastRow,
    })),
    getRange: vi.fn((row: number, col: number, numRows?: number, numCols?: number) => ({
      setValue: (value: unknown) => {
        writes.push({ row, col, value });
        if (row > lastRow) {
          lastRow = row;
        }
      },
      setValues: (values: Array<Array<unknown>>) => {
        values.forEach((rowVals, rIdx) => {
          rowVals.forEach((val, cIdx) => {
            writes.push({ row: row + rIdx, col: col + cIdx, value: val });
          });
        });
        const endRow = row + (numRows ?? values.length) - 1;
        if (endRow > lastRow) {
          lastRow = endRow;
        }
      },
      getValues: () => {
        if (row === 2 && col === 1 && numRows && numCols) {
          return untalliedValues;
        }

        return [];
      },
      getA1Notation: () => `${String.fromCharCode(64 + col)}${row}`,
    })),
  } as unknown as GoogleAppsScript.Spreadsheet.Sheet & { writes: Array<{ row: number; col: number; value: unknown }> };

  (sheet as typeof sheet & { writes: Array<{ row: number; col: number; value: unknown }> }).writes = writes;

  return sheet;
}

describe("main calculation flow", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.stubGlobal("Logger", { log: vi.fn() });
  });

  it("aggregates shares across untallied sheets and appends summary rows for each sheet", () => {
    const janRows = [
      ["Date", "Description", "Financial Source", "Amount", "Split", "Ratio", "Amit Share", "Shreya Share"],
      ["", "", "Amit's Wallet", "", "", "", "", 20],
      ["", "", "Shreya's Wallet", "", "", "", 15, ""],
    ];
    const febRows = [
      ["Date", "Description", "Financial Source", "Amount", "Split", "Ratio", "Amit Share", "Shreya Share"],
      ["", "", "Shreya's Wallet", "", "", "", 10, ""],
    ];

    const summarySheetRef = createSummarySheet({
      dataRows: [["Month", "Tallied", "Amit Owes", "Shreya Owes", "A to S", "S to A", "lastEdit"]],
      untalliedValues: [["Jan", false], ["Feb", false]],
    });
    const janSheet = createSheet("Jan", janRows);
    const febSheet = createSheet("Feb", febRows);

    const aSS = {
      getSheets: vi.fn(() => [janSheet, febSheet, summarySheetRef]),
      getSheetByName: vi.fn((name: string) => {
        if (name === "Jan") return janSheet;
        if (name === "Feb") return febSheet;
        if (name === summarySheet.name) return summarySheetRef;
        return undefined;
      }),
    } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;

    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => aSS),
    });

    calculateSummary();

    const writes = (summarySheetRef as typeof summarySheetRef & { writes: Array<{ row: number; col: number; value: unknown }> }).writes;
    expect(writes).toEqual(
      expect.arrayContaining([
        { row: 3, col: 1, value: "Jan" },
        { row: 4, col: 1, value: "Feb" },
        { row: 3, col: 3, value: 15 },
        { row: 3, col: 4, value: 20 },
        { row: 3, col: 5, value: 0 },
        { row: 3, col: 6, value: 5 },
        { row: 4, col: 3, value: 10 },
        { row: 4, col: 5, value: 10 },
      ]),
    );
  });

  it("updates an existing summary row instead of creating a duplicate", () => {
    const summarySheetRef = createSummarySheet({
      dataRows: [["Jan", false, 0, 0, 0, 0, "old"]],
      untalliedValues: [["Jan", false]],
    });
    const janSheet = createSheet("Jan", [
      ["Date", "Description", "Financial Source", "Amount", "Split", "Ratio", "Amit Share", "Shreya Share"],
      ["", "", "Amit's Wallet", "", "", "", "", 10],
    ]);

    const aSS = {
      getSheets: vi.fn(() => [janSheet, summarySheetRef]),
      getSheetByName: vi.fn((name: string) => {
        if (name === "Jan") return janSheet;
        if (name === summarySheet.name) return summarySheetRef;
        return undefined;
      }),
    } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;

    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => aSS),
    });

    calculateSummary();

    const writes = (summarySheetRef as typeof summarySheetRef & { writes: Array<{ row: number; col: number; value: unknown }> }).writes;
    expect(writes).toEqual(
      expect.arrayContaining([
        { row: 1, col: 3, value: 0 },
        { row: 1, col: 4, value: 10 },
      ]),
    );
    expect(writes).not.toEqual(expect.arrayContaining([{ row: 3, col: 1, value: "Jan" }]));
  });

  it("handles rows with missing or non-string financial sources without throwing", () => {
    const summarySheetRef = createSummarySheet({
      dataRows: [["Month", "Tallied", "Amit Owes", "Shreya Owes", "A to S", "S to A", "lastEdit"]],
      untalliedValues: [["Mar", false]],
    });
    const marSheet = createSheet("Mar", [
      ["Date", "Description", "Financial Source", "Amount", "Split", "Ratio", "Amit Share", "Shreya Share"],
      ["", "", 123, "", "", "", "", ""],
      ["", "", "Unknown Source", "", "", "", 5, 1],
      ["", "", "Shreya's Wallet", "", "", "", "", ""],
    ]);

    const aSS = {
      getSheets: vi.fn(() => [marSheet, summarySheetRef]),
      getSheetByName: vi.fn((name: string) => {
        if (name === "Mar") return marSheet;
        if (name === summarySheet.name) return summarySheetRef;
        return undefined;
      }),
    } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;

    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => aSS),
    });

    expect(() => calculateSummary()).not.toThrow();

    const writes = (summarySheetRef as typeof summarySheetRef & { writes: Array<{ row: number; col: number; value: unknown }> }).writes;
    expect(writes).toEqual(expect.arrayContaining([{ row: 3, col: 3, value: 0 }, { row: 3, col: 4, value: 0 }]));
  });


});
