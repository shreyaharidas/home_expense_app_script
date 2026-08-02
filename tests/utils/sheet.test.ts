import { beforeEach, describe, expect, it, vi } from "vitest";

import { getAndValidateSpreadSheet } from "../../src/utils/sheet";

describe("getAndValidateSpreadSheet", () => {
  const getName = vi.fn();
  const spreadsheet = { getName } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;

  beforeEach(() => {
    vi.clearAllMocks();
    vi.stubGlobal("SpreadsheetApp", {
      getActiveSpreadsheet: vi.fn(() => spreadsheet),
    });
  });

  it("returns the active spreadsheet when it is the Home Expenses workbook", () => {
    getName.mockReturnValue("Home Expenses");

    expect(getAndValidateSpreadSheet()).toBe(spreadsheet);
  });

  it("returns undefined for a different workbook", () => {
    getName.mockReturnValue("Personal Budget");

    expect(getAndValidateSpreadSheet()).toBeUndefined();
  });
});
