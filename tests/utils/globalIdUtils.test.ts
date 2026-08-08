import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { getGlobalIds, setGlobalIds } from "../../src/utils/globalIdUtils";

describe("globalIdUtils", () => {
  beforeEach(() => {
    delete (globalThis as any).workbookId;
    delete (globalThis as any).activeWorkSheetId;
  });

  afterEach(() => {
    delete (globalThis as any).workbookId;
    delete (globalThis as any).activeWorkSheetId;
  });

  describe("setGlobalIds", () => {
    it("sets workbookId on globalThis", () => {
      setGlobalIds({ type: "workbookId", id: "test-workbook-id" });

      expect((globalThis as any).workbookId).toBe("test-workbook-id");
    });

    it("sets activeWorkSheetId on globalThis", () => {
      setGlobalIds({ type: "activeWorkSheetId", id: "test-sheet-id" });

      expect((globalThis as any).activeWorkSheetId).toBe("test-sheet-id");
    });

    it("overwrites existing global id value when called again with same type", () => {
      setGlobalIds({ type: "workbookId", id: "initial-id" });
      expect((globalThis as any).workbookId).toBe("initial-id");

      setGlobalIds({ type: "workbookId", id: "updated-id" });
      expect((globalThis as any).workbookId).toBe("updated-id");
    });
  });

  describe("getGlobalIds", () => {
    it("returns undefined for workbookId and activeWorkSheetId when nothing is set", () => {
      const result = getGlobalIds();

      expect(result).toEqual({
        workbookId: undefined,
        activeWorkSheetId: undefined,
      });
    });

    it("returns workbookId and activeWorkSheetId when both are set", () => {
      setGlobalIds({ type: "workbookId", id: "wb-12345" });
      setGlobalIds({ type: "activeWorkSheetId", id: "sheet-67890" });

      const result = getGlobalIds();

      expect(result).toEqual({
        workbookId: "wb-12345",
        activeWorkSheetId: "sheet-67890",
      });
    });

    it("returns partial values when only one global id is set", () => {
      setGlobalIds({ type: "workbookId", id: "wb-only" });

      expect(getGlobalIds()).toEqual({
        workbookId: "wb-only",
        activeWorkSheetId: undefined,
      });

      delete (globalThis as any).workbookId;
      setGlobalIds({ type: "activeWorkSheetId", id: "sheet-only" });

      expect(getGlobalIds()).toEqual({
        workbookId: undefined,
        activeWorkSheetId: "sheet-only",
      });
    });
  });
});
