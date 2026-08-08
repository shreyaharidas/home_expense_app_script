export function setGlobalIds({ type, id }): void {
  (globalThis as any)[type] = id;
}

export function getGlobalIds(): { workbookId: string, activeWorkSheetId: string } {
  return {
    workbookId: (globalThis as any).workbookId,
    activeWorkSheetId: (globalThis as any).activeWorkSheetId
  };
}
