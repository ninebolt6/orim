import { abortError, acquireContainerLease, type RenderLease } from "../../common/lease";
import type { XlsxWorkbook } from "../model";
import { parseXlsx } from "../parse";
import { renderWorksheet, type RenderedWorksheet } from "./worksheet";

export interface RenderedWorkbook {
  readonly workbook: XlsxWorkbook;
  readonly sheet: number;
  update(next: { sheet?: number }): void;
  destroy(): void;
}

export interface RenderWorkbookOptions {
  readonly sheet?: number;
  readonly onSheetChange?: (index: number) => void;
}

export interface RenderXlsxOptions extends RenderWorkbookOptions {
  readonly signal?: AbortSignal;
}

const validateSheet = (workbook: XlsxWorkbook, index: number): void => {
  if (
    !Number.isInteger(index) ||
    index < 0 ||
    index >= workbook.sheets.length ||
    workbook.sheets[index]?.state !== "visible"
  )
    throw new RangeError("Worksheet index is out of range or hidden.");
};

export const renderWorkbook = (
  workbook: XlsxWorkbook,
  container: HTMLElement,
  options: RenderWorkbookOptions = {},
): RenderedWorkbook => {
  let sheet = options.sheet ?? workbook.activeSheet;
  validateSheet(workbook, sheet);
  let root: HTMLDivElement | undefined;
  let worksheetView: RenderedWorksheet | undefined;
  const lease = acquireContainerLease(container, () => {
    worksheetView?.destroy();
    if (root?.parentElement === container) container.replaceChildren();
  });
  try {
    const document = container.ownerDocument;
    root = document.createElement("div");
    root.className = "orim-xlsx-book";
    const worksheetContainer = document.createElement("div");
    worksheetContainer.className = "orim-xlsx-sheet";
    const tabs = document.createElement("div");
    tabs.className = "orim-xlsx-tabs";
    tabs.setAttribute("role", "tablist");
    tabs.setAttribute("aria-label", "Worksheets");
    const visibleSheets = workbook.sheets.flatMap((item, index) =>
      item.state === "visible" ? [{ item, index }] : [],
    );
    const buttons = new Map<number, HTMLButtonElement>();
    const setSelected = (index: number): void => {
      for (const [buttonIndex, button] of buttons) {
        const selected = buttonIndex === index;
        button.setAttribute("aria-selected", String(selected));
        button.tabIndex = selected ? 0 : -1;
      }
    };
    const selectSheet = (index: number, notify: boolean): void => {
      if (!lease.active) return;
      validateSheet(workbook, index);
      if (sheet === index && worksheetView) {
        setSelected(index);
        if (notify) options.onSheetChange?.(index);
        return;
      }
      worksheetView?.destroy();
      sheet = index;
      worksheetView = renderWorksheet(workbook, index, worksheetContainer);
      setSelected(index);
      if (notify) options.onSheetChange?.(index);
    };
    for (const { item, index } of visibleSheets) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "orim-xlsx-tab";
      button.setAttribute("role", "tab");
      button.setAttribute("aria-selected", String(index === sheet));
      button.tabIndex = index === sheet ? 0 : -1;
      button.append(document.createTextNode(item.name));
      button.addEventListener("click", () => selectSheet(index, true));
      buttons.set(index, button);
      tabs.append(button);
    }
    root.append(worksheetContainer, tabs);
    container.replaceChildren(root);
    worksheetView = renderWorksheet(workbook, sheet, worksheetContainer);
    return {
      workbook,
      get sheet() {
        return sheet;
      },
      update: (next) => {
        if (!lease.active || next.sheet === undefined) return;
        selectSheet(next.sheet, false);
      },
      destroy: () => lease.cancel(),
    };
  } catch (error) {
    lease.cancel();
    throw error;
  }
};

export const renderXlsx = async (
  input: ArrayBuffer | Uint8Array | Blob,
  container: HTMLElement,
  options: RenderXlsxOptions = {},
): Promise<RenderedWorkbook> => {
  const controller = new AbortController();
  let cancelled = false;
  let lease: RenderLease;
  const onAbort = (): void => lease.cancel();
  lease = acquireContainerLease(container, () => {
    cancelled = true;
    controller.abort();
    options.signal?.removeEventListener("abort", onAbort);
  });
  if (options.signal?.aborted) lease.cancel();
  else options.signal?.addEventListener("abort", onAbort, { once: true });
  const check = (): void => {
    if (cancelled || controller.signal.aborted || !lease.active) throw abortError();
  };
  let workbook: XlsxWorkbook;
  try {
    check();
    workbook = await parseXlsx(input);
    check();
  } catch (error) {
    const aborted = cancelled || controller.signal.aborted || !lease.active;
    lease.cancel();
    if (aborted) throw abortError();
    throw error;
  }
  try {
    return renderWorkbook(workbook, container, {
      sheet: options.sheet,
      onSheetChange: options.onSheetChange,
    });
  } catch (error) {
    lease.cancel();
    throw error;
  }
};
