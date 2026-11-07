import { useLayoutEffect, useRef, useState } from "react";
import type { HTMLAttributes, JSX } from "react";
import { renderWorksheet } from "@orim/core";
import type { XlsxWorkbook } from "@orim/core";

export interface XlsxViewerProps extends HTMLAttributes<HTMLDivElement> {
  workbook: XlsxWorkbook | null | undefined;
  sheet?: number;
  onSheetChange?: (index: number) => void;
}

export const XlsxViewer = ({
  workbook,
  sheet,
  onSheetChange,
  className,
  ...rest
}: XlsxViewerProps): JSX.Element => {
  const element = useRef<HTMLDivElement>(null);
  const [selection, setSelection] = useState<{ workbook: XlsxWorkbook; index: number }>();
  const visibleSheets = workbook?.sheets.flatMap((item, index) =>
    item.state === "visible" ? [{ item, index }] : [],
  );
  const fallback =
    visibleSheets?.find(({ index }) => index === workbook?.activeSheet)?.index ??
    visibleSheets?.[0]?.index ??
    -1;
  const requested =
    sheet ?? (selection && selection.workbook === workbook ? selection.index : fallback);
  const current = visibleSheets?.some(({ index }) => index === requested) ? requested : fallback;

  useLayoutEffect(() => {
    const container = element.current;
    if (!workbook || !container || current < 0) return;
    const rendered = renderWorksheet(workbook, current, container);
    return () => rendered.destroy();
  }, [workbook, current]);

  return (
    <div {...rest} className={["orim-xlsx-react", className].filter(Boolean).join(" ")}>
      <div className="orim-xlsx-sheet" key={current < 0 ? "empty" : "sheet"} ref={element}>
        {workbook && current < 0 && <div className="orim-xlsx-note">No visible worksheets.</div>}
      </div>
      {workbook && (
        <div className="orim-xlsx-tabs" role="tablist" aria-label="Worksheets">
          {visibleSheets?.map(({ item, index }, position) => (
            <button
              aria-selected={current === index}
              className="orim-xlsx-tab"
              key={index}
              onClick={() => {
                setSelection({ workbook, index });
                onSheetChange?.(index);
              }}
              role="tab"
              onKeyDown={(event) => {
                const next =
                  event.key === "ArrowLeft"
                    ? (position - 1 + visibleSheets.length) % visibleSheets.length
                    : event.key === "ArrowRight"
                      ? (position + 1) % visibleSheets.length
                      : event.key === "Home"
                        ? 0
                        : event.key === "End"
                          ? visibleSheets.length - 1
                          : undefined;
                if (next === undefined) return;
                event.preventDefault();
                const target = visibleSheets[next]!.index;
                setSelection({ workbook, index: target });
                onSheetChange?.(target);
                event.currentTarget.parentElement
                  ?.querySelectorAll<HTMLButtonElement>('[role="tab"]')
                  [next]?.focus();
              }}
              tabIndex={current === index ? 0 : -1}
              type="button"
            >
              {item.name}
            </button>
          ))}
        </div>
      )}
    </div>
  );
};
