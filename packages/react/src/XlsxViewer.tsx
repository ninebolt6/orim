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
  const current =
    sheet ??
    (selection && selection.workbook === workbook ? selection.index : (workbook?.activeSheet ?? 0));

  useLayoutEffect(() => {
    const container = element.current;
    if (!workbook || !container) return;
    const rendered = renderWorksheet(workbook, current, container);
    return () => rendered.destroy();
  }, [workbook, current]);

  const visibleSheets = workbook?.sheets.flatMap((item, index) =>
    item.state === "visible" ? [{ item, index }] : [],
  );

  return (
    <div {...rest} className={["orim-xlsx-react", className].filter(Boolean).join(" ")}>
      <div className="orim-xlsx-sheet" ref={element} />
      {workbook && (
        <div className="orim-xlsx-tabs" role="tablist" aria-label="Worksheets">
          {visibleSheets?.map(({ item, index }) => (
            <button
              aria-selected={current === index}
              className="orim-xlsx-tab"
              key={index}
              onClick={() => {
                setSelection({ workbook, index });
                onSheetChange?.(index);
              }}
              role="tab"
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
