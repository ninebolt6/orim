import { useLayoutEffect, useRef } from "react";
import type { HTMLAttributes, JSX } from "react";
import { renderPresentation } from "@orim/core";
import type { PptxPresentation } from "@orim/core";

export interface PptxViewerProps extends HTMLAttributes<HTMLDivElement> {
  presentation: PptxPresentation | null | undefined;
  /** Render slides marked hidden in the deck. */
  hiddenSlides?: boolean;
}

export const PptxViewer = ({
  presentation,
  hiddenSlides = false,
  ...rest
}: PptxViewerProps): JSX.Element => {
  const element = useRef<HTMLDivElement>(null);

  useLayoutEffect(() => {
    const container = element.current;
    if (!presentation || !container) return;
    const rendered = renderPresentation(presentation, container, { hiddenSlides });
    return () => rendered.destroy();
  }, [presentation, hiddenSlides]);

  return <div {...rest} ref={element} />;
};
