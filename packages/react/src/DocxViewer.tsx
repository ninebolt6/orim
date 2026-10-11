import { useLayoutEffect, useRef } from "react";
import type { HTMLAttributes, JSX } from "react";
import { renderDocument } from "@orim/core";
import type { ChangesMode, DocxDocument } from "@orim/core";

export interface DocxViewerProps extends HTMLAttributes<HTMLDivElement> {
  document: DocxDocument | null | undefined;
  changes?: ChangesMode;
  comments?: boolean;
}

export function DocxViewer({
  document,
  changes = "markup",
  comments = true,
  ...rest
}: DocxViewerProps): JSX.Element {
  const element = useRef<HTMLDivElement>(null);

  useLayoutEffect(() => {
    const container = element.current;
    if (!document || !container) return;
    const rendered = renderDocument(document, container, { changes, comments });
    return () => rendered.destroy();
  }, [document, changes, comments]);

  return <div {...rest} ref={element} />;
}
