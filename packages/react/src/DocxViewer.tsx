import { useEffect, useRef } from "react";
import type { HTMLAttributes, JSX } from "react";
import { renderDocx } from "@orim/core";
import type { ChangesMode, DocxDocument, RenderedDocument } from "@orim/core";

export interface DocxViewerProps extends Omit<HTMLAttributes<HTMLDivElement>, "onError"> {
  src: ArrayBuffer | Uint8Array | Blob | DocxDocument | null | undefined;
  changes?: ChangesMode;
  comments?: boolean;
  onRender?: (rendered: RenderedDocument) => void;
  onError?: (error: unknown) => void;
}

function isAbortError(error: unknown): boolean {
  return error instanceof DOMException && error.name === "AbortError";
}

export function DocxViewer({
  src,
  changes = "markup",
  comments = true,
  onRender,
  onError,
  ...rest
}: DocxViewerProps): JSX.Element {
  const element = useRef<HTMLDivElement>(null);
  const rendered = useRef<RenderedDocument | undefined>(undefined);
  const changesRef = useRef(changes);
  const commentsRef = useRef(comments);
  const onRenderRef = useRef(onRender);
  const onErrorRef = useRef(onError);

  useEffect(() => {
    changesRef.current = changes;
    commentsRef.current = comments;
    onRenderRef.current = onRender;
    onErrorRef.current = onError;
  }, [changes, comments, onRender, onError]);

  useEffect(() => {
    const container = element.current;
    if (!container) return;
    const controller = new AbortController();
    let active = true;
    rendered.current?.destroy();
    rendered.current = undefined;
    container.replaceChildren();
    if (src != null) {
      void renderDocx(src, container, {
        changes: changesRef.current,
        comments: commentsRef.current,
        signal: controller.signal,
      }).then(
        (result) => {
          if (!active || controller.signal.aborted) {
            result.destroy();
            return;
          }
          result.update({ changes: changesRef.current, comments: commentsRef.current });
          rendered.current = result;
          onRenderRef.current?.(result);
        },
        (error: unknown) => {
          if (active && !isAbortError(error)) onErrorRef.current?.(error);
        },
      );
    }
    return () => {
      active = false;
      controller.abort();
      rendered.current?.destroy();
      rendered.current = undefined;
    };
  }, [src]);

  useEffect(() => {
    rendered.current?.update({ changes, comments });
  }, [changes, comments]);

  return <div {...rest} ref={element} />;
}
