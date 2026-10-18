export interface RenderLease {
  cancel(): void;
  readonly active: boolean;
}

const renderLeases = new WeakMap<HTMLElement, RenderLease>();

export const acquireContainerLease = (
  container: HTMLElement,
  onCancel: () => void,
): RenderLease => {
  let cancelled = false;
  let lease: RenderLease;
  lease = {
    get active() {
      return !cancelled && renderLeases.get(container) === lease;
    },
    cancel() {
      if (cancelled) return;
      cancelled = true;
      onCancel();
      if (renderLeases.get(container) === lease) renderLeases.delete(container);
    },
  };
  const previous = renderLeases.get(container);
  renderLeases.set(container, lease);
  previous?.cancel();
  return lease;
};

export const abortError = (): DOMException => {
  return new DOMException("Rendering was aborted.", "AbortError");
};
