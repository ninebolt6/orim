import { abortError, acquireContainerLease, type RenderLease } from "./lease";

export const renderParsed = async <Model, Rendered>(
  load: () => Promise<Model>,
  container: HTMLElement,
  signal: AbortSignal | undefined,
  render: (model: Model) => Rendered,
): Promise<Rendered> => {
  let cancelled = false;
  let lease: RenderLease;
  const onAbort = (): void => lease.cancel();
  lease = acquireContainerLease(container, () => {
    cancelled = true;
    signal?.removeEventListener("abort", onAbort);
  });
  if (signal?.aborted) lease.cancel();
  else signal?.addEventListener("abort", onAbort, { once: true });
  const check = (): void => {
    if (cancelled || !lease.active) throw abortError();
  };
  let model: Model;
  try {
    check();
    model = await load();
    check();
  } catch (error) {
    const aborted = cancelled || !lease.active;
    lease.cancel();
    if (aborted) throw abortError();
    throw error;
  }
  try {
    return render(model);
  } catch (error) {
    lease.cancel();
    throw error;
  }
};
