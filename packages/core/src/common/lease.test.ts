import { describe, expect, it, vi } from "vite-plus/test";
import { acquireContainerLease } from "./lease";

describe("container render leases", () => {
  it("cancels the previous holder and clears its content when a new holder acquires the container", () => {
    const container = document.createElement("div");
    const cancelPrevious = vi.fn(() => container.replaceChildren());
    const previous = acquireContainerLease(container, cancelPrevious);
    container.append(document.createElement("span"));

    const current = acquireContainerLease(container, () => {});

    expect(cancelPrevious).toHaveBeenCalledOnce();
    expect(previous.active).toBe(false);
    expect(current.active).toBe(true);
    expect(container.childElementCount).toBe(0);
  });

  it("keeps newer content when a stale holder cancels and makes cancellation idempotent", () => {
    const container = document.createElement("div");
    const cancelPrevious = vi.fn(() => container.replaceChildren());
    const previous = acquireContainerLease(container, cancelPrevious);
    const current = acquireContainerLease(container, () => container.replaceChildren());
    const currentContent = document.createElement("span");
    container.append(currentContent);

    previous.cancel();
    expect(container.firstChild).toBe(currentContent);
    current.cancel();
    current.cancel();

    expect(container.childElementCount).toBe(0);
    expect(cancelPrevious).toHaveBeenCalledOnce();
    expect(previous.active).toBe(false);
    expect(current.active).toBe(false);
  });
});
