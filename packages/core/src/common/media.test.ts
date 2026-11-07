import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { createObjectUrls } from "./media";

const bytes = new Uint8Array([1, 2, 3]);

describe("media object URLs", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it.each(["png", "jpeg", "gif", "bmp", "webp", "avif", "tiff", "x-icon"])(
    "creates URLs for case-insensitive image/%s MIME types",
    (type) => {
      const create = vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:image");
      const urls = createObjectUrls();
      expect(urls.url("image", { bytes, mime: `IMAGE/${type.toUpperCase()}` })).toBe("blob:image");
      expect(create).toHaveBeenCalledOnce();
      expect(create.mock.calls[0][0]).toBeInstanceOf(Blob);
      expect((create.mock.calls[0][0] as Blob).type).toBe(`image/${type}`);
    },
  );

  it.each([
    "image/svg+xml",
    "IMAGE/SVG+XML",
    "image/png; charset=utf-8",
    "text/html",
    "application/octet-stream",
    "image/emf",
    "image/png-extra",
    "",
  ])("rejects unsupported MIME type %s before creating or reusing a URL", (mime) => {
    const create = vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:image");
    const urls = createObjectUrls();
    expect(urls.url("image", { bytes, mime })).toBeUndefined();
    expect(create).not.toHaveBeenCalled();
    urls.url("image", { bytes, mime: "image/png" });
    expect(urls.url("image", { bytes, mime })).toBeUndefined();
    expect(create).toHaveBeenCalledOnce();
  });

  it("returns undefined for missing parts or an unavailable object URL API", () => {
    const create = vi.spyOn(URL, "createObjectURL");
    expect(createObjectUrls().url("missing", undefined)).toBeUndefined();
    expect(create).not.toHaveBeenCalled();
    vi.stubGlobal("URL", { createObjectURL: undefined });
    expect(createObjectUrls().url("image", { bytes, mime: "image/png" })).toBeUndefined();
  });

  it("caches safe URLs and revokes them idempotently", () => {
    const create = vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:image");
    const revoke = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
    const urls = createObjectUrls();
    const part = { bytes, mime: "image/png" };
    expect(urls.url("image", part)).toBe("blob:image");
    expect(urls.url("image", part)).toBe("blob:image");
    expect(create).toHaveBeenCalledOnce();
    urls.revoke();
    urls.revoke();
    expect(revoke).toHaveBeenCalledExactlyOnceWith("blob:image");
    urls.url("image", part);
    expect(create).toHaveBeenCalledTimes(2);
  });
});
