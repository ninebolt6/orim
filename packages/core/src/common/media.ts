export interface ReadonlyBytes {
  readonly length: number;
  readonly [index: number]: number;
  slice(start?: number, end?: number): Uint8Array;
}

export interface MediaPart {
  readonly bytes: ReadonlyBytes;
  readonly mime: string;
}

export interface ObjectUrls {
  url(key: string, part: MediaPart | undefined): string | undefined;
  revoke(): void;
}

export const createObjectUrls = (): ObjectUrls => {
  const urls = new Map<string, string>();
  return {
    url(key, part) {
      if (!part || typeof URL.createObjectURL !== "function") return undefined;
      if (!/^image\/(?:png|jpeg|gif|bmp|webp|avif|tiff|x-icon)$/i.test(part.mime)) return undefined;
      let url = urls.get(key);
      if (!url) {
        url = URL.createObjectURL(
          new Blob([part.bytes.slice().buffer as ArrayBuffer], { type: part.mime }),
        );
        urls.set(key, url);
      }
      return url;
    },
    revoke() {
      for (const url of urls.values()) URL.revokeObjectURL(url);
      urls.clear();
    },
  };
};
