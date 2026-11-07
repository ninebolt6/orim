export const validColor = (value: string | undefined): string | undefined => {
  return value && /^#[\da-f]{6}(?:[\da-f]{2})?$/i.test(value) ? value : undefined;
};

export const safeHref = (value: string | undefined): string | undefined => {
  if (!value) return undefined;
  if (value.startsWith("#") && /^#[\w.-]+$/.test(value)) return value;
  try {
    const url = new URL(value);
    return ["http:", "https:", "mailto:"].includes(url.protocol) ? value : undefined;
  } catch {
    return undefined;
  }
};
