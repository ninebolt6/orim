const number = (value: string | undefined): number | undefined => {
  if (value === undefined) return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
};

/** English Metric Units (12,700 per point) to points. */
export const emu = (value: string | undefined): number | undefined => {
  const parsed = number(value);
  return parsed === undefined ? undefined : parsed / 12700;
};

/** Angles in 60,000ths of a degree to degrees. */
export const angle = (value: string | undefined): number | undefined => {
  const parsed = number(value);
  return parsed === undefined ? undefined : parsed / 60000;
};

/** Percentages in 1,000ths of a percent to fractions (100000 = 1). */
export const fraction = (value: string | undefined): number | undefined => {
  const parsed = number(value);
  return parsed === undefined ? undefined : parsed / 100000;
};
