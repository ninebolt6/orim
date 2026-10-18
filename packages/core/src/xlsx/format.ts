import type { CellValue } from "./model";
import { defaultIndexedPalette } from "./colors";

type Condition = { operator: string; value: number };
type FormatToken = {
  kind:
    | "literal"
    | "symbol"
    | "placeholder"
    | "text"
    | "elapsed"
    | "year"
    | "month"
    | "day"
    | "hour"
    | "second"
    | "period"
    | "japaneseDay"
    | "scientific"
    | "unknown";
  value: string;
  exponentMarker?: string;
  exponentSign?: string;
  exponentPadding?: number;
};
type FormatSection = {
  tokens: FormatToken[];
  condition?: Condition;
  color?: string;
  valid: boolean;
};

const namedColors: Readonly<Record<string, string>> = {
  black: "#000000",
  blue: "#0000ff",
  cyan: "#00ffff",
  green: "#008000",
  magenta: "#ff00ff",
  red: "#ff0000",
  white: "#ffffff",
  yellow: "#ffff00",
};

const maxFormatLength = 4096;

const tokenize = (format: string): FormatSection[] => {
  const sections: FormatSection[] = [{ tokens: [], valid: true }];
  let section = sections[0]!;
  const push = (kind: FormatToken["kind"], value: string): void => {
    section.tokens.push({ kind, value });
  };
  for (let index = 0; index < format.length;) {
    const rest = format.slice(index);
    const character = format[index]!;
    if (character === ";") {
      section = { tokens: [], valid: true };
      sections.push(section);
      index++;
      continue;
    }
    if (character === '"') {
      let literal = "";
      let closed = false;
      index++;
      while (index < format.length) {
        if (format[index] === '"' && format[index + 1] === '"') {
          literal += '"';
          index += 2;
        } else if (format[index] === '"') {
          index++;
          closed = true;
          break;
        } else {
          literal += format[index]!;
          index++;
        }
      }
      if (!closed) section.valid = false;
      push("literal", literal);
      continue;
    }
    if (character === "\\") {
      if (index + 1 >= format.length) {
        section.valid = false;
        break;
      }
      push("literal", format[index + 1]!);
      index += 2;
      continue;
    }
    if (character === "_" || character === "*") {
      if (index + 1 >= format.length) {
        section.valid = false;
        break;
      }
      if (character === "_") push("literal", " ");
      index += 2;
      continue;
    }
    if (character === "[") {
      const closing = format.indexOf("]", index + 1);
      if (closing < 0) {
        section.valid = false;
        break;
      }
      const directive = format.slice(index + 1, closing);
      const lower = directive.toLowerCase();
      const conditional = /^(<=|>=|<>|=|<|>)(-?\d+(?:\.\d+)?)$/.exec(directive);
      if (conditional) {
        section.condition ??= { operator: conditional[1]!, value: Number(conditional[2]) };
      } else if (lower in namedColors) {
        section.color = namedColors[lower];
      } else {
        const colorIndex = /^color(\d+)$/i.exec(directive);
        if (colorIndex) {
          const colorNumber = Number(colorIndex[1]);
          const paletteIndex = colorNumber + 7;
          if (colorNumber >= 1 && colorNumber <= 56 && paletteIndex < defaultIndexedPalette.length)
            section.color = defaultIndexedPalette[paletteIndex];
          else section.valid = false;
        } else if (directive.startsWith("$")) {
          const currency = /^\$([^-]*)-/.exec(directive)?.[1] ?? "";
          if (currency) push("literal", currency);
        } else if (/^(h{1,2}|m{1,2}|s{1,2})$/i.test(directive)) {
          push("elapsed", lower);
        } else {
          section.valid = false;
        }
      }
      index = closing + 1;
      continue;
    }
    const period = /^(AM\/PM|A\/P)/i.exec(rest);
    if (period) {
      push("period", period[1]!);
      index += period[1]!.length;
      continue;
    }
    const scientific = /^([Ee])([+-])([0#?]+)/.exec(rest);
    if (scientific) {
      section.tokens.push({
        kind: "scientific",
        value: scientific[0],
        exponentMarker: scientific[1]!,
        exponentSign: scientific[2]!,
        exponentPadding: (scientific[3]!.match(/0/g) ?? []).length,
      });
      index += scientific[0].length;
      continue;
    }
    const japaneseDay = /^a{3,4}/i.exec(rest);
    if (japaneseDay) {
      push("japaneseDay", japaneseDay[0]);
      index += japaneseDay[0].length;
      continue;
    }
    if ("0#?".includes(character)) {
      push("placeholder", character);
      index++;
      continue;
    }
    if ("ymdhs".includes(character.toLowerCase())) {
      const letter = character.toLowerCase();
      let end = index + 1;
      while (format[end]?.toLowerCase() === letter) end++;
      const kind =
        letter === "y"
          ? "year"
          : letter === "m"
            ? "month"
            : letter === "d"
              ? "day"
              : letter === "h"
                ? "hour"
                : "second";
      push(kind, format.slice(index, end));
      index = end;
      continue;
    }
    if (/[a-z]/i.test(character)) {
      let end = index + 1;
      while (/[a-z]/i.test(format[end] ?? "")) end++;
      push("unknown", format.slice(index, end));
      section.valid = false;
      index = end;
      continue;
    }
    push(character === "@" ? "text" : "symbol", character);
    index++;
  }
  return sections;
};

const matchesCondition = (condition: Condition, value: number): boolean => {
  switch (condition.operator) {
    case ">":
      return value > condition.value;
    case ">=":
      return value >= condition.value;
    case "<":
      return value < condition.value;
    case "<=":
      return value <= condition.value;
    case "=":
      return value === condition.value;
    case "<>":
      return value !== condition.value;
    default:
      return false;
  }
};

const chooseSection = (
  sections: readonly FormatSection[],
  value: number,
): { section: FormatSection; value: number; negative: boolean } => {
  if (sections.some((section) => section.condition)) {
    const matching = sections.find(
      (section) => section.condition && matchesCondition(section.condition, value),
    );
    const fallback = sections.find((section) => !section.condition);
    return {
      section: matching ?? fallback ?? { tokens: [], valid: true },
      value,
      negative: value < 0,
    };
  }
  const selected =
    sections.length === 1
      ? 0
      : sections.length === 2
        ? value < 0
          ? 1
          : 0
        : value > 0
          ? 0
          : value < 0
            ? 1
            : 2;
  const negative = value < 0;
  return {
    section: sections[selected] ?? { tokens: [], valid: true },
    value: negative && sections.length > 1 ? Math.abs(value) : value,
    negative: negative && sections.length === 1,
  };
};

const trimDecimal = (value: string): string => {
  if (!value.includes(".")) return value;
  return value.replace(/(\.\d*?)0+$/, "$1").replace(/\.$/, "");
};

const expandExponential = (value: string): string => {
  const match = /^(-?)(\d+)(?:\.(\d*))?[eE]([+-]?\d+)$/.exec(value);
  if (!match) return value;
  const sign = match[1]!;
  const digits = match[2]! + (match[3] ?? "");
  const point = match[2]!.length + Number(match[4]);
  if (point <= 0) return `${sign}0.${"0".repeat(-point)}${digits}`;
  if (point >= digits.length) return `${sign}${digits}${"0".repeat(point - digits.length)}`;
  return `${sign}${digits.slice(0, point)}.${digits.slice(point)}`;
};

const general = (value: number): string => {
  if (!Number.isFinite(value)) return String(value);
  if (Number.isInteger(value) && Math.abs(value).toString().replace(".", "").length <= 11)
    return String(value);
  const absolute = Math.abs(value);
  if (absolute >= 1e11 || (absolute > 0 && absolute < 1e-4)) {
    const match = value.toExponential(5).match(/^(-?)(\d+(?:\.\d+)?)[eE]([+-]?\d+)$/);
    if (!match) return String(value);
    const exponent = Number(match[3]);
    return `${match[1]}${trimDecimal(match[2]!)}E${exponent < 0 ? "-" : "+"}${String(Math.abs(exponent)).padStart(2, "0")}`;
  }
  for (let precision = 11; precision >= 1; precision--) {
    const rendered = trimDecimal(expandExponential(value.toPrecision(precision)));
    if (rendered.replace("-", "").length <= 11) return rendered;
  }
  return String(value);
};

const decimalParts = (value: number): { sign: string; whole: string; fraction: string } => {
  const raw = value.toString();
  const match = /^(-?)(\d+)(?:\.(\d*))?(?:[eE]([+-]?\d+))?$/.exec(raw);
  if (!match) return { sign: value < 0 ? "-" : "", whole: "0", fraction: "" };
  const sign = match[1]!;
  const integer = match[2]!;
  const fraction = match[3] ?? "";
  const exponent = Number(match[4] ?? "0");
  const digits = integer + fraction;
  const point = integer.length + exponent;
  if (point <= 0) return { sign, whole: "0", fraction: `${"0".repeat(-point)}${digits}` };
  if (point >= digits.length)
    return { sign, whole: `${digits}${"0".repeat(point - digits.length)}`, fraction: "" };
  return { sign, whole: digits.slice(0, point), fraction: digits.slice(point) };
};

const incrementDigits = (digits: string): string => {
  return (BigInt(digits || "0") + 1n).toString();
};

const roundedDecimal = (value: number, places: number): string => {
  const { sign, whole, fraction } = decimalParts(Math.abs(value));
  const kept = `${whole}${fraction.padEnd(places + 1, "0").slice(0, places)}` || "0";
  const discarded = fraction[places] ?? "0";
  const rounded = discarded >= "5" ? incrementDigits(kept) : kept;
  const digits = rounded.padStart(places + 1, "0");
  const integer = places === 0 ? digits : digits.slice(0, -places);
  const decimal = places === 0 ? "" : `.${digits.slice(-places)}`;
  return `${sign}${integer}${decimal}`;
};

const tokenText = (token: FormatToken): string => {
  return token.value;
};

const joinTokens = (tokens: readonly FormatToken[]): string => tokens.map(tokenText).join("");

const isDateToken = (token: FormatToken): boolean => {
  return (
    token.kind === "elapsed" ||
    token.kind === "year" ||
    token.kind === "month" ||
    token.kind === "day" ||
    token.kind === "hour" ||
    token.kind === "second" ||
    token.kind === "period" ||
    token.kind === "japaneseDay"
  );
};

const dateFormat = (tokens: readonly FormatToken[]): boolean => {
  return tokens.some(isDateToken);
};

const pad = (value: number | string, length: number): string => {
  return String(value).padStart(length, "0");
};

const dateDetails = (
  serial: number,
  date1904: boolean,
  subSecondDigits: number,
): {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
  fraction: string;
  weekday: number;
  dayCount: number;
} => {
  const scale = 10 ** subSecondDigits;
  const unitsPerDay = 86400 * scale;
  const totalUnits = Math.round(serial * unitsPerDay);
  const dayCount = Math.floor(totalUnits / unitsPerDay);
  const remainder = totalUnits - dayCount * unitsPerDay;
  const totalSeconds = Math.floor(remainder / scale);
  const hour = Math.floor(totalSeconds / 3600);
  const minute = Math.floor((totalSeconds % 3600) / 60);
  const second = totalSeconds % 60;
  const fraction = subSecondDigits ? pad(remainder % scale, subSecondDigits) : "";
  if (!date1904 && dayCount === 60) {
    return {
      year: 1900,
      month: 2,
      day: 29,
      hour,
      minute,
      second,
      fraction,
      weekday: 3,
      dayCount,
    };
  }
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, dayCount < 60 ? 31 : 30);
  const adjustedDays = !date1904 && dayCount > 60 ? dayCount : dayCount;
  const date = new Date(epoch + adjustedDays * 86400000);
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    hour,
    minute,
    second,
    fraction,
    weekday: date.getUTCDay(),
    dayCount,
  };
};

const dateValue = (tokens: readonly FormatToken[], serial: number, date1904: boolean): string => {
  if (serial < 0 || serial > 2958465) return "########";
  const fractionalSeconds = tokens.findIndex(
    (token, index) =>
      token.value === "." &&
      tokens[index - 1]?.kind === "second" &&
      tokens[index + 1]?.kind === "placeholder",
  );
  let subSecondDigits = 0;
  if (fractionalSeconds >= 0) {
    for (
      let index = fractionalSeconds + 1;
      tokens[index]?.kind === "placeholder" && tokens[index]?.value === "0";
      index++
    )
      subSecondDigits++;
  }
  const details = dateDetails(serial, date1904, subSecondDigits);
  const hasPeriod = tokens.some((token) => token.kind === "period");
  const weekdays = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const shortWeekdays = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const japaneseWeekdays = ["日", "月", "火", "水", "木", "金", "土"];
  const japaneseFullWeekdays = [
    "日曜日",
    "月曜日",
    "火曜日",
    "水曜日",
    "木曜日",
    "金曜日",
    "土曜日",
  ];
  const months = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];
  const shortMonths = months.map((month) => month.slice(0, 3));
  const nearestDateType = (index: number, direction: -1 | 1): string | undefined => {
    for (let cursor = index + direction; tokens[cursor]; cursor += direction) {
      const token = tokens[cursor]!;
      if (token.kind === "elapsed") continue;
      if (
        token.kind === "year" ||
        token.kind === "month" ||
        token.kind === "day" ||
        token.kind === "hour" ||
        token.kind === "second"
      )
        return token.kind;
      if (token.kind === "unknown" || token.kind === "period" || token.kind === "japaneseDay")
        return undefined;
    }
    return undefined;
  };
  const elapsed = (unit: string): string => {
    const totalSeconds =
      details.dayCount * 86400 + details.hour * 3600 + details.minute * 60 + details.second;
    const amount =
      unit[0] === "h"
        ? Math.floor(totalSeconds / 3600)
        : unit[0] === "m"
          ? Math.floor(totalSeconds / 60)
          : totalSeconds;
    return unit.length > 1 ? pad(amount, 2) : String(amount);
  };
  return tokens
    .map((token, index) => {
      if (fractionalSeconds >= 0 && index === fractionalSeconds) return `.${details.fraction}`;
      if (
        fractionalSeconds >= 0 &&
        index > fractionalSeconds &&
        index <= fractionalSeconds + subSecondDigits
      )
        return "";
      if (token.kind === "literal") return token.value;
      if (token.kind === "elapsed") return elapsed(token.value);
      if (token.kind === "year")
        return token.value.length <= 2 ? pad(details.year % 100, 2) : pad(details.year, 4);
      if (token.kind === "month") {
        const previous = nearestDateType(index, -1);
        const next = nearestDateType(index, 1);
        const isMinute = previous === "hour" || next === "second";
        if (isMinute)
          return token.value.length === 1 ? String(details.minute) : pad(details.minute, 2);
        if (token.value.length === 1) return String(details.month);
        if (token.value.length === 2) return pad(details.month, 2);
        if (token.value.length === 3) return shortMonths[details.month - 1]!;
        if (token.value.length === 4) return months[details.month - 1]!;
        return months[details.month - 1]![0]!;
      }
      if (token.kind === "day") {
        if (token.value.length === 1) return String(details.day);
        if (token.value.length === 2) return pad(details.day, 2);
        if (token.value.length === 3) return shortWeekdays[details.weekday]!;
        return weekdays[details.weekday]!;
      }
      if (token.kind === "japaneseDay")
        return token.value.length === 3
          ? japaneseWeekdays[details.weekday]!
          : japaneseFullWeekdays[details.weekday]!;
      if (token.kind === "hour") {
        const hour = hasPeriod ? details.hour % 12 || 12 : details.hour;
        return token.value.length === 1 ? String(hour) : pad(hour, 2);
      }
      if (token.kind === "second")
        return token.value.length === 1 ? String(details.second) : pad(details.second, 2);
      if (token.kind === "period")
        return token.value.length === 3
          ? details.hour < 12
            ? "A"
            : "P"
          : details.hour < 12
            ? "AM"
            : "PM";
      return tokenText(token);
    })
    .join("");
};

const placeholders = (tokens: readonly FormatToken[]): number[] => {
  return tokens.flatMap((token, index) => (token.kind === "placeholder" ? [index] : []));
};

const fractionFormat = (tokens: readonly FormatToken[], value: number): string | undefined => {
  const slash = tokens.findIndex((token) => token.kind === "symbol" && token.value === "/");
  if (slash < 0) return undefined;
  const allPlaces = placeholders(tokens);
  let numeratorStart = 0;
  for (let index = 0; index < slash; index++) {
    if (tokens[index]!.value === " " || tokens[index]!.kind === "literal")
      numeratorStart = index + 1;
  }
  const wholePositions = allPlaces.filter((index) => index < numeratorStart);
  const numeratorPositions = allPlaces.filter((index) => index >= numeratorStart && index < slash);
  const denominatorPositions = allPlaces.filter((index) => index > slash);
  const fixedDenominator = joinTokens(tokens.slice(slash + 1, denominatorPositions[0]));
  const fixed = /^\d+$/.test(fixedDenominator);
  if (numeratorPositions.length === 0 || (!fixed && denominatorPositions.length === 0))
    return undefined;
  if (!fixed && denominatorPositions.length > 4) return undefined;
  const maxDenominator = fixed ? Number(fixedDenominator) : 10 ** denominatorPositions.length - 1;
  if (!Number.isFinite(maxDenominator) || maxDenominator < 1) return undefined;
  const mixed = wholePositions.length > 0;
  const absolute = Math.abs(value);
  let whole = mixed ? Math.floor(absolute) : 0;
  const fractional = mixed ? absolute - whole : absolute;
  let numerator = 0;
  let denominator = maxDenominator;
  if (fixed) {
    numerator = Math.round(fractional * denominator);
  } else {
    let best = Number.POSITIVE_INFINITY;
    for (let candidate = 1; candidate <= maxDenominator; candidate++) {
      const candidateNumerator = Math.round(fractional * candidate);
      const error = Math.abs(fractional - candidateNumerator / candidate);
      if (error < best) {
        best = error;
        numerator = candidateNumerator;
        denominator = candidate;
      }
    }
  }
  if (mixed && numerator >= denominator) {
    whole++;
    numerator = 0;
  }
  const formatPart = (positions: readonly number[], valueText: string): string => {
    const digits = valueText.replace(/^0+(?=\d)/, "");
    const padding = positions.length - digits.length;
    return positions
      .map((position, index) => {
        const placeholder = tokens[position]!.value;
        const digitIndex = index - padding;
        if (digitIndex >= 0) return digits[digitIndex] ?? "";
        return placeholder === "0" ? "0" : placeholder === "?" ? " " : "";
      })
      .join("");
  };
  const numeratorText = formatPart(numeratorPositions, String(numerator));
  const prefix = joinTokens(tokens.slice(0, wholePositions[0] ?? numeratorPositions[0]!));
  const wholeText = mixed
    ? (whole === 0 && wholePositions.every((position) => tokens[position]!.value !== "0")
        ? wholePositions.map((position) => (tokens[position]!.value === "?" ? " " : "")).join("")
        : formatPart(wholePositions, String(whole))) +
      joinTokens(tokens.slice(wholePositions.at(-1)! + 1, numeratorPositions[0]))
    : "";
  const numeratorGap = joinTokens(tokens.slice(numeratorPositions.at(-1)! + 1, slash));
  const denominatorGap = fixed ? "" : joinTokens(tokens.slice(slash + 1, denominatorPositions[0]));
  const denominatorOutput = fixed
    ? fixedDenominator
    : formatPart(denominatorPositions, String(denominator));
  const denominatorEnd = fixed ? slash + fixedDenominator.length : denominatorPositions.at(-1)!;
  const suffix = joinTokens(tokens.slice(denominatorEnd + 1));
  if (mixed && numerator === 0) {
    const fractionStart = wholePositions.at(-1)! + 1;
    const fractionWidth = joinTokens(tokens.slice(fractionStart, denominatorEnd + 1)).length;
    return `${prefix}${formatPart(wholePositions, String(whole))}${" ".repeat(fractionWidth)}${suffix}`;
  }
  return `${prefix}${wholeText}${numeratorText}${numeratorGap}/${denominatorGap}${denominatorOutput}${suffix}`;
};

const integerText = (
  tokens: readonly FormatToken[],
  first: number,
  last: number,
  digits: string,
  grouping: boolean,
): string => {
  const integerTokens = tokens.slice(first, last + 1);
  const positions = integerTokens.flatMap((token, index) =>
    token.kind === "placeholder" ? [index] : [],
  );
  const extra =
    digits.length > positions.length ? digits.slice(0, digits.length - positions.length) : "";
  const remaining = digits.slice(extra.length);
  const missing = positions.length - remaining.length;
  let positionIndex = 0;
  const filled = integerTokens
    .map((token) => {
      if (token.kind !== "placeholder")
        return grouping && token.kind === "symbol" && token.value === "," ? "" : token.value;
      const digitIndex = positionIndex++ - missing;
      const digit = digitIndex >= 0 ? remaining[digitIndex] : "";
      if (digit) return digit;
      return token.value === "0" ? "0" : token.value === "?" ? " " : "";
    })
    .join("");
  const result = `${extra}${filled}`;
  return grouping && !/[^\d\s]/.test(result)
    ? result.replace(/\B(?=(\d{3})+(?!\d))/g, ",")
    : result;
};

const decimalFormat = (
  tokens: readonly FormatToken[],
  value: number,
): { text: string; valid: boolean } => {
  const places = placeholders(tokens);
  if (!places.length) return { text: "", valid: false };
  const percentCount = tokens.filter(
    (token) => token.kind === "symbol" && token.value === "%",
  ).length;
  const scientificIndex = tokens.findIndex((token) => token.kind === "scientific");
  if (scientificIndex >= 0) {
    const scientific = tokens[scientificIndex]!;
    const mantissaTokens = tokens.slice(0, scientificIndex);
    const integerDigits = placeholders(mantissaTokens).filter(
      (index) => !mantissaTokens.slice(0, index).some((token) => token.value === "."),
    ).length;
    const absolute = Math.abs(value) * 100 ** percentCount;
    const exponent =
      absolute === 0
        ? 0
        : integerDigits > 1
          ? Math.floor(Math.floor(Math.log10(absolute)) / integerDigits) * integerDigits
          : Math.floor(Math.log10(absolute));
    let formatted = decimalFormat(mantissaTokens, absolute === 0 ? 0 : absolute / 10 ** exponent);
    if (!formatted.valid) return formatted;
    let displayExponent = exponent;
    const displayedValue = Number(formatted.text.replace(/[^\d.-]/g, ""));
    if (displayedValue >= 10 ** integerDigits) {
      displayExponent += integerDigits;
      formatted = decimalFormat(mantissaTokens, absolute / 10 ** displayExponent);
    }
    const exponentText = `${displayExponent < 0 ? "-" : scientific.exponentSign === "+" ? "+" : ""}${pad(
      Math.abs(displayExponent),
      Math.max(1, scientific.exponentPadding ?? 0),
    )}`;
    return {
      text: `${formatted.text}${scientific.exponentMarker}${exponentText}${joinTokens(
        tokens.slice(scientificIndex + 1),
      )}`,
      valid: formatted.valid,
    };
  }
  const slashFormat = fractionFormat(tokens, value);
  if (slashFormat !== undefined) return { text: slashFormat, valid: true };

  const decimalIndex = tokens.findIndex(
    (token, index) =>
      token.kind === "symbol" &&
      token.value === "." &&
      tokens.slice(index + 1).some((following) => following.kind === "placeholder"),
  );
  const beforeDecimal = decimalIndex < 0 ? tokens : tokens.slice(0, decimalIndex);
  const integerPlaces = placeholders(beforeDecimal);
  const firstPlace = integerPlaces[0]!;
  const lastPlace = integerPlaces.at(-1)!;
  const decimalTokens = decimalIndex < 0 ? [] : tokens.slice(decimalIndex + 1);
  const decimalPlaces = decimalTokens.filter((token) => token.kind === "placeholder");
  const integerCommas = beforeDecimal.flatMap((token, index) =>
    token.kind === "symbol" && token.value === "," ? [index] : [],
  );
  let trailingScale = 0;
  for (
    let index = lastPlace + 1;
    beforeDecimal[index]?.kind === "symbol" && beforeDecimal[index]?.value === ",";
    index++
  )
    trailingScale++;
  const grouping = integerCommas.some((index) => index < lastPlace && index > firstPlace);
  const scaled = (Math.abs(value) * 100 ** percentCount) / 1000 ** trailingScale;
  const rounded = roundedDecimal(scaled, decimalPlaces.length);
  const roundedParts = /^(-?)(\d+)(?:\.(\d+))?$/.exec(rounded);
  if (!roundedParts) return { text: "", valid: false };
  const rawInteger = roundedParts[2]!.replace(/^0+(?=\d)/, "");
  const digits = rawInteger || "0";
  const rawFraction = roundedParts[3] ?? "";
  let integer = integerText(tokens, firstPlace, lastPlace, digits, grouping);
  if (/^0+$/.test(digits) && integerPlaces.every((index) => tokens[index]!.value !== "0"))
    integer = integer.replace(/0/g, "");

  let fraction = rawFraction.split("");
  for (let index = decimalPlaces.length - 1; index >= 0; index--) {
    const placeholder = decimalPlaces[index]!.value;
    if (fraction[index] !== "0") break;
    if (placeholder === "#") fraction[index] = "";
    else if (placeholder === "?") fraction[index] = " ";
    else break;
  }
  const fractionText = fraction.join("");
  const prefix = joinTokens(tokens.slice(0, firstPlace));
  const decimalPoint = decimalIndex >= 0 && fractionText.length ? "." : "";
  let lastDecimalPlace = -1;
  decimalTokens.forEach((token, index) => {
    if (token.kind === "placeholder") lastDecimalPlace = index;
  });
  const finalPlace = decimalIndex < 0 ? lastPlace : decimalIndex + 1 + lastDecimalPlace;
  const suffix = joinTokens(
    tokens.slice(finalPlace + 1).filter((token) => token.kind !== "symbol" || token.value !== ","),
  );
  const text = `${prefix}${integer}${decimalPoint}${fractionText}${suffix}`;
  return { text, valid: true };
};

const formatNumber = (
  section: FormatSection,
  value: number,
  negativePrefix: boolean,
  fallbackValue: number,
  date1904: boolean,
): string => {
  if (!section.valid) return general(fallbackValue);
  if (dateFormat(section.tokens)) return dateValue(section.tokens, value, date1904);
  const normalized = Number(value.toPrecision(15));
  const rendered = decimalFormat(section.tokens, normalized);
  if (!rendered.valid) {
    if (!section.tokens.some((token) => token.kind === "placeholder"))
      return joinTokens(section.tokens);
    return general(fallbackValue);
  }
  return `${negativePrefix ? "-" : ""}${rendered.text}`;
};

export const formatCellValue = (
  value: CellValue,
  format: string,
  date1904: boolean,
): { text: string; color?: string } => {
  try {
    if (value.type === "empty") return { text: "" };
    if (value.type === "boolean") return { text: value.value ? "TRUE" : "FALSE" };
    if (value.type === "error") return { text: value.value };
    if (format.length > maxFormatLength)
      return { text: value.type === "string" ? value.value : general(value.value) };
    if (value.type === "string") {
      const sections = tokenize(format);
      const textFormat =
        sections.length >= 4
          ? sections[3]
          : sections.length === 1 && sections[0]!.tokens.some((token) => token.kind === "text")
            ? sections[0]
            : undefined;
      if (textFormat === undefined) return { text: value.value };
      return {
        text: textFormat.tokens
          .map((token) => (token.kind === "text" ? value.value : tokenText(token)))
          .join(""),
        ...(textFormat.color ? { color: textFormat.color } : {}),
      };
    }
    if (!Number.isFinite(value.value)) return { text: general(value.value) };
    if (!format || format.toLowerCase() === "general") return { text: general(value.value) };
    const sections = tokenize(format);
    const selected = chooseSection(sections, Number(value.value.toPrecision(15)));
    return {
      text: formatNumber(
        selected.section,
        selected.value,
        selected.negative,
        value.value,
        date1904,
      ),
      ...(selected.section.valid && selected.section.color
        ? { color: selected.section.color }
        : {}),
    };
  } catch {
    return { text: value.type === "number" ? general(value.value) : "" };
  }
};
