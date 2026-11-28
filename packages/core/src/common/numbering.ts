export const alphabetic = (number: number, upper: boolean): string => {
  if (!Number.isSafeInteger(number)) return String(number);
  let result = "";
  for (let n = Math.max(1, number); n > 0; n = Math.floor((n - 1) / 26)) {
    result = String.fromCharCode((upper ? 65 : 97) + ((n - 1) % 26)) + result;
  }
  return result;
};

const romanDigits: readonly (readonly [number, string])[] = [
  [1000, "M"],
  [900, "CM"],
  [500, "D"],
  [400, "CD"],
  [100, "C"],
  [90, "XC"],
  [50, "L"],
  [40, "XL"],
  [10, "X"],
  [9, "IX"],
  [5, "V"],
  [4, "IV"],
  [1, "I"],
];

export const roman = (number: number): string => {
  if (!Number.isInteger(number) || number < 1 || number > 3999) return String(number);
  let remaining = number;
  let output = "";
  for (const [value, symbol] of romanDigits) {
    output += symbol.repeat(Math.floor(remaining / value));
    remaining %= value;
  }
  return output;
};
