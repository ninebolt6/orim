import { attr, children, elements, first, xmlBoolean } from "../../common/xml";
import { angle, emu } from "../../drawingml/units";
import type { Geometry, GeometryPath, PathCommand, Transform } from "../model";

export const parseTransform = (xfrm: Element | undefined): Transform | undefined => {
  const offset = first(xfrm, "off");
  const extent = first(xfrm, "ext");
  if (!offset || !extent) return undefined;
  return {
    x: emu(attr(offset, "x")) ?? 0,
    y: emu(attr(offset, "y")) ?? 0,
    width: emu(attr(extent, "cx")) ?? 0,
    height: emu(attr(extent, "cy")) ?? 0,
    rotation: angle(attr(xfrm, "rot")) ?? 0,
    flipH: xmlBoolean(attr(xfrm, "flipH")) ?? false,
    flipV: xmlBoolean(attr(xfrm, "flipV")) ?? false,
  };
};

const adjustValues = (list: Element | undefined): Readonly<Record<string, number>> =>
  Object.fromEntries(
    elements(list, "gd").flatMap((guide) => {
      const match = /^val\s+(-?\d+)$/.exec(attr(guide, "fmla") ?? "");
      return match ? [[attr(guide, "name") ?? "", Number(match[1])] as const] : [];
    }),
  );

const point = (element: Element | undefined): readonly [number, number] => [
  Number(attr(element, "x")) || 0,
  Number(attr(element, "y")) || 0,
];

const customPath = (path: Element): GeometryPath => {
  const width = Number(attr(path, "w")) || 0;
  const height = Number(attr(path, "h")) || 0;
  let current: readonly [number, number] = [0, 0];
  const commands = children(path).flatMap((command): PathCommand[] => {
    const points = elements(command, "pt").map(point);
    switch (command.localName) {
      case "moveTo":
      case "lnTo":
        current = points[0] ?? current;
        return [{ type: command.localName === "moveTo" ? "M" : "L", values: current }];
      case "cubicBezTo":
      case "quadBezTo":
        current = points.at(-1) ?? current;
        return [{ type: command.localName === "cubicBezTo" ? "C" : "Q", values: points.flat() }];
      case "arcTo": {
        const radiusX = Number(attr(command, "wR")) || 0;
        const radiusY = Number(attr(command, "hR")) || 0;
        const start = ((angle(attr(command, "stAng")) ?? 0) * Math.PI) / 180;
        const sweep = angle(attr(command, "swAng")) ?? 0;
        const end = start + (sweep * Math.PI) / 180;
        const centerX = current[0] - radiusX * Math.cos(start);
        const centerY = current[1] - radiusY * Math.sin(start);
        current = [centerX + radiusX * Math.cos(end), centerY + radiusY * Math.sin(end)];
        const large = Math.abs(sweep) > 180 ? 1 : 0;
        return [{ type: "A", values: [radiusX, radiusY, 0, large, sweep > 0 ? 1 : 0, ...current] }];
      }
      case "close":
        return [{ type: "Z", values: [] }];
    }
    return [];
  });
  return {
    width,
    height,
    commands,
    fill: attr(path, "fill") !== "none",
    stroke: attr(path, "stroke") !== "0" && attr(path, "stroke") !== "false",
  };
};

/** Reads `a:prstGeom`/`a:custGeom` from `spPr`; `undefined` when neither is present. */
export const parseGeometry = (properties: Element | undefined): Geometry | undefined => {
  const preset = first(properties, "prstGeom");
  if (preset)
    return { preset: attr(preset, "prst") ?? "rect", adjust: adjustValues(first(preset, "avLst")) };
  const custom = first(properties, "custGeom");
  if (!custom) return undefined;
  return {
    preset: "custom",
    adjust: {},
    paths: elements(first(custom, "pathLst"), "path").map(customPath),
  };
};
