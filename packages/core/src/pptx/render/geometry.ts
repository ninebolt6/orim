import type { Geometry } from "../model";

type Adjust = (name: string, fallback: number) => number;
type Preset = (width: number, height: number, adjust: Adjust) => string;

const polygon = (points: readonly (readonly [number, number])[]): string =>
  `M${points.map((point) => point.join(" ")).join("L")}Z`;

const ellipseAt = (x: number, y: number, width: number, height: number): string =>
  `M${x} ${y + height / 2}A${width / 2} ${height / 2} 0 1 0 ${x + width} ${y + height / 2}A${width / 2} ${height / 2} 0 1 0 ${x} ${y + height / 2}Z`;

const ellipse: Preset = (width, height) => ellipseAt(0, 0, width, height);

const rounded = (width: number, height: number, radius: number, corners: readonly boolean[]) => {
  const [topLeft, topRight, bottomRight, bottomLeft] = corners.map((round) =>
    round ? Math.min(radius, width / 2, height / 2) : 0,
  ) as [number, number, number, number];
  const arc = (r: number, x: number, y: number) => (r ? `A${r} ${r} 0 0 1 ${x} ${y}` : "");
  return `M${topLeft} 0H${width - topRight}${arc(topRight, width, topRight)}V${height - bottomRight}${arc(bottomRight, width - bottomRight, height)}H${bottomLeft}${arc(bottomLeft, 0, height - bottomLeft)}V${topLeft}${arc(topLeft, topLeft, 0)}Z`;
};

const arrow =
  (direction: "right" | "left" | "up" | "down"): Preset =>
  (width, height, adjust) => {
    const horizontal = direction === "right" || direction === "left";
    const [long, short] = horizontal ? [width, height] : [height, width];
    const half = (short * adjust("adj1", 50000)) / 200000;
    const head = Math.min(long, (Math.min(width, height) * adjust("adj2", 50000)) / 100000);
    const points: [number, number][] = [
      [0, short / 2 - half],
      [long - head, short / 2 - half],
      [long - head, 0],
      [long, short / 2],
      [long - head, short],
      [long - head, short / 2 + half],
      [0, short / 2 + half],
    ];
    return polygon(
      points.map(([along, across]) =>
        direction === "right"
          ? [along, across]
          : direction === "left"
            ? [width - along, across]
            : direction === "down"
              ? [across, along]
              : [across, height - along],
      ),
    );
  };

const star =
  (count: number, ratio: number): Preset =>
  (width, height) =>
    polygon(
      Array.from({ length: count * 2 }, (_, index) => {
        const radius = index % 2 ? ratio : 1;
        const theta = (Math.PI * index) / count - Math.PI / 2;
        return [
          (width / 2) * (1 + radius * Math.cos(theta)),
          (height / 2) * (1 + radius * Math.sin(theta)),
        ];
      }),
    );

const presets: Readonly<Record<string, Preset>> = {
  rect: (width, height) => `M0 0H${width}V${height}H0Z`,
  flowChartProcess: (width, height) => `M0 0H${width}V${height}H0Z`,
  roundRect: (width, height, adjust) =>
    rounded(width, height, (Math.min(width, height) * adjust("adj", 16667)) / 100000, [
      true,
      true,
      true,
      true,
    ]),
  round2SameRect: (width, height, adjust) =>
    rounded(width, height, (Math.min(width, height) * adjust("adj1", 16667)) / 100000, [
      true,
      true,
      false,
      false,
    ]),
  flowChartTerminator: (width, height) =>
    rounded(width, height, Math.min(width, height) / 2, [true, true, true, true]),
  ellipse,
  flowChartConnector: ellipse,
  line: (width, height) => `M0 0L${width} ${height}`,
  straightConnector1: (width, height) => `M0 0L${width} ${height}`,
  bentConnector2: (width, height) => `M0 0H${width}V${height}`,
  bentConnector3: (width, height, adjust) => {
    const x = (width * adjust("adj1", 50000)) / 100000;
    return `M0 0H${x}V${height}H${width}`;
  },
  curvedConnector3: (width, height, adjust) => {
    const x = (width * adjust("adj1", 50000)) / 100000;
    return `M0 0C${x} 0 ${x} ${height} ${width} ${height}`;
  },
  triangle: (width, height, adjust) =>
    polygon([
      [(width * adjust("adj", 50000)) / 100000, 0],
      [width, height],
      [0, height],
    ]),
  rtTriangle: (width, height) =>
    polygon([
      [0, 0],
      [width, height],
      [0, height],
    ]),
  flowChartDecision: (width, height) =>
    polygon([
      [width / 2, 0],
      [width, height / 2],
      [width / 2, height],
      [0, height / 2],
    ]),
  diamond: (width, height) =>
    polygon([
      [width / 2, 0],
      [width, height / 2],
      [width / 2, height],
      [0, height / 2],
    ]),
  parallelogram: (width, height, adjust) => {
    const offset = (Math.min(width, height) * adjust("adj", 25000)) / 100000;
    return polygon([
      [offset, 0],
      [width, 0],
      [width - offset, height],
      [0, height],
    ]);
  },
  trapezoid: (width, height, adjust) => {
    const offset = (Math.min(width, height) * adjust("adj", 25000)) / 100000;
    return polygon([
      [0, height],
      [offset, 0],
      [width - offset, 0],
      [width, height],
    ]);
  },
  pentagon: (width, height) =>
    polygon([
      [width / 2, 0],
      [width, height * 0.382],
      [width * 0.809, height],
      [width * 0.191, height],
      [0, height * 0.382],
    ]),
  hexagon: (width, height, adjust) => {
    const offset = (Math.min(width, height) * adjust("adj", 25000)) / 100000;
    return polygon([
      [offset, 0],
      [width - offset, 0],
      [width, height / 2],
      [width - offset, height],
      [offset, height],
      [0, height / 2],
    ]);
  },
  octagon: (width, height, adjust) => {
    const offset = (Math.min(width, height) * adjust("adj", 29289)) / 100000;
    return polygon([
      [offset, 0],
      [width - offset, 0],
      [width, offset],
      [width, height - offset],
      [width - offset, height],
      [offset, height],
      [0, height - offset],
      [0, offset],
    ]);
  },
  homePlate: (width, height, adjust) => {
    const offset = (Math.min(width, height) * adjust("adj", 50000)) / 100000;
    return polygon([
      [0, 0],
      [width - offset, 0],
      [width, height / 2],
      [width - offset, height],
      [0, height],
    ]);
  },
  chevron: (width, height, adjust) => {
    const offset = (Math.min(width, height) * adjust("adj", 50000)) / 100000;
    return polygon([
      [0, 0],
      [width - offset, 0],
      [width, height / 2],
      [width - offset, height],
      [0, height],
      [offset, height / 2],
    ]);
  },
  snip1Rect: (width, height, adjust) => {
    const offset = (Math.min(width, height) * adjust("adj", 16667)) / 100000;
    return polygon([
      [0, 0],
      [width - offset, 0],
      [width, offset],
      [width, height],
      [0, height],
    ]);
  },
  plus: (width, height, adjust) => {
    const offset = (Math.min(width, height) * adjust("adj", 25000)) / 100000;
    return polygon([
      [offset, 0],
      [width - offset, 0],
      [width - offset, offset],
      [width, offset],
      [width, height - offset],
      [width - offset, height - offset],
      [width - offset, height],
      [offset, height],
      [offset, height - offset],
      [0, height - offset],
      [0, offset],
      [offset, offset],
    ]);
  },
  donut: (width, height, adjust) => {
    const inset = (Math.min(width, height) * adjust("adj", 25000)) / 100000;
    return (
      ellipse(width, height, adjust) +
      ellipseAt(inset, inset, width - inset * 2, height - inset * 2)
    );
  },
  rightArrow: arrow("right"),
  leftArrow: arrow("left"),
  upArrow: arrow("up"),
  downArrow: arrow("down"),
  star4: star(4, 0.25),
  star5: star(5, 0.382),
  star6: star(6, 0.5),
};

export const isLinePreset = (preset: string): boolean => /^(?:line|.*Connector\d?)$/.test(preset);

/** SVG path data for `geometry` in a `width` × `height` point box. */
export const geometryPaths = (
  geometry: Geometry,
  width: number,
  height: number,
): readonly { readonly data: string; readonly fill: boolean; readonly stroke: boolean }[] => {
  if (geometry.paths?.length) {
    return geometry.paths.map((path) => {
      const scaleX = path.width ? width / path.width : 1;
      const scaleY = path.height ? height / path.height : 1;
      const data = path.commands
        .map(
          (command) =>
            command.type +
            command.values
              .map((value, index) =>
                command.type === "A" && index >= 2 && index <= 4
                  ? value
                  : value *
                    ((command.type === "A" && index > 4 ? index - 5 : index) % 2 ? scaleY : scaleX),
              )
              .join(" "),
        )
        .join("");
      return { data, fill: path.fill, stroke: path.stroke };
    });
  }
  const preset = presets[geometry.preset] ?? presets.rect!;
  const adjust: Adjust = (name, fallback) =>
    geometry.adjust[name] ?? (name === "adj1" ? geometry.adjust.adj : undefined) ?? fallback;
  return [
    {
      data: preset(width, height, adjust),
      fill: !isLinePreset(geometry.preset),
      stroke: true,
    },
  ];
};
