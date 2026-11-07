import { describe, expect, it } from "vite-plus/test";
import { geometryPaths, isLinePreset } from "./geometry";

describe("PPTX geometry", () => {
  it("builds preset paths with adjustment values", () => {
    expect(geometryPaths({ preset: "rect", adjust: {} }, 10, 20)).toEqual([
      { data: "M0 0H10V20H0Z", fill: true, stroke: true },
    ]);
    expect(geometryPaths({ preset: "triangle", adjust: { adj: 0 } }, 10, 20)[0]!.data).toBe(
      "M0 0L10 20L0 20Z",
    );
    expect(geometryPaths({ preset: "roundRect", adjust: {} }, 100, 50)[0]!.data).toContain(
      "A8.3335 8.3335",
    );
    expect(geometryPaths({ preset: "unknownShape", adjust: {} }, 1, 1)[0]!.data).toBe(
      "M0 0H1V1H0Z",
    );
  });

  it("does not fill connector presets", () => {
    expect(isLinePreset("straightConnector1")).toBe(true);
    expect(isLinePreset("bentConnector3")).toBe(true);
    expect(isLinePreset("rect")).toBe(false);
    expect(geometryPaths({ preset: "line", adjust: {} }, 5, 5)[0]).toMatchObject({ fill: false });
  });

  it("scales custom paths, including arc radii and endpoints", () => {
    const [path] = geometryPaths(
      {
        preset: "rect",
        adjust: {},
        paths: [
          {
            width: 10,
            height: 10,
            fill: true,
            stroke: false,
            commands: [
              { type: "M", values: [0, 5] },
              { type: "A", values: [5, 5, 0, 0, 1, 10, 5] },
              { type: "Z", values: [] },
            ],
          },
        ],
      },
      20,
      40,
    );
    expect(path).toEqual({ data: "M0 20A10 20 0 0 1 20 20Z", fill: true, stroke: false });
  });
});
