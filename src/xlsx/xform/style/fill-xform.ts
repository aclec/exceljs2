import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { FillGradientPath, FillPatterns } from "../../../../index";
import type { FillModel, GradientStopModel } from "../../../types";
import BaseXform, { ownEntry } from "../base-xform";

import ColorXform from "./color-xform";

/** A pattern fill model. */
type PatternFillModel = Extract<FillModel, { type: "pattern" }>;
/** A gradient fill model. */
type GradientFillModel = Extract<FillModel, { type: "gradient" }>;
/**
 * A gradient fill model as GradientFillXform parses it: without `type`, which FillXform adds; `degree` is set with
 * `gradient: "angle"`, `center` with `gradient: "path"`.
 */
interface ParsedGradientFillModel {
    type?: "gradient";
    gradient?: "angle" | "path";
    degree?: number;
    center?: FillGradientPath["center"];
    stops: GradientStopModel[];
}

/**
 * The model of a {@link FillXform}: a fill, or `{}` (a `<fill>` without a known child parses as `{}`; rendering a model
 * without a known `type` writes nothing).
 */
export type FillXformModel = FillModel | { type?: undefined };

/** A gradient stop, `<stop position><color/></stop>`. */
class StopXform extends BaseXform<GradientStopModel> {
    declare map: { readonly color: ColorXform };

    constructor() {
        super();

        this.map = {
            color: new ColorXform(),
        };
    }

    get tag(): string {
        return "stop";
    }

    override render(xmlStream: XmlStream, model: GradientStopModel): void {
        xmlStream.openNode("stop");
        xmlStream.addAttribute("position", model.position);
        this.map.color.render(xmlStream, model.color);
        xmlStream.closeNode();
    }

    /** @returns `true` for `<stop>` and its `<color>` */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "stop":
                this.model = {
                    position: BaseXform.toFloatValue(node.attributes.position, NaN),
                };
                return true;
            case "color":
                this.parser = this.map.color;
                this.parser.parseOpen(node);
                return true;
            default:
                return false;
        }
    }

    override parseText(): void {}

    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // <color> is inside <stop>, which created the model; an auto color is undefined
                this.model!.color = this.map.color.model ?? undefined;
                this.parser = undefined;
            }
            return true;
        }
        return false;
    }
}

/** A pattern fill, `<patternFill patternType>` with its `<fgColor>` / `<bgColor>`. */
class PatternFillXform extends BaseXform<PatternFillModel> {
    declare map: { readonly fgColor: ColorXform; readonly bgColor: ColorXform };

    constructor() {
        super();

        this.map = {
            fgColor: new ColorXform("fgColor"),
            bgColor: new ColorXform("bgColor"),
        };
    }

    /** The `type` of the fill model. */
    get name(): "pattern" {
        return "pattern";
    }

    get tag(): string {
        return "patternFill";
    }

    override render(xmlStream: XmlStream, model: PatternFillModel): void {
        xmlStream.openNode("patternFill");
        xmlStream.addAttribute("patternType", model.pattern);
        if (model.fgColor) {
            this.map.fgColor.render(xmlStream, model.fgColor);
        }
        if (model.bgColor) {
            this.map.bgColor.render(xmlStream, model.bgColor);
        }
        xmlStream.closeNode();
    }

    /** @returns `true` for `<patternFill>` and its colors */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "patternFill":
                this.model = {
                    type: "pattern",
                    // unvalidated (and undefined when the attribute is absent)
                    pattern: node.attributes.patternType as FillPatterns,
                };
                return true;
            default: {
                const parser = ownEntry(this.map, node.name);
                this.parser = parser;
                if (parser) {
                    parser.parseOpen(node);
                    return true;
                }
                return false;
            }
        }
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // the colors are inside <patternFill>, which created the model
                const color = this.parser === this.map.fgColor ? this.map.fgColor.model : this.map.bgColor.model;
                if (color) {
                    this.model![this.parser === this.map.fgColor ? "fgColor" : "bgColor"] = color;
                }
                this.parser = undefined;
            }
            return true;
        }
        return false;
    }
}

/** A gradient fill, `<gradientFill degree | type="path" left right top bottom>` with its `<stop>`s. */
class GradientFillXform extends BaseXform<ParsedGradientFillModel> {
    declare map: { readonly stop: StopXform };

    constructor() {
        super();

        this.map = {
            stop: new StopXform(),
        };
    }

    /** The `type` of the fill model. */
    get name(): "gradient" {
        return "gradient";
    }

    get tag(): string {
        return "gradientFill";
    }

    override render(xmlStream: XmlStream, model: GradientFillModel): void {
        xmlStream.openNode("gradientFill");
        switch (model.gradient) {
            case "angle":
                xmlStream.addAttribute("degree", model.degree);
                break;
            case "path": {
                xmlStream.addAttribute("type", "path");
                const { center } = model;
                if (center) {
                    if (center.left) {
                        xmlStream.addAttribute("left", center.left);
                        if (center.right === undefined) {
                            xmlStream.addAttribute("right", center.left);
                        }
                    }
                    if (center.right) {
                        xmlStream.addAttribute("right", center.right);
                    }
                    if (center.top) {
                        xmlStream.addAttribute("top", center.top);
                        if (center.bottom === undefined) {
                            xmlStream.addAttribute("bottom", center.top);
                        }
                    }
                    if (center.bottom) {
                        xmlStream.addAttribute("bottom", center.bottom);
                    }
                }
                break;
            }

            case undefined:
            default:
                break;
        }

        const stopXform = this.map.stop;
        for (const stopModel of model.stops) {
            stopXform.render(xmlStream, stopModel);
        }

        xmlStream.closeNode();
    }

    /** @returns `true` for `<gradientFill>` and its stops */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "gradientFill": {
                const { degree, type, left, right, top, bottom } = node.attributes;
                const model: ParsedGradientFillModel = {
                    stops: [],
                };
                if (degree) {
                    model.gradient = "angle";
                    model.degree = parseInt(degree, 10);
                } else if (type === "path") {
                    model.gradient = "path";
                    const center: FillGradientPath["center"] = {
                        left: left ? parseFloat(left) : 0,
                        top: top ? parseFloat(top) : 0,
                    };
                    if (right !== left) {
                        center.right = right ? parseFloat(right) : 0;
                    }
                    if (bottom !== top) {
                        center.bottom = bottom ? parseFloat(bottom) : 0;
                    }
                    model.center = center;
                }
                this.model = model;
                return true;
            }

            case "stop":
                this.parser = this.map.stop;
                this.parser.parseOpen(node);
                return true;

            default:
                return false;
        }
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // <stop> is inside <gradientFill>, which created the model
                this.model!.stops.push(this.map.stop.model!);
                this.parser = undefined;
            }
            return true;
        }
        return false;
    }
}

// the pattern types of ECMA-376 ST_PatternType (as in the original list, with its duplicates)
const VALID_PATTERN_VALUES: Readonly<Record<string, true>> = Object.assign(
    Object.create(null) as Record<string, true>,
    Object.fromEntries(
        [
            "none",
            "solid",
            "darkVertical",
            "darkGray",
            "mediumGray",
            "lightGray",
            "gray125",
            "gray0625",
            "darkHorizontal",
            "darkVertical",
            "darkDown",
            "darkUp",
            "darkGrid",
            "darkTrellis",
            "lightHorizontal",
            "lightVertical",
            "lightDown",
            "lightUp",
            "lightGrid",
            "lightTrellis",
            "lightGrid",
        ].map((v) => [v, true] as const)
    )
);

/**
 * A fill of the styles part, `<fill>` with a `<patternFill>` or a `<gradientFill>` (see {@link FillModel}). A model
 * without a known `type` renders nothing.
 */
class FillXform extends BaseXform<FillXformModel> {
    /** The pattern types (ST_PatternType), as a set of keys. Unused by the library. */
    static readonly validPatternValues = VALID_PATTERN_VALUES;

    /** The xform of a gradient stop. */
    static readonly StopXform = StopXform;
    /** The xform of a pattern fill. */
    static readonly PatternFillXform = PatternFillXform;
    /** The xform of a gradient fill. */
    static readonly GradientFillXform = GradientFillXform;

    declare map: { readonly patternFill: PatternFillXform; readonly gradientFill: GradientFillXform };

    constructor() {
        super();

        this.map = {
            patternFill: new PatternFillXform(),
            gradientFill: new GradientFillXform(),
        };
    }

    get tag(): string {
        return "fill";
    }

    override render(xmlStream: XmlStream, model: FillXformModel): void {
        switch (model.type) {
            case "pattern":
                xmlStream.openNode("fill");
                this.map.patternFill.render(xmlStream, model);
                xmlStream.closeNode();
                break;
            case "gradient":
                xmlStream.openNode("fill");
                this.map.gradientFill.render(xmlStream, model);
                xmlStream.closeNode();
                break;
            case undefined:
            default:
                // a fill of an unknown type is not written
                break;
        }
    }

    /** @returns `true` for `<fill>` and its pattern or gradient */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "fill":
                this.model = {};
                return true;
            default: {
                const parser = ownEntry(this.map, node.name);
                this.parser = parser;
                if (parser) {
                    parser.parseOpen(node);
                    return true;
                }
                return false;
            }
        }
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // the fill is the model of its child, with the type of the child
                // (its model was created when its element opened)
                if (this.parser === this.map.patternFill) {
                    this.model = Object.assign(this.map.patternFill.model!, { type: this.map.patternFill.name });
                } else {
                    // a parsed path gradient has its center (see ParsedGradientFillModel)
                    this.model = Object.assign(this.map.gradientFill.model!, { type: this.map.gradientFill.name }) as GradientFillModel;
                }
                this.parser = undefined;
            }
            return true;
        }
        return false;
    }

    /** @returns `true` for a pattern type of ST_PatternType, else `undefined`. Unused by the library. */
    validStyle(value: string): true | undefined {
        return VALID_PATTERN_VALUES[value];
    }
}

export default FillXform;
