import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { BorderEdgeModel, BorderModel, ColorModel } from "../../../types";
import type { BorderStyle } from "../../../../index";
import BaseXform, { isOneOf, ownEntry } from "../base-xform";
import { parseBoolean } from "../../../utils/utils";

import ColorXform from "./color-xform";

// the border styles of ECMA-376 ST_BorderStyle (except none), as a set of keys
const VALID_STYLE_VALUES: Readonly<Record<BorderStyle, true>> = {
    thin: true,
    dashed: true,
    dotted: true,
    dashDot: true,
    hair: true,
    dashDotDot: true,
    slantDashDot: true,
    mediumDashed: true,
    mediumDashDotDot: true,
    mediumDashDot: true,
    medium: true,
    double: true,
    thick: true,
};

/**
 * One edge of a border (`<left>`, `<right>`, `<top>`, `<bottom>`, `<diagonal>`): its style attribute and its color.
 * The color is written only with a style. Parsing gives `undefined` for an edge with neither style nor color (a style
 * `none` or not one of ST_BorderStyle is no style).
 */
class EdgeXform extends BaseXform<BorderEdgeModel> {
    /** The border styles of ST_BorderStyle, as a set of keys. Unused by the library. */
    static readonly validStyleValues = VALID_STYLE_VALUES;

    /** The name of the element. */
    readonly name: string;
    /** A default color of the edge; never set by the library. */
    declare defaultColor?: ColorModel;

    declare map: { readonly color: ColorXform };

    constructor(name: string) {
        super();

        this.name = name;
        this.map = {
            color: new ColorXform(),
        };
    }

    get tag(): string {
        return this.name;
    }

    /** @param defaultColor - the color written when the edge has a style but no color */
    override render(xmlStream: XmlStream, model: BorderEdgeModel | undefined, defaultColor?: ColorModel): void {
        const color = (model && model.color) || defaultColor || this.defaultColor;
        xmlStream.openNode(this.name);
        if (model && model.style) {
            xmlStream.addAttribute("style", model.style);
            if (color) {
                this.map.color.render(xmlStream, color);
            }
        }
        xmlStream.closeNode();
    }

    /** @returns `true` for its element and its `<color>` */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case this.name: {
                // none (the default) or an invalid style: no style
                const { style } = node.attributes;
                this.model = isOneOf(VALID_STYLE_VALUES, style) ? { style } : undefined;
                return true;
            }
            case "color":
                this.parser = this.map.color;
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
                this.parser = undefined;
            }
            return true;
        }

        if (name === this.name) {
            const color = this.map.color.model;
            if (color) {
                const model = this.model || (this.model = {});
                model.color = color;
            }
        }

        return false;
    }

    /** @returns `true` for a style of ST_BorderStyle, else `undefined`. Unused by the library. */
    validStyle(value: string): true | undefined {
        return isOneOf(VALID_STYLE_VALUES, value) ? true : undefined;
    }
}

/** The extensions of the diagonal edge of a parsed border. */
type DiagonalExtensions = { up?: boolean; down?: boolean };

/**
 * A border of the styles part, `<border>` with its edges in the CT_Border order (left, right, top, bottom, diagonal),
 * see {@link BorderModel}. `diagonalUp` / `diagonalDown` are written when the diagonal has a style.
 */
class BorderXform extends BaseXform<BorderModel> {
    declare map: {
        readonly top: EdgeXform;
        readonly left: EdgeXform;
        readonly bottom: EdgeXform;
        readonly right: EdgeXform;
        readonly diagonal: EdgeXform;
    };

    /** While parsing: the `diagonalUp` attribute of the current border. */
    declare diagonalUp?: boolean;
    /** While parsing: the `diagonalDown` attribute of the current border. */
    declare diagonalDown?: boolean;

    constructor() {
        super();

        this.map = {
            top: new EdgeXform("top"),
            left: new EdgeXform("left"),
            bottom: new EdgeXform("bottom"),
            right: new EdgeXform("right"),
            diagonal: new EdgeXform("diagonal"),
        };
    }

    override render(xmlStream: XmlStream, model: BorderModel): void {
        const { color } = model;
        xmlStream.openNode("border");
        if (model.diagonal && model.diagonal.style) {
            if (model.diagonal.up) {
                xmlStream.addAttribute("diagonalUp", "1");
            }
            if (model.diagonal.down) {
                xmlStream.addAttribute("diagonalDown", "1");
            }
        }
        function add(edgeModel: BorderEdgeModel | undefined, edgeXform: EdgeXform): void {
            if (edgeModel && !edgeModel.color && model.color) {
                // don't mess with incoming models
                edgeModel = {
                    ...edgeModel,
                    color: model.color,
                };
            }
            edgeXform.render(xmlStream, edgeModel, color);
        }
        add(model.left, this.map.left);
        add(model.right, this.map.right);
        add(model.top, this.map.top);
        add(model.bottom, this.map.bottom);
        add(model.diagonal, this.map.diagonal);

        xmlStream.closeNode();
    }

    /** @returns `true` for `<border>` and its edges */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "border":
                this.reset();
                this.diagonalUp = parseBoolean(node.attributes.diagonalUp);
                this.diagonalDown = parseBoolean(node.attributes.diagonalDown);
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
                this.parser = undefined;
            }
            return true;
        }
        if (name === "border") {
            const model: BorderModel = (this.model = {});
            const add = function (
                key: "left" | "right" | "top" | "bottom" | "diagonal",
                edgeModel: BorderEdgeModel | null | undefined,
                extensions?: DiagonalExtensions
            ): void {
                if (edgeModel) {
                    if (extensions) {
                        Object.assign(edgeModel, extensions);
                    }
                    model[key] = edgeModel;
                }
            };
            add("left", this.map.left.model);
            add("right", this.map.right.model);
            add("top", this.map.top.model);
            add("bottom", this.map.bottom.model);
            // the attributes were read when <border> opened
            add("diagonal", this.map.diagonal.model, { up: this.diagonalUp, down: this.diagonalDown });
        }
        return false;
    }
}

export default BorderXform;
