import type { CellErrorValue } from "../../../../index";
import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { XmlAttributes, XmlValue } from "../../../utils/xml-stream";
import type {
    CellModel,
    CellPrepareOptions,
    CellReconcileOptions,
    FormulaCellModel,
    HyperlinkCellModel,
    ParsedCellModel,
    RichTextValue,
    SharedStringValue,
} from "../../../types";
import utils from "../../../utils/utils";
import { parseIsoDate } from "../../../utils/iso-date";
import BaseXform from "../base-xform";
import Range from "../../../doc/range";
import { ValueType } from "../../../doc/enums";

import RichTextXform from "../strings/rich-text-xform";

/**
 * The type of a value (of a formula result), told as Cell tells it (src/doc/cell.ts): Null for null / undefined; the
 * markers of an object value by their presence and type (an empty text, formula or error code, a shared string 0 are
 * values).
 *
 * @throws Error for a value of no known type
 */
function getValueType(v: unknown): ValueType {
    if (v === null || v === undefined) {
        return ValueType.Null;
    }
    if (v instanceof String || typeof v === "string") {
        return ValueType.String;
    }
    if (typeof v === "number") {
        return ValueType.Number;
    }
    if (typeof v === "boolean") {
        return ValueType.Boolean;
    }
    if (v instanceof Date) {
        return ValueType.Date;
    }
    if (typeof v === "object") {
        if ("hyperlink" in v && typeof v.hyperlink === "string" && "text" in v && v.text !== undefined && v.text !== null) {
            return ValueType.Hyperlink;
        }
        if (("formula" in v && typeof v.formula === "string") || ("sharedFormula" in v && typeof v.sharedFormula === "string")) {
            return ValueType.Formula;
        }
        if ("richText" in v && Array.isArray(v.richText)) {
            return ValueType.RichText;
        }
        if ("sharedString" in v && typeof v.sharedString === "number") {
            return ValueType.SharedString;
        }
        if ("error" in v && typeof v.error === "string") {
            return ValueType.Error;
        }
    }
    throw new Error("I could not understand type of value");
}

/**
 * The type of a cell for its number format: the type of the result of a formula, Date for a hyperlink on a date (both are
 * written as serial numbers: without a number format of their own, they get the default date format).
 */
function getEffectiveCellType(cell: CellModel): ValueType {
    if (cell.type === ValueType.Formula) {
        return getValueType(cell.result);
    }
    if (cell.type === ValueType.Hyperlink && cell.text instanceof Date) {
        return ValueType.Date;
    }
    return cell.type;
}

function isRichText(value: unknown): value is RichTextValue {
    return typeof value === "object" && value !== null && "richText" in value && !!value.richText;
}

/** An error value (`{ error }`: told by the presence and type of `error`). */
function isErrorValue(value: unknown): value is CellErrorValue {
    return typeof value === "object" && value !== null && "error" in value && typeof value.error === "string";
}

/**
 * The shared string of the text of a hyperlink: a string or a rich text. `undefined` for a number, Date, boolean or error
 * text, written with its own type.
 */
function sharedStringText(text: HyperlinkCellModel["text"] | null | undefined): SharedStringValue | undefined {
    return typeof text === "string" || isRichText(text) ? text : undefined;
}

/** A value converted to text as XmlStream and the string functions do (`toString()`): any value but null / undefined. */
interface Stringifiable {
    toString(): string;
}

/**
 * The text of a `<v>`: what XmlStream writes for the value (its `toString()`: "[object Object]" for an error or rich text
 * value without shared strings), an empty element for `undefined`.
 */
function valueText(value: XmlValue | Stringifiable | undefined): XmlValue | undefined {
    return value === undefined || typeof value === "string" || typeof value === "number" || typeof value === "boolean"
        ? value
        : value.toString();
}

/**
 * The text of the `<c>` being parsed: a parsed value is text, unless `<r>` runs made it rich text (malformed with a type
 * other than inlineStr: then "[object Object]", as the string functions read it).
 */
function rawText(value: Stringifiable): string {
    return typeof value === "string" ? value : value.toString();
}

/**
 * A cell, `<c r="A1" s="1" t="…">` with its value (`<v>`), formula (`<f>`) or inline string (`<is>`). The hottest xform
 * of a workbook: it runs for every cell written (prepare, render) and read (parse, reconcile).
 *
 * Write: {@link prepare} registers the style, the shared string, the hyperlink, the note, the merge and the shared formula
 * of the cell (in `options`), then {@link render} writes it (nothing for a Null cell without style). Read: the parse gives
 * a {@link ParsedCellModel} (raw value), which {@link reconcile} resolves into a CellModel (shared string, date format,
 * shared formula, hyperlink, note).
 */
class CellXform extends BaseXform<ParsedCellModel | CellModel> {
    /** The parsed cell (prepare and render take the document {@link CellModel}: see the type parameter). */
    declare model: ParsedCellModel | null | undefined;

    /** The xform of the runs of an inline rich string. */
    readonly richTextXForm = new RichTextXform();

    /** While parsing: the `t` attribute (the type of the value) of the cell. */
    declare t: string | undefined;
    /** While parsing: whether the cell has a `<v>` (an empty one is an empty string result). */
    declare hasValueNode: boolean;
    /** While parsing: the element whose text is being read. */
    declare currentNode: "f" | "v" | "t" | undefined;

    get tag(): string {
        return "c";
    }

    /**
     * Registers what the cell needs in the workbook and worksheet (mutates `options` and the model: `styleId`, `ssId`,
     * `date1904`, `si`; a shared formula master gets `shareType` / `si` / `range` from its first clone).
     *
     * @throws Error for a shared formula clone whose master is not above or left of it (not prepared yet)
     */
    override prepare(model: CellModel, options: CellPrepareOptions): void {
        const styleId = options.styles.addStyleModel(model.style || {}, getEffectiveCellType(model));
        if (styleId) {
            model.styleId = styleId;
        }

        if (model.comment) {
            options.comments.push({ ...model.comment, ref: model.address });
        }

        switch (model.type) {
            case ValueType.String:
            case ValueType.RichText:
                if (options.sharedStrings) {
                    model.ssId = options.sharedStrings.add(model.value);
                }
                break;

            case ValueType.Date:
                if (options.date1904) {
                    model.date1904 = true;
                }
                break;

            case ValueType.Hyperlink:
                // a link on a number, Date, boolean or error cell (as read from a file) keeps its value type
                if (options.date1904 && model.text instanceof Date) {
                    model.date1904 = true;
                }
                if (options.sharedStrings) {
                    const text = sharedStringText(model.text);
                    if (text !== undefined) {
                        model.ssId = options.sharedStrings.add(text);
                    }
                }
                options.hyperlinks.push({
                    address: model.address,
                    target: model.hyperlink,
                    tooltip: model.tooltip,
                });
                break;

            case ValueType.Merge:
                options.merges.add(model);
                break;

            case ValueType.Formula:
                if (options.date1904) {
                    // in case valueType is date
                    model.date1904 = true;
                }

                if (model.shareType === "shared") {
                    model.si = options.siFormulae++;
                }

                if (model.formula) {
                    options.formulae[model.address] = model;
                } else if (model.sharedFormula) {
                    const master = options.formulae[model.sharedFormula];
                    if (!master) {
                        throw new Error(`Shared Formula master must exist above and or left of clone for cell ${model.address}`);
                    }
                    if (master.si === undefined) {
                        master.shareType = "shared";
                        master.si = options.siFormulae++;
                        master.range = new Range(master.address, model.address);
                    } else if (master.range) {
                        master.range.expandToAddress(model.address);
                    }
                    model.si = master.si;
                }
                break;

            case ValueType.Null:
            case ValueType.Number:
            case ValueType.SharedString:
            case ValueType.Boolean:
            case ValueType.Error:
            default:
                break;
        }
    }

    /**
     * Writes the `<f>` and `<v>` of a formula cell, and the `t` attribute of its result. A cell without formula text but a
     * clone of a shared formula (an orphan clone read from a file, an empty formula) is written as its cached value: an
     * empty `<f>` (CT_CellFormula) is the formula of a shared formula clone only.
     *
     * @throws Error for a result that is not a string, number, boolean, Date or error value, or a shared formula master
     *   without `ref` (or clones)
     */
    renderFormula(xmlStream: XmlStream, model: FormulaCellModel): void {
        let attrs: XmlAttributes | undefined;
        // a clone of a shared formula has no formula text of its own
        let clone = false;
        switch (model.shareType) {
            case "shared": {
                const ref = model.ref || (model.range && model.range.range);
                if (!ref) {
                    throw new Error(`The shared formula of cell ${model.address} needs a ref`);
                }
                attrs = {
                    t: "shared",
                    ref,
                    si: model.si,
                };
                break;
            }

            case "array":
                attrs = {
                    t: "array",
                    ref: model.ref,
                };
                break;

            case undefined:
            default:
                if (model.si !== undefined) {
                    clone = true;
                    attrs = {
                        t: "shared",
                        si: model.si,
                    };
                }
                break;
        }

        const { result } = model;
        // the type (t attribute) and text of the cached value
        let type: "str" | "b" | "e" | undefined;
        let value: XmlValue | undefined;
        if (result === null || result === undefined) {
            value = undefined;
        } else if (result instanceof String || typeof result === "string") {
            // oddly, formula results don't ever use shared strings
            type = "str";
            value = result.toString();
        } else if (typeof result === "number") {
            value = result;
        } else if (typeof result === "boolean") {
            type = "b";
            value = result ? 1 : 0;
        } else if (result instanceof Date) {
            value = utils.dateToExcel(result, model.date1904);
        } else if (getValueType(result) === ValueType.Error) {
            // getValueType throws for a value of no known type
            type = "e";
            value = result.error;
        } else {
            // a hyperlink, formula, rich text or shared string value
            throw new Error(`Cell ${model.address}: the result of a formula must be a string, number, boolean, Date or error value`);
        }

        if (type !== undefined) {
            xmlStream.addAttribute("t", type);
        }
        if (clone || (model.formula !== undefined && model.formula !== "")) {
            xmlStream.leafNode("f", attrs, model.formula);
        }
        if (value !== undefined) {
            xmlStream.leafNode("v", undefined, value);
        }
    }

    /** Writes the cell (nothing for a Null cell without style). */
    override render(xmlStream: XmlStream, model: CellModel): void {
        if (model.type === ValueType.Null && !model.styleId) {
            // if null and no style, exit
            return;
        }

        xmlStream.openNode("c");
        xmlStream.addAttribute("r", model.address);

        if (model.styleId) {
            xmlStream.addAttribute("s", model.styleId);
        }

        switch (model.type) {
            case ValueType.Null:
                break;

            case ValueType.Number:
                xmlStream.leafNode("v", undefined, model.value);
                break;

            case ValueType.Boolean:
                xmlStream.addAttribute("t", "b");
                xmlStream.leafNode("v", undefined, model.value ? "1" : "0");
                break;

            case ValueType.Error:
                xmlStream.addAttribute("t", "e");
                xmlStream.leafNode("v", undefined, model.value.error);
                break;

            case ValueType.String:
            case ValueType.RichText: {
                const { value } = model;
                if (model.ssId !== undefined) {
                    xmlStream.addAttribute("t", "s");
                    xmlStream.leafNode("v", undefined, model.ssId);
                } else if (isRichText(value)) {
                    xmlStream.addAttribute("t", "inlineStr");
                    xmlStream.openNode("is");
                    value.richText.forEach((text) => {
                        this.richTextXForm.render(xmlStream, text);
                    });
                    xmlStream.closeNode();
                } else {
                    xmlStream.addAttribute("t", "str");
                    xmlStream.leafNode("v", undefined, valueText(value));
                }
                break;
            }

            case ValueType.Date:
                xmlStream.leafNode("v", undefined, utils.dateToExcel(model.value, model.date1904));
                break;

            case ValueType.Hyperlink:
                if (typeof model.text === "number") {
                    xmlStream.leafNode("v", undefined, model.text);
                } else if (model.text instanceof Date) {
                    xmlStream.leafNode("v", undefined, utils.dateToExcel(model.text, model.date1904));
                } else if (typeof model.text === "boolean") {
                    xmlStream.addAttribute("t", "b");
                    xmlStream.leafNode("v", undefined, model.text ? "1" : "0");
                } else if (isErrorValue(model.text)) {
                    xmlStream.addAttribute("t", "e");
                    xmlStream.leafNode("v", undefined, model.text.error);
                } else if (model.ssId !== undefined) {
                    xmlStream.addAttribute("t", "s");
                    xmlStream.leafNode("v", undefined, model.ssId);
                } else {
                    xmlStream.addAttribute("t", "str");
                    xmlStream.leafNode("v", undefined, valueText(model.text));
                }
                break;

            case ValueType.Formula:
                this.renderFormula(xmlStream, model);
                break;

            case ValueType.Merge:
                // nothing to add
                break;

            case ValueType.SharedString:
            default:
                break;
        }

        xmlStream.closeNode(); // </c>
    }

    /** @returns `true` for `<c>` and its descendants (`<f>`, `<v>`, `<is>`, `<t>`, `<r>`) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "c": {
                // `type` is set by the end tag (Null until then). NB: through this.model, not a local alias of the new
                // object: measured 20% faster to parse and reconcile (V8 builds faster objects), see the micro-benchmark
                // of the sheet xforms
                this.model = {
                    address: node.attributes.r,
                    type: ValueType.Null,
                };
                this.t = node.attributes.t;
                this.hasValueNode = false;
                if (node.attributes.s) {
                    this.model.styleId = parseInt(node.attributes.s, 10);
                }
                return true;
            }

            case "f": {
                // <f> is read inside <c>, which created the model
                const { model } = this;
                this.currentNode = "f";
                if (model) {
                    model.si = node.attributes.si;
                    model.shareType = node.attributes.t;
                    model.ref = node.attributes.ref;
                }
                return true;
            }

            case "v":
                this.currentNode = "v";
                this.hasValueNode = true;
                return true;

            case "t":
                this.currentNode = "t";
                return true;

            case "r":
                this.parser = this.richTextXForm;
                this.parser.parseOpen(node);
                return true;

            default:
                return false;
        }
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
            return;
        }
        // the text elements are read inside <c>, which created the model
        const { model } = this;
        if (!model) {
            return;
        }
        switch (this.currentNode) {
            case "f":
                model.formula = model.formula ? model.formula + text : text;
                break;
            case "v":
            case "t":
                // a <t> after the runs of an <is> (not CT_Rst) is ignored
                if (!isRichText(model.value)) {
                    model.value = model.value ? rawText(model.value) + text : text;
                }
                break;
            case undefined:
            default:
                break;
        }
    }

    /** @returns `false` at the end of `<c>` (the type and value of the cell are then set), else `true` */
    override parseClose(name: string): boolean {
        switch (name) {
            case "c": {
                // the end tag of the element that created the model
                const { model } = this;
                if (!model) {
                    return false;
                }
                const { t } = this;

                // first guess on cell type
                if (model.formula || model.shareType) {
                    model.type = ValueType.Formula;
                    if (model.value === undefined && this.hasValueNode && t === "str") {
                        // <v></v>: the formula evaluated to an empty string
                        model.value = "";
                    }
                    const { value } = model;
                    if (value !== undefined) {
                        if (t === "str") {
                            model.result = utils.xmlDecode(rawText(value));
                        } else if (t === "b") {
                            model.result = parseInt(rawText(value), 10) !== 0;
                        } else if (t === "e") {
                            // the error code of the file (ST_Xstring)
                            model.result = { error: rawText(value) };
                        } else if (t === "d") {
                            // the text, when it is not a date
                            model.result = parseIsoDate(value) || rawText(value);
                        } else {
                            model.result = parseFloat(rawText(value));
                        }
                        model.value = undefined;
                    }
                } else if (model.value !== undefined) {
                    const { value } = model;
                    switch (t) {
                        case "s":
                            model.type = ValueType.String;
                            model.value = parseInt(rawText(value), 10);
                            break;
                        case "str":
                            model.type = ValueType.String;
                            model.value = utils.xmlDecode(rawText(value));
                            break;
                        case "inlineStr":
                            model.type = ValueType.String;
                            break;
                        case "b":
                            model.type = ValueType.Boolean;
                            model.value = parseInt(rawText(value), 10) !== 0;
                            break;
                        case "e":
                            model.type = ValueType.Error;
                            // the error code of the file (ST_Xstring)
                            model.value = { error: rawText(value) };
                            break;
                        case "d": {
                            const date = parseIsoDate(value);
                            if (date) {
                                model.type = ValueType.Date;
                                model.value = date;
                            } else {
                                model.type = ValueType.String;
                            }
                            break;
                        }
                        case undefined:
                        default:
                            model.type = ValueType.Number;
                            model.value = parseFloat(rawText(value));
                            break;
                    }
                } else if (model.styleId) {
                    model.type = ValueType.Null;
                } else {
                    model.type = ValueType.Merge;
                }
                return false;
            }

            case "f":
            case "v":
            case "is":
                this.currentNode = undefined;
                return true;

            case "t":
                if (this.parser) {
                    this.parser.parseClose(name);
                    return true;
                }
                this.currentNode = undefined;
                return true;

            case "r": {
                // <r> is read inside <c>, which created the model; the run xform has its model once its element is closed
                const { model } = this;
                const run = this.richTextXForm.model;
                this.parser = undefined;
                this.currentNode = undefined;
                if (!model || !run) {
                    return true;
                }
                let { value } = model;
                if (typeof value === "string") {
                    // <is> is <t>?<r>*: a leading <t> is an unformatted first run
                    value = { richText: [{ text: value }] };
                } else if (!isRichText(value)) {
                    value = { richText: [] };
                }
                model.value = value;
                value.richText.push(run);
                return true;
            }

            default:
                if (this.parser) {
                    this.parser.parseClose(name);
                    return true;
                }
                return false;
        }
    }

    /**
     * Resolves a parsed cell (mutates it into a CellModel): its style, a shared string, a date (number with a date
     * format), a shared formula (masters are recorded in `options.formulae` by `si`, clones get their master's address),
     * its hyperlink and its note.
     *
     * @throws Error for a shared string index that is not in the shared strings table
     */
    override reconcile(model: ParsedCellModel, options: CellReconcileOptions): void {
        const style = model.styleId && options.styles && options.styles.getStyleModel(model.styleId);
        if (style) {
            model.style = style;
        }
        if (model.styleId !== undefined) {
            model.styleId = undefined;
        }

        switch (model.type) {
            case ValueType.String:
                if (typeof model.value === "number") {
                    if (!options.sharedStrings) {
                        // a t="s" cell of a workbook without shared strings part: its string is missing, the cell is
                        // empty (as the document always showed it)
                        model.type = ValueType.Null;
                        delete model.value;
                        break;
                    }
                    const index = model.value;
                    const value = options.sharedStrings.getString(index);
                    if (value === undefined) {
                        throw new Error(`Cell ${String(model.address)}: shared string ${index} is not in the shared strings table`);
                    }
                    model.value = value;
                }
                if (isRichText(model.value)) {
                    model.type = ValueType.RichText;
                }
                break;

            case ValueType.Number:
                if (style && utils.isDateFmt(style.numFmt) && typeof model.value === "number") {
                    model.type = ValueType.Date;
                    model.value = utils.excelToDate(model.value, options.date1904);
                }
                break;

            case ValueType.Formula:
                // only a numeric result is a serial date: keep strings, booleans, errors and t="d" dates
                if (typeof model.result === "number" && style && utils.isDateFmt(style.numFmt)) {
                    model.result = utils.excelToDate(model.result, options.date1904);
                }
                if (model.shareType === "shared") {
                    // the si of the file (a missing one is the key "undefined", as for an object key)
                    const si = String(model.si);
                    if (model.ref) {
                        // master
                        options.formulae[si] = model.address;
                    } else {
                        // slave
                        model.sharedFormula = options.formulae[si];
                        delete model.shareType;
                    }
                    delete model.si;
                }
                break;

            case ValueType.Null:
            case ValueType.Merge:
            case ValueType.Date:
            case ValueType.Hyperlink:
            case ValueType.SharedString:
            case ValueType.RichText:
            case ValueType.Boolean:
            case ValueType.Error:
            default:
                break;
        }

        // look for hyperlink (own keys only: the cell reference comes from the file). The worksheet gives every cell an
        // address before reconcile (a missing one would be the key "undefined", as for an object key)
        const address = String(model.address);
        const { hyperlinkMap } = options;
        const hyperlink = Object.hasOwn(hyperlinkMap, address) ? hyperlinkMap[address] : undefined;
        if (hyperlink) {
            if (model.type === ValueType.Formula) {
                model.text = model.result;
                model.result = undefined;
            } else {
                model.text = model.value;
                model.value = undefined;
            }
            model.type = ValueType.Hyperlink;
            model.hyperlink = hyperlink.target;
            if (hyperlink.tooltip) {
                model.tooltip = hyperlink.tooltip;
            }
        }

        const comment = options.commentsMap && options.commentsMap[address];
        if (comment) {
            model.comment = comment;
        }
    }
}

export default CellXform;
