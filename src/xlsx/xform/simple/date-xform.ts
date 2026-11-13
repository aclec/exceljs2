import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { XmlAttributes } from "../../../utils/xml-stream";
import BaseXform, { type ValueXformOptions } from "../base-xform";

/** Formats a date for the XML. */
export type DateFormatter = (date: Date) => string;

/** Parses the XML value of a date (`undefined`: the attribute is absent). `undefined` result: no date. */
export type DateParser = (text: string | undefined) => Date | undefined;

/** The options of a {@link DateXform}. */
export interface DateXformOptions extends ValueXformOptions {
    /** Default: ISO 8601 (`toISOString()`), `""` for an invalid date. */
    readonly format?: DateFormatter;
    /** Default: `new Date(text)` (an Invalid Date when the text is not a date). */
    readonly parse?: DateParser;
}

// the default format: an invalid date (or a value that is not a Date) is written as ""
function formatDate(date: Date): string {
    try {
        if (Number.isNaN(date.getTime())) return "";
        return date.toISOString();
    } catch {
        return "";
    }
}

// the default parse: new Date(undefined) is an Invalid Date, as is new Date(NaN)
function parseDate(text: string | undefined): Date {
    return new Date(text ?? NaN);
}

/** An element holding a date, in an attribute or as its text. Rendered when the model is truthy. */
class DateXform extends BaseXform<Date> {
    /** The element. */
    readonly tag: string;
    /** The attribute holding the value; `undefined`: the text of the element. */
    readonly attr: string | undefined;
    /** Fixed attributes written on the element. */
    readonly attrs: XmlAttributes | undefined;
    private readonly _format: DateFormatter;
    private readonly _parse: DateParser;
    /** The chunks of text of the element being parsed (text mode). */
    protected text: string[] = [];

    constructor(options: DateXformOptions) {
        super();

        this.tag = options.tag;
        this.attr = options.attr;
        this.attrs = options.attrs;
        this._format = options.format || formatDate;
        this._parse = options.parse || parseDate;
    }

    override render(xmlStream: XmlStream, model: Date | undefined): void {
        if (model) {
            xmlStream.openNode(this.tag);
            if (this.attrs) {
                xmlStream.addAttributes(this.attrs);
            }
            if (this.attr) {
                xmlStream.addAttribute(this.attr, this._format(model));
            } else {
                xmlStream.writeText(this._format(model));
            }
            xmlStream.closeNode();
        }
    }

    /** Starts reading its element (attribute mode: the model is the parsed attribute). */
    override parseOpen(node: SaxTag): void {
        if (node.name === this.tag) {
            if (this.attr) {
                this.model = this._parse(node.attributes[this.attr]);
            } else {
                this.text = [];
            }
        }
    }

    override parseText(text: string): void {
        if (!this.attr) {
            this.text.push(text);
        }
    }

    override parseClose(): boolean {
        if (!this.attr) {
            this.model = this._parse(this.text.join(""));
        }
        return false;
    }
}

export default DateXform;
