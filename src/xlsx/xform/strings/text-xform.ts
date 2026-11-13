import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { XmlValue } from "../../../utils/xml-stream";
import BaseXform from "../base-xform";

//   <t xml:space="preserve"> is </t>

// ECMA-376 ST_Xstring: "_xHHHH_" stands for the character U+HHHH (decoded by model). A literal one is written with its "_"
// escaped as "_x005F_", so that it reads back as written: every "_" that a reader would take as the start of an escape.
const ESCAPE_START = /_(?=x[0-9A-Fa-f]{4}_)/g;
const ESCAPE = /_x([0-9A-F]{4})_/g;
// text that needs xml:space="preserve": leading or trailing white space, or a line feed
const NEEDS_PRESERVE = /^\s|\n|\s$/;

function decodeEscape(match: string, hex: string): string {
    return String.fromCharCode(parseInt(hex, 16));
}

/**
 * Decodes the ST_Xstring escapes (`_xHHHH_`) of a text read from a file.
 *
 * @param text - the text of the element
 * @returns the text, with the escapes replaced by their characters
 */
export function decodeXstring(text: string): string {
    return text.replace(ESCAPE, decodeEscape);
}

/**
 * Escapes the `_` of a literal `_xHHHH_` (as `_x005F_`), so that an ST_Xstring reads back as written.
 *
 * @param text - the text to write
 * @returns the text to write in the element
 */
export function encodeXstring(text: string): string {
    return text.replace(ESCAPE_START, "_x005F_");
}

/**
 * A `<t>` text element (of a shared string, a rich text run or a phonetic run): the model is its text, with the ST_Xstring
 * escapes (`_xHHHH_`) decoded. Rendering escapes a literal `_xHHHH_` and adds `xml:space="preserve"` when the text has
 * leading / trailing white space or a line feed.
 *
 * The model is set when the element closes ({@link parseClose}).
 */
class TextXform extends BaseXform<string> {
    /** The chunks of text of the element being parsed. */
    private _text: string[] = [];

    /** The element: `t`. */
    get tag(): string {
        return "t";
    }

    /** @param model - the text (a number or boolean is written as its string) */
    override render(xmlStream: XmlStream, model: XmlValue): void {
        xmlStream.openNode("t");
        if (NEEDS_PRESERVE.test(String(model))) {
            xmlStream.addAttribute("xml:space", "preserve");
        }
        xmlStream.writeText(typeof model === "string" ? encodeXstring(model) : model);
        xmlStream.closeNode();
    }

    /** @returns `true` for `<t>` */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case "t":
                this._text = [];
                return true;
            default:
                return false;
        }
    }

    override parseText(text: string): void {
        this._text.push(text);
    }

    /** Completes the model (the text, decoded). */
    override parseClose(): boolean {
        this.model = decodeXstring(this._text.join(""));
        return false;
    }
}

export default TextXform;
