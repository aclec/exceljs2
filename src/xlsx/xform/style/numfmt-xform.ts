import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { NumFmtModel } from "../../../types";
import defaultNumFormats, { type DefaultNumFormat } from "../../defaultnumformats";

import BaseXform from "../base-xform";

/** The built-in formats by id, as a plain record (any number may be looked up). */
const builtInFormats: Readonly<Record<number, DefaultNumFormat | undefined>> = defaultNumFormats;

// the id of each built-in format, by format code (a Map: a format such as "constructor" is not found)
function hashDefaultFormats(): ReadonlyMap<string, number> {
    const hash = new Map<string, number>();
    for (const [id, dnf] of Object.entries(defaultNumFormats)) {
        if ("f" in dnf) {
            hash.set(dnf.f, parseInt(id, 10));
        }
        // at some point, add the other cultures here...
    }
    return hash;
}
const defaultFmtHash = hashDefaultFormats();

/**
 * A custom number format, `<numFmt numFmtId formatCode>` of the styles part. A format code longer than
 * {@link NumFmtXform.MAX_FORMAT_CODE_LENGTH} characters (or a missing one) parses as `formatCode: undefined`.
 */
class NumFmtXform extends BaseXform<NumFmtModel> {
    /** Excel's limit on the length of a format code. */
    static readonly MAX_FORMAT_CODE_LENGTH = 255;

    /** Unused: kept for the constructor signature of the original code. */
    readonly id: number | undefined;
    /** Unused: kept for the constructor signature of the original code. */
    readonly formatCode: string | undefined;

    constructor(id?: number, formatCode?: string) {
        super();

        this.id = id;
        this.formatCode = formatCode;
    }

    get tag(): string {
        return "numFmt";
    }

    override render(xmlStream: XmlStream, model: NumFmtModel): void {
        xmlStream.leafNode("numFmt", { numFmtId: model.id, formatCode: model.formatCode });
    }

    /** @returns `true` for `<numFmt>` */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case "numFmt": {
                // kept as written: a backslash escapes a literal character ("\-", "\ ", "\m")
                const { formatCode } = node.attributes;
                this.model = {
                    id: BaseXform.toIntValue(node.attributes.numFmtId, NaN),
                    // Excel limits format codes to 255 characters: a longer one is not a format Excel wrote, and
                    // is ignored (its cells read as General) instead of being scanned for every cell using it
                    formatCode:
                        formatCode !== undefined && formatCode.length <= NumFmtXform.MAX_FORMAT_CODE_LENGTH ? formatCode : undefined,
                };
                return true;
            }
            default:
                return false;
        }
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }

    /** @returns the id of the built-in format with this code, or `undefined` when it is not a built-in one */
    static getDefaultFmtId(this: void, formatCode: string): number | undefined {
        return defaultFmtHash.get(formatCode);
    }

    /** @returns the code of a built-in format (default locale), or `undefined` when there is none for this id */
    static getDefaultFmtCode(this: void, numFmtId: number): string | undefined {
        return builtInFormats[numFmtId]?.f;
    }
}

export default NumFmtXform;
