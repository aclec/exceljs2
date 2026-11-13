import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { DefinedNameModel } from "../../../types";
import BaseXform from "../base-xform";
import colCache from "../../../utils/col-cache";

/**
 * A `<definedName name localSheetId>` of `xl/workbook.xml`: its ranges, comma-separated. Parsing keeps the ranges that
 * are references (`colCache.decodeEx` accepts them); a comma inside a quoted sheet name does not split a range. A defined
 * name without `name` (CT_DefinedName requires it) parses as `null` (the list of the defined names skips it).
 */
class DefinedNamesXform extends BaseXform<DefinedNameModel> {
    /** While parsing: the `name` attribute. */
    private _parsedName?: string;
    /** While parsing: the `localSheetId` attribute. */
    private _parsedLocalSheetId?: string;
    /** While parsing: the chunks of text of the element. */
    private _parsedText: string[] = [];

    override render(xmlStream: XmlStream, model: DefinedNameModel): void {
        // <definedNames>
        //   <definedName name="name">name.ranges.join(',')</definedName>
        //   <definedName name="_xlnm.Print_Area" localSheetId="0">name.ranges.join(',')</definedName>
        // </definedNames>
        xmlStream.openNode("definedName", {
            name: model.name,
            localSheetId: model.localSheetId,
        });
        xmlStream.writeText(model.ranges.join(","));
        xmlStream.closeNode();
    }

    /** @returns `true` for `<definedName>` */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case "definedName":
                this._parsedName = node.attributes.name;
                this._parsedLocalSheetId = node.attributes.localSheetId;
                this._parsedText = [];
                return true;
            default:
                return false;
        }
    }

    override parseText(text: string): void {
        this._parsedText.push(text);
    }

    override parseClose(): boolean {
        const name = this._parsedName;
        if (name === undefined) {
            this.model = null;
            return false;
        }
        const model: DefinedNameModel = {
            name,
            ranges: extractRanges(this._parsedText.join("")),
        };
        if (this._parsedLocalSheetId !== undefined) {
            model.localSheetId = parseInt(this._parsedLocalSheetId, 10);
        }
        this.model = model;
        return false;
    }
}

function isValidRange(range: string): boolean {
    try {
        colCache.decodeEx(range);
        return true;
    } catch {
        return false;
    }
}

function extractRanges(parsedText: string): string[] {
    const ranges: string[] = [];
    let quotesOpened = false;
    let last = "";
    for (const item of parsedText.split(",")) {
        if (!item) {
            continue;
        }
        const quotes = (item.match(/'/g) || []).length;

        if (!quotes) {
            if (quotesOpened) {
                last += `${item},`;
            } else if (isValidRange(item)) {
                ranges.push(item);
            }
            continue;
        }
        const quotesEven = quotes % 2 === 0;

        if (!quotesOpened && quotesEven && isValidRange(item)) {
            ranges.push(item);
        } else if (quotesOpened && !quotesEven) {
            quotesOpened = false;
            if (isValidRange(last + item)) {
                ranges.push(last + item);
            }
            last = "";
        } else {
            quotesOpened = true;
            last += `${item},`;
        }
    }
    return ranges;
}

export default DefinedNamesXform;
