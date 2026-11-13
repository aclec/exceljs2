import { xmlEncode } from "../../utils/utils";
import RelType from "../../xlsx/rel-type";
import HyperlinkXform from "../../xlsx/xform/sheet/hyperlink-xform";
import type { ZipEntry } from "../../utils/zip-stream";
import type { SheetHyperlinkModel } from "../../types";

/** A relationship to write: `Id` is assigned by the writer. */
export interface NewRelationship {
    Type: string;
    Target: string;
    TargetMode?: string;
}

/** What the sheet writers need from the workbook writer. */
export interface SheetPartsWorkbook {
    /** a zip entry written as a stream */
    _openStream(path: string): ZipEntry;
}

/** The arguments of the {@link SheetRelsWriter} constructor (those of the WorksheetWriter). */
export interface SheetRelsWriterOptions {
    /** the number of the worksheet in the workbook */
    id: number;
    workbook: SheetPartsWorkbook;
}

/** The `hyperlinks` collector of the cell prepare options: forwards to {@link SheetRelsWriter.addHyperlink}. */
class HyperlinksProxy {
    readonly writer: SheetRelsWriter;

    constructor(sheetRelsWriter: SheetRelsWriter) {
        this.writer = sheetRelsWriter;
    }

    push(hyperlink: SheetHyperlinkModel): void {
        this.writer.addHyperlink(hyperlink);
    }
}

/**
 * Writes the relationships part of a streamed worksheet (`xl/worksheets/_rels/sheetN.xml.rels`) as the relationships
 * are added (hyperlinks, images, comments); keeps the hyperlinks for the `<hyperlinks>` of the worksheet. The part is
 * written only when it has a relationship.
 */
class SheetRelsWriter {
    /** The number of the worksheet in the workbook. */
    readonly id: number;

    /** The number of relationships written. */
    count: number;

    /** The hyperlinks of the worksheet, with the relationship id of the external ones. */
    _hyperlinks: SheetHyperlinkModel[];

    private _workbook: SheetPartsWorkbook;
    private _stream?: ZipEntry;
    private _hyperlinksProxy?: HyperlinksProxy;

    constructor(options: SheetRelsWriterOptions) {
        // in a workbook, each sheet will have a number
        this.id = options.id;

        // count of all relationships
        this.count = 0;

        // keep record of all hyperlinks
        this._hyperlinks = [];

        this._workbook = options.workbook;
    }

    /** The zip entry of the part, opened on first use. */
    get stream(): ZipEntry {
        if (!this._stream) {
            this._stream = this._workbook._openStream(`/xl/worksheets/_rels/sheet${this.id}.xml.rels`);
        }
        return this._stream;
    }

    /** The number of hyperlinks. */
    get length(): number {
        return this._hyperlinks.length;
    }

    /** Calls `fn` for each hyperlink. */
    each(fn: (hyperlink: SheetHyperlinkModel, index: number) => void): void {
        this._hyperlinks.forEach(fn);
    }

    /** The collector of hyperlinks passed to the cells' prepare (`options.hyperlinks`). */
    get hyperlinksProxy(): HyperlinksProxy {
        return this._hyperlinksProxy || (this._hyperlinksProxy = new HyperlinksProxy(this));
    }

    /** Adds a hyperlink of a cell (cell-xform prepare: `target` is set): an external one gets a relationship. */
    addHyperlink(hyperlink: SheetHyperlinkModel): void {
        const { address, target, tooltip } = hyperlink;
        // internal links are a location in the sheet, without a relationship
        const rId = HyperlinkXform.isInternalLink(target)
            ? undefined
            : this._writeRelationship({
                  // the cells always give their hyperlink (String(undefined) would be written for a missing one)
                  Target: String(target),
                  Type: RelType.Hyperlink,
                  TargetMode: "External",
              });

        // store sheet stuff for later
        this._hyperlinks.push({ rId, address, target, tooltip });
    }

    /** @returns the relationship id of an image */
    addMedia(media: NewRelationship): string {
        return this._writeRelationship(media);
    }

    /** @returns the relationship id */
    addRelationship(rel: NewRelationship): string {
        return this._writeRelationship(rel);
    }

    /** Ends the part, if it has a relationship. */
    commit(): void {
        if (this.count) {
            // write xml utro
            this._writeClose();
            // and close stream
            this.stream.end();
        }
    }

    // ================================================================================
    private _writeOpen(): void {
        this.stream.write(
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
       <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
        );
    }

    private _writeRelationship(relationship: NewRelationship): string {
        if (!this.count) {
            this._writeOpen();
        }

        const rId = `rId${++this.count}`;

        if (relationship.TargetMode) {
            this.stream.write(
                `<Relationship Id="${rId}"` +
                    ` Type="${relationship.Type}"` +
                    ` Target="${xmlEncode(relationship.Target)}"` +
                    ` TargetMode="${relationship.TargetMode}"` +
                    "/>"
            );
        } else {
            this.stream.write(`<Relationship Id="${rId}" Type="${relationship.Type}" Target="${xmlEncode(relationship.Target)}"/>`);
        }

        return rId;
    }

    private _writeClose(): void {
        this.stream.write("</Relationships>");
    }
}

export default SheetRelsWriter;
