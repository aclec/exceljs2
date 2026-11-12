import { EventEmitter } from "events";
import parseSax, { type SaxInput } from "../../utils/parse-sax";

import { RelationshipType } from "../../doc/enums";
import RelType from "../../xlsx/rel-type";
import type WorkbookReader from "./workbook-reader";
import type { WorkbookReaderOptions } from "./workbook-reader";

/**
 * A hyperlink relationship of a worksheet (`xl/worksheets/_rels/sheetN.xml.rels`), as read. The values come from the
 * file (untrusted); an absent attribute is `undefined`.
 */
export interface HyperlinkRelationship {
    /** Always `RelationshipType.Styles` (5), as in ExcelJS < 5 (kept for compatibility). */
    type: RelationshipType;
    /** The relationship id (`Id`), referenced by the `r:id` of a `<hyperlink>` of the worksheet. */
    rId: string | undefined;
    /** The link (`Target`). */
    target: string | undefined;
    /** `External` for a link outside the package. */
    targetMode: string | undefined;
}

/** The arguments of the {@link HyperlinkReader} constructor. */
export interface HyperlinkReaderOptions {
    workbook: WorkbookReader;
    /** The number of the worksheet part (`sheet3.xml.rels` → "3"). */
    id: string;
    /** The XML of the relationships part. */
    iterator: SaxInput;
    /** The options of the workbook reader: `hyperlinks` selects what {@link HyperlinkReader.read} does. */
    options: Pick<WorkbookReaderOptions, "hyperlinks">;
}

/** The events of a {@link HyperlinkReader}. */
interface HyperlinkReaderEvents {
    /** a hyperlink relationship (`hyperlinks: "emit"`) */
    hyperlink: [relationship: HyperlinkRelationship];
    /** the part was read */
    finished: [];
    /** the part could not be read */
    error: [error: unknown];
}

/**
 * Reads the hyperlink relationships of a worksheet, for the streaming {@link WorkbookReader}: emitted one by one
 * (`hyperlinks: "emit"`), or cached in {@link hyperlinks} keyed by relationship id (`hyperlinks: "cache"`).
 */
class HyperlinkReader extends EventEmitter<HyperlinkReaderEvents> {
    readonly workbook: WorkbookReader;
    /** The number of the worksheet part. */
    readonly id: string;
    readonly iterator: SaxInput;
    readonly options: Pick<WorkbookReaderOptions, "hyperlinks">;

    /** `hyperlinks: "cache"`: the relationships read, keyed by relationship id (no prototype: the ids come from the file). */
    hyperlinks?: Record<string, HyperlinkRelationship>;

    private _reading?: Promise<void>;

    constructor({ workbook, id, iterator, options }: HyperlinkReaderOptions) {
        super();

        this.workbook = workbook;
        this.id = id;
        this.iterator = iterator;
        this.options = options;
    }

    /** The number of relationships cached (0 before {@link read} and when they are not cached). */
    get count(): number {
        return this.hyperlinks ? Object.keys(this.hyperlinks).length : 0;
    }

    /** Calls `fn` for each cached relationship, in the order of the part. */
    each(fn: (relationship: HyperlinkRelationship, index: number) => void): void {
        if (this.hyperlinks) {
            Object.values(this.hyperlinks).forEach(fn);
        }
    }

    /**
     * Reads the relationships once: WorkbookReader.read() awaits it after emitting the reader, whose listener may have
     * started it already (the entry must be read to the end before the zip moves on to the next one).
     *
     * @returns resolves once the part is read; errors are emitted (`error`), not thrown
     */
    read(): Promise<void> {
        if (!this._reading) {
            this._reading = this._read();
        }
        return this._reading;
    }

    private async _read(): Promise<void> {
        const { iterator, options } = this;
        const emitHyperlinks = options.hyperlinks === "emit";
        let hyperlinks: Record<string, HyperlinkRelationship> | null = null;
        if (options.hyperlinks === "cache") {
            // keyed by relationship ids from the file
            this.hyperlinks = hyperlinks = Object.create(null) as Record<string, HyperlinkRelationship>;
        }

        if (!emitHyperlinks && !hyperlinks) {
            this.emit("finished");
            return;
        }

        try {
            for await (const events of parseSax(iterator)) {
                for (const event of events) {
                    if (event.eventType === "opentag") {
                        const node = event.value;
                        if (node.name === "Relationship" && node.attributes.Type === RelType.Hyperlink) {
                            const relationship: HyperlinkRelationship = {
                                type: RelationshipType.Styles,
                                rId: node.attributes.Id,
                                target: node.attributes.Target,
                                targetMode: node.attributes.TargetMode,
                            };
                            if (hyperlinks) {
                                // String(): the key of a relationship without Id is "undefined", as before
                                hyperlinks[String(relationship.rId)] = relationship;
                            } else {
                                this.emit("hyperlink", relationship);
                            }
                        }
                    }
                }
            }
            this.emit("finished");
        } catch (error) {
            this.emit("error", error);
        }
    }
}

export default HyperlinkReader;
