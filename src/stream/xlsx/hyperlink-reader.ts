/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import { EventEmitter } from "events";
import parseSax from "../../utils/parse-sax";

import Enums from "../../doc/enums";
import RelType from "../../xlsx/rel-type";

class HyperlinkReader extends EventEmitter {
    constructor({ workbook, id, iterator, options }) {
        super();

        this.workbook = workbook;
        this.id = id;
        this.iterator = iterator;
        this.options = options;
    }

    get count() {
        return (this.hyperlinks && this.hyperlinks.length) || 0;
    }

    each(fn) {
        return this.hyperlinks.forEach(fn);
    }

    // Reads the relationships once: WorkbookReader.read() awaits it after emitting the reader, whose listener may have
    // started it already (the entry must be read to the end before the zip moves on to the next one)
    read() {
        if (!this._reading) {
            this._reading = this._read();
        }
        return this._reading;
    }

    async _read() {
        const { iterator, options } = this;
        let emitHyperlinks = false;
        let hyperlinks = null;
        switch (options.hyperlinks) {
            case "emit":
                emitHyperlinks = true;
                break;
            case "cache":
                // keyed by relationship ids from the file
                this.hyperlinks = hyperlinks = Object.create(null);
                break;
            default:
                break;
        }

        if (!emitHyperlinks && !hyperlinks) {
            this.emit("finished");
            return;
        }

        try {
            for await (const events of parseSax(iterator)) {
                for (const { eventType, value } of events) {
                    if (eventType === "opentag") {
                        const node = value;
                        if (node.name === "Relationship") {
                            const rId = node.attributes.Id;
                            switch (node.attributes.Type) {
                                case RelType.Hyperlink:
                                    {
                                        const relationship = {
                                            type: Enums.RelationshipType.Styles,
                                            rId,
                                            target: node.attributes.Target,
                                            targetMode: node.attributes.TargetMode,
                                        };
                                        if (emitHyperlinks) {
                                            this.emit("hyperlink", relationship);
                                        } else {
                                            hyperlinks[relationship.rId] = relationship;
                                        }
                                    }
                                    break;

                                default:
                                    break;
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
