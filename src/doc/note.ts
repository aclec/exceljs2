// The note (comment) of a cell: `cell.note`.

import type { Comment } from "../../index";
import { deepMerge } from "../utils/under-dash";
import type { NoteModel } from "../types";

/** The note of a cell: a plain string, or a {@link Comment} (rich text, margins, protection...). */
class Note {
    /** The settings every note model gets unless the note sets them. */
    static DEFAULT_CONFIGS: { readonly note: Comment } = {
        note: {
            margins: {
                insetmode: "auto",
                inset: [0.13, 0.13, 0.25, 0.25],
            },
            protection: {
                locked: "True",
                lockText: "True",
            },
            editAs: "absolute",
        },
    };

    /** The note: a string, or a comment. */
    note: string | Comment | undefined;

    constructor(note?: string | Comment) {
        this.note = note;
    }

    /**
     * The model of the note: a string note becomes a single text run; the {@link Note.DEFAULT_CONFIGS} fill in what the
     * note does not set. A new object (deep copy) on each read.
     */
    get model(): NoteModel {
        let value: { type: "note"; note: Comment | undefined };
        if (typeof this.note === "string") {
            value = {
                type: "note",
                note: {
                    texts: [
                        {
                            text: this.note,
                        },
                    ],
                },
            };
        } else {
            value = {
                type: "note",
                note: this.note,
            };
        }
        // Suitable for all cell comments
        return deepMerge({}, Note.DEFAULT_CONFIGS, value);
    }

    /**
     * Sets the note from a model: a single text run without font becomes a string note, anything else the comment
     * itself (by reference).
     * @throws Error when the comment has no `texts`
     */
    set model(value: NoteModel) {
        const { note } = value;
        const { texts } = note;
        if (!texts) {
            throw new Error("The model of a note needs the texts of its comment");
        }
        const [first] = texts;
        if (texts.length === 1 && first && Object.keys(first).length === 1) {
            this.note = first.text;
        } else {
            this.note = note;
        }
    }

    /** A note set from a model (see the model setter). */
    static fromModel(model: NoteModel): Note {
        const note = new Note();
        note.model = model;
        return note;
    }
}

export default Note;
