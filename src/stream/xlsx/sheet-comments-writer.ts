import XmlStream from "../../utils/xml-stream";
import RelType from "../../xlsx/rel-type";
import colCache from "../../utils/col-cache";
import CommentsXform from "../../xlsx/xform/comment/comments-xform";
import VmlShapeXform from "../../xlsx/xform/comment/vml-shape-xform";
import VmlNotesXform from "../../xlsx/xform/comment/vml-notes-xform";
import type { ZipEntry } from "../../utils/zip-stream";
import type { CommentRefModel, CommentRefsEntry, NoteModel } from "../../types";
import type SheetRelsWriter from "./sheet-rels-writer";
import type { SheetPartsWorkbook } from "./sheet-rels-writer";

/** What the comments writer needs from the workbook writer. */
export interface CommentsWorkbook extends SheetPartsWorkbook {
    /** the comments parts of the workbook (for the content types) */
    commentRefs: CommentRefsEntry[];
    /** the next block of VML shape ids that the notes of a worksheet can claim (VmlNotesXform) */
    _nextShapeBlock: number;
}

/** What the comments writer needs from the worksheet writer. */
export interface CommentsWorksheet {
    /** set to `[]` when the first note is written (compatibility) */
    comments?: CommentRefModel[];
}

/** The options of the {@link SheetCommentsWriter} constructor (those of the WorksheetWriter). */
export interface SheetCommentsWriterOptions {
    /** the number of the worksheet in the workbook */
    id: number;
    workbook: CommentsWorkbook;
}

/**
 * Writes the notes of a streamed worksheet: the comments part (`xl/commentsN.xml`) and its VML drawing
 * (`xl/drawings/vmlDrawingN.vml`), when the worksheet is committed. Nothing is written for a worksheet without notes.
 */
class SheetCommentsWriter {
    /** The number of the worksheet in the workbook. */
    readonly id: number;
    /** The number of notes written. */
    count: number;
    /** The relationship id of the VML drawing, once the first note is written. */
    vmlRelId?: string;
    /** True once the first note is written. */
    startedData?: boolean;

    // The notes are kept until the worksheet is committed: the authors of the comments part are listed before the
    // notes, and so are the blocks of VML shape ids a worksheet claims (VmlNotesXform, unique across the workbook)
    private _comments: CommentRefModel[];
    private _worksheet: CommentsWorksheet;
    private _workbook: CommentsWorkbook;
    private _sheetRelsWriter: SheetRelsWriter;
    private _commentsStream?: ZipEntry;
    private _vmlStream?: ZipEntry;

    constructor(worksheet: CommentsWorksheet, sheetRelsWriter: SheetRelsWriter, options: SheetCommentsWriterOptions) {
        // in a workbook, each sheet will have a number
        this.id = options.id;
        this.count = 0;
        this._comments = [];
        this._worksheet = worksheet;
        this._workbook = options.workbook;
        this._sheetRelsWriter = sheetRelsWriter;
    }

    /** The zip entry of the comments part, opened on first use. */
    get commentsStream(): ZipEntry {
        if (!this._commentsStream) {
            this._commentsStream = this._workbook._openStream(`/xl/comments${this.id}.xml`);
        }
        return this._commentsStream;
    }

    /** The zip entry of the VML drawing, opened on first use. */
    get vmlStream(): ZipEntry {
        if (!this._vmlStream) {
            this._vmlStream = this._workbook._openStream(`xl/drawings/vmlDrawing${this.id}.vml`);
        }
        return this._vmlStream;
    }

    private _addRelationships(): void {
        const commentRel = {
            Type: RelType.Comments,
            Target: `../comments${this.id}.xml`,
        };
        this._sheetRelsWriter.addRelationship(commentRel);

        const vmlDrawingRel = {
            Type: RelType.VmlDrawing,
            Target: `../drawings/vmlDrawing${this.id}.vml`,
        };
        this.vmlRelId = this._sheetRelsWriter.addRelationship(vmlDrawingRel);
    }

    private _addCommentRefs(): void {
        this._workbook.commentRefs.push({
            commentName: `comments${this.id}`,
            vmlDrawing: `vmlDrawing${this.id}`,
        });
    }

    private _writeVml(): void {
        const firstBlock = this._workbook._nextShapeBlock;
        this._workbook._nextShapeBlock += VmlNotesXform.blockCount(this.count);
        this.vmlStream.write(
            '<?xml version="1.0" encoding="UTF-8"?>' +
                '<xml xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
                '<o:shapelayout v:ext="edit">' +
                `<o:idmap v:ext="edit" data="${VmlNotesXform.idmap(firstBlock, this.count)}" />` +
                "</o:shapelayout>" +
                '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
                '<v:stroke joinstyle="miter" />' +
                '<v:path gradientshapeok="t" o:connecttype="rect" />' +
                "</v:shapetype>"
        );
        this._comments.forEach((comment, index) => {
            const vmlXmlStream = new XmlStream();
            new VmlShapeXform().render(vmlXmlStream, comment, VmlNotesXform.shapeIndex(firstBlock, index));
            this.vmlStream.write(vmlXmlStream.xml);
        });
        this.vmlStream.write("</xml>");
    }

    private _writeComments(): void {
        const commentsXmlStream = new XmlStream();
        new CommentsXform().render(commentsXmlStream, { comments: this._comments });
        this.commentsStream.write(commentsXmlStream.xml);
    }

    /**
     * Adds the notes of a row (cell-xform prepare collects them with their `ref`). Mutates them: adds `refAddress`.
     *
     * @throws Error for a `ref` that is not a cell address, or a note whose author is not a string
     */
    addComments(comments: readonly (NoteModel & { ref: string })[] | undefined): void {
        if (comments && comments.length) {
            if (!this.startedData) {
                this._worksheet.comments = [];
                this._addRelationships();
                this._addCommentRefs();
                this.startedData = true;
            }

            // the notes themselves, with their refAddress (as worksheet-xform prepare does)
            const refComments = comments.map((item): CommentRefModel =>
                Object.assign(item, { refAddress: colCache.decodeAddress(item.ref) })
            );

            for (const comment of refComments) {
                // checked now, not when the worksheet is committed
                CommentsXform.authorOf(comment);
                this._comments.push(comment);
                this.count += 1;
            }
        }
    }

    /** Writes the comments part and the VML drawing, if the worksheet has notes. */
    commit(): void {
        if (this.count) {
            this._writeComments();
            this._writeVml();
            this._comments = [];
            this.commentsStream.end();
            this.vmlStream.end();
        }
    }
}

export default SheetCommentsWriter;
