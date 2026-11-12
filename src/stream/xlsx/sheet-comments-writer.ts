import XmlStream from "../../utils/xml-stream";
import RelType from "../../xlsx/rel-type";
import colCache from "../../utils/col-cache";
import CommentXformJs from "../../xlsx/xform/comment/comment-xform";
import VmlShapeXform from "../../xlsx/xform/comment/vml-shape-xform";
import VmlNotesXformJs from "../../xlsx/xform/comment/vml-notes-xform";
import type { ZipEntry } from "../../utils/zip-stream";
import type { CommentRefModel, CommentRefsEntry, NoteModel } from "../../types";
import type SheetRelsWriter from "./sheet-rels-writer";
import type { SheetPartsWorkbook } from "./sheet-rels-writer";

// TS-WAVE2: the comment xforms are typed in parallel; their APIs as used here, asserted once (guide rule 15)
const CommentXform = CommentXformJs as unknown as new () => { render(xmlStream: XmlStream, model: CommentRefModel): void };
const VmlNotesXform = VmlNotesXformJs as unknown as {
    /** the number of blocks of VML shape ids that `count` notes take */
    blockCount(this: void, count: number): number;
    /** the `idmap` data of a worksheet whose `count` notes use the blocks from `firstBlock` */
    idmap(this: void, firstBlock: number, count: number): string;
    /** the index of the VML shape of the note at `index` */
    shapeIndex(this: void, firstBlock: number, index: number): number;
};

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
 * Writes the notes of a streamed worksheet: the comments part (`xl/commentsN.xml`), as the rows are written, and its
 * VML drawing (`xl/drawings/vmlDrawingN.vml`), when the worksheet is committed. Nothing is written for a worksheet
 * without notes.
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

    // VML shape ids must be unique across the workbook, and the blocks of ids a worksheet claims (VmlNotesXform) are
    // listed before its shapes: the notes are kept until the worksheet is committed, when their number is known
    private _vmlComments: CommentRefModel[];
    private _worksheet: CommentsWorksheet;
    private _workbook: CommentsWorkbook;
    private _sheetRelsWriter: SheetRelsWriter;
    private _commentsStream?: ZipEntry;
    private _vmlStream?: ZipEntry;

    constructor(worksheet: CommentsWorksheet, sheetRelsWriter: SheetRelsWriter, options: SheetCommentsWriterOptions) {
        // in a workbook, each sheet will have a number
        this.id = options.id;
        this.count = 0;
        this._vmlComments = [];
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

    private _writeOpen(): void {
        this.commentsStream.write(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
                "<authors><author>Author</author></authors>" +
                "<commentList>"
        );
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
        this._vmlComments.forEach((comment, index) => {
            const vmlXmlStream = new XmlStream();
            new VmlShapeXform().render(vmlXmlStream, comment, VmlNotesXform.shapeIndex(firstBlock, index));
            this.vmlStream.write(vmlXmlStream.xml);
        });
        this._vmlComments = [];
        this.vmlStream.write("</xml>");
    }

    private _writeComment(comment: CommentRefModel): void {
        const commentXform = new CommentXform();
        const commentsXmlStream = new XmlStream();
        commentXform.render(commentsXmlStream, comment);
        this.commentsStream.write(commentsXmlStream.xml);
        this._vmlComments.push(comment);
    }

    private _writeClose(): void {
        this.commentsStream.write("</commentList></comments>");
        this._writeVml();
    }

    /**
     * Writes the notes of a row (cell-xform prepare collects them with their `ref`). Mutates them: adds `refAddress`.
     *
     * @throws Error for a `ref` that is not a cell address
     */
    addComments(comments: readonly (NoteModel & { ref: string })[] | undefined): void {
        if (comments && comments.length) {
            if (!this.startedData) {
                this._worksheet.comments = [];
                this._writeOpen();
                this._addRelationships();
                this._addCommentRefs();
                this.startedData = true;
            }

            const refComments = comments.map((item): CommentRefModel => {
                const comment = item as CommentRefModel;
                comment.refAddress = colCache.decodeAddress(item.ref);
                return comment;
            });

            for (const comment of refComments) {
                this._writeComment(comment);
                this.count += 1;
            }
        }
    }

    /** Ends the comments part and writes the VML drawing, if the worksheet has notes. */
    commit(): void {
        if (this.count) {
            this._writeClose();
            this.commentsStream.end();
            this.vmlStream.end();
        }
    }
}

export default SheetCommentsWriter;
