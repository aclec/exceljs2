// Internal models of a workbook and of its package parts: the JSON exchanged between src/doc/workbook.ts, xlsx.ts (which
// orchestrates the zip parts), the book / core xforms and the stream WorkbookWriter / WorkbookReader.
//
// PARTIAL: the read model accumulated by xlsx.ts holds the part models typed in ./sheet.ts at the level xlsx.ts uses
// them; pivot tables are left `unknown`.

import type { CalculationProperties, WorkbookProperties } from "../../index";
import type { DrawingModel, NoteModel, TableModel, WorksheetModel, ParsedWorksheetModel } from "./sheet";
import type { SharedStringValue } from "./strings";

// ---------------------------------------------------------------------------------------------------------------------
// Package parts

/**
 * A relationship of an OPC part (`_rels/*.rels`, relationship-xform). Rendered as the attributes of `<Relationship>`;
 * parsed as the raw attributes of the element (other attributes are kept). `Type` is one of the URIs of
 * `src/xlsx/rel-type.ts`.
 */
export interface RelationshipModel {
    Id: string;
    Type: string;
    Target: string;
    TargetMode?: "External";
    /** Parse: any other attribute of the element, as written. */
    [attribute: string]: string | undefined;
}

/**
 * Extended properties, `docProps/app.xml` (app-xform). Render: `worksheets` (their names: TitlesOfParts, their count:
 * HeadingPairs), `company`, `manager`. Parse: `{ worksheets, company, manager }`, `worksheets` always `undefined` (the
 * titles are not parsed).
 */
export interface AppModel {
    worksheets?: readonly { readonly name: string }[];
    company?: string;
    manager?: string;
}

/**
 * Core properties, `docProps/core.xml` (core-xform): Dublin Core terms. An invalid date is neither written nor read
 * (`undefined`). `version` is rendered but not parsed.
 */
export interface CoreModel {
    creator?: string;
    title?: string;
    subject?: string;
    description?: string;
    identifier?: string;
    language?: string;
    keywords?: string;
    category?: string;
    lastModifiedBy?: string;
    lastPrinted?: Date;
    /** parsed as an integer (`NaN` if not a number) */
    revision?: number;
    version?: string;
    contentStatus?: string;
    contentType?: string;
    created?: Date;
    modified?: Date;
}

/**
 * A medium of the workbook (an image). From `workbook.addImage`: `extension` and one of `filename` / `buffer` / `base64`
 * (a data URL or plain base64). On write, WorkbookXform.prepare sets `name` (`image1`...). On read: `{ type, name,
 * extension, buffer }`, then `index` (WorkbookXform.reconcile). The stream writer's names include the extension.
 *
 * Differs from index.d.ts `Media` (required `buffer`, no `base64` / `filename` / `index`).
 */
export interface MediaModel {
    type: "image";
    name?: string;
    extension: string;
    buffer?: Uint8Array | ArrayBuffer;
    base64?: string;
    filename?: string;
    index?: number;
}

/** What `[Content_Types].xml` (content-types-xform, render only) reads from the write model. */
export interface ContentTypesModel {
    worksheets: readonly { readonly id: number }[];
    media?: readonly Pick<MediaModel, "type" | "extension">[];
    pivotTables?: readonly unknown[];
    /** an Override is written when `count` is non-zero */
    sharedStrings?: { readonly count: number };
    tables?: readonly { readonly target?: string }[];
    drawings?: readonly { readonly name?: string }[];
    commentRefs?: readonly CommentRefsEntry[];
}

/** A comments part of the write model: `comments<N>` and its VML drawing. */
export interface CommentRefsEntry {
    commentName: string;
    vmlDrawing: string;
}

// ---------------------------------------------------------------------------------------------------------------------
// Workbook part

/** A `<sheet>` of `xl/workbook.xml` (book/sheet-xform). `name` is XML-decoded; `id` is `NaN` if `sheetId` is not a number. */
export interface SheetDefModel {
    id: number;
    name: string;
    state?: string;
    rId: string;
}

/**
 * A defined name (book/defined-name-xform, `DefinedNames.model`): `ranges` are `$`-absolute references
 * (`Sheet1!$A$1:$B$2`). `localSheetId` only in the xform model (print areas / titles).
 */
export interface DefinedNameModel {
    name: string;
    ranges: string[];
    localSheetId?: number;
}

/**
 * A workbook view (book/workbook-view-xform): index.d.ts `WorkbookView`; `activeTab` / `firstSheet` are absent when not
 * set.
 */
export interface WorkbookViewModel {
    x: number;
    y: number;
    width: number;
    height: number;
    visibility: string;
    activeTab?: number;
    firstSheet?: number;
}

/**
 * `xl/workbook.xml` as WorkbookXform parses it (read by xlsx.ts and the stream WorkbookReader). `sheets` / `views` are
 * `undefined` when their element is absent; `calcProperties` is always `{}` (not parsed).
 */
export interface WorkbookPartModel {
    sheets?: SheetDefModel[];
    properties: Partial<WorkbookProperties>;
    views?: WorkbookViewModel[];
    calcProperties: Partial<CalculationProperties>;
    definedNames?: DefinedNameModel[];
}

// ---------------------------------------------------------------------------------------------------------------------
// Workbook

/**
 * A workbook as the document produces it (`Workbook.model`, src/doc/workbook.ts) and consumes it (`Workbook.model =`).
 *
 * Getter only: `sheets` (the worksheet models again, separate objects), `themes` may be undefined. Setter: reads
 * `sheets` only to order the worksheets; ignores `identifier`, `version`, `contentType`.
 * Index.d.ts `WorkbookModel` differs (`worksheets: Worksheet[]`, `revision: Date`, `themes: string[]`).
 */
export interface WorkbookModel extends CoreModel {
    properties: Partial<WorkbookProperties>;
    calcProperties: Partial<CalculationProperties>;
    worksheets: WorksheetModel[];
    sheets?: WorksheetModel[] | SheetDefModel[];
    definedNames: DefinedNameModel[];
    views: WorkbookViewModel[];
    company?: string;
    manager?: string;
    /** theme part name (`theme1`) → XML */
    themes?: Record<string, string>;
    media: MediaModel[];
    /** PARTIAL: the document's PivotTable objects. */
    pivotTables: unknown[];
}

/**
 * The workbook model while writing (xlsx.ts prepareModel): the document model plus the write state. `styles` /
 * `sharedStrings` are the managers (StylesXform or its Mock, SharedStringsXform), typed by the files that create them.
 */
export interface WorkbookWriteModel extends WorkbookModel {
    useSharedStrings: boolean;
    useStyles: boolean;
    /** every worksheet, for workbook.xml (WorkbookXform.prepare: `sheets = worksheets`) */
    sheets: WorksheetModel[];
    drawings: DrawingModel[];
    commentRefs: CommentRefsEntry[];
    /** every table of every worksheet, with its `target` and `id` */
    tables: TableModel[];
}

/**
 * The model accumulated by xlsx.ts while reading the parts of a file, before reconcile. Hashes keyed by part names come
 * from the file: they must not be plain `{}` objects (`Object.create(null)` or `Map`). reconcile deletes the read-only
 * state (`worksheetHash`, `worksheetRels`, `globalRels`, `sharedStrings`, `workbookRels`, `styles`, `mediaIndex`,
 * `drawings`, `drawingRels`, `vmlDrawings`, `loadBudget`) and leaves a {@link WorkbookModel}-like object.
 * PARTIAL: the values of the part hashes are typed at the level xlsx.ts uses them.
 */
export interface WorkbookReadModel extends CoreModel {
    worksheets: ParsedWorksheetModel[];
    /** part name (`xl/worksheets/sheet1.xml`) → worksheet */
    worksheetHash: Record<string, ParsedWorksheetModel>;
    /** indexed by sheet number (from the part name) */
    worksheetRels: RelationshipModel[][];
    /** theme name → XML */
    themes: Record<string, string>;
    media: MediaModel[];
    /** media file name and base name → index in `media` */
    mediaIndex: Record<string, number>;
    /** drawing name (`drawing1`) → drawing */
    drawings: Record<string, DrawingModel>;
    drawingRels: Record<string, RelationshipModel[]>;
    /** comments part name → its notes */
    comments: Record<string, { comments: NoteModel[] }>;
    /** table part name → table */
    tables: Record<string, TableModel>;
    /** VML part name → its shapes (PARTIAL) */
    vmlDrawings: Record<string, { comments: unknown[] }>;
    globalRels?: RelationshipModel[];
    workbookRels?: RelationshipModel[];
    sheets?: SheetDefModel[];
    definedNames?: DefinedNameModel[];
    views?: WorkbookViewModel[];
    properties?: Partial<WorkbookProperties>;
    calcProperties?: Partial<CalculationProperties>;
    company?: string;
    manager?: string;
    /** the parsed shared strings manager (SharedStringsXform), typed by xlsx.ts */
    sharedStrings?: { getString(index: number): SharedStringValue | undefined };
}
