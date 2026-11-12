// Type-level checks (no runtime content): the classes of src/stream/xlsx against the public API of index.d.ts
// (`stream.xlsx`). See conformance-helpers.ts for how a check compares a class with its declaration.
//
// The lists of exceptions are the remaining known mismatches. Every entry must say why it stays.

import type * as Public from "../../index";
import type WorkbookWriter from "../stream/xlsx/workbook-writer";
import type WorksheetWriter from "../stream/xlsx/worksheet-writer";
import type WorkbookReader from "../stream/xlsx/workbook-reader";
import type { WorkbookReaderEvents } from "../stream/xlsx/workbook-reader";
import type WorksheetReader from "../stream/xlsx/worksheet-reader";
import type { WorksheetReaderEvents } from "../stream/xlsx/worksheet-reader";
import type HyperlinkReader from "../stream/xlsx/hyperlink-reader";
import type { HyperlinkReaderEvents } from "../stream/xlsx/hyperlink-reader";
import type Row from "../doc/row";
import type Cell from "../doc/cell";
import type Column from "../doc/column";
import type DefinedNames from "../doc/defined-names";
import type { Assert, Check } from "./conformance-helpers";

type PublicXlsx = typeof Public.stream.xlsx;

/**
 * The declared types compared as their runtime class wherever they appear: each one is checked on its own (Row, Cell,
 * Column and DefinedNames in doc-conformance.ts, the stream classes below), so that a check reports only its own members
 * (and the `this` returned by the EventEmitter methods is the class itself).
 */
type Substitutions =
    | [Public.Row, Row]
    | [Public.Cell, Cell]
    | [Public.Column, Column]
    | [Public.DefinedNames, DefinedNames]
    | [Public.stream.xlsx.WorkbookWriter, WorkbookWriter]
    | [Public.stream.xlsx.WorksheetWriter, WorksheetWriter]
    | [Public.stream.xlsx.WorkbookReader, WorkbookReader]
    | [Public.stream.xlsx.WorksheetReader, WorksheetReader]
    | [Public.stream.xlsx.HyperlinkReader, HyperlinkReader];

/** `true` when the declared constructor arguments `P` are accepted by the runtime constructor (arguments `I`). */
type AcceptsArguments<I extends unknown[], P extends unknown[]> = [P] extends [I] ? true : false;

/** The runtime WorkbookWriter conforms to index.d.ts `stream.xlsx.WorkbookWriter`. */
export type WorkbookWriterConforms = Assert<Check<WorkbookWriter, Public.stream.xlsx.WorkbookWriter, never, Substitutions>>;
/** The runtime WorkbookWriter constructor accepts the declared options. */
export type WorkbookWriterConstructorConforms = Assert<
    AcceptsArguments<ConstructorParameters<typeof WorkbookWriter>, ConstructorParameters<PublicXlsx["WorkbookWriter"]>>
>;
/** The runtime WorksheetWriter conforms to index.d.ts `stream.xlsx.WorksheetWriter`, but for the members listed. */
export type WorksheetWriterConforms = Assert<
    Check<
        WorksheetWriter,
        Public.stream.xlsx.WorksheetWriter,
        // `undefined` only for an invalid argument (a number below 1): declared `Column` (rule 2)
        "getColumn",
        Substitutions
    >
>;
/** The runtime WorkbookReader conforms to index.d.ts `stream.xlsx.WorkbookReader`. */
export type WorkbookReaderConforms = Assert<Check<WorkbookReader, Public.stream.xlsx.WorkbookReader, never, Substitutions>>;
/** The runtime WorkbookReader constructor accepts the declared arguments. */
export type WorkbookReaderConstructorConforms = Assert<
    AcceptsArguments<ConstructorParameters<typeof WorkbookReader>, ConstructorParameters<PublicXlsx["WorkbookReader"]>>
>;
/** The static `WorkbookReader.Options` conforms to its declaration. */
export type WorkbookReaderOptionsConforms = Assert<Check<typeof WorkbookReader.Options, PublicXlsx["WorkbookReader"]["Options"]>>;
/** The runtime WorksheetReader conforms to index.d.ts `stream.xlsx.WorksheetReader`, but for the members listed. */
export type WorksheetReaderConforms = Assert<
    Check<
        WorksheetReader,
        Public.stream.xlsx.WorksheetReader,
        // `undefined` only for an invalid argument (a number that is not a positive integer): declared `Column` (rule 2)
        "getColumn",
        Substitutions
    >
>;
/** The runtime HyperlinkReader conforms to index.d.ts `stream.xlsx.HyperlinkReader`. */
export type HyperlinkReaderConforms = Assert<Check<HyperlinkReader, Public.stream.xlsx.HyperlinkReader, never, Substitutions>>;

/** The events of the runtime WorkbookReader conform to index.d.ts `stream.xlsx.WorkbookReaderEvents`. */
export type WorkbookReaderEventsConform = Assert<
    Check<WorkbookReaderEvents, Public.stream.xlsx.WorkbookReaderEvents, never, Substitutions>
>;
/** The events of the runtime WorksheetReader conform to index.d.ts `stream.xlsx.WorksheetReaderEvents`. */
export type WorksheetReaderEventsConform = Assert<
    Check<WorksheetReaderEvents, Public.stream.xlsx.WorksheetReaderEvents, never, Substitutions>
>;
/** The events of the runtime HyperlinkReader conform to index.d.ts `stream.xlsx.HyperlinkReaderEvents`. */
export type HyperlinkReaderEventsConform = Assert<
    Check<HyperlinkReaderEvents, Public.stream.xlsx.HyperlinkReaderEvents, never, Substitutions>
>;
