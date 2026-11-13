// Type-level checks (no runtime content): the classes of src/stream/xlsx against the public API of index.d.ts
// (`stream.xlsx`). See conformance-helpers.ts for how a check compares a class with its declaration.
//
// The known deviations (the rules of doc-conformance.ts, the declarations given to a check) each say why they stay.

import type * as Public from "../../index";
import type WorkbookWriter from "../stream/xlsx/workbook-writer";
import type WorksheetWriter from "../stream/xlsx/worksheet-writer";
import type { WorksheetWriterChart } from "../stream/xlsx/stream-chart";
import type WorkbookReader from "../stream/xlsx/workbook-reader";
import type { WorkbookReaderEvents } from "../stream/xlsx/workbook-reader";
import type WorksheetReader from "../stream/xlsx/worksheet-reader";
import type { WorksheetReaderEvents } from "../stream/xlsx/worksheet-reader";
import type HyperlinkReader from "../stream/xlsx/hyperlink-reader";
import type { HyperlinkReaderEvents } from "../stream/xlsx/hyperlink-reader";
import type { Assert, Check } from "./conformance-helpers";
import type { DocRules } from "./doc-conformance";

type PublicXlsx = typeof Public.stream.xlsx;

/**
 * The rules of the checks (see conformance-helpers.ts): those of doc-conformance.ts (Row, Cell, Column and DefinedNames
 * are checked there), and the stream classes compared as themselves wherever they appear, each one checked on its own
 * below (so that the `this` returned by the EventEmitter methods is the class itself).
 */
type StreamRules =
    | DocRules
    | [Public.stream.xlsx.WorkbookWriter, WorkbookWriter]
    | [Public.stream.xlsx.WorksheetWriter, WorksheetWriter]
    // a chart of a streamed worksheet is declared WorksheetChart, as a chart of a document worksheet
    | [Public.WorksheetChart, WorksheetWriterChart]
    | [Public.stream.xlsx.WorkbookReader, WorkbookReader]
    | [Public.stream.xlsx.WorksheetReader, WorksheetReader]
    | [Public.stream.xlsx.HyperlinkReader, HyperlinkReader];

/** `true` when the declared constructor arguments `P` are accepted by the runtime constructor (arguments `I`). */
type AcceptsArguments<I extends unknown[], P extends unknown[]> = [P] extends [I] ? true : false;

/** The runtime WorkbookWriter conforms to index.d.ts `stream.xlsx.WorkbookWriter`. */
export type WorkbookWriterConforms = Assert<Check<WorkbookWriter, Public.stream.xlsx.WorkbookWriter, StreamRules>>;
/** The runtime WorkbookWriter constructor accepts the declared options. */
export type WorkbookWriterConstructorConforms = Assert<
    AcceptsArguments<ConstructorParameters<typeof WorkbookWriter>, ConstructorParameters<PublicXlsx["WorkbookWriter"]>>
>;
/** The runtime WorksheetWriter conforms to index.d.ts `stream.xlsx.WorksheetWriter`, but for the known deviations given. */
export type WorksheetWriterConforms = Assert<
    Check<
        WorksheetWriter,
        Public.stream.xlsx.WorksheetWriter,
        StreamRules,
        {
            // `undefined` only for an invalid argument (a number below 1): declared `Column`, as its TSDoc says (rule 2)
            getColumn(indexOrKey: number | string): Public.Column | undefined;
        }
    >
>;
/** The runtime chart of a streamed worksheet conforms to index.d.ts `WorksheetChart`. */
export type WorksheetWriterChartConforms = Assert<Check<WorksheetWriterChart, Public.WorksheetChart, StreamRules>>;
/** The runtime WorkbookReader conforms to index.d.ts `stream.xlsx.WorkbookReader`. */
export type WorkbookReaderConforms = Assert<Check<WorkbookReader, Public.stream.xlsx.WorkbookReader, StreamRules>>;
/** The runtime WorkbookReader constructor accepts the declared arguments. */
export type WorkbookReaderConstructorConforms = Assert<
    AcceptsArguments<ConstructorParameters<typeof WorkbookReader>, ConstructorParameters<PublicXlsx["WorkbookReader"]>>
>;
/** The static `WorkbookReader.Options` conforms to its declaration. */
export type WorkbookReaderOptionsConforms = Assert<Check<typeof WorkbookReader.Options, PublicXlsx["WorkbookReader"]["Options"]>>;
/** The runtime WorksheetReader conforms to index.d.ts `stream.xlsx.WorksheetReader`, but for the known deviations given. */
export type WorksheetReaderConforms = Assert<
    Check<
        WorksheetReader,
        Public.stream.xlsx.WorksheetReader,
        StreamRules,
        {
            // `undefined` only for an invalid argument (a number that is not a positive integer): declared `Column`, as
            // its TSDoc says (rule 2)
            getColumn(indexOrKey: number | string): Public.Column | undefined;
        }
    >
>;
/** The runtime HyperlinkReader conforms to index.d.ts `stream.xlsx.HyperlinkReader`. */
export type HyperlinkReaderConforms = Assert<Check<HyperlinkReader, Public.stream.xlsx.HyperlinkReader, StreamRules>>;

/** The events of the runtime WorkbookReader conform to index.d.ts `stream.xlsx.WorkbookReaderEvents`. */
export type WorkbookReaderEventsConform = Assert<Check<WorkbookReaderEvents, Public.stream.xlsx.WorkbookReaderEvents, StreamRules>>;
/** The events of the runtime WorksheetReader conform to index.d.ts `stream.xlsx.WorksheetReaderEvents`. */
export type WorksheetReaderEventsConform = Assert<Check<WorksheetReaderEvents, Public.stream.xlsx.WorksheetReaderEvents, StreamRules>>;
/** The events of the runtime HyperlinkReader conform to index.d.ts `stream.xlsx.HyperlinkReaderEvents`. */
export type HyperlinkReaderEventsConform = Assert<Check<HyperlinkReaderEvents, Public.stream.xlsx.HyperlinkReaderEvents, StreamRules>>;
