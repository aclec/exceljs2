// Type-level checks (no runtime content): the classes of src/doc against the public API of index.d.ts.
//
// Each check compiles only when every public member of the index.d.ts type conforms to the runtime class; on failure
// the error names the members that do not conform. See conformance-helpers.ts for what is compared (and what is not).
//
// The known deviations are listed with their reason: the rules of `DocRules` for the types that appear in several
// members, the declarations given to a check for the members of one class. Aligning index.d.ts with the runtime removes
// them; a member deviation that no longer deviates fails its check (`stale:name`).

import type * as Public from "../../index";
import type Anchor from "../doc/anchor";
import type Cell from "../doc/cell";
import type { FormulaValueObject } from "../doc/cell";
import type Row from "../doc/row";
import type Column from "../doc/column";
import type Range from "../doc/range";
import type DataValidations from "../doc/data-validations";
import type DefinedNames from "../doc/defined-names";
import type Table from "../doc/table";
import type Worksheet from "../doc/worksheet";
import type Workbook from "../doc/workbook";
import type ModelContainer from "../doc/modelcontainer";
import type { ColorScaleRuleModel, IconSetRuleModel, ImageHyperlinkModel } from "./sheet";
import type { Assert, Check, ComparedAs } from "./conformance-helpers";

/**
 * The declared types compared as their runtime class wherever they appear (see conformance-helpers.ts): each one is
 * checked on its own below.
 */
export type DocSubstitutions =
    | [Public.Anchor, Anchor]
    | [Public.Range, Range]
    | [Public.DataValidations, DataValidations]
    | [Public.Cell, Cell]
    | [Public.Row, Row]
    | [Public.Column, Column]
    | [Public.DefinedNames, DefinedNames]
    | [Public.Table, Table]
    | [Public.Worksheet, Worksheet]
    | [Public.Workbook, Workbook]
    | [Public.ModelContainer, ModelContainer];

/**
 * The rules of the checks of src/doc (see conformance-helpers.ts): the substitutions, and the known deviations of the
 * types that appear in several members (values that only invalid input or a malformed file gives: rule 2).
 */
export type DocRules =
    | DocSubstitutions
    // The JSON value of a cell: the `object` member of the runtime `CellValue` (src/doc/cell.ts), any object that is not
    // a cell value. index.d.ts `CellValue` keeps it out on purpose: it comes only from a value outside `CellValue` (an
    // untyped one), or from the falsy-marker quirk (`{ text: "", hyperlink }`, `{ sharedString: 0 }`...); describing it
    // would make every cell value an untyped `object`. Not compared wherever cell values appear (`Cell.value`,
    // `Row.values`, `Column.values`, `getSheetValues`...).
    | ComparedAs<object, never>
    // A formula value has its own formula or the address of its shared formula master, but for a formula read from a
    // malformed file (a shared formula clone whose `si` has no master): its value is `{ result }`.
    | ComparedAs<FormulaValueObject, FormulaValueObject & ({ formula: string } | { sharedFormula: string })>
    // A colorScale / iconSet rule read from a malformed file (`<cfRule type="colorScale">` without its `<colorScale>`,
    // which CT_ColorScale requires) has no `cfvo` / `color`: declared required, as `addConditionalFormatting` needs them.
    | ComparedAs<ColorScaleRuleModel, ColorScaleRuleModel & Required<Pick<ColorScaleRuleModel, "cfvo" | "color">>>
    | ComparedAs<IconSetRuleModel, IconSetRuleModel & Required<Pick<IconSetRuleModel, "cfvo">>>
    // The link of an image (`getImages()[].range.hyperlinks`) has no `hyperlink` only when read from a file whose
    // `<a:hlinkClick>` has no relationship (no `r:id`: Excel writes one): declared required in `WorksheetImage` (and
    // given so to `addImage`), optional in `WorksheetMediumModel`.
    | ComparedAs<ImageHyperlinkModel, ImageHyperlinkModel & { hyperlink: string }>;

/** The runtime Anchor conforms to index.d.ts `Anchor`. */
export type AnchorConforms = Assert<Check<Anchor, Public.Anchor, DocRules>>;
/** The runtime Range conforms to index.d.ts `Range`. */
export type RangeConforms = Assert<Check<Range, Public.Range, DocRules>>;
/** The runtime DataValidations conforms to index.d.ts `DataValidations`. */
export type DataValidationsConforms = Assert<Check<DataValidations, Public.DataValidations, DocRules>>;
/** The runtime Cell conforms to index.d.ts `Cell`. */
export type CellConforms = Assert<Check<Cell, Public.Cell, DocRules>>;
/** The runtime Row conforms to index.d.ts `Row`. */
export type RowConforms = Assert<Check<Row, Public.Row, DocRules>>;
/** The runtime Column conforms to index.d.ts `Column`. */
export type ColumnConforms = Assert<Check<Column, Public.Column, DocRules>>;
/** The runtime DefinedNames conforms to index.d.ts `DefinedNames`. */
export type DefinedNamesConforms = Assert<Check<DefinedNames, Public.DefinedNames, DocRules>>;
/** The runtime Table conforms to index.d.ts `Table`. */
export type TableConforms = Assert<Check<Table, Public.Table, DocRules>>;
/** The runtime Worksheet conforms to index.d.ts `Worksheet`, but for the known deviations given. */
export type WorksheetConforms = Assert<
    Check<
        Worksheet,
        Public.Worksheet,
        DocRules,
        {
            // `undefined` only for an invalid number (not an integer >= 1): declared `Column`, as its TSDoc says (rule 2)
            getColumn(indexOrKey: number | string): Public.Column | undefined;
        }
    >
>;
/** The runtime Workbook conforms to index.d.ts `Workbook`. */
export type WorkbookConforms = Assert<Check<Workbook, Public.Workbook, DocRules>>;
/** The runtime ModelContainer conforms to index.d.ts `ModelContainer`. */
export type ModelContainerConforms = Assert<Check<ModelContainer, Public.ModelContainer, DocRules>>;
