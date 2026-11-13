// The enumerations of the public API (index.d.ts declares them as `enum`s with the same names, keys and values). They
// are plain objects, not TypeScript enums: a numeric enum would add the reverse mappings (ValueType[2] ===
// "Number") to the objects that the package entry point spreads into its exports. Each one is exported both as a value
// and as the union type of its values (like an enum): `ValueType.Number` is `2`, `ValueType` is `0 | 1 | ... | 10`.

/** The type of a cell value (`cell.type`, `type` of the cell models). */
export const ValueType = {
    Null: 0,
    Merge: 1,
    Number: 2,
    String: 3,
    Date: 4,
    Hyperlink: 5,
    Formula: 6,
    SharedString: 7,
    RichText: 8,
    Boolean: 9,
    Error: 10,
} as const;
/** A {@link ValueType} value. */
export type ValueType = (typeof ValueType)[keyof typeof ValueType];

/** The kind of formula of a cell (`cell.formulaType`). */
export const FormulaType = {
    None: 0,
    Master: 1,
    Shared: 2,
} as const;
/** A {@link FormulaType} value. */
export type FormulaType = (typeof FormulaType)[keyof typeof FormulaType];

/** Relationship types (legacy: the relationships of the parts use the URIs of `xlsx/rel-type`). */
export const RelationshipType = {
    None: 0,
    OfficeDocument: 1,
    Worksheet: 2,
    CalcChain: 3,
    SharedStrings: 4,
    Styles: 5,
    Theme: 6,
    Hyperlink: 7,
} as const;
/** A {@link RelationshipType} value. */
export type RelationshipType = (typeof RelationshipType)[keyof typeof RelationshipType];

/** Document types. */
export const DocumentType = {
    Xlsx: 1,
} as const;
/** A {@link DocumentType} value. */
export type DocumentType = (typeof DocumentType)[keyof typeof DocumentType];

/** The `readingOrder` attribute of an alignment (the model uses `"ltr"` / `"rtl"`). */
export const ReadingOrder = {
    LeftToRight: 1,
    RightToLeft: 2,
} as const;
/** A {@link ReadingOrder} value. */
export type ReadingOrder = (typeof ReadingOrder)[keyof typeof ReadingOrder];

/** The error values of a cell (`{ error: ErrorValue.Ref }`). */
export const ErrorValue = {
    NotApplicable: "#N/A",
    Ref: "#REF!",
    Name: "#NAME?",
    DivZero: "#DIV/0!",
    Null: "#NULL!",
    Value: "#VALUE!",
    Num: "#NUM!",
} as const;
/** An {@link ErrorValue} value. */
export type ErrorValue = (typeof ErrorValue)[keyof typeof ErrorValue];

/**
 * Common paper sizes of `pageSetup.paperSize` (the ST_PaperSize numbers of ECMA-376; any other number of the standard is
 * accepted too).
 */
export const PaperSize = {
    Legal: 5,
    Executive: 7,
    A4: 9,
    A5: 11,
    B5: 13,
    Envelope_10: 20,
    Envelope_DL: 27,
    Envelope_C5: 28,
    Envelope_B5: 34,
    Envelope_Monarch: 37,
    Double_Japan_Postcard_Rotated: 82,
    K16_197x273_mm: 119,
} as const;
/** A {@link PaperSize} value. */
export type PaperSize = (typeof PaperSize)[keyof typeof PaperSize];

/** All the enumerations, as one object (the package entry point spreads it into its exports). */
const Enums = {
    ValueType,
    FormulaType,
    RelationshipType,
    DocumentType,
    ReadingOrder,
    ErrorValue,
    PaperSize,
};

export default Enums;
