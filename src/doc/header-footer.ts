// The text of a header or footer (`worksheet.headerFooter.oddHeader`...) from its three sections, and back
// (ECMA-376 Part 1 18.3.1.36 headerFooter: the formatting codes). Pure functions, no state: the package exports them as
// `headerFooter` (index.d.ts `headerFooter` namespace).
//
// The texts of the file use `&` codes: `&L` / `&C` / `&R` start the left / center / right section, `&P` is the page
// number, `&&` a literal ampersand, `&B` toggles bold, `&"Arial,Bold"` sets the font... Excel's dialog shows the fields
// with friendly names instead (`&[Page]`): `build` takes them and turns them into the file codes, `parse` gives them
// back. Both scan the text once with indexOf (linear, no regular expression).

/** The three sections of a header or footer: the text of each one, as `headerFooter.build` takes it. */
export interface HeaderFooterSections {
    /** The text on the left. */
    left?: string | undefined;
    /** The text in the center. */
    center?: string | undefined;
    /** The text on the right. */
    right?: string | undefined;
}

/**
 * The longest header or footer that Excel takes: 255 characters (the text of the file, all sections and codes included;
 * Excel's dialog does not accept more).
 */
export const MAX_HEADER_FOOTER_LENGTH = 255;

/** The fields of Excel's dialog (`&[Page]`) and their codes in the file (`&P`). */
const FIELDS: ReadonlyMap<string, string> = new Map([
    ["[Page]", "P"],
    ["[Pages]", "N"],
    ["[Date]", "D"],
    ["[Time]", "T"],
    ["[Path]", "Z"],
    ["[File]", "F"],
    ["[Tab]", "A"],
]);

/** The field of a picture: the library cannot write the image part it needs. */
const PICTURE = "[Picture]";

/** The file codes of the fields, back to their friendly names (`P` → `&[Page]`), the picture included. */
const FIELD_NAMES: ReadonlyMap<string, string> = new Map([
    ...[...FIELDS].map(([name, code]): [string, string] => [code, name]),
    ["G", PICTURE],
]);

/**
 * The letters of the formatting codes kept as they are: bold, italic, underline, double underline, strikethrough,
 * superscript, subscript, outline, shadow, color (`&KFF0000`), and the field codes (`&P`...).
 */
const FORMAT_CODES = new Set(["B", "I", "U", "E", "S", "X", "Y", "O", "H", "K", "P", "N", "D", "T", "Z", "F", "A"]);

/**
 * Turns the text of one section into its text in the file: the fields of Excel's dialog (`&[Page]`) become their codes
 * (`&P`); the formatting codes (`&B`, `&"Font,Style"`, `&14`, `&KFF0000`, `&P`...) and `&&` are kept; any other `&` is a
 * literal ampersand, written `&&`.
 * @throws Error for `&[Picture]` or `&G` (a picture needs an image part that is not written), and for a section code
 * (`&L`, `&C`, `&R`: give the sections as `left`, `center`, `right`)
 */
function buildSection(text: string, name: string): string {
    const parts: string[] = [];
    let from = 0;
    for (let at = text.indexOf("&"); at !== -1; at = text.indexOf("&", from)) {
        parts.push(text.slice(from, at));
        const next = text.charAt(at + 1);
        from = at + 2;
        if (next === "[") {
            // a constant number of comparisons at this position (no search for the "]": linear)
            let field: string | undefined;
            for (const candidate of FIELDS.keys()) {
                if (text.startsWith(candidate, at + 1)) {
                    field = candidate;
                }
            }
            const code = field === undefined ? undefined : FIELDS.get(field);
            if (field !== undefined && code !== undefined) {
                parts.push(`&${code}`);
                from = at + 1 + field.length;
            } else if (text.startsWith(PICTURE, at + 1)) {
                throw new Error(`Unsupported header/footer field in ${name}: &[Picture] (pictures are not written)`);
            } else {
                // not a field: a literal ampersand
                parts.push("&&[");
            }
        } else if (next === "&" || next === '"' || FORMAT_CODES.has(next) || (next >= "0" && next <= "9")) {
            parts.push(`&${next}`);
        } else if (next === "G") {
            throw new Error(`Unsupported header/footer code in ${name}: &G (pictures are not written)`);
        } else if (next === "L" || next === "C" || next === "R") {
            throw new Error(`Invalid header/footer code in ${name}: &${next} (give the sections as left, center and right)`);
        } else {
            // a literal ampersand (the character after it, if any, is read next)
            parts.push("&&");
            from = at + 1;
        }
    }
    parts.push(text.slice(from));
    return parts.join("");
}

/**
 * The text of a header or footer (for `worksheet.headerFooter.oddHeader`...) from its sections: `&L`, `&C`, `&R` then
 * the text of each section given (absent or empty sections are left out). In a section, the fields of Excel's dialog
 * become their codes (`&[Page]` → `&P`, `&[Pages]` → `&N`, `&[Date]` → `&D`, `&[Time]` → `&T`, `&[Path]` → `&Z`,
 * `&[File]` → `&F`, `&[Tab]` → `&A`); the formatting codes (`&B`, `&I`, `&U`, `&"Font,Style"`, `&14`, `&KFF0000`...),
 * the field codes (`&P`...) and `&&` are kept; any other `&` is a literal ampersand (written `&&`). A text that may hold
 * `&` followed by a code letter (`AT&T`) goes through {@link escape} first.
 * @throws Error for `&[Picture]` (`&G`): a picture needs an image part that the library does not write; for a section
 * code in a section (`&L`, `&C`, `&R`); for a section that is not a string; and for a result longer than
 * {@link MAX_HEADER_FOOTER_LENGTH} characters, which Excel does not take
 */
export function build(sections: HeaderFooterSections): string {
    let text = "";
    for (const [name, code] of [
        ["left", "L"],
        ["center", "C"],
        ["right", "R"],
    ] as const) {
        const section: unknown = sections[name];
        if (section === undefined || section === "") {
            continue;
        }
        if (typeof section !== "string") {
            throw new Error(`Invalid header/footer section ${name}: a string is expected`);
        }
        text += `&${code}${buildSection(section, name)}`;
    }
    if (text.length > MAX_HEADER_FOOTER_LENGTH) {
        throw new Error(`Header/footer too long: ${text.length} characters (Excel allows at most ${MAX_HEADER_FOOTER_LENGTH})`);
    }
    return text;
}

/**
 * The sections of the text of a header or footer (as read from a file): the text before any section code is the center
 * (as in Excel), a section given twice has its texts joined. Each section is given in the form that {@link build} takes:
 * the field codes as the fields of Excel's dialog (`&P` → `&[Page]`, `&G` → `&[Picture]`), `&&` and the formatting codes
 * as they are. Sections without text are left out. Any length is read.
 */
export function parse(text: string): HeaderFooterSections {
    const sections: { left: string[]; center: string[]; right: string[] } = { left: [], center: [], right: [] };
    let current = sections.center;
    let from = 0;
    for (let at = text.indexOf("&"); at !== -1; at = text.indexOf("&", from)) {
        current.push(text.slice(from, at));
        const next = text.charAt(at + 1);
        from = at + 2;
        if (next === "L") {
            current = sections.left;
        } else if (next === "C") {
            current = sections.center;
        } else if (next === "R") {
            current = sections.right;
        } else {
            const field = FIELD_NAMES.get(next);
            // a trailing "&" stays a single character
            current.push(field ? `&${field}` : `&${next}`);
        }
    }
    current.push(text.slice(from));
    const result: HeaderFooterSections = {};
    for (const name of ["left", "center", "right"] as const) {
        const section = sections[name].join("");
        if (section) {
            result[name] = section;
        }
    }
    return result;
}

/**
 * A text with its ampersands doubled (`AT&T` → `AT&&T`): a text shown as it is in a section given to {@link build}
 * (otherwise `&T` would be the time). Also for the text of `worksheet.headerFooter.oddHeader`... set without build.
 */
export function escape(text: string): string {
    return text.split("&").join("&&");
}

/** The helpers of the header and footer texts (the `headerFooter` export of the package). */
const headerFooter = { build, parse, escape };

export default headerFooter;
