# UPGRADE FROM 4.x to 5.0

Version 5.0 modernises the runtime requirements and hardens the parsing of untrusted files. Most of the API is unchanged; the sections below
list everything that can break an existing application. See [CHANGELOG.md](CHANGELOG.md) for the full list of changes.

## Package name

5.0 is published by the maintained fork as **`@aclec/exceljs2`** (the `exceljs` name on npm belongs to the original, unmaintained project).
Either change your imports to `@aclec/exceljs2`, or keep `require('exceljs')` / `import 'exceljs'` unchanged by installing the fork under an
alias:

```shell
npm install exceljs@npm:@aclec/exceljs2
```

## Runtime requirements

|          | ExcelJS 4.x                               | ExcelJS 5.0                                                                     |
| -------- | ----------------------------------------- | ------------------------------------------------------------------------------- |
| Node.js  | >= 8.3 (ES5 build for older versions)     | **>= 22.12**                                                                    |
| Bun      | not supported                             | **>= 1.4**                                                                      |
| Browsers | ES5 bundles with core-js polyfills, IE 11 | **ES2020 syntax, ES2022 built-ins** (Chrome / Edge 93, Firefox 92, Safari 15.4) |

Applications that must run on an older Node.js or on Internet Explorer should stay on 4.x.

## ES5 build removed

`exceljs/dist/es5` no longer exists, and the browser bundles no longer include or require the `core-js` / `regenerator-runtime` polyfills.

| ExcelJS 4.x                                                                       | ExcelJS 5.0                     |
| --------------------------------------------------------------------------------- | ------------------------------- |
| `require('exceljs/dist/es5')`                                                     | `require('exceljs')`            |
| `require('core-js/...')`, `require('regenerator-runtime/runtime')` before ExcelJS | not needed                      |
| `<script src="polyfill.js">` + `<script src="exceljs.js">`                        | `<script src="exceljs.min.js">` |

## Package exports

`package.json` now has an `exports` map. Two consequences:

- **Deep imports** are limited to `exceljs/lib/...` (with or without `.js`; `.json` and `.xml` files with their extension),
  `exceljs/dist/...`, `exceljs/excel(.js)` and `exceljs/package.json`. Other paths (`exceljs/index.d.ts`, `exceljs/lib/doc/data/theme1`
  without `.json`, ...) no longer resolve. The internal modules `lib/csv/stream-converter.js`, `lib/utils/auto-drain`,
  `browser-buffer-encode`, `iterate-stream`, `stream-base64`, `string-builder`, `stuttered-pipe` and `typed-stack` were removed.
- **Internal modules are compiled from TypeScript.** The library is written in TypeScript (`src/`) and `lib/` holds the compiled CommonJS
  output with its `.d.ts` files. The package entry point is unchanged (`require('exceljs')` returns the same object), but a deep import of
  an internal module that used to be `module.exports = X` now exposes `X` as its default export: use
  `require('exceljs/lib/doc/workbook').default`. TypeScript compiled to CommonJS can write
  `import Workbook from 'exceljs/lib/doc/workbook'`; in native ES modules the default import is the CommonJS exports object, so the class is
  `(await import('exceljs/lib/doc/workbook')).default.default`. Internal modules are not part of the public API and may change in any
  release.
- **The `browser` condition** resolves the package to the browser bundle `dist/exceljs.min.js` (bundlers targeting browsers did so through
  the `browser` field already). The bundle has no `stream` namespace and no file APIs. Test runners that apply the `browser` condition in
  Node now load the bundle too: **Jest with the jsdom environment** does by default (`customExportConditions: ['browser']`). To test with
  the Node build, either change the conditions or map the package to its Node entry point:

```javascript
// jest.config.js
module.exports = {
    testEnvironment: "jsdom",
    testEnvironmentOptions: { customExportConditions: ["node", "node-addons"] },
    // or: moduleNameMapper: { "^exceljs$": "<rootDir>/node_modules/exceljs/excel.js" },
};
```

## Dependencies

Runtime dependencies were upgraded, three were removed and the three zip libraries were replaced by one. They are internal to ExcelJS, but
you may notice them in your lockfile or bundle:

| Dependency      | 4.x      | 5.0                                                                            |
| --------------- | -------- | ------------------------------------------------------------------------------ |
| dayjs           | ^1.8.34  | ^1.11.23                                                                       |
| fast-csv        | ^4.3.1   | ^5.0.8                                                                         |
| saxes           | ^5.0.1   | ^6.0.0                                                                         |
| archiver        | ^5.0.0   | removed: see [Zip library](#zip-library-fflate)                                |
| jszip           | ^3.10.1  | removed: see [Zip library](#zip-library-fflate)                                |
| unzipper        | ^0.10.11 | removed: see [Zip library](#zip-library-fflate)                                |
| fflate          | -        | ^0.8.3 (new, no dependencies)                                                  |
| readable-stream | ^3.6.0   | removed: `node:stream` is used (`stream-browserify` inside the browser bundle) |
| tmp             | ^0.2.0   | removed: `fs.promises.mkdtemp`                                                 |
| uuid            | ^8.3.0   | removed: `crypto.randomUUID` / `crypto.getRandomValues`                        |

Side effects:

- ExcelJS no longer calls `tmp.setGracefulCleanup()` when it is loaded, so it no longer changes the behaviour of the `tmp` package for the
  rest of your application.
- The streaming `WorkbookReader` spools worksheets that come before the shared strings to a temporary directory (`os.tmpdir()/exceljs-*`)
  that is removed when the iteration ends, including when you `break` out of it early or an error is thrown. 4.x leaked the files and their
  file descriptors in that case.
- Streams created by ExcelJS are `node:stream` streams instead of `readable-stream` ones. The streaming `WorkbookReader` still accepts any
  readable stream as input, including `readable-stream` streams created by other libraries (through2, duplexify, ...).

## Limits when reading files

To make parsing of untrusted files safe, a few inputs that used to be accepted (and could exhaust memory or CPU) are now rejected or stored
differently. See also the [Security & limits](README.md#security--limits) section of the README and [SECURITY.md](SECURITY.md).

### Rows and columns beyond the Excel grid throw

Cell references must be within rows 1 to 1,048,576 and columns A to XFD (16,384), the size of an Excel worksheet. A reference with a larger
row (for example `A1048577`) now throws `Out of bounds. Excel supports rows from 1 to 1048576`; 4.x only checked the column. `<col min max>`
definitions read from a file are clamped to column 16,384, and a `<row r>` outside 1 to 1,048,576 rejects the file (document and streaming
readers).

### Sheet ids are renumbered when out of range

Worksheets are stored by `sheetId`. A `sheetId` read from a file is kept when it is an integer from 1 to 65,535 that no earlier sheet uses;
otherwise (a huge, zero, non-numeric or duplicate id) the sheet gets the lowest free id. 4.x kept huge ids, which made `workbook.worksheets`
and `getWorksheet()` slow, and a duplicate id replaced the earlier worksheet.

### Long number formats are ignored

Excel limits number format codes to 255 characters. A custom `formatCode` longer than that is ignored when a file is read: its cells read as
General (`cell.numFmt` is `undefined`).

### Merged cells are capped

Every cell of a merge is created in memory (300 to 600 bytes each), so merges are capped, with the same rules when a workbook is built and
when it is read (a workbook ExcelJS writes with the defaults can be read back with the defaults):

- A worksheet can have at most **1,048,576 merged cells** in total (the size of one full column). `worksheet.mergeCells()`, a splice that
  moves or widens merges, the streaming `WorkbookWriter` and loading a file throw
  `Cannot merge <range>: worksheet would exceed 1048576 merged cells` beyond that. This cap is not configurable.
- A workbook can have at most **`workbook.maxMergedCells`** merged cells in all its worksheets together, **1,048,576** (one full worksheet,
  about 600 MB of memory) by default. Beyond it they throw
  `Cannot merge <range>: the workbook would exceed 1048576 merged cells (see the maxMergedCells option)`. Raise it with the `maxMergedCells`
  option of `xlsx.load` / `read` / `readFile` (the loaded workbook keeps that value), by setting `workbook.maxMergedCells` before merging or
  loading, or with the `maxMergedCells` option of the streaming `WorkbookWriter`:

```javascript
const workbook = new ExcelJS.Workbook();
await workbook.xlsx.load(data, { maxMergedCells: 16 * 1048576 }); // trusted input only
```

The check for overlapping merges now takes time proportional to the area being merged instead of the number of existing merges, so files
with tens of thousands of merges load quickly. The `Cannot merge already merged cells` error is unchanged.

### Large data-validation ranges stay ranges

In 4.x, every `sqref` range read from a file was expanded to one entry per cell in `worksheet.dataValidations.model`, so
`sqref="A1:XFD1048576"` could exhaust memory. In 5.0, ranges are expanded until 100,000 cells have been expanded for the workbook; ranges
that do not fit are stored under a single **range key**:

```js
// 4.x: {B1: {...}, B2: {...}, ..., B1048576: {...}}
// 5.0: {'B1:B1048576': {...}}
worksheet.dataValidations.model;
```

`cell.dataValidation` resolves both kinds of keys and both are written back to the file, so code that goes through cells is not affected:
setting or removing the validation of a cell inside a range key replaces that key by the rectangles around the cell. Code that reads
`dataValidations.model` directly must handle keys of the form `'A1:B2'`. Range keys are looked up through an index (a lookup does not scan
every key). Range keys may also be written to the model directly: write through a fresh `worksheet.dataValidations.model` (reading the
property refreshes the index), or use `dataValidations.add(range, validation)`; a reference to the model kept across lookups is not
re-indexed.

Also:

- `sqref` tokens that are not cell references (`A1` or `A1:B2`, with or without `$`) are ignored.
- Validations removed with `dataValidations.remove(address)` are no longer written as empty entries.

### Large defined-name ranges stay ranges

Defined names are stored one entry per cell so that they follow row and column splices. A defined-name range covering more than **100,000
cells** (for example `Sheet1!$A$1:$XFD$1048576`) is now kept as a single rectangle instead, and so are the ranges read from a file once
100,000 cells have been expanded for all its names. Such ranges:

- are returned by `workbook.definedNames.model`, `getRanges(name)` and `cell.names`;
- move or shrink with `spliceRows` / `spliceColumns`;
- are **not** enumerated cell by cell by `workbook.definedNames.forEach()`;
- are removed as a whole when the name is removed from any of their cells (`cell.removeName(name)`, `definedNames.remove(...)`).

### Formulas longer than 8,192 characters

Excel limits formulas to 8,192 characters. ExcelJS now throws `Formula exceeds the Excel limit of 8192 characters` when it has to analyse a
longer formula: translating a shared formula to another cell (`cell.formula` on a shared-formula clone) and `cell.dependencies`.

## Hyperlinks to a place in the workbook

Internal links (Excel's "Place in This Document": a `location` without a relationship) are now read. 4.x ignored them and returned the plain
cell value. Such cells now have type `ValueType.Hyperlink` and a `{text, hyperlink}` value, with `#` before the location:

```js
// 4.x: 42                                   (type Number)
// 5.0: {text: 42, hyperlink: '#Data!A1'}    (type Hyperlink)
worksheet.getCell("A1").value;
```

`String(cell.value)` then gives `[object Object]`; use `cell.text`. An external link with a location (an anchor) is read as
`'https://example.com/doc#part2'`.

On write, a hyperlink that starts with `#` (`'#Sheet2!A1'`, `"#'My Sheet'!B2:C3"`, `'#MyDefinedName'`) is written as Excel writes internal
links: `location="Sheet2!A1"`, without `#` and without an external relationship. 4.x wrote a relationship to `#...` (and, for `Sheet!A1`
targets, a location with the `#`). The legacy form without `#` (`'Sheet2!A1'`, `"'My Sheet'!B2:C3"`: a sheet name, `!`, then a cell or a
range) is also written as an internal link, so it reads back as `'#Sheet2!A1'`. An unquoted sheet name in that form cannot contain `:`, `/`,
`\` or `?`, nor start or end with `'` (as in Excel), so URLs and paths whose end looks like a sheet reference
(`'https://example.com/#!page1'`, `'../book.xlsx#Sheet1!A1'`) stay external links, while `"Bob's!A1"` and `'Q#1!B2'` are internal; prefix
the target with `#` (`"#'Q#1'!B2"`), or quote the sheet name, for any other sheet name.

## Styles are copied on write

Style setters (`cell.font = ...`, `row.fill = ...`, `column.alignment = ...`, `cell.style = ...`) no longer modify a style object that is
shared with other cells, rows or columns, or that your code passed in. In 4.x, cells that shared a style object (loaded from a file, or
assigned with `b.style = a.style`) were all restyled:

```js
b.style = a.style;
a.font = { bold: true };
b.font; // 4.x: {bold: true}; 5.0: unchanged
```

Set the property on each cell that must change. Merged cells still share the style of their master cell. Copying costs time and memory when
restyling many cells: about twice as slow per setter, and about 26 MB of heap for restyling 200,000 loaded cells.

## Other behaviour changes

- **Number formats keep their `\` escapes.** 4.x removed every backslash of the number formats it read (cells, columns, rows and the
  differential formats of conditional formatting), so formats were written back changed, `\#0` became a digit placeholder and `0.0\ \m` was
  taken for a date format. Excel writes escaped formats routinely; code that compares `cell.numFmt` with a string may need the escaped form:

    | Format in the file | `cell.numFmt` in 4.x | `cell.numFmt` in 5.0 |
    | ------------------ | -------------------- | -------------------- |
    | `yyyy\-mm\-dd;@`   | `yyyy-mm-dd;@`       | `yyyy\-mm\-dd;@`     |
    | `d\ mmmm\ yyyy`    | `d mmmm yyyy`        | `d\ mmmm\ yyyy`      |
    | `#,##0.00\ "€"`    | `#,##0.00 "€"`       | `#,##0.00\ "€"`      |

- **Conditional formatting read model.** Rules read from files now match what ExcelJS writes:
    - the `numFmt` of a rule `style` is the format code string (`'0.00%'`), like any style (4.x: `{id, formatCode}`);
    - `beginsWith`, `endsWith` and `notContainsText` rules are read as `{type: 'containsText', operator: 'beginsWith', text: 'x'}` (4.x:
      `type: 'beginsWith'` without `text`, and they were dropped on write); `text` is read for all text rules;
    - x14 data bars without explicit attributes get the schema defaults: `direction: 'context'`, `negativeBarColorSameAsPositive: false`,
      `minLength: 10`, `maxLength: 90` (4.x: `'leftToRight'`, `true`, `0`, `100`). New data bars are still written with Excel's 0 / 100;
    - threshold values (`cfvo.value`) that are formulas or cell references are kept as strings (4.x: `NaN` or a partial number).

- **Merged cells follow splices.** `spliceRows`, `insertRow(s)`, `duplicateRow` and `spliceColumns` move, shrink or widen the merges with
  the cells (4.x left the merges at their old place, see the [README](README.md#merged-cells)). As in Excel, rows or columns inserted inside
  a merge widen it, and only its master cell holds a value: inserting values into the cells of a merge now throws
  `Cannot insert a value in B3: it would be inside the merged cells B2:C5, unmerge them first`, before anything changes (4.x kept the values
  in a broken merge). Empty inserted cells (`null`, `undefined`, `''`) are accepted. Call `worksheet.unMergeCells()` first to insert values
  there.

- **Invalid XML characters are removed on write.** U+FFFE, U+FFFF and unpaired surrogates (which are not allowed in XML 1.0) are stripped
  from text when writing xlsx, like control characters already were. Valid surrogate pairs (emoji, etc.) are kept.
- **CSV special values.** When reading CSV, cells whose text is `__proto__`, `constructor`, `toString`, ... are read as plain strings. 4.x
  returned `Object.prototype` or built-in functions.
- **CSV numbers.** Whitespace-only cells (`' '`) are read as strings (4.x: `0`). Integers beyond `Number.MAX_SAFE_INTEGER` (20-digit ids)
  are read as strings instead of rounded numbers.
- **CSV headers.** `parserOptions.headers` is supported (4.x threw). The header names become the column keys and row 1 of the worksheet; the
  data starts at row 2.
- **Tables.** When a file omits `headerRowCount` (default 1), `table.headerRow` and `filterButton` are read as `true` (4.x: `false`). A
  table without a header row is written without `<autoFilter>`.
- **Alignment.** `wrapText` and `shrinkToFit` attributes equal to `0` (the default) are no longer read into `alignment` as `false`.
- **`worksheet.tables`** is still a plain object keyed by table name. Use `worksheet.getTable(name)`, which only returns tables (a table may
  be named `constructor` or `__proto__`).
- **Streaming `WorksheetReader.id`** is the `sheetId` of the worksheet (a number) whenever the worksheet can be matched to `workbook.xml`.
  4.x failed to match more files (absolute relationship targets, worksheets stored before `workbook.xml`) and then gave the number of the
  sheet file as a string (`'1'` for `sheet1.xml`).
- **Streaming `parse()` with `sharedStrings: 'emit'`** yields `{eventType: 'shared-strings', value: {index, text}}`, like the other events.
  4.x yielded a bare `{index, text}`.
- **ISO 8601 date cells (`t="d"`)** are read as dates by both readers, as UTC when they have no offset. `T24:00` is 00:00 of the next day;
  values that are not valid dates stay strings.
- **Object keys `__proto__`, `constructor` and `prototype`** in objects given to ExcelJS (for example `cell.note = JSON.parse(...)`) are
  ignored when they are merged into internal models.
- **Pivot tables** (experimental): field names and values are XML-escaped.
- **Rows and cells that splices do not need are not created.** `spliceColumns` no longer creates the rows that get no inserted value (4.x
  gave them empty cells, written as empty `<row>` elements); the rows moved by `spliceRows`, `insertRow(s)` and `duplicateRow` get the cells
  their source has, not an empty styled cell in every column up to the last one; `duplicateRow` of a row that does not exist duplicates an
  empty row (4.x threw a `TypeError`). Values, styles, notes and `cellCount` are unchanged.
- **Truncated XML parts are rejected.** A part that ends before its root element is closed (or has no root element) makes `xlsx.load` and
  the streaming reader reject with the error of the XML parser (`unclosed tag: sheetData`). 4.x read it as if it were complete (the
  streaming reader returned the rows before the cut) or failed with a `TypeError`.
- **`_xHHHH_` in strings.** A literal `_xHHHH_` in a string is written escaped as `_x005F_xHHHH_` (ECMA-376), so that Excel and ExcelJS read
  it back as written. 4.x wrote it as is, and it was read as the character U+HHHH (`x_x0041_y` became `xAy`).
- **Print areas and print titles** are written as absolute references (`'Sheet'!$A$1:$G$20`), whether `pageSetup.printArea` is given with
  `$` or not. 4.x wrote `$A1:$G20`, and `$$A$1:$$B$2` for `$A$1:$B$2`. They are still read without `$`.
- **Streaming reader shared strings** are read like the document reader reads them: rich text runs keep their whole font (4.x lost the
  colour, name and charset of most runs), phonetic runs are no longer read as the text, and `_xHHHH_` escapes are decoded.
- **`worksheet.model = model`** replaces the merges of the worksheet (4.x kept the previous merges next to the new rows).
- **Streaming writer backpressure.** `row.commit()` and `worksheet.commit()` of the streaming `WorkbookWriter` return a promise (4.x:
  `undefined`) that resolves once the output can take more data. `await worksheet.addRow(values).commit()` keeps memory bounded on large
  sheets; code that ignores the promise works as before (and keeps all the pending data in memory until it yields to the event loop). See
  the [README](README.md#streaming-xlsx). In TypeScript, `commit()` is typed `void | Promise<void>`.
- **Invalid references throw.** `getCell`, `getRow`, `getRows`, `mergeCells`, `mergeCellsWithoutStyle`, `unMergeCells` and `fillFormula`
  (document and streaming worksheets) throw a clear `Error` for an address or range they cannot place: row 0, NaN or fractional numbers,
  bounds outside the grid, no arguments, unparsable or lowercase references (`"a1:b2"` merged A1:A2 in 4.x), or a sheet name other than the
  worksheet's. The worksheet's own sheet name is accepted and dropped (`mergeCells("Sheet1!A1:B2")` on `Sheet1`). `findCell` returns
  `undefined` in those cases. `fillFormula` stores its `ref` without sheet name or `$`, and fills a single cell (`"C3"`).
- **Views and header/footer after a load.** `worksheet.views` and `workbook.views` are `[]`, and `worksheet.headerFooter` the default
  object, for a file without them (4.x: `null` / `undefined`, so `views.push(...)` threw).
- **Page setup.** Row breaks are read (at most 1,023, Excel's limit; invalid `<brk>` attributes reject the load). `firstPageNumber` is not
  set when the file uses automatic numbering, and `useFirstPageNumber` is read and written as given (4.x read `firstPageNumber: 1`).
- **Icon sets.** A rule read without `iconSet` has the schema default `"3TrafficLights1"` (4.x: `"3TrafficLights"`, which is not a valid
  value; it is still accepted on write, as the default). Custom icons are written with their `iconId`, which is required (except for
  `NoIcons`): 4.x replaced it with the icon position. A rule with `icons` is written in the x14 extension, also without `custom`. An unknown
  icon set, an `iconId` outside its set or a number of icons different from the number of thresholds throws on write.
- **Colours and fonts.** An automatic colour (`auto="1"`) reads as `{ auto: true }` (4.x: `undefined`); font sizes keep their decimals.
- **Data-validation formulas** that are not plain numbers or dates (`$B$1`, `TODAY()`, `A1+1`) are read and written as text (4.x: `NaN`). A
  `date` rule takes a `Date`, an ISO 8601 date string or a formula string; an invalid date or a non-finite number throws on write.
- **Cell value markers** are recognised by their type, not their truthiness: `{ text: "", hyperlink }`, `{ sharedString: 0 }`,
  `{ formula: "", result }` are hyperlink, shared string and formula values (4.x stored them as JSON text), and an object whose marker has
  the wrong type (`{ formula: 1 }`, `{ richText: "x" }`) is a JSON value. `cell.text` is always a string.
- **Defined names.** Removing a cell or a range from a range kept as a rectangle (more than 100,000 cells) keeps the rest of the range, cut
  into at most four rectangles (4.x removed the whole range). Invalid range bounds throw.
- **Tables.** `addTable`, `table.addRow`, `addColumn` and `commit` validate the whole table first and throw a clear `Error` (a row with more
  values than columns, ...) without writing any cell; `getColumn(index)` out of range throws. A table read from a file without style returns
  `null` / `undefined` for `theme` and the `show*` properties.
- **Streaming reader hyperlinks.** With `hyperlinks: "cache"`, `worksheetReader.hyperlinks` entries have their `target` and `targetMode`.

## TypeScript

The typings were corrected to match the runtime. Code that relied on the old (wrong) types may need small changes:

| Typing                                                        | 4.x                                         | 5.0                                                       |
| ------------------------------------------------------------- | ------------------------------------------- | --------------------------------------------------------- |
| `Address.col`, `Address.row` (so also `cell.col`, `cell.row`) | `string`                                    | `number`                                                  |
| `Row.dimensions`                                              | `number`                                    | `{min: number; max: number} \| null`                      |
| `Row.values`                                                  | `CellValue[] \| {[key: string]: CellValue}` | getter: `CellValue[]`; setter: unchanged                  |
| `Column.values`                                               | `CellValue`                                 | `CellValue[]`                                             |
| `addImage(..., {editAs})`                                     | `string`                                    | `'oneCell' \| 'twoCell' \| 'absolute'` (`ImageEditAs`)    |
| `TableStyleProperties.theme`                                  | theme names                                 | theme names or `null` (no style)                          |
| `stream.xlsx.WorksheetReader` `dimensions`, `columns`         | methods                                     | getters; `columns` may be `null`                          |
| `csv.createInputStream()`                                     | declared                                    | removed (it has thrown since 4.0: use `csv.read(stream)`) |
| `writeBuffer()` result (`ExcelJS.Buffer`)                     | an `ArrayBuffer`-like interface             | a Node `Buffer` backed by an `ArrayBuffer`                |

`editAs` is now checked: an object literal stored in a variable before the call widens `editAs` to `string` and no longer compiles. Annotate
it instead:

```ts
const position = { tl: { col: 0, row: 0 }, ext: { width: 100, height: 50 }, editAs: "oneCell" as const };
// or: const editAs: ExcelJS.ImageEditAs = 'oneCell';
worksheet.addImage(imageId, position);
```

`(await workbook.xlsx.writeBuffer()) as Buffer` failed with TS2352 in 4.x; the result is now typed as what it is at run time (a Node
`Buffer`, or the Buffer polyfill in the browser bundle): it can be passed to `fs.writeFile` or cast to `Buffer`, and still to
`new Blob([data])`. Assigning it to an `ArrayBuffer` variable no longer compiles (it never was one); copy it with
`data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength)` when an `ArrayBuffer` is needed. `xlsx.load()` is typed
`ArrayBuffer | ArrayBufferView | Blob | number[]` (Node Buffers included), or base64 text with `{ base64: true }`.

New or widened types: `worksheet.dataValidations` (`DataValidations`: `add`, `find`, `remove`, `model`), `DataValidationType`,
`ImageEditAs`, all `cellIs` operators (`notEqual`, `greaterThanOrEqual`, `lessThanOrEqual`, `notBetween`), optional conditional-formatting
`priority`, data bar colours, optional `worksheet.protect()` arguments, optional `WorkbookReader` options, `WorksheetReader`
`id`/`name`/`state`, and the EventEmitter methods of the streaming readers. `xlsx.load()` accepts Node Buffers and `Uint8Array`, and
`Image.buffer` accepts Node Buffers, with TypeScript 5.7+ typed arrays.

### Typings checked against the runtime

`index.d.ts` is now checked at compile time against the TypeScript sources (`src/types/doc-conformance.ts`,
`src/types/stream-conformance.ts`): every public member must exist at run time with a compatible type. Aligning it changed the following;
each item can break code that compiled with 4.x.

Values that are `undefined` or `null` in normal use now say so:

| Member                                                                                                          | 5.0                                                  |
| --------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------- |
| `cell.numFmt`, `font`, `alignment`, `protection`, `border`, `fill` (also on rows and columns)                   | `... \| undefined`                                   |
| `cell.hyperlink`, `name`, `dataValidation`, `note`, `formula`, `result`, `formulaType`                          | `... \| undefined`                                   |
| `row.height`                                                                                                    | optional                                             |
| `row.model`                                                                                                     | getter `RowModel \| null`; the setter rejects `null` |
| `worksheet.getSheetValues()`                                                                                    | `(CellValue[] \| undefined)[]` (sparse)              |
| `pageSetup.margins`                                                                                             | `\| null` for a sheet read from a file without them  |
| `worksheet.columns`                                                                                             | getter `Column[] \| null`                            |
| `worksheet.getTable()`, `findRows()` items, `insertRows()`, `getBackgroundImageId()`                            | `\| undefined`                                       |
| `workbook.title`, `creator`, `created` and the other document properties                                        | `\| undefined` (absent in the file)                  |
| `workbook.getImage()`; streaming `getWorksheet()`, `getImage()`, `lastRow`, `findRow()`, `findCell()`           | `\| undefined`                                       |
| `outlineProperties.summaryBelow` / `summaryRight`, `RowBreak.min`, `WorkbookView.firstSheet` / `activeTab`      | optional                                             |
| Conditional-formatting `rule.style` parts (`DifferentialStyle`), gradient `degree` / stop `color`, `top10.rank` | `null` / optional as read from a file                |

Removed (they did not exist, or had no effect): the `Anchor` class (now an interface), `config.setValue`, `cell.release()` and the other
internal merge helpers, `worksheet.commit()` (on document worksheets), `Table.columns` / `rows` / `style`, `DefinedNames.getMatrix()`,
`WorksheetProperties.showGridLines` (set `views[].showGridLines`), `date1904` on formula values, `stream` / `filename` in
`XlsxWriteOptions`, the `Workbook` base of `stream.xlsx.WorkbookWriter` / `WorkbookReader` (they are not workbooks: no `xlsx`, `csv`,
`worksheets`...), the constructor of `stream.xlsx.WorksheetReader` (an interface: it is not exported) and the fields of `ZlibOptions` other
than `level`.

Corrected signatures:

- `getCell(address)` / `getCell(row, col)` and `findCell(address)` / `findCell(row, col)`: a row number with a column letter
  (`getCell(2, "B")`) threw or returned `undefined` and no longer compiles.
- `fillFormula(range, ...)` takes an address string (a `Range` or `Location` threw); its `results` may hold strings and booleans; new
  `shareType`.
- `mergeCells` / `unMergeCells` take the range arguments the runtime accepts (`RangeArguments`).
- `Range.containsEx()` needs `row` and `col`.
- `mergeCells()` / `unMergeCells()` without arguments no longer compile (they merged A1 or threw); `IconSetRuleType.iconSet` no longer
  accepts `"3TrafficLights"` (not a schema value; still written as the default at run time); new `Color.auto`.
- `addRow` / `insertRow` / `addRows` take `RowValues` and a `RowStyleOption` (`"n" | "i" | "i+" | "o" | "o+"`); `row.splice()` takes
  `CellValue`s; `TableProperties.rows` is `CellValue[][]`; conditional-formatting `formulae` are `CfFormula[]`
  (`string | number | boolean`), and `colorScale` / `iconSet` rules require `cfvo` (writing threw without them).
- `CellHyperlinkValue.text` may be a number, boolean, date, error or rich text (as read back).
- `removeConditionalFormatting(filter?)` takes an index or a filter function.
- `Media` describes `workbook.media` (`type: "image"`, optional `name` / `buffer`, `base64`, `filename`); `WorkbookModel` and
  `WorksheetModel` describe the real models (`WorksheetModel.media` holds worksheet images, `revision` is a number, `themes` a record,
  `mergeCells` for the setter).
- `workbook.worksheets` is `readonly` (a getter).
- `xlsx.write(output)` takes any object with `write` / `end` or a Node writable stream and resolves with the `Xlsx` object; `csv.read()`
  takes a Node readable stream, `csv.write()` a writable one; CSV `map` options take and return `CellValue`, and the fast-csv row types hold
  `unknown` instead of `any`.
- Streaming: `stream.xlsx.WorkbookWriter.commit()` resolves with the writer; the `stream` option is a `NodeJS.WritableStream`; reader events
  are typed (`WorkbookReaderEvents`, `WorksheetReaderEvents`): unknown event names and wrong listener arguments no longer compile; `parse()`
  yields typed events (`WorkbookReaderEvent`, `WorksheetReaderEvent[]`) instead of `any`.

New at run time: `ExcelJS.PaperSize` (`ExcelJS.PaperSize.A4`), an object of the common paper sizes; it was a `const enum`, unusable with
`isolatedModules`. `pageSetup.paperSize` accepts any paper size number of the standard (`8` for A3).

`Range` addresses and `count` are `readonly` (they are getters: setting them threw). `xlsx.load(data, { base64: true })` takes a string or a
Node `Buffer` of the text (a plain `Uint8Array` was not decoded as text). `new ModelContainer()` takes an optional model; `container.model`
is a `WorkbookModelInput | undefined` (a model read from a file has no `rowBreaks` or `merges`, and its worksheet `dimensions` is a string).
`worksheet.model` and `workbook.model` accept the narrower `WorksheetModelInput` / `WorkbookModelInput` when set. The `style` of a column of
`WorksheetModel.tables` is a `TableColumnStyle` (its parts may be `null`, as read from a file).

Added (they existed at run time): `ModelContainer`; `cell.html`; `row.style`; `CellFormulaValue.shareType` / `ref`; `comment.anchor`;
`Range.setTLBR()` / `forEachAddress()`; `worksheet.sheetProtection`, `rowBreaks`, `conditionalFormattings`; `PageSetup.copies`,
`usePrinterDefaults`, `useFirstPageNumber`; the `autoFilter` option of `addWorksheet`; `addImage` corners given as `{ col, row }` or native
offsets (`ImageAnchorPosition`); `TableColumnProperties.totalsRowResult`; `Color.tint` / `indexed`; `Font.condense` / `extend` / `shadow`;
`Borders.color`; icon-set `icons` (`CfIcon`); `workbook.language`, `revision`, `contentStatus`, `media`; `xlsx.load()` with a `Blob`, any
`ArrayBufferView`, a byte array or base64 text; `xlsx.read()` with an async iterable; the full `stream.xlsx.WorksheetWriter` and
`WorksheetReader` APIs, the `HyperlinkReader` and the document properties of `WorkbookWriter`.

`index.d.ts` now starts with `/// <reference types="node" />`, so that it works with TypeScript 6 and 7, which no longer load `@types/*`
packages automatically. Install `@types/node` as a dev dependency (also in browser projects), or enable `skipLibCheck`. As a consequence,
Node's globals (`process`, `setTimeout` returning `NodeJS.Timeout`, ...) are visible in projects that compile `index.d.ts`, including
browser projects that set `"types": []`.

The typings use `get`/`set` accessor declarations in interfaces, which require TypeScript 4.3 or later; they are tested with TypeScript 5.9,
6.0 and 7.0.

## Zip library: fflate

`archiver`, `jszip` and `unzipper` were replaced by ExcelJS's own zip code (records, ZIP64, limits) on top of one DEFLATE implementation per
platform: `node:zlib` (native) on Node.js and Bun, and [fflate](https://github.com/101arrowz/fflate) (^0.8.3, no dependencies) in the
browser bundle. The two produce standard archives that each reads from the other.

### `zip` option of the writers

The `zip` option of `xlsx.write` / `writeBuffer` / `writeFile` was passed to JSZip, the one of the streaming `WorkbookWriter` to archiver.
Both writers now take the same options, and the old shapes keep working:

| ExcelJS 4.x                                                                        | ExcelJS 5.0                    |
| ---------------------------------------------------------------------------------- | ------------------------------ |
| `{compression: 'DEFLATE', compressionOptions: {level: 9}}` (JSZip)                 | `{level: 9}`                   |
| `{zlib: {level: 9}}` (archiver)                                                    | `{level: 9}`                   |
| `{compression: 'STORE'}` (JSZip), `{store: true}` (archiver)                       | `{level: 0}`                   |
| `{forceZip64: true}`, `{comment}` (archiver)                                       | unchanged, now in both writers |
| other JSZip / archiver / zlib options (`type`, `platform`, `memLevel`, `strategy`) | ignored                        |
| an invalid `level`                                                                 | throws a `TypeError`           |

### Behaviour changes

- **Decompression limits.** Both readers reject files whose content inflates beyond `maxEntrySize`, `maxUncompressedSize` or
  `maxCompressionRatio` (see the [Security & limits](README.md#security--limits) section of the README for the defaults). A legitimate file
  above the defaults needs the option raised, e.g. `xlsx.load(buffer, {maxEntrySize: 4 * 1024 ** 3})`.
- **ZIP64** records are written when an entry or the archive exceeds 4 GiB, or with 65,535 entries or more. In 4.x the streaming writer
  (archiver) did so but `writeBuffer` (JSZip) did not.
- **`xlsx.load` input**: a Buffer, a `Uint8Array`, an `ArrayBuffer` or a `Blob`. Strings (JSZip's binary strings) and other values are
  rejected with `Can't read the zip file: expected a Buffer, a Uint8Array, an ArrayBuffer or a Blob`; use `{base64: true}` for base64.
- **Error messages** of malformed archives changed: a truncated stream rejects with `Unexpected end of zip data` (unzipper: `FILE_ENDED`),
  other errors start with `Invalid or unsupported zip file:`. Part names are decoded as UTF-8, like JSZip did.
- **Streaming reader input**: besides file names and `stream.Readable`, any object with a `pipe` method (e.g. a `readable-stream` or older
  stream), a web `ReadableStream` and an async iterable of byte chunks are accepted.
- **Image extensions** must be 1 to 5 letters or digits (`png`, `jpeg`, ...). Another value (which would end up in a zip entry name and in
  XML) makes `writeBuffer` / `write` reject and `WorkbookWriter.addImage` throw.
- A failing output stream (or an unreadable image file) now rejects `write` / `writeFile` / `WorkbookWriter.commit()` instead of leaving the
  promise pending or raising an unhandled rejection.

## Images do not create rows or columns

In 4.x, `ws.addImage(id, 'B2:D5')` (or reading `getImages()[i].range.tl.col` / `.row`) created empty rows and columns up to the anchor, so
`ws.addRow()` after `addImage` started below the image and `ws.rowCount` counted those rows. Anchors no longer create anything: add the rows
first, or use `ws.getRow(n)` / `ws.insertRow(n, values)` to place them. Anchors read from a file are clamped to the grid (column XFD, row
1,048,576).
