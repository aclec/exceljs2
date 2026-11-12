# Changelog

Changes up to 4.4.0 are listed in the [Release History](README.md#release-history) section of the README.

Issue and pull request numbers (#NNNN) refer to the upstream repository, [exceljs/exceljs](https://github.com/exceljs/exceljs). Fixes were
re-implemented and tested here; upstream patches were used as bug reports only.

## 5.0.0

Read [UPGRADE-5.0.md](UPGRADE-5.0.md) before upgrading.

### Breaking changes

- Requires Node.js >= 22.12. Bun >= 1.4 is supported.
- The ES5 build (`exceljs/dist/es5`) was removed; the browser bundles use ES2020 syntax and ES2022 built-ins (`Object.hasOwn`) and need no
  polyfills on Chrome / Edge 93, Firefox 92, Safari 15.4 or later.
- Hardened parsing of untrusted files changes a few behaviours: grid bounds are enforced, merged cells are capped per worksheet (1,048,576)
  and per workbook (`maxMergedCells`, 1,048,576 by default, the same when building, streaming and loading a workbook), large data-validation
  and defined-name ranges are kept as ranges, and out-of-range sheet ids are renumbered. See
  [UPGRADE-5.0.md](UPGRADE-5.0.md#limits-when-reading-files).
- The library is written in TypeScript: `src/**/*.ts` is compiled to CommonJS in `lib/` (with `.d.ts` and source maps). The package entry
  point keeps its shape; deep imports of internal `lib/` modules now expose their main value as the default export. See
  [UPGRADE-5.0.md](UPGRADE-5.0.md#package-exports).
- TypeScript typings corrected to match the runtime (for example `cell.col` / `cell.row` are numbers) and now checked against it at compile
  time: nullable values say so, members that did not exist are removed, signatures follow what the code accepts, the streaming classes are
  no longer declared as workbooks, and missing members are declared. See
  [UPGRADE-5.0.md](UPGRADE-5.0.md#typings-checked-against-the-runtime).
- Behaviour changes when reading and writing: internal hyperlinks are read as hyperlink cells and written as a `location`; style setters
  copy shared styles instead of modifying them; CSV whitespace-only cells, large integers and `parserOptions.headers`; table header
  defaults; streaming `WorksheetReader.id` and `sharedStrings: 'emit'` event shape. See [UPGRADE-5.0.md](UPGRADE-5.0.md).

### Security

All fixes have regression tests in `spec/integration/security`.

- Merged cells: a merge covering the whole grid (`<mergeCell ref="A1:XFD1048576"/>`) exhausted the heap and aborted the process. The merged
  area is now capped at 1,048,576 cells per worksheet (#3114).
- Cell references: rows beyond 1,048,576 are rejected when decoding addresses, so that malformed references cannot become huge ranges
  (#3103, #3114).
- Merges: the overlap check was quadratic in the number of merges (CPU denial of service with many small merges); it is now proportional to
  the merged area (#2920).
- Data validations: a huge `sqref` was expanded to one object per cell (memory and CPU exhaustion). Large ranges are now kept as range keys;
  tokens that are not cell references are ignored (#2977).
- Defined names: a huge range was expanded to one entry per cell (memory exhaustion); ranges over 100,000 cells are kept as rectangles.
- Defined names: a sheet named `__proto__` in a defined name polluted `Object.prototype` (null-prototype maps are now used).
- Columns: `<col min max>` was not bounded, so `max="2000000000"` allocated two billion columns; it is clamped to the last Excel column.
- Images: parsing a drawing relationship `Target` used a regular expression with cubic complexity (ReDoS); it is now linear.
- Formulas over Excel's 8,192-character limit are rejected before quadratic regular-expression scans (shared-formula translation,
  `cell.dependencies`).
- `deepMerge` (used for cell notes) followed `__proto__`, `constructor` and `prototype` keys and polluted `Object.prototype` (#3083, #3101).
- CSV: cells named like `Object.prototype` members (`__proto__`, `constructor`, ...) were read as `Object.prototype` or functions instead of
  text (#3089).
- XML output: U+FFFE, U+FFFF and unpaired surrogates produced files that Excel and ExcelJS could not read back; they are stripped (#3113).
- Pivot tables: field names and values were written into XML without escaping (#2996).
- Streaming reader: temporary files (and their file descriptors) of deferred worksheets leaked when the iteration stopped early; the
  temporary directory is now always removed.
- Conditional formatting: an `<x14:cfRule id="__proto__">` polluted `Object.prototype` with attacker-chosen `type`, `operator`, `formulae`
  and `style` for the life of the process. Other maps keyed by ids read from the file (relationships, shared formulas, media, streamed
  hyperlinks) are prototype-less too, so ids such as `constructor` are treated as unknown instead of throwing.
- CSV: with `parserOptions.headers`, a header row of more than 16,384 names threw an uncaught exception that crashed the process.
- Number formats: `isDateFmt`, run for every numeric cell, was quadratic on `[[[[...` (ReDoS); it is linear and cached, and format codes
  over Excel's 255-character limit are ignored.
- A huge `sheetId` or `<row r>` made every traversal of the worksheets or rows take minutes. Sheet ids are renumbered unless they are unique
  integers from 1 to 65,535, and row numbers beyond 1,048,576 are rejected (document and streaming readers).
- The merged-cell, data-validation and defined-name caps applied per worksheet, element or range and were multiplied by the number of sheets
  or names; they are now shared by the whole workbook being loaded.
- Shared-formula translation was quadratic per clone; `DataValidations.find` and `DefinedNames.model` / `cell.names` were quadratic in the
  number of validations or names. All are linear.
- Shared parts: a comments part referenced by the relationships of W worksheets was materialised W times (a 61 KiB file gave 400,000 notes
  and 391 MB of heap). A comments, notes VML, drawing or table part now belongs to the first worksheet that refers to it; the references of
  other worksheets are ignored.
- Data validations: every `cell.dataValidation` lookup scanned all the range keys (an unbounded number once 100,000 cells are expanded: 0.75
  ms per lookup with 250,000 keys). Range keys are indexed by position and size; a lookup no longer depends on their number.
- Second review (super-linear time or memory on small hostile files, regression tests in `spec/integration/security/complexity.spec.js`):
    - Conditional formatting: the number check of `<cfvo val>` / x14 `<xm:f>` thresholds was quadratic (ReDoS, 6 KB file: 15 s); linear.
    - Loading matched each worksheet part to its `<sheet>` and re-joined the `Print_Area` names of a sheet quadratically; both linear.
    - Columns: one `<col min="1" max="16384">` (or one cell in column XFD) created 16,384 Column objects per worksheet (57 KB file: 534 MB).
      Columns of loaded ranges are created when they are used; `ws.columns` still lists them all.
    - Defined names on far single cells (`$XFD$1048576`) cost a million steps per name on every write or splice; now proportional to the
      cells.
    - Image anchors beyond the grid are clamped, and reading `range.tl.col` / `.row` no longer creates the row or columns up to the anchor.
    - Re-saving a rich text referenced by many cells serialised it once per cell; once per text.
    - CSV: `csv.read` ran four date parses on every text field (~100 times slower); only fields shaped like a default date are parsed.
      `csv.write` / `writeFile` crashed the process on a destination error (ENOENT, ENOSPC); the promise rejects.
    - Pivot tables: header names such as `constructor` were looked up in plain objects (corrupt pivot cache).
    - An invalid `docProps/core.xml` date loaded as an Invalid Date that made every later write reject; it is read as undefined.
- Third review:
    - Rows and cells are kept in arrays indexed by row / column number, and every traversal visited each index: a worksheet whose only row
      is 1,048,576 cost 14 ms per traversal (400 such sheets in a 114 KiB file: 10.8 s to write), a row whose only cell is in column XFD 0.2
      ms (3,000 such rows: 6.2 s), and moving such a row created its 16,384 cells. Traversals (`eachRow`, `eachCell`, `values`,
      `dimensions`, `actualRowCount`, splices, the written model, defined names) now visit the existing entries only
      (`lib/utils/sparse-array.js`), as fast as before on dense rows.
    - Streaming writer: committing a row after a gap was quadratic in the gap (a row 10,000 rows below the last one: 2 s; row 1,048,576
      never finished), and committing rows one by one before a far row was quadratic in the rows; both are proportional to the rows. A row
      whose writing throws is no longer written twice.
    - Streaming reader: one `<col min="1" max="16384">` created 16,384 Column objects per streamed worksheet, and a cell in column XFD did
      the same; columns are created when they are used, as in the document model.
    - Merged cells: the default workbook budget (`maxMergedCells`) is one full worksheet (1,048,576 merged cells, about 600 MB) again: with
      four worksheets, a 7.6 KB file reached 2.5 GB of heap. Checking it summed the worksheets over their sheet ids on every merge (8,000
      merges on a worksheet whose `sheetId` is 65535: 7.6 s); the workbook keeps a running total.
    - XML: a truncated part (an element left open) is rejected by both readers instead of being read as if complete (the streaming reader
      returned the first rows of a cut worksheet without an error).
- Hostile names (found while porting to TypeScript): XML elements named after `Object.prototype` members (`<constructor/>`, `<toString/>`,
  `__proto__`…) are skipped as unknown instead of crashing the load; zip parts named `constructor` / `toString` (themes, drawings, media,
  tables, comments, worksheets) are kept apart from the prototype; a `<comment>` with `note=`, `type=` or `texts=` attributes no longer
  crashes the load; the x14 data bar / icon set extensions no longer write a model property named after the file; `_.keyBy` returns a
  null-prototype object.

### Dependencies

- archiver, jszip and unzipper removed: `node:zlib` on Node.js and Bun, fflate ^0.8.3 (no dependencies) in the browser bundle, see
  [Zip library](#zip-library-fflate) below
- dayjs ^1.8.34 → ^1.11.23
- fast-csv ^4.3.1 → ^5.0.8
- saxes ^5.0.1 → ^6.0.0
- Removed `readable-stream` (replaced by `node:stream`), `tmp` (replaced by `fs.promises.mkdtemp`) and `uuid` (replaced by
  `crypto.randomUUID` / `crypto.getRandomValues`).
- Development: typescript ^6.0.3 and typescript-eslint ^8.72.0 (type-aware lint of `src/`, `any` banned).
- Development: bun is the package manager (`bun.lock`); mocha 12, chai 6 and tsx replace mocha 7, chai 4 and ts-node; `bun audit` reports no
  known vulnerability. No dependency runs install scripts.

### Bug fixes

Upstream numbers refer to issues / pull requests of exceljs/exceljs that reported the bug. Every fix was re-implemented in this fork (no
upstream PR code was merged) and has a regression test.

- Files reported as corrupt by Excel:
    - worksheet elements written in `CT_Worksheet` order, by both writers (`legacyDrawing` #2079; streaming writer #2685, #2803); `sheetPr`
      children order (#2807); data bar extension colour order (#2651);
    - multiple print areas written as one `_xlnm.Print_Area` name, whole-column/row print areas, apostrophes in sheet names (#1516, #3071);
    - `theme1.xml` always written (#2009); comment VML shape ids unique across the workbook, both writers, also for worksheets with more
      than 1,023 notes, which claim several id blocks (#3070; the streaming writer writes the notes VML when the worksheet is committed);
      unique picture `creationId` (#3092); the right image when pictures repeat (#3094);
    - empty rich-text runs (#2737); distinct rich-text shared strings in the streaming writer (#2588); dxf number format written as
      `[object Object]` (#2698);
    - internal hyperlinks (`#Sheet!A1`, `#DefinedName`) written as a `location` without an external relationship, by both writers (#2002);
      URLs and paths ending like a sheet reference (`https://example.com/#!page1`) stay external links, while legacy links to sheets named
      with `'` or `#` (`Bob's!A1`, `Q#1!B2`) stay internal; `inset="NaNmm,..."` written for notes whose VML textbox had no inset;
    - print areas and print titles written as absolute references (`$A$1:$G$20`; 4.x wrote `$A1:$G20`, and `$$A$1` for `$A$1`), by both
      writers;
    - a literal `_xHHHH_` in a string (`x_x0041_y`) is written escaped (`_x005F_x0041_`), as ECMA-376 requires, so that it reads back as
      written instead of as the character U+HHHH.
- Data lost or wrong on read:
    - rows/cells without an `r` attribute (#2962, both readers); `t="d"` ISO 8601 date cells (#2702, both readers);
    - falsy formula results (#3048); only numeric formula results converted to dates (#3111);
    - leading plain text of mixed rich-text strings (#3091); explicit `val="0"`/`"false"` boolean attributes (#2095, #2851, #3057);
    - hyperlink `location` (#2002); table columns after a column with child elements (#1936); table `filterColumn` matched by `colId`
      (#2978); filter buttons kept visible on round trip; `headerRowCount` default and no `autoFilter` without a header row (#1345); x14
      `expression`/`cellIs` conditional formats (#1767); `stopIfTrue` (#2736);
    - column width 9 when the sheet default differs, also when set after `worksheet.columns` (#3088); sheet names longer than 31 characters
      (#3087);
    - both readers parse `t="d"` values with the same rules (`T24:00`, offsets, years 0-99); hyperlinks on number or boolean cells keep
      their value type on write.
- Files that failed to load: unusual `docProps` parts (#2846); relationships to missing or absolute-path parts and VML notes (#3051, #3105);
  relationships without a `Target` (drawing, background image, worksheet); unmodelled table filter elements such as `dynamicFilter` (#2973);
  bare data bars and rules without formulae on write (#3019, #2883).
- Streaming reader: rejects instead of hanging or crashing on truncated/corrupt input and always removes its temp files (#2148); waits for
  `workbook.xml`/styles before parsing sheets stored earlier in the zip (#2790, #2915, #3095); UTF-8 characters split across chunks (#1665,
  #2894, #3102); column definitions, hidden rows, sheet format properties (#2800); absolute relationship targets (#2671); options passed to
  `read()`/`parse()` merged with the defaults; `sharedStrings: 'emit'` emits again; shared strings read like the document reader does (rich
  text fonts lost their colour, name and charset, phonetic runs replaced the text); a failed close of the temp file of a skipped worksheet
  is no longer an unhandled rejection.
- Document model: setting a style property on a cell no longer restyles every cell sharing a loaded style, while merged cells still share
  the master style (#2055, #2781); `spliceRows` stale rows (#1600); column keys / table names / CSV headers named like `Object.prototype`
  members (#2080), with `worksheet.tables` kept a plain object; assigning the model of a merged cell keeps its master, and releasing a merge
  reference twice no longer throws; a value inserted inside a merge by column key is checked with the keys that `row.values` uses (a value
  could be lost); setting `worksheet.model` replaces the merges of the worksheet.
- `writeBuffer()` of a small workbook no longer waits for a timer, so it settles under fake timers (Vitest, Jest, sinon); see the README
  (Known Issues > Fake timers) for what still does.
- Browser bundle: loading it in Node.js 25+ no longer prints a `localStorage` ExperimentalWarning (from the `util-deprecate` and `util`
  polyfills, replaced by small shims); the bundle is about 6% smaller.
- CSV: whitespace-only cells stay strings (#2752); integers beyond `Number.MAX_SAFE_INTEGER` stay strings (#1743); `parserOptions.headers`
  support (#2080), keeping the values of columns with empty or repeated header names; `read()` rejects instead of crashing or hanging when a
  listener throws (e.g. a header row wider than 16,384 columns) or the input stream fails.

- Conditional formatting:
    - `beginsWith`, `endsWith` and `notContainsText` rules were read as unknown types and dropped on write; they are read as `containsText`
      rules with that `operator`, their `text` is read and written, and they get Excel's formula when none is given (656fd9bd);
    - rules that Excel stores only in the x14 extension (cross-sheet `expression` / `cellIs` / text / time period rules, x14-only rules)
      were written back as main rules, which Excel may repair; they are written in the extension again (d32353ee, f3eb6d94);
    - data bars use the schema defaults of `CT_DataBar` (`direction: 'context'`, `negativeBarColorSameAsPositive: false`, `minLength` 10,
      `maxLength` 90 when absent), `showValue` is read and written, and data bars read from a file keep their x14 settings (negative
      colours, axis, border, lengths, `autoMin` / `autoMax` thresholds) on write (1ffd02ba, 503c07cf);
    - rules without `priority` were numbered from the highest given priority instead of the next one, so two rules shared it (9c3245b4);
    - formula and cell-reference threshold values (`<cfvo val="$A$1">`, `xm:f`) were read with `parseFloat` (`NaN` or a partial number);
      they are kept as written (8da19e03);
    - a differential format with both a border and a protection was written in the wrong element order (c897f516);
    - the streaming writer without `useStyles` threw on any rule with a style (ddeacb2b).
- Number formats read from files keep their `\` escapes (`yyyy\-mm\-dd`, `#,##0.00\ "€"`): 4.x removed every backslash, so formats were read
  and written back changed (`\#0` even turned a literal `#` into a digit placeholder) and `0.0\ \m` was taken for a date format. Escaped,
  padding and fill characters are ignored when detecting date formats, quoted texts included (d30edbd0, 5347e2f3).
- Streaming writer: `<sst count>` is the number of references to the shared strings, not the number of strings (ddb85b85).
- Streaming reader: worksheets and shared strings were silently dropped when the zip stream ended before all entries had been read (#2790,
  #2915, #3095).
- TypeScript typings:
    - `index.d.ts` references `@types/node`, so it compiles with TypeScript 6 and 7.
    - `xlsx.load()` and `Image.buffer` accept Node Buffers with TypeScript 5.7+.
    - `Address.col` / `Address.row` are numbers (#2720); `Row.dimensions` can be `null` (#2596); `Column.values` is an array and the
      `Row.values` getter returns an array (#2587).
    - `ImageEditAs` (#2983), all `cellIs` operators and optional rule `priority` (#2736), data bar colours (#2655), `DataValidationType`
      (#2697), optional `protect()` arguments (#2664), `null` table theme (#2061).
    - Streaming readers: optional `WorkbookReader` options (#1922), `WorksheetReader` properties and getters (#1901, #3097), EventEmitter
      methods.
    - The `csv.createInputStream()` typing was removed (it has thrown since 4.0). Comment typos (#3003, #3081).
- Data validations: a range key written directly into `dataValidations.model` after a lookup was ignored by later lookups.
- Data validations on whole columns or rows (`dataValidations.add("A:A", ...)`, `"1:3"`, `sqref="D:E"` in a file) crashed the writers, were
  ignored by `find()` and were dropped on read; they are kept as the rectangle they cover (`A1:A1048576`), never expanded per cell.
- Merged cells after `spliceRows` / `insertRow(s)` / `duplicateRow` / `spliceColumns`: the merge index used by the writers was not moved
  with the cells (merges written at their old place), `spliceRows` with deletions and insertions shifted merges by the number of inserted
  rows instead of the net offset, and slaves could point at a replaced master. Merges now move as rectangles: they shrink when some of their
  rows / columns are removed, widen when rows / columns are inserted inside them (as in Excel), disappear when reduced to one cell, and
  `duplicateRow` copies the single-row merges of the duplicated row. Inserting values into the cells of a merge (which would be lost, only
  the master holds a value) throws before anything changes.
- Notes (comments) were dropped from the rows that `spliceRows` / `insertRow(s)` moved; they move with their cells, as `spliceColumns`
  already did.
- Notes on cells without a value were lost on read: Excel (and ExcelJS) write no `<c>` element for such cells. The cells, and their rows
  when missing, are now created.
- Streaming writer: notes on the cells of a row without values were dropped (the row was skipped).
- Streaming reader: error and boolean formula results (`t="e"`, `t="b"`) were read as `NaN` and `1` / `0`; they are read as
  `{error: '#N/A'}` and `true` / `false`, as the document reader does.
- Hyperlink tooltips were lost on read (document reader). The streaming reader's `hyperlink` events now carry `location`, `tooltip` and
  `display` too (an internal link was an event without target).
- Streaming writer: memory grew with the number of rows (about 0.5 KB per row: 170 MB for 300,000 rows) because rows were committed faster
  than they were compressed and written, with no way to wait. `row.commit()` and `worksheet.commit()` now return a promise that resolves
  once the output has drained; awaiting it keeps memory bounded (about 7 MB of live memory whatever the number of rows). Not awaiting it
  keeps the previous behaviour.
- Streaming reader: the temp file of a worksheet spooled before the shared strings stayed open for the life of the process when the consumer
  skipped that worksheet (or with `worksheets: 'ignore'`); it is closed once its reader is garbage collected. Removing the temp dir no
  longer waits forever for a spooled file that was closed before it opened.
- `xlsx.load` / `read` / `readFile` and `xlsx.write` / `writeBuffer` / `writeFile` never gave the event loop a turn (2 s and 4 s for a
  200,000-row sheet, during which a server answered nothing). They now pause about every 25 ms while parsing, reconciling, preparing and
  rendering worksheets, at no measurable cost. Building the workbook from the parsed model and compressing one part are still done at once
  (about 0.5 s and 1.5 s for that sheet).

- An empty shared string (`<si/>`) reads as `""` in both readers; it used to be re-saved as `[object Object]`.
- A link to another workbook such as `"Book2.xlsx#Sheet1!A1"` is written as an external link again (not as a place in this workbook).

Found while porting to TypeScript (each with a regression test that fails on the JavaScript code):

- Streaming writer: distinct shared formulas were all written with `si="0"` (the counter was copied per row); `findCell` never found a cell.
- Streaming reader: with `hyperlinks: "cache"`, a `<hyperlinks>` before `<sheetData>` erased the cell values; a stray `</row>` emitted a
  `null` row; `HyperlinkReader.each` threw; a row with `s="0"` threw with `styles: "ignore"` (the default).
- A styles part without dxfs or without cellXfs, or a workbook without styles part whose rows, columns or conditional formattings are
  styled, made the load throw.
- `cell.text` of a merged cell threw when its master was empty and gave `[object Object]` for a rich text or formula master; `cell.text` and
  `toCsvString()` of a formula error gave `[object Object]`; `toCsvString()` of rich text threw.
- `table.displayName` always returned the name (typo) and could not be set; `table.theme` read and wrote the wrong property.
- `worksheet.addPivotTable()` threw on a loaded workbook (`pivotTables` was undefined).
- `pageSetup.cellComments: "asDisplayed"` (the documented value) was not written.
- The internal `StreamBuf` is a real `Duplex`: `on('data')`, `pause()`, `destroy()` behave as in Node.
- Malformed files: a part without `<workbook>` / `<worksheet>`, a conditional formatting rule without range, an unknown shared string index,
  an invalid data-validation range or a colorScale / iconSet without `cfvo` throw an explicit `Error` instead of a `TypeError`.

Found while aligning the typings with the runtime (each with a regression test that fails before the fix):

- Files Excel may reject: a sheet name in a merge reference (`mergeCells("Sheet1!A1:B2")`) was written into `<mergeCell>`; the streaming
  writer wrote merge references without XML escaping; a gradient stop or a colour scale entry with an automatic colour lost its `<color>`;
  `addPageBreak()` wrote `max="16838"` (the last column is 16383); `NoIcons` was written in the main `<iconSet>`; icon set names that are
  not `ST_IconSetType` values were written as is.
- Data lost or wrong on round trip: custom icon sets lost their icons, and `iconId` was overwritten with the icon position; font sizes with
  decimals were read as integers (10.5 → 10); automatic colours were dropped; data-validation formulas that are references or formulas
  (`$B$1`, `TODAY()`) were read as `NaN` / `Invalid Date` and written back as `NaN`; row breaks were not read; automatic page numbering
  became "start at 1"; `date1904="true"` (a valid xsd:boolean) was read as false; `worksheet.model = other.model` dropped the merges.
- Cell values: hyperlink, shared-string and formula values with an empty or zero marker (`{ text: "", hyperlink }`, `{ sharedString: 0 }`,
  `{ formula: "", result }`) were stored as JSON text; `cell.text` of a hyperlink whose text is a number or rich text was not a string; a
  formula with a falsy result had an empty text.
- Worksheets: `getCell("A0")`, `getCell(0, 1)` created a cell outside the grid, `getCell("foo")` threw an unrelated error,
  `mergeCells(10, 10)` and `mergeCells()` registered broken merges, lowercase references merged the wrong cells, `fillFormula` ignored a
  single cell or an invalid range and threw a `TypeError` with fewer result rows than the range: they throw a clear `Error` or work.
- Defined names: removing a range (`remove("S!$A$1:$B$2", name)`, `removeEx(range)`) did nothing; removing one cell from a huge range kept
  as a rectangle removed the whole range.
- Tables: a row longer than the columns threw a `TypeError` and left a half-written table; a second `commit()` after a move cleared the
  wrong cells; tables read from a file threw `TypeError`s from `height`, their mutators and style setters.
- Streaming reader: `hyperlinks: "cache"` never read the hyperlink targets.
- xsd:boolean attributes written `"true"` / `"false"` (valid, written by other tools) were read as false in the workbook properties, page
  setup, print options, sheet views, sheet protection, cell protection, outline properties, header / footer, table styles and conditional
  formatting (`reverse`, `showValue`, `custom`...).
- A table without `<tableStyleInfo>` loaded but could not be written back (`TypeError`); it round-trips without style. Tables read from a
  file are edited with the table API like created ones. `totalsRowCount` / `headerRowCount` are read as counts.
- `tint` is read and written with any colour (rgb, indexed, auto), not only theme colours.
- Conditional formatting rules of an unsupported type (`duplicateValues`, `uniqueValues`...) were kept and written back as an empty
  `<conditionalFormatting>` (a corrupt file); they are not read. A hyperlink whose text is an error value was written as the shared string
  `[object Object]`; it is written as an error cell. A formula cell without formula text (an orphan shared-formula clone) was written with
  an empty `<f/>`; it is written as its cached value.
- Data validation dates follow the 1904 date system of the workbook, on read and write; `type="time"` rules are kept.
- Streaming reader: shared-formula clones get the formula of their master, translated as `workbook.xlsx.load` does (they had an empty
  formula).
- `spliceColumns` dropped the values inserted for rows past the last one, and inserted `0`, `false` and `""` as empty cells.
- `xlsx.load(data, { base64: true })` decodes base64 text given as bytes (a `Uint8Array`, `ArrayBuffer`, `Blob`...), not only as a string or
  a Buffer. The streaming reader gives the numeric result of a formula in a date-formatted cell as a Date, and the `t="s"` cells of a
  workbook without shared strings part an empty value, as `xlsx.load` does. A hyperlink on a date is written as a date (it was written as
  its locale text). Defined names no longer mutate the cell addresses that the address cache shares.
- `<formula1>` outside `<dataValidation>`, a `<sheet>` or table without name, a table without `<tableColumns>` and other malformed parts
  threw `TypeError`s: they are skipped or throw a clear `Error`. Invalid enumerated attribute values (page orientation, pane, view style,
  validation type or operator, cfvo type, data bar axis, border style, pattern type, underline, image `editAs`, sheet state...) read as the
  schema default instead of being kept as is.
- Internal: no type assertion is left to silence the compiler in `src/` (`no-non-null-assertion` and `no-unsafe-type-assertion` are lint
  errors); the 410 assertions were replaced by checks, which turned several `TypeError`s on invalid input into clear errors.

### Performance

The TypeScript port was measured against the JavaScript code on every hot path (`bun run benchmark:suite`, median of 7 runs, Node 26):
writing a workbook (`writeBuffer`) is 30 % faster, the streaming writer 45 % faster, the streaming reader 11 % faster (cells are written
with one string per start tag, the style registry skips the cells that only have a number format, the reader scans canonical cell references
without decoding them); reading and CSV are unchanged. The minified browser bundle is 448 KiB instead of 471 KiB.

### Documentation

- README: installation with any package manager, requirements, a "Security & limits" section, streaming reader defaults (#1531), tab colour
  example (#2577), formulas without a leading `=` (#3011).
- New [UPGRADE-5.0.md](UPGRADE-5.0.md) and [SECURITY.md](SECURITY.md).
- README examples are type-checked against `index.d.ts` in CI (`bun run check:readme`); the examples that failed at run time were fixed
  (`ExcelJS.` instead of the undefined `Excel` / `Enums`, note protection values, a table row with too many values, CSV `quote: null`).
- Every exported symbol and public member of `src/` has TSDoc (`bun run check:tsdoc`).

### Zip library (fflate)

- `archiver` (streaming writer), `unzipper` (streaming reader) and `jszip` (document reader / writer, browser bundle) were replaced by
  ExcelJS's own zip container code (`lib/utils/zip-stream.js`, `lib/utils/zip-reader.js`) on top of `node:zlib` on Node.js and Bun
  (`lib/utils/deflate.js`; the streaming writer compresses in the thread pool, like archiver did) and
  [fflate](https://github.com/101arrowz/fflate) in the browser bundle. Runtime dependencies go from 3 zip libraries (about 80 packages) to
  one package without dependencies; the minified browser bundle shrinks from 529.6 KiB to 482.9 KiB (gzip 150.4 KiB → 139.5 KiB).
- Security: both readers enforce decompression limits while inflating (`maxEntrySize`, `maxUncompressedSize`, `maxCompressionRatio`), so a
  zip bomb is rejected before it is materialised in memory or in the streaming reader's temporary files.
- ZIP64: both writers add ZIP64 records past 4 GiB or 65,535 entries (`writeBuffer` could not before), `forceZip64` works in both.
- The `zip` writer option takes `{level, forceZip64, comment}`; the archiver and JSZip shapes are still accepted. See
  [UPGRADE-5.0.md](UPGRADE-5.0.md#zip-library-fflate).
- Fixed: a string part with a character outside the BMP (emoji) cut by JSZip's 16 KiB encoding slices was written as two U+FFFD.
- Fixed: the streaming reader stopped an entry written with a data descriptor at the first bytes of its data that looked like a data
  descriptor signature (unzipper); descriptors are now checked against the size of the data read.
- Fixed: write errors (failing output stream, unreadable image file) reject `write` / `writeFile` / `WorkbookWriter.commit()` instead of
  leaving them pending or raising an unhandled rejection.
- Security: an image `extension` was used unchecked in the zip entry name (`xl/media/image1.<extension>`, zip slip for whoever extracts the
  file) and, by the streaming writer, in the relationship XML (relationship injection). Both writers now require 1 to 5 letters or digits,
  and the streaming writer XML-encodes relationship targets.
- Removed the internal modules that no longer had a user (`lib/utils/auto-drain`, `browser-buffer-encode`, `iterate-stream`,
  `stream-base64`, `string-builder`, `stuttered-pipe`, `typed-stack`, and `lib/csv/stream-converter`); they were only reachable through deep
  imports (`exceljs/lib/...`).
- `package.json` has an `exports` map: deep imports are limited to `lib/`, `dist/`, `excel(.js)` and `package.json`, and the `browser`
  condition resolves the package to the browser bundle (see [UPGRADE-5.0.md](UPGRADE-5.0.md#package-exports)).
- The streaming `WorkbookReader` accepts any stream with a `pipe` method (e.g. `readable-stream`), web `ReadableStream`s and async
  iterables, not only `stream.Readable` instances.
