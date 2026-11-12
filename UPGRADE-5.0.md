# UPGRADE FROM 4.x to 5.0

Version 5.0 modernises the runtime requirements and hardens the parsing of untrusted files. Most of the API is unchanged; the sections below
list everything that can break an existing application. See [CHANGELOG.md](CHANGELOG.md) for the full list of changes.

## Runtime requirements

|          | ExcelJS 4.x                               | ExcelJS 5.0                       |
| -------- | ----------------------------------------- | --------------------------------- |
| Node.js  | >= 8.3 (ES5 build for older versions)     | **>= 22.12**                      |
| Bun      | not supported                             | **>= 1.4**                        |
| Browsers | ES5 bundles with core-js polyfills, IE 11 | **ES2020** (all current browsers) |

Applications that must run on an older Node.js or on Internet Explorer should stay on 4.x.

## ES5 build removed

`exceljs/dist/es5` no longer exists, and the browser bundles no longer include or require the `core-js` / `regenerator-runtime` polyfills.

| ExcelJS 4.x                                                                       | ExcelJS 5.0                     |
| --------------------------------------------------------------------------------- | ------------------------------- |
| `require('exceljs/dist/es5')`                                                     | `require('exceljs')`            |
| `require('core-js/...')`, `require('regenerator-runtime/runtime')` before ExcelJS | not needed                      |
| `<script src="polyfill.js">` + `<script src="exceljs.js">`                        | `<script src="exceljs.min.js">` |

## Dependencies

Runtime dependencies were upgraded and three were removed. They are internal to ExcelJS, but you may notice them in your lockfile or bundle:

| Dependency      | 4.x      | 5.0                                                                            |
| --------------- | -------- | ------------------------------------------------------------------------------ |
| archiver        | ^5.0.0   | ^7.0.1                                                                         |
| dayjs           | ^1.8.34  | ^1.11.23                                                                       |
| fast-csv        | ^4.3.1   | ^5.0.8                                                                         |
| jszip           | ^3.10.1  | ^3.10.2                                                                        |
| saxes           | ^5.0.1   | ^6.0.0                                                                         |
| unzipper        | ^0.10.11 | ^0.12.5                                                                        |
| readable-stream | ^3.6.0   | removed: `node:stream` is used (`stream-browserify` inside the browser bundle) |
| tmp             | ^0.2.0   | removed: `fs.promises.mkdtemp`                                                 |
| uuid            | ^8.3.0   | removed: `crypto.randomUUID` / `crypto.getRandomValues`                        |

Side effects:

- ExcelJS no longer calls `tmp.setGracefulCleanup()` when it is loaded, so it no longer changes the behaviour of the `tmp` package for the
  rest of your application.
- The streaming `WorkbookReader` spools worksheets that come before the shared strings to a temporary directory (`os.tmpdir()/exceljs-*`)
  that is removed when the iteration ends, including when you `break` out of it early or an error is thrown. 4.x leaked the files and their
  file descriptors in that case.
- Streams created by ExcelJS are `node:stream` streams instead of `readable-stream` ones.

## Limits when reading files

To make parsing of untrusted files safe, a few inputs that used to be accepted (and could exhaust memory or CPU) are now rejected or stored
differently. See also the [Security & limits](README.md#security--limits) section of the README and [SECURITY.md](SECURITY.md).

### Rows and columns beyond the Excel grid throw

Cell references must be within rows 1 to 1,048,576 and columns A to XFD (16,384), the size of an Excel worksheet. A reference with a larger
row (for example `A1048577`) now throws `Out of bounds. Excel supports rows from 1 to 1048576`; 4.x only checked the column. `<col min max>`
definitions read from a file are clamped to column 16,384.

### Merged cells are capped per worksheet

A worksheet can have at most **1,048,576 merged cells** in total (the size of one full column). `worksheet.mergeCells()` throws
`Cannot merge <range>: worksheet would exceed 1048576 merged cells` beyond that, and `xlsx.load` / `xlsx.read` reject a file that contains
such merges.

The check for overlapping merges now takes time proportional to the area being merged instead of the number of existing merges, so files
with tens of thousands of merges load quickly. The `Cannot merge already merged cells` error is unchanged.

### Large data-validation ranges stay ranges

In 4.x, every `sqref` range read from a file was expanded to one entry per cell in `worksheet.dataValidations.model`, so
`sqref="A1:XFD1048576"` could exhaust memory. In 5.0, ranges are expanded until 100,000 cells have been expanded for the worksheet; ranges
that do not fit are stored under a single **range key**:

```js
// 4.x: {B1: {...}, B2: {...}, ..., B1048576: {...}}
// 5.0: {'B1:B1048576': {...}}
worksheet.dataValidations.model;
```

`cell.dataValidation` resolves both kinds of keys and both are written back to the file, so code that goes through cells is not affected.
Code that reads `dataValidations.model` directly must handle keys of the form `'A1:B2'`.

Also:

- `sqref` tokens that are not cell references (`A1` or `A1:B2`, with or without `$`) are ignored.
- Validations removed with `dataValidations.remove(address)` are no longer written as empty entries.

### Large defined-name ranges stay ranges

Defined names are stored one entry per cell so that they follow row and column splices. A defined-name range covering more than **100,000
cells** (for example `Sheet1!$A$1:$XFD$1048576`) is now kept as a single rectangle instead. Such ranges:

- are returned by `workbook.definedNames.model`, `getRanges(name)` and `cell.names`;
- move or shrink with `spliceRows` / `spliceColumns`;
- are **not** enumerated cell by cell by `workbook.definedNames.forEach()`;
- are removed as a whole when the name is removed from any of their cells (`cell.removeName(name)`, `definedNames.remove(...)`).

### Formulas longer than 8,192 characters

Excel limits formulas to 8,192 characters. ExcelJS now throws `Formula exceeds the Excel limit of 8192 characters` when it has to analyse a
longer formula: translating a shared formula to another cell (`cell.formula` on a shared-formula clone) and `cell.dependencies`.

## Other behaviour changes

- **Invalid XML characters are removed on write.** U+FFFE, U+FFFF and unpaired surrogates (which are not allowed in XML 1.0) are stripped
  from text when writing xlsx, like control characters already were. Valid surrogate pairs (emoji, etc.) are kept.
- **CSV special values.** When reading CSV, cells whose text is `__proto__`, `constructor`, `toString`, ... are read as plain strings. 4.x
  returned `Object.prototype` or built-in functions.
- **Object keys `__proto__`, `constructor` and `prototype`** in objects given to ExcelJS (for example `cell.note = JSON.parse(...)`) are
  ignored when they are merged into internal models.
- **Pivot tables** (experimental): field names and values are XML-escaped.

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

`editAs` is now checked: an object literal stored in a variable before the call widens `editAs` to `string` and no longer compiles. Annotate
it instead:

```ts
const position = { tl: { col: 0, row: 0 }, ext: { width: 100, height: 50 }, editAs: "oneCell" as const };
// or: const editAs: ExcelJS.ImageEditAs = 'oneCell';
worksheet.addImage(imageId, position);
```

New or widened types: `DataValidationType`, `ImageEditAs`, all `cellIs` operators (`notEqual`, `greaterThanOrEqual`, `lessThanOrEqual`,
`notBetween`), optional conditional-formatting `priority`, data bar colours, optional `worksheet.protect()` arguments, optional
`WorkbookReader` options, `WorksheetReader` `id`/`name`/`state`, and the EventEmitter methods of the streaming readers. `xlsx.load()`
accepts Node Buffers and `Uint8Array`, and `Image.buffer` accepts Node Buffers, with TypeScript 5.7+ typed arrays.

`index.d.ts` now starts with `/// <reference types="node" />`, so that it works with TypeScript 6 and 7, which no longer load `@types/*`
packages automatically. Install `@types/node` as a dev dependency (also in browser projects), or enable `skipLibCheck`. As a consequence,
Node's globals (`process`, `setTimeout` returning `NodeJS.Timeout`, ...) are visible in projects that compile `index.d.ts`, including
browser projects that set `"types": []`.

The typings use `get`/`set` accessor declarations in interfaces, which require TypeScript 4.3 or later; they are tested with TypeScript 5.9,
6.0 and 7.0.

## Zip library: fflate

`archiver`, `jszip` and `unzipper` were replaced by a single zero-dependency library, [fflate](https://github.com/101arrowz/fflate)
(^0.8.3), in the browser bundle and on Node.js and Bun. This supersedes their rows in the [Dependencies](#dependencies) table above. The zip
records (including ZIP64) are written and read by ExcelJS itself; fflate compresses and decompresses.

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
