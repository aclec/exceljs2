# Security Policy

## Supported versions

| Version | Supported                                                                                  |
| ------- | ------------------------------------------------------------------------------------------ |
| 5.0.x   | Yes                                                                                        |
| < 5.0   | No (4.x and older are no longer maintained; upgrade with [UPGRADE-5.0.md](UPGRADE-5.0.md)) |

Security fixes are released as patch versions of the latest minor version.

## Reporting a vulnerability

Please do **not** open a public issue or pull request for a vulnerability.

Report it privately with GitHub's private vulnerability reporting: open the **Security** tab of this repository and choose **Report a
vulnerability**. Include:

- the ExcelJS version and runtime (Node.js, Bun or browser, with versions);
- the API that is called (for example `workbook.xlsx.load`, `csv.read`, `stream.xlsx.WorkbookReader`);
- a minimal reproduction: preferably a script that builds the malicious file, or the file itself;
- the impact you observed (crash, memory or CPU use, corrupted output, ...).

We will acknowledge the report, keep you informed of the fix, and agree on a disclosure date with you once a fix is available.

## Threat model

The main use case we defend is an application (typically a server) that **parses files supplied by untrusted users**: `workbook.xlsx.load` /
`read` / `readFile`, `workbook.csv.read` / `readFile`, and the streaming `stream.xlsx.WorkbookReader`. We also care about the files ExcelJS
**writes** from untrusted data.

### In scope

- Small inputs that cause disproportionate memory or CPU use, or that crash the process in a way the application cannot catch (out-of-memory
  aborts, catastrophic regular-expression backtracking, quadratic algorithms). Since 5.0, ExcelJS bounds cell references to the Excel grid,
  caps merged cells per worksheet, keeps large data-validation and defined-name ranges as ranges, clamps column definitions and rejects
  over-long formulas; see [Security & limits](README.md#security--limits).
- Prototype pollution or other corruption of the application's global state.
- Malformed or injected XML in written files (for example from cell values or names).
- Reading or writing files outside the paths given by the application, and code execution.

Errors thrown for malformed input (a rejected promise) are not vulnerabilities in themselves, but reports of unhelpful errors are welcome as
regular issues.

### Out of scope: application responsibilities

ExcelJS reads and writes data; it does not decide whether that data is safe to use elsewhere.

- **CSV / formula injection.** CSV output is written as is. Values starting with `=`, `+`, `-`, `@`, tab or carriage return can be
  interpreted as formulas by spreadsheet applications. Neutralise untrusted values before writing CSV.
- **Hyperlinks and other URLs** read from a file (`javascript:`, `file:`, ...) are returned unchanged. Validate them before rendering or
  following them.
- **Decompressed size (zip bombs).** ExcelJS does not limit the decompressed size of the parts of an xlsx file, and memory use grows with
  the size of legitimate workbooks. Limit upload sizes, and parse untrusted files in a worker thread or a separate process with memory and
  time limits.
- **Formulas are not evaluated** by ExcelJS, but they are kept and will be evaluated by Excel when a user opens a file that you generated
  from untrusted input.
- **Content passed through unchanged**, such as theme parts and media, is not sanitised.
- **Worksheet protection** (`worksheet.protect`) is an editing restriction understood by Excel, not encryption or access control.
