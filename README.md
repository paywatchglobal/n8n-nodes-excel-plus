# @paywatchglobal/n8n-nodes-excel-plus

An n8n community node for working with Excel files. The node accepts up to 20 binary inputs and produces a single merged workbook. More Excel operations will be added over time.

[n8n](https://n8n.io/) is a [fair-code licensed](https://docs.n8n.io/sustainable-use-license/) workflow automation platform.

## Installation

Follow the [installation guide](https://docs.n8n.io/integrations/community-nodes/installation/) in the n8n community nodes documentation. Install `@paywatchglobal/n8n-nodes-excel-plus` from the in-app community nodes installer.

## Operations

### Merge Sheets

Combines every sheet from each input workbook into one output workbook (one tab per source sheet). Cell values, basic styles, column widths and merged cells are preserved.

Properties:

- **Number of Inputs** — how many input connections the node accepts (1-20). Reconnect upstream nodes after changing.
- **Binary Field Mode**
  - *Same Across All Inputs*: every input is expected to expose the same binary property name (default `data`).
  - *Different per Input*: a JSON object keyed by input number, e.g. `{ "1": "data", "2": "report" }`.
- **Binary Field Name** — the binary property name when mode is *Same* (default `data`).
- **Binary Field Names per Input** — JSON map when mode is *Different*.
- **Sheet Name Conflict**
  - *Auto-Suffix With Counter*: duplicate names get ` (1)`, ` (2)`, … appended.
  - *Prefix With File Name*: each output sheet is prefixed with the source file name (no extension).

Output:

- Single item with binary property `data` containing `merged.xlsx`.
- `json` payload includes `operation`, `mergedFiles`, `mergedSheets`, `sheetNames`, and `skipped` (when "Continue On Fail" is enabled and some inputs failed).

## Compatibility

- Tested against n8n 1.x.
- Bundles `exceljs` v4.4.0 into the dist artifact, so no additional runtime dependencies need to be installed inside n8n.

## Resources

- [n8n community nodes documentation](https://docs.n8n.io/integrations/#community-nodes)
- [ExcelJS](https://github.com/exceljs/exceljs) — the underlying Excel library

## Version history

- **0.1.0** — initial release with the `Merge Sheets` operation.
