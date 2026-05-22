# Changelog

## Unreleased

### Added
- `Options > Output File Name` on the `Merge Sheets` operation, allowing the merged workbook's file name to be customized. The `.xlsx` extension is enforced automatically.

## 0.1.0 - 2026-05-22

### Added
- Initial `Excel Plus` node with a `Merge Sheets` operation.
- Dynamic input count (1-20) configurable from the node UI.
- Binary field name configuration with two modes: a single shared name across all inputs, or a JSON map of input number -> binary field name.
- Sheet name conflict handling: auto-suffix with counter, or prefix with source file name.
- Output is a single binary workbook (`merged.xlsx`) under the `data` property.
