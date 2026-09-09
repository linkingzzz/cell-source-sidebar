# Tests / Example Integration

This file records that `examples/示例.xlsx` is included as a demo workbook for Stage 2 development and testing.

Steps to use the example for local testing:

1. Download the example workbook from: https://github.com/linkingzzz/cell-source-sidebar/raw/main/examples/%E7%A4%BA%E4%BE%8B.xlsx
2. Use the Office Add-in local sideload instructions (see `office-addin/README.md`) to load the add-in in Excel.
3. Open the example workbook and verify the sidebar shows metadata (after Stage 2 implementation). If the workbook lacks `__metadata`, use `examples/sample-metadata.json` as a template to initialize metadata.

Notes:
- The example workbook currently appears not to contain a `__metadata` sheet; you can import `examples/sample-metadata.json` using the add-in import feature (once implemented) or manually add a `__metadata` sheet.
