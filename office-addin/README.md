# Office Add-in Sideload & Development

This folder contains the Office Add-in prototype and development files. Use these steps to run the taskpane locally for testing.

Quick start (development):
1. Install dependencies: npm install
2. Start a static server that serves the `office-addin/src` folder (e.g., `npx http-server office-addin/src -p 3000`) or run the dev server of your choice.
3. Edit `office-addin/manifest.xml` to point <SourceLocation DefaultValue="https://localhost:3000/taskpane.html" /> or the URL where you host files.
4. Sideload the add-in into Excel (Windows/Mac) following the Office sideload instructions. For quick testing use Excel Online by loading the manifest in the Office Add-ins dialog from URL.

Notes:
- This prototype uses worksheet-based metadata (hidden sheet `__metadata`) for read/write compatibility across hosts. CustomXMLParts support is left as an enhancement.
- See `tools/init_metadata.js` to inject sample metadata into an existing workbook (generates `examples/demo-initialized.xlsx`).
