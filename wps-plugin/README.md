# WPS plugin

This folder will contain the WPS plugin source. Goal: provide a side-panel in WPS that mirrors the Office Add-in functionality and reads/writes the same `__metadata` sheets in the workbook.

Notes:
- WPS plugin dev uses the WPS SDK; we will attempt to embed the same web UI (from office-addin/src) inside a WebView or Browser control inside WPS plugin.
- Primary target: WPS for Windows (recent versions). WPS for Mac compatibility to be evaluated later.

