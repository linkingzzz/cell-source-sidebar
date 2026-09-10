# WPS plugin（WPS 原生加载项）

A WPS-native add-in (JSAPI) mirroring the Office task pane: it shows and edits a cell's data source in a side panel and reads/writes the same `__metadata` sheets, so one workbook keeps working across Office and WPS.

WPS does **not** load Office.js add-ins — its core modules contain no `appsforoffice` / `Office.onReady` support (verified against WPS Office 12.1.0.28505). It uses its own JSAPI (`window.Application`), so this is a separate implementation rather than a re-packaged `office-addin`.

## Layout

| Path | Purpose |
| --- | --- |
| `index.html` | Entry file WPS loads at startup |
| `main.js` | Pulls in `js/util.js` and `js/ribbon.js` |
| `ribbon.xml` | Ribbon tab 单元格来源 with the 显示/隐藏侧边栏 button |
| `manifest.xml` | Add-in manifest (`JsPlugin`) |
| `js/ribbon.js` | `OnAddinLoad` / `OnAction` / ribbon callbacks; creates the task pane |
| `js/metadata.js` | Metadata layer over the WPS API (`__metadata`, `__metadata_attachments`) |
| `ui/taskpane.html`, `js/taskpane.js` | The side panel itself |

## Install (Windows)

WPS reads add-ins from `%APPDATA%\kingsoft\wps\jsaddons\`:

1. Copy this folder to `%APPDATA%\kingsoft\wps\jsaddons\cell-source-sidebar_1.0.0\`.
2. Register it in `%APPDATA%\kingsoft\wps\jsaddons\publish.xml`:

       <jsplugins>
           <jsplugin name="cell-source-sidebar" type="et" url="cell-source-sidebar_1.0.0" version="1.0.0" enable="enable_dev" install="null" customDomain=""/>
       </jsplugins>

3. Restart WPS 表格 (`et.exe`). A 单元格来源 tab appears; click 显示/隐藏侧边栏 to open the panel.
4. Uninstall: delete the addon folder and its `<jsplugin>` entry.

## Verification

Set `AUTO_OPEN_TASKPANE = true` in `js/ribbon.js` to open the panel on startup. If the workbook also has a sheet named `__selftest__`, the panel runs a self-test on load: it reads `__metadata`, writes a probe row, saves the workbook and prints `SELFTEST PASS` in the panel. `../03-执行与产出/make-wps-test.js` generates such a workbook.

## Notes
- Keys use `Sheet!A1` on both hosts. WPS reports `Selection.Address` differently, so `js/metadata.js` falls back to building A1 from `Row` / `Column`.
- Attachments (base64 in `__metadata_attachments`) and JSON import/export are not implemented yet; the Office add-in has the attachment plumbing.
- Only WPS 表格 (ET) is targeted; Writer/Presentation are out of scope.
