# WPS plugin（WPS 原生加载项）

A WPS-native add-in (JSAPI) mirroring the Office task pane: it stores **per-cell** annotations — attachments, a 9×4 detail table and a note — inside the workbook, so the data travels with the file.

WPS does **not** load Office.js add-ins — its core modules contain no `appsforoffice` / `Office.onReady` support (verified against WPS Office 12.1.0.28505). It uses its own JSAPI (`window.Application`), so this is a separate implementation rather than a re-packaged `office-addin`.

## Ribbon

Tab 数据源 → group 侧边栏批注:

| Button | Behaviour |
| --- | --- |
| 显示侧边栏批注 | Opens the sidebar for the currently selected cell |
| 隐藏侧边栏批注 | Hides the sidebar (annotation data is never deleted) |

The sidebar is always about the **currently selected cell**. It re-reads when the selection changes (1.5s poll) and also picks up changes made from the ribbon.

## Sidebar

1. **附件** — 上传附件 opens a native file picker. Each attachment is listed by filename; click a name to download it, 删除 to drop it. Payloads are stored base64-in-workbook; single file ≤ 10 MB.
2. **明细表（9 × 4）** — a 9-row × 4-column editable grid.
3. **备注** — free text.

「保存到工作簿」 writes the grid + note (plus the attachment ids) and saves the workbook. Uploading an attachment saves immediately.

## Storage contract

Same two sheets as the Office add-in, so one workbook works in both hosts.

- `__metadata` — one row per annotated cell, keyed `Sheet!A1`:
  A key · B type · C source · D url · E note · F attachments (JSON array of ids) · G author · H last_modified · I version · **J grid (JSON 9×4 array)**
- `__metadata_attachments`:
  A id · B filename · C mime · D size · E base64 · F uploaded_by · G uploaded_at · **H part (1-based index)**

A WPS cell holds at most 32767 characters, so base64 is split into 30000-character chunks — one row per chunk, same `id`, numbered in column `H`; reads concatenate them in `part` order. Columns J and H were added for this feature; workbooks written by the previous version still read fine (missing header or cell → empty grid).

Both sheets are created **hidden** (`Visible = 0`) and the previously active sheet is re-activated before hiding, so writing data never adds a visible tab or jumps the view. Unhide them from the sheet-tab context menu when you want to inspect the raw rows.

## Layout

| Path | Purpose |
| --- | --- |
| `index.html` | Entry file WPS loads at startup |
| `main.js` | Pulls in `js/util.js`, `js/metadata.js`, `js/ribbon.js` |
| `ribbon.xml` | Ribbon tab 数据源 with the two buttons |
| `manifest.xml` | Add-in manifest (`JsPlugin`) |
| `js/ribbon.js` | `OnAddinLoad` / `OnAction` / `ShowTaskPane`; creates and hides the task pane |
| `js/metadata.js` | Storage layer over the WPS API (`__metadata`, `__metadata_attachments`) |
| `ui/taskpane.html`, `js/taskpane.js` | The side panel itself |

## Install (Windows)

WPS reads add-ins from `%APPDATA%\kingsoft\wps\jsaddons\`:

1. Copy this folder to `%APPDATA%\kingsoft\wps\jsaddons\cell-source-sidebar_1.0.0\`.
2. Register it in `%APPDATA%\kingsoft\wps\jsaddons\publish.xml`:

       <jsplugins>
           <jsplugin name="cell-source-sidebar" type="et" url="cell-source-sidebar_1.0.0" version="1.0.0" enable="enable_dev" install="null" customDomain=""/>
       </jsplugins>

3. Restart WPS 表格 (`et.exe`). A 数据源 tab appears.
4. Uninstall: delete the addon folder and its `<jsplugin>` entry.

## Verification

Set `AUTO_OPEN_TASKPANE = true` in `js/ribbon.js` to open the panel on startup. If the workbook also has a sheet named `__selftest__`, the panel runs a self-test on load: it writes a 9×4 grid and a small base64 attachment, reads both back, saves, and prints `SELFTEST PASS` in the panel. `../03-执行与产出/make-wps-test.js` generates such a workbook.

## Notes

- Keys use `Sheet!A1` on both hosts. WPS reports `Selection.Address` as a **method**, not a property, so `js/metadata.js` falls back to building A1 from `Row` / `Column`.
- `PluginStorage` persists `taskpane_id` across sessions, so `ShowTaskPane` recreates the pane when a stored id no longer resolves.
- Ribbon actions write their result to `PluginStorage["css_status"]`, which the sidebar shows — so ribbon actions give visible feedback even though the ribbon and the pane are separate web views.
- Deleting a row falls back to clearing its cells if `Range.Delete()` is unavailable; empty keys are skipped on read.
- JSON import/export is not implemented yet.
- Only WPS 表格 (ET) is targeted; Writer/Presentation are out of scope.
