# Office Add-in - Run locally and Sideload into Excel

This folder contains the Office Add-in prototype and development files. Use these steps to run the taskpane locally for testing.

## Prerequisites
- Node.js (tested with v24) and Microsoft Excel desktop (Microsoft 365) on Windows.

## Run the pane locally

1. `npm install` — install dependencies.
2. `npx office-addin-dev-certs install` — create and trust a localhost certificate (one time). The manifest points at `https://localhost:3000`, so the pane must be served over HTTPS with a certificate Excel trusts.
3. `npm start` — serves `src/` over HTTPS on port 3000 using those certs. Sanity check: open `https://localhost:3000/taskpane.html`; the status line should become `Office ready`.

## Sideload into Excel (Windows)

1. Excel → 文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录.
2. Add this folder (`...\cell-source-sidebar\office-addin`) as a catalog, tick 显示在菜单中, then restart Excel.
3. 插入 → 我的加载项 → 共享文件夹 → **Cell Source Sidebar** → 添加.
4. Open any workbook; the pane appears and writes metadata to a hidden `__metadata` sheet.

The manifest needs a unique `<Id>` GUID, and its `SourceLocation` must match the URL the dev server actually serves.

Notes:
- This prototype uses worksheet-based metadata (hidden sheet `__metadata`) for read/write compatibility across hosts. CustomXMLParts support is left as an enhancement.
- See `tools/init_metadata.js` to inject sample metadata into an existing workbook (generates `examples/demo-initialized.xlsx`).
- On macOS: install the certs the same way and swap `%USERPROFILE%` in the `start` script for `$HOME`. Outside Excel the selection poll logs `Excel is not defined`; that is expected and disappears inside the host.
