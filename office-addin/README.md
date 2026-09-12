# Office Add-in - Cell Source Sidebar (Office.js)

Excel 任务窗格加载项（Office.js）。侧边栏是**针对选中单元格**的批注，分三块：1 附件（上传 / 点击下载 / 删除）、2 明细表（9 × 4 可编辑表格）、3 备注，底部「保存到工作簿」写回。功能区页签「数据源」下有 **添加侧边栏批注** / **隐藏侧边栏批注** 两个按钮，与 `wps-plugin/` 同款、同存储契约。

## 文件结构

| 文件 | 作用 |
| --- | --- |
| `manifest.xml` | 加载项清单：任务窗格 + `VersionOverrides`（自定义页签、两个功能区按钮、长生命周期共享运行时） |
| `src/taskpane.html` / `src/taskpane.js` | 侧边栏 UI：附件 / 9×4 明细表 / 备注；功能区动作 `hideAnnotation` 也在这里 |
| `src/metadata.js` | 元数据层（`window.metadataApi`），读写隐藏表 `__metadata` / `__metadata_attachments` |
| `src/assets/icon-{16,32,80}.png` | 功能区与清单图标 |

> **为什么没有 `commands.html`**：`Office.addin.hide()` 属于 **SharedRuntime 1.1** 要求集。没有共享运行时时，`Office.addin` 对象虽然存在，但 `hide()` 会抛 `There was an internal error while processing the request.`——按钮点了毫无反应。因此 `manifest.xml` 声明了 `<Set Name="SharedRuntime" MinVersion="1.1" />`，并用 `<Runtimes><Runtime resid="Taskpane.Url" lifetime="long" /></Runtimes>` 让功能区命令与侧边栏跑在同一个运行时里，`<FunctionFile>` 随之指向 `Taskpane.Url`，`hideAnnotation` 改用 `Office.actions.associate` 注册。

## 存储契约（与 WPS 版一致）

- `__metadata`：A key(`Sheet!A1`) · B type · C source · D url · E note · F attachments(JSON 数组) · G author · H last_modified · I version · **J grid（9×4 JSON）**
- `__metadata_attachments`：A id · B filename · C mime · D size · E base64 · F uploaded_by · G uploaded_at · **H part（1-based 分片序号）**
- base64 超过 30000 字符时按片独占一行，同 id 多行，读时按 `part` 拼接（Excel 单元格上限 32767）。单文件上限 10MB。

## Prerequisites
- Node.js (tested with v24) and Microsoft Excel desktop (Microsoft 365) on Windows.

## Run the pane locally

1. `npm install` — install dependencies.
2. `npx office-addin-dev-certs install` — create and trust a localhost certificate (one time). The manifest points at `https://localhost:3000`, so the pane must be served over HTTPS with a certificate Excel trusts.
3. `npm start` — serves `src/` over HTTPS on port 3000 using those certs. Sanity check: open `https://localhost:3000/taskpane.html`; the status line should become `Office ready`.

> PowerShell 提示：如果 `npm` 报「无法加载文件 npm.ps1，因为在此系统上禁止运行脚本」，那是 `ExecutionPolicy` 拦的 PSCmdlet 包装，不是报错也不是卡住。改用 `& "C:\Program Files\nodejs\npm.cmd" install` 即可。

## Sideload into Excel (Windows)

1. Excel → 文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录.
2. Add this folder (`...\cell-source-sidebar\office-addin`) as a catalog, tick 显示在菜单中, then restart Excel.
3. 插入 → 我的加载项 → 共享文件夹 → **Cell Source Sidebar** → 添加.
4. Open any workbook; the pane appears and writes metadata to a hidden `__metadata` sheet. 功能区「数据源」页签下应出现两个按钮。

The manifest needs a unique `<Id>` GUID, and its `SourceLocation` must match the URL the dev server actually serves.

## 在 Excel 里加载（实测有效的做法，2026-09-12）

只写 `HKCU\Software\Microsoft\Office\16.0\Wef\Developer` 注册表**不够**——那只是登记，功能区不会出现页签。真正让加载项加载的是 **sideload 文档**：官方工具会生成一个内嵌 web-extension 引用的工作簿，用 Excel 打开它即可挂载加载项。

```powershell
# 1) 起 HTTPS 开发服务器（另开一个终端）
& "C:\Program Files\nodejs\npm.cmd" start     # https://localhost:3000

# 2) 生成 sideload 文档并启动 Excel（二选一）
& "C:\Program Files\nodejs\npm.cmd" install --no-save office-addin-debugging
node -e "require('office-addin-dev-settings/lib/sideload.js').sideloadAddIn(require('path').resolve('manifest.xml'),'excel',false,false,'desktop')"
```

也可以走手工路径：把本文件夹加入 Excel「受信任的加载项目录」，再用 插入 → 我的加载项 → 共享文件夹 添加。

## 验证状态

- ✅ **已在真实 Excel 中实测通过**（Excel 16.0.20326.20132 / Microsoft 365 家庭版，Windows）：
  - 加载项加载：功能区出现「数据源」页签 + 两个按钮，右侧任务窗格打开并显示 `Office ready`。
  - 三模块读写：明细表（9×4）与备注填好后点「保存到工作簿」→ `__metadata` 出现 `Sheet1!C2` 行，J 列是 9×4 的 grid JSON（读盘核对过）；切走再切回该单元格能正确读回。
  - 附件：上传 `attach-sample.txt` 后自动建 `__metadata_attachments` 表并写入 base64；点附件名触发下载，落盘文件与源文件 **MD5 一致**（45 字节）。
  - 功能区「隐藏侧边栏批注」：隐藏右侧任务窗格（`Office.addin.hide()`，依赖上面的共享运行时）。
  - 功能区「添加侧边栏批注」：关掉侧边栏后点它可重新打开。
  - 2026-09-12 复测（改共享运行时之后，截图 `../_附件/excel-v3-*.png`）：页签 `数据源`、按钮 `隐藏侧边栏批注`（悬停提示「隐藏右侧的批注侧边栏」）→ 点它侧边栏消失 → 再点「添加侧边栏批注」回来 → 备注保存后状态变「已有批注」。
- ✅ `office-addin-manifest validate manifest.xml`（Microsoft 官方校验器）→ `The manifest is valid.`
- ✅ 静态契约一致性脚本（WPS 版 vs Office 版：常量、`parseGrid`、`emptyGrid`、`keyFromAddress`、DOM id、manifest resid）
- ⚠️ 已知观感差异：Excel 会把按钮标签折成两行（「添加侧 / 边栏批注」），WPS 同一标签是一行。功能不受影响。
- ℹ️ 功能区已不含「删除批注」：按钮现为「隐藏侧边栏批注」（只隐藏窗格），批注数据不会被删除；如需删除能力需另外补回。

Notes:
- This prototype uses worksheet-based metadata (hidden sheet `__metadata`) for read/write compatibility across hosts. CustomXMLParts support is left as an enhancement.
- See `tools/init_metadata.js` to inject sample metadata into an existing workbook (generates `examples/demo-initialized.xlsx`).
- 对照实现见 `../wps-plugin/`（WPS 原生 JSAPI 版，已在真实 WPS 12.1.0.28505 实测通过）。
