# 安装包（表格批注插件）

一键安装包：`output/表格批注插件-Setup.exe`（单文件，双击即装，不需要管理员权限）。

## 安装内容

| 宿主 | 安装动作 |
| --- | --- |
| Excel（Microsoft 365 / 2016+） | 复制 `manifest.xml` 到 `%LOCALAPPDATA%\表格批注插件\`；写 `HKCU\...\Office\16.0\WEF\Developer`（开发者加载项）与 `WEF\TrustedCatalogs`（插入 → 我的加载项 → 共享文件夹）；桌面建「表格批注插件」快捷方式，指向已内嵌加载项引用的空白工作簿 |
| WPS 表格 | 复制插件到 `%APPDATA%\kingsoft\wps\jsaddons\cell-source-sidebar_1.0.0\`，并在同目录 `publish.xml` 注册（`enable="enable_dev"`） |
| 两者 | 写卸载入口 `HKCU\...\Uninstall\表格批注插件`（设置 → 应用 → 已安装的应用） |

安装目录：`%LOCALAPPDATA%\表格批注插件\`（`manifest.xml`、`表格批注插件.xlsx`、`uninstall.exe`、`setup.log`）。
安装器会自动探测本机装了哪个宿主，装了就配，没装就跳过。

## 使用前提（重要）

- Excel 侧的侧边栏页面托管在 GitHub Pages：`https://linkingzzz.github.io/cell-source-sidebar/office-addin/src/taskpane.html`，**Excel 侧需要联网**（WPS 侧全部本地文件，离线可用）。
- Excel 的加载项靠「工作簿内嵌加载项引用 + 注册表开发者加载项」两件事一起生效：**请从安装时创建的「表格批注插件.xlsx」（桌面快捷方式）开始用**；已有的其它工作簿要插入该加载项，走「插入 → 我的加载项 → 共享文件夹 → 表格批注插件」。
- 仓库 `main` 分支根目录就是 Pages 站点，改完 `office-addin/src/` 推送后约 1 分钟自动生效，无需重新打包。
- WPS 首次加载 jsAddons 较慢（实测 1~2 分钟），之后在功能区最右侧出现「数据源」选项卡；Excel 重启后可见。
- 卸载：设置 → 应用 → 已安装的应用 → 表格批注插件；或 `%LOCALAPPDATA%\表格批注插件\uninstall.exe /uninstall`。静默安装/卸载加 `--silent`。

## 构建

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```

只用 Windows 自带的 .NET Framework 4.x `csc.exe`，不需要管理员权限，不需要 Inno Setup / NSIS 等第三方工具。
产物：`output/表格批注插件-Setup.exe`（安装程序把 payload 压缩成 `payload.zip` 内嵌进 exe，运行时解到临时目录再安装）。

## 文件说明

- `manifest.xml` — Office 加载项清单（发布版，指向 GitHub Pages）
- `assets/表格批注插件.xlsx` — 空白工作簿，内嵌加载项引用。做法：取
  `office-addin/node_modules/office-addin-dev-settings/templates/ExcelWorkbookWithTaskPane.xlsx`，
  把 `xl/webextensions/webextension.xml` 里的 `00000000-0000-0000-0000-000000000000` 换成 manifest 里的 `<Id>`。
  **`<Id>` 或 `<Version>` 变了必须重新生成这个工作簿**，否则 Excel 找不到加载项清单。
- `assets/app.ico` — 安装程序图标（由 `office-addin/src/assets/icon-80.png` 生成）
- `src/installer.cs` — 安装/卸载程序源码（WinForms；同时是 `uninstall.exe` 的源码，靠 `/uninstall` 参数区分）
- `build.ps1` — 打包脚本
- `output/` — 构建产物（提交进仓库，方便直接下载）