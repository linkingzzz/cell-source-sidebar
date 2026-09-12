# cell-source-sidebar

Excel/WPS 插件：在侧边栏显示单元格数据来源，支持 Office & WPS（Windows / Mac），元数据与附件随文件保存（MVP）。

Status: Initial repository skeleton created. Development will proceed in stages.

Repository structure:
- office-addin/ — Office.js task pane add-in source skeleton (manifest + frontend)
- wps-plugin/ — WPS 插件 placeholder and implementation notes
- installer/ — Windows 一键安装包（源码 + 表格批注插件-Setup.exe，Excel/WPS 自动探测注册）
- examples/ — demo template and sample metadata
- docs/ — design & metadata schema

Next steps (immediate):
1. Implement Office Add-in MVP (selection listener, metadata read/write, attachments base64 store fallback to __metadata, UI skeleton).
2. Implement WPS plugin (Windows first), reusing metadata format and UI where possible.
3. ~~Build Windows installer~~ 已完成：`installer/output/表格批注插件-Setup.exe`（双击安装，含卸载；Mac 脚本暂缺）。

How to contribute:
- Open issues for features/bugs.
- I will push milestones and issues for the Stage 1 plan within 48 hours.

