# cell-source-sidebar

Excel/WPS 插件：在侧边栏显示单元格数据来源，支持 Office & WPS（Windows / Mac），元数据与附件随文件保存（MVP）。

Status: Initial repository skeleton created. Development will proceed in stages.

Repository structure:
- office-addin/ — Office.js task pane add-in source skeleton (manifest + frontend)
- wps-plugin/ — WPS 插件 placeholder and implementation notes
- installer/ — Installer scripts and notes (Windows / Mac)
- examples/ — demo template and sample metadata
- docs/ — design & metadata schema

Next steps (immediate):
1. Implement Office Add-in MVP (selection listener, metadata read/write, attachments base64 store fallback to __metadata, UI skeleton).
2. Implement WPS plugin (Windows first), reusing metadata format and UI where possible.
3. Build Windows installer that registers the add-in manifest and installs WPS plugin; provide Mac installer scripts.

How to contribute:
- Open issues for features/bugs.
- I will push milestones and issues for the Stage 1 plan within 48 hours.

