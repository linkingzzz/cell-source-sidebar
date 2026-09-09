# High-level Design (MVP)

See design decisions summarized here. This document is a concise extract of the technical design agreed with the stakeholder.

Goals
- Provide a user-installable plugin that integrates with Excel (Office) and WPS, on Windows and Mac, to show and edit "data source" metadata for cells in a task-pane sidebar.
- Metadata and attachments must be saved inside the workbook so they travel with the file.
- MVP attachments limit: single file <= 10MB (base64 embedded in workbook or stored in CustomXMLParts when possible).

Architecture
- Office Add-in (Office.js) for Excel Desktop (Win/Mac) and Excel Online. Task Pane UI built in HTML/JS (React optional).
- WPS plugin using WPS SDK / plugin interface. UI should mirror the Office add-in to maximize code reuse (embed the same web UI inside WPS webview/browser control if possible).
- Metadata persisted in workbook via hidden worksheet `__metadata` (Entries) and `__metadata_attachments`. Office Add-in uses CustomXMLParts when available and keeps a `__metadata` fallback for WPS compatibility.

Key features
- Selection-changed listener -> lookup metadata (cell/column/header/sheet precedence) -> render editable record in sidebar.
- Attachments upload/download (client-side base64 <10MB) -> saved into workbook attachments table.
- Metadata editing and immediate save back to workbook.
- Import/export metadata JSON; initialize skeleton on new workbooks.

Installer & Distribution
- Windows: .msi/.exe installer that registers the Office add-in manifest (local catalog or per-user registration) and installs WPS plugin files.
- Mac: .pkg + scripts to automate manifest registration where possible; otherwise provide scripted one-click instructions to add the add-in.

Milestones (MVP)
- Stage 1: Repo skeleton & design (this commit)
- Stage 2: Office add-in MVP (10-14 working days)
- Stage 3: WPS plugin MVP (Windows-first) (7-12 working days)
- Stage 4: Windows installer (3-5 days)
- Stage 5: Mac installer scripting (3-7 days)
- Stage 6: Integration, testing, docs (4-7 days)

