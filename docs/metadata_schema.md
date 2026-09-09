# Metadata schema (__metadata)

Two sheets in workbook used by MVP:

1) Sheet: __metadata (Entries)
Columns:
- A: key (string) e.g. "Sheet1!A1", "Sheet1!B", "Sheet1!Header:Price", "Sheet1"
- B: type (cell/column/header/sheet/table)
- C: source (text)
- D: url (text)
- E: note (text)
- F: attachments (JSON array of attachment ids) e.g. ["att-0001"]
- G: author
- H: last_modified (ISO8601)
- I: version

2) Sheet: __metadata_attachments (Attachments)
Columns:
- A: id (string e.g. att-0001)
- B: filename
- C: mime
- D: size (bytes)
- E: base64 (attachment payload)
- F: uploaded_by
- G: uploaded_at

Notes & limits
- Excel cell text length ~32k characters. Storing large base64 values in a single cell may exceed limits.
- Office Add-in will prefer CustomXMLParts for attachments when available; fallback to splitting base64 across cells in __metadata_attachments or storing small attachments directly.
- MVP enforces single-file limit 10MB via client-side validation.

