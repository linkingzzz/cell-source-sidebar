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
- J: grid (JSON 2D array of strings, the sidebar detail table; 9 rows x 4 columns)

2) Sheet: __metadata_attachments (Attachments)
Columns:
- A: id (string e.g. att-0001)
- B: filename
- C: mime
- D: size (bytes)
- E: base64 (attachment payload)
- F: uploaded_by
- G: uploaded_at
- H: part (1-based chunk index; one payload spans several rows when its base64 exceeds one cell)

Notes & limits
- Excel cell text length ~32k characters. Storing large base64 values in a single cell may exceed limits.
- Both hosts implement the split the same way: base64 is chunked at 30000 characters, one row per chunk sharing the same id and numbered in column H; readers concatenate in `part` order (see `wps-plugin/js/metadata.js` and `office-addin/src/metadata.js`).
- CustomXMLParts remain an option for attachments but are not used: both hosts store base64 in __metadata_attachments.
- Columns J (grid) and H (part) were added with the WPS sidebar feature; readers must treat them as optional so workbooks written before that stay readable.
- MVP enforces single-file limit 10MB via client-side validation.

