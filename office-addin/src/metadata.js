// office-addin/src/metadata.js
// Workbook storage layer. Mirrors the WPS add-in contract exactly:
//   __metadata             : one row per annotated cell, key = "Sheet!A1", column J = 9x4 grid (JSON)
//   __metadata_attachments : base64 payload split into 30000-char chunks, column H = 1-based part index

const META_SHEET = "__metadata";
const ATT_SHEET = "__metadata_attachments";
const META_HEADER = ["key", "type", "source", "url", "note", "attachments", "author", "last_modified", "version", "grid"];
const ATT_HEADER = ["id", "filename", "mime", "size", "base64", "uploaded_by", "uploaded_at", "part"];
const GRID_ROWS = 9;
const GRID_COLS = 4;
// Excel/WPS both cap a cell at 32767 characters; stay under it with room to spare.
const CELL_LIMIT = 30000;
const MAX_FILE_BYTES = 10 * 1024 * 1024;

function emptyGrid() {
  const g = [];
  for (let r = 0; r < GRID_ROWS; r++) {
    const row = [];
    for (let c = 0; c < GRID_COLS; c++) row.push("");
    g.push(row);
  }
  return g;
}

function parseGrid(raw) {
  const out = emptyGrid();
  if (!raw) return out;
  let g = null;
  try {
    g = JSON.parse(String(raw));
  } catch (e) {
    return out;
  }
  if (!g || !g.length) return out;
  for (let r = 0; r < GRID_ROWS && r < g.length; r++) {
    for (let c = 0; c < GRID_COLS && c < g[r].length; c++) {
      out[r][c] = g[r][c] === null || g[r][c] === undefined ? "" : String(g[r][c]);
    }
  }
  return out;
}

function chunkBase64(b64) {
  const parts = [];
  const s = String(b64 || "");
  for (let i = 0; i < s.length; i += CELL_LIMIT) parts.push(s.substring(i, i + CELL_LIMIT));
  if (!parts.length) parts.push("");
  return parts;
}

// Office reports "Sheet1!A1" but quotes sheet names containing spaces ("'My Sheet'!A1").
// WPS builds keys as name + "!" + address without quotes, so unquote here to keep one key format.
function keyFromAddress(address) {
  const raw = String(address || "").replace(/\$/g, "");
  const bang = raw.lastIndexOf("!");
  if (bang < 0) return raw;
  let sheet = raw.substring(0, bang);
  const cell = raw.substring(bang + 1);
  if (sheet.length > 1 && sheet.charAt(0) === "'" && sheet.charAt(sheet.length - 1) === "'") {
    sheet = sheet.substring(1, sheet.length - 1).replace(/''/g, "'");
  }
  return sheet + "!" + cell;
}

function entryFromRow(r, rowIndex) {
  let attachments = [];
  try {
    attachments = r[5] ? JSON.parse(r[5]) : [];
  } catch (e) {
    attachments = [];
  }
  if (!Array.isArray(attachments)) attachments = [];
  return {
    key: String(r[0] || ""),
    type: r[1] || "cell",
    source: r[2] || "",
    url: r[3] || "",
    note: r[4] || "",
    attachments,
    author: r[6] || "",
    last_modified: r[7] || "",
    version: r[8] || "",
    grid: parseGrid(r[9]),
    row: rowIndex + 1
  };
}

// Returns the sheet, creating it (with its header) when missing. Also tops up header cells that a
// workbook written by an older version lacks, so the grid/part columns appear on first write.
async function ensureSheet(context, name, header) {
  let sheet = context.workbook.worksheets.getItemOrNullObject(name);
  sheet.load("name");
  await context.sync();
  if (sheet.isNullObject) {
    sheet = context.workbook.worksheets.add(name);
    sheet.getRangeByIndexes(0, 0, 1, header.length).values = [header];
    await context.sync();
    return sheet;
  }
  const head = sheet.getRangeByIndexes(0, 0, 1, header.length);
  head.load("values");
  await context.sync();
  const current = (head.values && head.values[0]) || [];
  const fixed = [];
  let changed = false;
  for (let c = 0; c < header.length; c++) {
    const v = current[c];
    if (v === "" || v === null || v === undefined) {
      fixed.push(header[c]);
      changed = true;
    } else {
      fixed.push(v);
    }
  }
  if (changed) {
    head.values = [fixed];
    await context.sync();
  }
  return sheet;
}

async function usedValues(context, sheet) {
  const used = sheet.getUsedRangeOrNullObject();
  used.load("values");
  await context.sync();
  return used.isNullObject ? [] : used.values;
}

// Selection key + all entries in one round trip (the pane polls this).
async function readState() {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    const sheet = context.workbook.worksheets.getItemOrNullObject(META_SHEET);
    const used = sheet.getUsedRangeOrNullObject();
    used.load("values");
    const attSheet = context.workbook.worksheets.getItemOrNullObject(ATT_SHEET);
    const attUsed = attSheet.getUsedRangeOrNullObject();
    attUsed.load("values");
    await context.sync();
    let entries = [];
    if (!sheet.isNullObject && !used.isNullObject) {
      const rows = used.values;
      for (let i = 1; i < rows.length; i++) {
        if (!rows[i] || !rows[i][0]) continue;
        entries.push(entryFromRow(rows[i], i));
      }
    }
    // One record per attachment id (a payload spans several rows); the pane shows these names.
    const seen = {};
    const attachments = [];
    if (!attSheet.isNullObject && !attUsed.isNullObject) {
      const rows = attUsed.values;
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r || !r[0] || seen[r[0]]) continue;
        seen[r[0]] = true;
        attachments.push({ id: String(r[0]), filename: r[1] || "", mime: r[2] || "", size: r[3] || 0 });
      }
    }
    return { key: keyFromAddress(range.address), entries, attachments };
  });
}

async function readEntries() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(META_SHEET);
    const rows = await usedValues(context, sheet);
    if (sheet.isNullObject) return [];
    const entries = [];
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i] || !rows[i][0]) continue;
      entries.push(entryFromRow(rows[i], i));
    }
    return entries;
  });
}

async function findEntry(key) {
  const entries = await readEntries();
  return entries.find((e) => e.key === String(key)) || null;
}

async function writeEntry(entry) {
  return Excel.run(async (context) => {
    const sheet = await ensureSheet(context, META_SHEET, META_HEADER);
    const rows = await usedValues(context, sheet);
    let target = -1;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i] && String(rows[i][0]) === String(entry.key)) {
        target = i + 1;
        break;
      }
    }
    if (target < 0) target = rows.length < 1 ? 2 : rows.length + 1;
    sheet.getRangeByIndexes(target - 1, 0, 1, META_HEADER.length).values = [[
      entry.key,
      entry.type || "cell",
      entry.source || "",
      entry.url || "",
      entry.note || "",
      JSON.stringify(entry.attachments || []),
      entry.author || "",
      new Date().toISOString(),
      entry.version || "",
      JSON.stringify(entry.grid || emptyGrid())
    ]];
    await context.sync();
    return target;
  });
}

async function deleteEntry(key) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(META_SHEET);
    const rows = await usedValues(context, sheet);
    if (sheet.isNullObject) return 0;
    let target = -1;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i] && String(rows[i][0]) === String(key)) {
        target = i + 1;
        break;
      }
    }
    if (target < 0) return 0;
    sheet.getRangeByIndexes(target - 1, 0, 1, META_HEADER.length).delete(Excel.DeleteShiftDirection.up);
    await context.sync();
    return target;
  });
}

async function writeAttachment(attId, att) {
  return Excel.run(async (context) => {
    const sheet = await ensureSheet(context, ATT_SHEET, ATT_HEADER);
    const rows = await usedValues(context, sheet);
    const stale = [];
    for (let i = 1; i < rows.length; i++) {
      if (rows[i] && String(rows[i][0]) === String(attId)) stale.push(i + 1);
    }
    for (let k = stale.length - 1; k >= 0; k--) {
      sheet.getRangeByIndexes(stale[k] - 1, 0, 1, ATT_HEADER.length).delete(Excel.DeleteShiftDirection.up);
    }
    if (stale.length) await context.sync();

    const parts = chunkBase64(att.base64);
    const remaining = stale.length ? await usedValues(context, sheet) : rows;
    const start = remaining.length < 1 ? 2 : remaining.length + 1;
    const stamp = new Date().toISOString();
    const block = parts.map((p, i) => [
      attId,
      att.filename || "",
      att.mime || "",
      att.size || 0,
      p,
      att.uploaded_by || "",
      stamp,
      i + 1
    ]);
    sheet.getRangeByIndexes(start - 1, 0, block.length, ATT_HEADER.length).values = block;
    await context.sync();
    return block.length;
  });
}

async function readAttachmentMeta(attId) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(ATT_SHEET);
    const rows = await usedValues(context, sheet);
    if (sheet.isNullObject) return null;
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (r && String(r[0]) === String(attId)) {
        return { id: String(attId), filename: r[1] || "", mime: r[2] || "", size: r[3] || 0 };
      }
    }
    return null;
  });
}

async function readAttachment(attId) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(ATT_SHEET);
    const rows = await usedValues(context, sheet);
    if (sheet.isNullObject) return null;
    const chunks = [];
    let meta = null;
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || String(r[0]) !== String(attId)) continue;
      if (!meta) meta = { id: String(attId), filename: r[1] || "", mime: r[2] || "", size: r[3] || 0 };
      chunks.push({ part: Number(r[7]) || 1, data: String(r[4] === null || r[4] === undefined ? "" : r[4]) });
    }
    if (!meta) return null;
    chunks.sort((a, b) => a.part - b.part);
    meta.base64 = chunks.map((c) => c.data).join("");
    return meta;
  });
}

async function deleteAttachment(attId) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(ATT_SHEET);
    const rows = await usedValues(context, sheet);
    if (sheet.isNullObject) return 0;
    const hits = [];
    for (let i = 1; i < rows.length; i++) {
      if (rows[i] && String(rows[i][0]) === String(attId)) hits.push(i + 1);
    }
    for (let k = hits.length - 1; k >= 0; k--) {
      sheet.getRangeByIndexes(hits[k] - 1, 0, 1, ATT_HEADER.length).delete(Excel.DeleteShiftDirection.up);
    }
    if (hits.length) await context.sync();
    return hits.length;
  });
}

async function initMetadataSheetFromJson(metaJson) {
  return Excel.run(async (context) => {
    const metaSheet = await ensureSheet(context, META_SHEET, META_HEADER);
    metaSheet.getRange().clear();
    const rows = [META_HEADER];
    (metaJson.entries || []).forEach((e) => {
      rows.push([
        e.key || "", e.type || "cell", e.source || "", e.url || "", e.note || "",
        JSON.stringify(e.attachments || []), e.author || "", e.last_modified || "", e.version || "",
        JSON.stringify(e.grid || emptyGrid())
      ]);
    });
    metaSheet.getRangeByIndexes(0, 0, rows.length, META_HEADER.length).values = rows;

    const attSheet = await ensureSheet(context, ATT_SHEET, ATT_HEADER);
    attSheet.getRange().clear();
    const attRows = [ATT_HEADER];
    (metaJson.attachments || []).forEach((a) => {
      chunkBase64(a.base64).forEach((p, i) => {
        attRows.push([a.id || "", a.filename || "", a.mime || "", a.size || 0, p, a.uploaded_by || "", a.uploaded_at || "", i + 1]);
      });
    });
    attSheet.getRangeByIndexes(0, 0, attRows.length, ATT_HEADER.length).values = attRows;

    metaSheet.visibility = "Hidden";
    attSheet.visibility = "Hidden";
    await context.sync();
    return true;
  });
}

window.metadataApi = {
  META_SHEET,
  ATT_SHEET,
  GRID_ROWS,
  GRID_COLS,
  CELL_LIMIT,
  MAX_FILE_BYTES,
  emptyGrid,
  parseGrid,
  chunkBase64,
  keyFromAddress,
  readState,
  readEntries,
  findEntry,
  writeEntry,
  deleteEntry,
  writeAttachment,
  readAttachmentMeta,
  readAttachment,
  deleteAttachment,
  initMetadataSheetFromJson
};
