// WPS(ET) 版元数据层：__metadata（条目 + 9x4 明细表）与 __metadata_attachments（附件，base64 分片存）
var META_SHEET = "__metadata"
var ATT_SHEET = "__metadata_attachments"
var META_HEADER = ["key", "type", "source", "url", "note", "attachments", "author", "last_modified", "version", "grid"]
var ATT_HEADER = ["id", "filename", "mime", "size", "base64", "uploaded_by", "uploaded_at", "part"]
var KEY_COL = 1
var NOTE_COL = 5
var ATT_IDS_COL = 6
var GRID_COL = 10
var ATT_PART_COL = 8
var CELL_LIMIT = 30000
var GRID_ROWS = 9
var GRID_COLS = 4
var MAX_FILE_BYTES = 10 * 1024 * 1024

function getApp() {
    return window.Application
}

function getSheet(name) {
    let wb = getApp().ActiveWorkbook
    if (!wb) return null
    let sheets = wb.Worksheets
    for (let i = 1; i <= sheets.Count; i++) {
        let s = sheets.Item(i)
        if (s.Name == name) return s
    }
    return null
}

function ensureSheet(name, header) {
    let s = getSheet(name)
    if (!s) {
        s = getApp().ActiveWorkbook.Worksheets.Add()
        s.Name = name
    }
    for (let c = 0; c < header.length; c++) {
        if (!cellText(s, 1, c + 1)) s.Cells.Item(1, c + 1).Value2 = header[c]
    }
    return s
}

function rowCount(sheet) {
    try {
        return sheet.UsedRange.Rows.Count
    } catch (e) {
        return 0
    }
}

function cellText(sheet, r, c) {
    let v = sheet.Cells.Item(r, c).Value2
    return v === null || v === undefined ? "" : String(v)
}

function cellRange(sheet, r, cols) {
    return sheet.Range(sheet.Cells.Item(r, 1), sheet.Cells.Item(r, cols))
}

// 删行；WPS 的行删除接口不保证可用，失败时退化为清空内容（readMetadata 会跳过空 key 行）
function dropRow(sheet, r, cols) {
    let row = cellRange(sheet, r, cols)
    try {
        row.Delete()
    } catch (e) {
        row.ClearContents()
    }
}

function emptyGrid() {
    let g = []
    for (let r = 0; r < GRID_ROWS; r++) {
        let row = []
        for (let c = 0; c < GRID_COLS; c++) row.push("")
        g.push(row)
    }
    return g
}

function parseGrid(raw) {
    let out = emptyGrid()
    if (!raw) return out
    let g = null
    try {
        g = JSON.parse(String(raw))
    } catch (e) {
        return out
    }
    if (!g || !g.length) return out
    for (let r = 0; r < GRID_ROWS && r < g.length; r++) {
        for (let c = 0; c < GRID_COLS && c < g[r].length; c++) {
            out[r][c] = g[r][c] === null || g[r][c] === undefined ? "" : String(g[r][c])
        }
    }
    return out
}

function findEntryRow(ms, key) {
    let rows = rowCount(ms)
    for (let r = 2; r <= rows; r++) {
        if (cellText(ms, r, KEY_COL) == String(key)) return r
    }
    return 0
}

function entryFromRow(ms, r) {
    let atts = []
    let raw = cellText(ms, r, ATT_IDS_COL)
    if (raw) {
        try {
            atts = JSON.parse(raw)
        } catch (e) {
            atts = []
        }
    }
    if (!(atts instanceof Array)) atts = []
    return {
        key: cellText(ms, r, KEY_COL),
        type: cellText(ms, r, 2) || "cell",
        source: cellText(ms, r, 3),
        url: cellText(ms, r, 4),
        note: cellText(ms, r, NOTE_COL),
        attachments: atts,
        last_modified: cellText(ms, r, 8),
        grid: parseGrid(cellText(ms, r, GRID_COL)),
        row: r
    }
}

function readEntries() {
    let entries = []
    let ms = getSheet(META_SHEET)
    if (!ms) return entries
    let rows = rowCount(ms)
    for (let r = 2; r <= rows; r++) {
        if (!cellText(ms, r, KEY_COL)) continue
        entries.push(entryFromRow(ms, r))
    }
    return entries
}

function findEntry(key) {
    let entries = readEntries()
    for (let i = 0; i < entries.length; i++) {
        if (entries[i].key == String(key)) return entries[i]
    }
    return null
}

function writeEntry(entry) {
    let ms = ensureSheet(META_SHEET, META_HEADER)
    let rows = rowCount(ms)
    let target = findEntryRow(ms, entry.key)
    if (!target) target = rows < 2 ? 2 : rows + 1
    let vals = [
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
    ]
    for (let c = 0; c < vals.length; c++) {
        ms.Cells.Item(target, c + 1).Value2 = vals[c]
    }
    return target
}

function deleteEntryRow(key) {
    let ms = getSheet(META_SHEET)
    if (!ms) return 0
    let r = findEntryRow(ms, key)
    if (!r) return 0
    dropRow(ms, r, META_HEADER.length)
    return r
}

function newAttachmentId() {
    return "att-" + Date.now().toString(36) + "-" + Math.floor(Math.random() * 1679616).toString(36)
}

function findAttachmentRows(id) {
    let as = getSheet(ATT_SHEET)
    let hit = []
    if (!as) return hit
    let rows = rowCount(as)
    for (let r = 2; r <= rows; r++) {
        if (cellText(as, r, 1) == String(id)) hit.push(r)
    }
    return hit
}

function writeAttachment(id, filename, mime, size, base64) {
    let as = ensureSheet(ATT_SHEET, ATT_HEADER)
    for (let r = rowCount(as); r >= 2; r--) {
        if (cellText(as, r, 1) == String(id)) dropRow(as, r, ATT_HEADER.length)
    }
    let parts = []
    for (let i = 0; i < base64.length; i += CELL_LIMIT) parts.push(base64.substring(i, i + CELL_LIMIT))
    if (!parts.length) parts.push("")
    let stamp = new Date().toISOString()
    for (let p = 0; p < parts.length; p++) {
        let r = rowCount(as) + 1
        if (r < 2) r = 2
        let vals = [id, filename, mime || "", size, parts[p], "", stamp, p + 1]
        for (let c = 0; c < vals.length; c++) as.Cells.Item(r, c + 1).Value2 = vals[c]
    }
    return parts.length
}

function deleteAttachmentRows(id) {
    let as = getSheet(ATT_SHEET)
    if (!as) return 0
    let n = 0
    for (let r = rowCount(as); r >= 2; r--) {
        if (cellText(as, r, 1) == String(id)) {
            dropRow(as, r, ATT_HEADER.length)
            n++
        }
    }
    return n
}

function readAttachmentMeta(id) {
    let as = getSheet(ATT_SHEET)
    if (!as) return null
    let rows = rowCount(as)
    for (let r = 2; r <= rows; r++) {
        if (cellText(as, r, 1) == String(id)) {
            return {
                id: String(id),
                filename: cellText(as, r, 2),
                mime: cellText(as, r, 3),
                size: cellText(as, r, 4)
            }
        }
    }
    return null
}

function readAttachmentBase64(id) {
    let as = getSheet(ATT_SHEET)
    if (!as) return null
    let chunks = []
    let rows = rowCount(as)
    for (let r = 2; r <= rows; r++) {
        if (cellText(as, r, 1) == String(id)) {
            let p = parseInt(cellText(as, r, ATT_PART_COL), 10)
            chunks.push({ p: isNaN(p) ? 1 : p, data: cellText(as, r, 5) })
        }
    }
    if (!chunks.length) return null
    chunks.sort(function (a, b) { return a.p - b.p })
    let out = ""
    for (let i = 0; i < chunks.length; i++) out += chunks[i].data
    return out
}

// 列号转字母（1 -> A, 27 -> AA），用于把 WPS 的 R1C1 回退值统一成 A1 形式
function colName(n) {
    let s = ""
    while (n > 0) {
        let m = (n - 1) % 26
        s = String.fromCharCode(65 + m) + s
        n = Math.floor((n - m) / 26)
    }
    return s
}

// 当前选区对应的 key，格式 工作表名!地址（A1 形式，去掉 $ 与工作表名前缀）
function getSelectionKey() {
    let sh = getApp().ActiveSheet
    let sel = getApp().Selection
    let addr = ""
    try {
        let a = sel.Address
        if (typeof a === "function") a = a()
        addr = String(a || "")
    } catch (e) {
        addr = ""
    }
    if (!addr) {
        try {
            addr = colName(sel.Column) + sel.Row
        } catch (e) {
            addr = ""
        }
    }
    addr = addr.replace(/\$/g, "")
    if (addr.indexOf("!") >= 0) {
        addr = addr.substring(addr.lastIndexOf("!") + 1)
    }
    return sh.Name + "!" + addr
}
