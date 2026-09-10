// WPS(ET) 版元数据读写，沿用 Office 版的 __metadata / __metadata_attachments 两张表
var META_SHEET = "__metadata"
var ATT_SHEET = "__metadata_attachments"
var META_HEADER = ["key", "type", "source", "url", "note", "attachments", "author", "last_modified", "version"]
var ATT_HEADER = ["id", "filename", "mime", "size", "base64", "uploaded_by", "uploaded_at"]

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
    if (s) return s
    s = getApp().ActiveWorkbook.Worksheets.Add()
    s.Name = name
    for (let c = 0; c < header.length; c++) {
        s.Cells.Item(1, c + 1).Value2 = header[c]
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

function readMetadata() {
    let result = { entries: [], attachments: [] }
    let ms = getSheet(META_SHEET)
    if (ms) {
        let rows = rowCount(ms)
        for (let r = 2; r <= rows; r++) {
            let key = ms.Cells.Item(r, 1).Value2
            if (key === null || key === undefined || key === "") continue
            let atts = []
            let raw = ms.Cells.Item(r, 6).Value2
            if (raw) {
                try { atts = JSON.parse(String(raw)) } catch (e) { atts = [] }
            }
            result.entries.push({
                key: String(key),
                type: String(ms.Cells.Item(r, 2).Value2 || "cell"),
                source: String(ms.Cells.Item(r, 3).Value2 || ""),
                url: String(ms.Cells.Item(r, 4).Value2 || ""),
                note: String(ms.Cells.Item(r, 5).Value2 || ""),
                attachments: atts,
                last_modified: String(ms.Cells.Item(r, 8).Value2 || "")
            })
        }
    }
    let as = getSheet(ATT_SHEET)
    if (as) {
        let rows = rowCount(as)
        for (let r = 2; r <= rows; r++) {
            let id = as.Cells.Item(r, 1).Value2
            if (id === null || id === undefined || id === "") continue
            result.attachments.push({ id: String(id), filename: String(as.Cells.Item(r, 2).Value2 || "") })
        }
    }
    return result
}

function writeEntry(entry) {
    let ms = ensureSheet(META_SHEET, META_HEADER)
    let rows = rowCount(ms)
    let target = 0
    for (let r = 2; r <= rows; r++) {
        let k = ms.Cells.Item(r, 1).Value2
        if (k !== null && k !== undefined && String(k) == String(entry.key)) {
            target = r
            break
        }
    }
    if (target === 0) {
        target = rows < 2 ? 2 : rows + 1
    }
    let vals = [
        entry.key,
        entry.type || "cell",
        entry.source || "",
        entry.url || "",
        entry.note || "",
        JSON.stringify(entry.attachments || []),
        entry.author || "",
        new Date().toISOString(),
        entry.version || ""
    ]
    for (let c = 0; c < vals.length; c++) {
        ms.Cells.Item(target, c + 1).Value2 = vals[c]
    }
    return target
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
