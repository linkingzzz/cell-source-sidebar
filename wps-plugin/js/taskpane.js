// 侧边栏逻辑：选中单元格 -> 该单元格的三块批注（附件 / 9x4 明细表 / 备注）
var POLL_MS = 1500
var lastKey = null
var lastSig = null
var lastStatus = null
var currentEntry = null
var gridInputs = []

function log(msg) {
    let el = document.getElementById("log")
    if (el) el.innerText = msg + "\n" + el.innerText
}

function $(id) {
    return document.getElementById(id)
}

window.onload = function () {
    try {
        let wb = getApp().ActiveWorkbook
        $("wbName").innerText = wb ? wb.Name : "(无)"
        buildGrid()
        $("uploadBtn").onclick = function () { $("fileInput").click() }
        $("fileInput").onchange = onUpload
        $("saveBtn").onclick = onSave
        refresh(true)
        setInterval(refresh, POLL_MS)
        log("ready")
        if (getSheet("__selftest__")) runSelfTest()
    } catch (e) {
        log("init error: " + e)
    }
}

function buildGrid() {
    let tbody = $("gridBody")
    tbody.innerHTML = ""
    gridInputs = []
    for (let r = 0; r < GRID_ROWS; r++) {
        let tr = document.createElement("tr")
        let row = []
        for (let c = 0; c < GRID_COLS; c++) {
            let td = document.createElement("td")
            let inp = document.createElement("input")
            inp.type = "text"
            td.appendChild(inp)
            tr.appendChild(td)
            row.push(inp)
        }
        tbody.appendChild(tr)
        gridInputs.push(row)
    }
}

function readGridInputs() {
    let g = emptyGrid()
    for (let r = 0; r < GRID_ROWS; r++) {
        for (let c = 0; c < GRID_COLS; c++) g[r][c] = gridInputs[r][c].value
    }
    return g
}

function fillGrid(g) {
    let src = g || emptyGrid()
    for (let r = 0; r < GRID_ROWS; r++) {
        for (let c = 0; c < GRID_COLS; c++) gridInputs[r][c].value = src[r][c] || ""
    }
}

function refresh(force) {
    try {
        reportStatus()
        let key = getSelectionKey()
        let entry = findEntry(key)
        let sig = JSON.stringify(entry)
        $("sel").innerText = key
        if (!force && key == lastKey && sig == lastSig) return
        lastKey = key
        lastSig = sig
        currentEntry = entry
        $("state").innerText = entry ? "（已有批注）" : "（无批注）"
        $("note").value = entry ? entry.note : ""
        fillGrid(entry ? entry.grid : null)
        renderAttachments(entry ? entry.attachments : [])
        $("metaCount").innerText = readEntries().length
    } catch (e) {
        log("refresh error: " + e)
    }
}

// 显示功能区按钮最近一次操作的结果（在功能区点击后由 ribbon.js 写入 PluginStorage）
function reportStatus() {
    let msg = ""
    try {
        msg = getApp().PluginStorage.getItem("css_status") || ""
    } catch (e) {
        msg = ""
    }
    if (msg && msg != lastStatus) {
        lastStatus = msg
        log(msg)
    }
}

function renderAttachments(ids) {
    let box = $("attList")
    box.innerHTML = ""
    if (!ids || !ids.length) {
        box.innerText = "（暂无附件）"
        return
    }
    for (let i = 0; i < ids.length; i++) {
        let id = ids[i]
        let meta = readAttachmentMeta(id)
        let div = document.createElement("div")
        div.className = "att"
        let a = document.createElement("a")
        a.innerText = meta ? meta.filename : id
        a.title = "点击下载"
        a.onclick = makeDownloadHandler(id)
        div.appendChild(a)
        let del = document.createElement("span")
        del.className = "del"
        del.innerText = "删除"
        del.onclick = makeDeleteAttachmentHandler(id)
        div.appendChild(del)
        box.appendChild(div)
    }
}

function makeDownloadHandler(id) {
    return function () { downloadAttachment(id) }
}

function makeDeleteAttachmentHandler(id) {
    return function () { removeAttachment(id) }
}

function base64ToBytes(b64) {
    let bin = atob(b64)
    let out = new Uint8Array(bin.length)
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
    return out
}

function downloadAttachment(id) {
    try {
        let meta = readAttachmentMeta(id) || { filename: id, mime: "" }
        let b64 = readAttachmentBase64(id)
        if (b64 === null) {
            log("附件数据缺失：" + id)
            return
        }
        let blob = new Blob([base64ToBytes(b64)], { type: meta.mime || "application/octet-stream" })
        let url = URL.createObjectURL(blob)
        let a = document.createElement("a")
        a.href = url
        a.download = meta.filename || id
        document.body.appendChild(a)
        a.click()
        document.body.removeChild(a)
        setTimeout(function () { URL.revokeObjectURL(url) }, 10000)
        log("已触发下载：" + a.download)
    } catch (e) {
        log("download error: " + e)
    }
}

function onUpload(ev) {
    let input = ev.target
    let f = input.files && input.files[0]
    if (!f) return
    if (f.size > MAX_FILE_BYTES) {
        log("附件超过 10MB 上限，已取消：" + f.name)
        input.value = ""
        return
    }
    let reader = new FileReader()
    reader.onload = function () {
        try {
            let b64 = String(reader.result).split(",")[1] || ""
            let id = newAttachmentId()
            writeAttachment(id, f.name, f.type, f.size, b64)
            let entry = currentEntry || { key: lastKey, type: "cell", source: "", url: "", note: "", attachments: [] }
            entry.note = $("note").value
            entry.grid = readGridInputs()
            entry.attachments = (entry.attachments || []).concat([id])
            writeEntry(entry)
            getApp().ActiveWorkbook.Save()
            refresh(true)
            log("已上传 " + f.name + "（" + (f.size / 1024).toFixed(1) + " KB）")
        } catch (e) {
            log("upload error: " + e)
        }
        input.value = ""
    }
    reader.readAsDataURL(f)
}

function removeAttachment(id) {
    try {
        deleteAttachmentRows(id)
        if (currentEntry) {
            let rest = []
            let ids = currentEntry.attachments || []
            for (let i = 0; i < ids.length; i++) {
                if (ids[i] != id) rest.push(ids[i])
            }
            currentEntry.attachments = rest
            writeEntry(currentEntry)
            getApp().ActiveWorkbook.Save()
        }
        refresh(true)
        log("已删除附件 " + id)
    } catch (e) {
        log("del attachment error: " + e)
    }
}

function onSave() {
    try {
        let entry = {
            key: lastKey,
            type: "cell",
            source: currentEntry ? currentEntry.source : "",
            url: currentEntry ? currentEntry.url : "",
            note: $("note").value,
            attachments: currentEntry ? (currentEntry.attachments || []) : [],
            grid: readGridInputs()
        }
        let row = writeEntry(entry)
        getApp().ActiveWorkbook.Save()
        currentEntry = entry
        refresh(true)
        log("已保存 " + entry.key + "（第 " + row + " 行）")
    } catch (e) {
        log("save error: " + e)
    }
}

// 仅当工作簿里存在 __selftest__ 表时运行，用于无人值守验证
function runSelfTest() {
    try {
        let key = "SELFTEST!" + new Date().toISOString()
        let grid = emptyGrid()
        grid[0][0] = "r1c1"
        grid[8][3] = "r9c4"
        let b64 = "aGVsbG8td3BzLXNlbGZ0ZXN0"
        let aid = "att-selftest"
        deleteAttachmentRows(aid)
        writeAttachment(aid, "selftest.txt", "text/plain", 19, b64)
        let back = readAttachmentBase64(aid)
        let row = writeEntry({ key: key, type: "cell", note: "selftest", attachments: [aid], grid: grid })
        let rt = findEntry(key)
        let ok = !!rt && rt.grid[0][0] == "r1c1" && rt.grid[8][3] == "r9c4" &&
            rt.attachments.length == 1 && back == b64
        getApp().ActiveWorkbook.Save()
        $("selftest").innerText = "SELFTEST " + (ok ? "PASS" : "FAIL") +
            " grid=" + (rt ? rt.grid[0][0] + "/" + rt.grid[8][3] : "n/a") +
            " att=" + (back ? back.length + "chars" : "null") +
            " row=" + row + " sel=" + getSelectionKey()
        log("selftest " + (ok ? "pass" : "fail"))
    } catch (e) {
        $("selftest").innerText = "SELFTEST FAIL: " + e
        log("selftest fail: " + e)
    }
}
