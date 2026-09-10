// 侧边栏逻辑：选区监听 + 元数据读写
var POLL_MS = 1500
var lastKey = null
var currentEntry = null

function log(msg) {
    let el = document.getElementById("log")
    if (el) el.innerText = msg + "\n" + el.innerText
}

window.onload = function () {
    try {
        let wb = getApp().ActiveWorkbook
        document.getElementById("wbName").innerText = wb ? wb.Name : "(无)"
        document.getElementById("saveBtn").onclick = onSave
        refresh()
        setInterval(refresh, POLL_MS)
        log("ready")
        if (getSheet("__selftest__")) {
            runSelfTest()
        }
    } catch (e) {
        log("init error: " + e)
    }
}

function refresh() {
    try {
        let key = getSelectionKey()
        document.getElementById("sel").innerText = key
        if (key == lastKey) return
        lastKey = key
        let meta = readMetadata()
        let entry = null
        for (let i = 0; i < meta.entries.length; i++) {
            if (meta.entries[i].key == key) {
                entry = meta.entries[i]
                break
            }
        }
        currentEntry = entry
        document.getElementById("src").value = entry ? entry.source : ""
        document.getElementById("url").value = entry ? entry.url : ""
        document.getElementById("note").value = entry ? entry.note : ""
        document.getElementById("metaCount").innerText = meta.entries.length
    } catch (e) {
        log("refresh error: " + e)
    }
}

function onSave() {
    try {
        let key = getSelectionKey()
        let entry = {
            key: key,
            type: "cell",
            source: document.getElementById("src").value,
            url: document.getElementById("url").value,
            note: document.getElementById("note").value,
            attachments: currentEntry ? (currentEntry.attachments || []) : []
        }
        let row = writeEntry(entry)
        currentEntry = entry
        log("saved " + key + " -> row " + row)
    } catch (e) {
        log("save error: " + e)
    }
}

// 仅当工作簿里存在 __selftest__ 表时运行，用于无人值守验证
function runSelfTest() {
    try {
        let meta = readMetadata()
        let selKey = getSelectionKey()
        let row = writeEntry({ key: "SELFTEST!" + new Date().toISOString(), type: "cell", source: "wps-selftest", url: "", note: selKey, attachments: [] })
        getApp().ActiveWorkbook.Save()
        document.getElementById("selftest").innerText = "SELFTEST PASS: entries=" + meta.entries.length + " wroteRow=" + row + " selection=" + getSelectionKey()
        log("selftest pass")
    } catch (e) {
        document.getElementById("selftest").innerText = "SELFTEST FAIL: " + e
        log("selftest fail: " + e)
    }
}
