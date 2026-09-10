// true 时启动即弹出任务窗格（联调/无人值守验证用）；日常使用保持 false，点功能区按钮打开
var AUTO_OPEN_TASKPANE = false

// 加载项最先执行的函数
function OnAddinLoad(ribbonUI) {
    if (typeof (window.Application.ribbonUI) != "object") {
        window.Application.ribbonUI = ribbonUI
    }
    if (typeof (window.Application.Enum) != "object") {
        window.Application.Enum = WPS_Enum
    }
    if (AUTO_OPEN_TASKPANE) {
        ShowTaskPane(true)
    }
    return true
}

// 功能区按钮的结果写进 PluginStorage，侧边栏轮询后显示（侧边栏没开时也有据可查）
function setStatus(msg) {
    try {
        window.Application.PluginStorage.setItem("css_status", new Date().toLocaleTimeString() + "  " + msg)
    } catch (e) {
    }
}

function OnAction(control) {
    if (control.Id == "btnAddNote") {
        ShowTaskPane(true)
        setStatus("已打开侧边栏，可编辑 " + safeSelectionKey())
    } else if (control.Id == "btnDelNote") {
        DeleteAnnotation()
    }
    return true
}

function safeSelectionKey() {
    try {
        return getSelectionKey()
    } catch (e) {
        return "（取当前选区失败：" + e + "）"
    }
}

// 打开/收起侧边栏；show 为 true 时保证可见
function ShowTaskPane(show) {
    let tsId = window.Application.PluginStorage.getItem("taskpane_id")
    let tskpane = tsId ? window.Application.GetTaskPane(tsId) : null
    if (!tskpane) {
        tskpane = window.Application.CreateTaskPane(GetUrlPath() + "/ui/taskpane.html")
        window.Application.PluginStorage.setItem("taskpane_id", tskpane.ID)
        tskpane.DockPosition = window.Application.Enum.msoCTPDockPositionRight
    }
    tskpane.Visible = show !== false
}

// 删除当前单元格的侧边栏批注（连同它的附件），对齐 Excel 原生「删除批注」的即时行为
function DeleteAnnotation() {
    let key = ""
    try {
        key = getSelectionKey()
        let entry = findEntry(key)
        if (!entry) {
            setStatus("无批注可删：" + key)
        } else {
            let ids = entry.attachments || []
            for (let i = 0; i < ids.length; i++) {
                deleteAttachmentRows(ids[i])
            }
            let row = deleteEntryRow(key)
            window.Application.ActiveWorkbook.Save()
            setStatus("已删除 " + key + " 的批注（原第 " + row + " 行，附件 " + ids.length + " 个）")
        }
    } catch (e) {
        setStatus("删除失败 " + key + "：" + e)
    }
    ShowTaskPane(true)
}

function GetImage(control) {
    return "images/icon.svg"
}

function OnGetEnabled(control) {
    return true
}

function OnGetVisible(control) {
    return true
}

function OnGetLabel(control) {
    return ""
}