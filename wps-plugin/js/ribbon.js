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
        ShowTaskPane()
    }
    return true
}

function ShowTaskPane() {
    let tsId = window.Application.PluginStorage.getItem("taskpane_id")
    let tskpane = tsId ? window.Application.GetTaskPane(tsId) : null
    if (!tskpane) {
        tskpane = window.Application.CreateTaskPane(GetUrlPath() + "/ui/taskpane.html")
        window.Application.PluginStorage.setItem("taskpane_id", tskpane.ID)
        tskpane.DockPosition = window.Application.Enum.msoCTPDockPositionRight
        tskpane.Visible = true
    } else {
        tskpane.Visible = !tskpane.Visible
    }
}

function OnAction(control) {
    if (control.Id == "btnTogglePane") {
        ShowTaskPane()
    }
    return true
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
