// WPS 暂未内置全部枚举值，先人工定义（官方模板同做法）
var WPS_Enum = {
    msoCTPDockPositionLeft: 0,
    msoCTPDockPositionRight: 2
}

// 取当前页面所在目录，用于拼出加载项内的绝对 URL
function GetUrlPath() {
    let e = document.location.toString()
    return -1 != (e = decodeURI(e)).indexOf("/") && (e = e.substring(0, e.lastIndexOf("/"))), e
}
