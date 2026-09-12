// office-addin/src/taskpane.js
// Sidebar for the selected cell: attachments / 9x4 detail table / note.
// Mirrors the WPS add-in UI and storage contract.

const POLL_MS = 2000;

let lastKey = null;
let lastSig = null;
let currentEntry = null;
let gridInputs = [];
let attMeta = {};
let busy = false;

function $(id) {
  return document.getElementById(id);
}

function log(msg) {
  const el = $("log");
  if (el) el.innerText = msg + "\n" + el.innerText;
}

// 功能区「隐藏侧边栏批注」：注册在共享运行时里的动作（见 manifest.xml 的 <Runtimes>）。
async function hideAnnotation(event) {
  try {
    await Office.addin.hide();
  } catch (e) {
    log("hide error: " + e);
  } finally {
    event.completed();
  }
}

Office.onReady((info) => {
  $("status").innerText = "Office ready";
  $("wbName").innerText = info && info.host ? String(info.host) : "";
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("hideAnnotation", hideAnnotation);
  } else {
    log("Office.actions 不可用，隐藏按钮将无效");
  }
  buildGrid();
  $("uploadBtn").addEventListener("click", () => $("fileInput").click());
  $("fileInput").addEventListener("change", onUpload);
  $("saveBtn").addEventListener("click", onSave);
  refresh(true).catch((e) => log("init error: " + e));
  // Polling covers both selection changes and edits made elsewhere (e.g. another surface).
  setInterval(() => refresh(false).catch((e) => log("poll error: " + e)), POLL_MS);
});

function buildGrid() {
  const tbody = $("gridBody");
  tbody.innerHTML = "";
  gridInputs = [];
  for (let r = 0; r < window.metadataApi.GRID_ROWS; r++) {
    const tr = document.createElement("tr");
    const row = [];
    for (let c = 0; c < window.metadataApi.GRID_COLS; c++) {
      const td = document.createElement("td");
      const inp = document.createElement("input");
      inp.type = "text";
      td.appendChild(inp);
      tr.appendChild(td);
      row.push(inp);
    }
    tbody.appendChild(tr);
    gridInputs.push(row);
  }
}

function readGridInputs() {
  const g = window.metadataApi.emptyGrid();
  for (let r = 0; r < g.length; r++) {
    for (let c = 0; c < g[r].length; c++) g[r][c] = gridInputs[r][c].value;
  }
  return g;
}

function fillGrid(grid) {
  const src = grid || window.metadataApi.emptyGrid();
  for (let r = 0; r < gridInputs.length; r++) {
    for (let c = 0; c < gridInputs[r].length; c++) gridInputs[r][c].value = src[r][c] || "";
  }
}

async function refresh(force) {
  if (busy) return;
  busy = true;
  try {
    const state = await window.metadataApi.readState();
    const entry = state.entries.find((e) => e.key === state.key) || null;
    const sig = JSON.stringify(entry);
    $("sel").innerText = state.key;
    if (!force && state.key === lastKey && sig === lastSig) return;
    lastKey = state.key;
    lastSig = sig;
    currentEntry = entry;
    attMeta = {};
    state.attachments.forEach((a) => { attMeta[a.id] = a; });
    $("state").innerText = entry ? "（已有批注）" : "（无批注）";
    $("metaCount").innerText = state.entries.length;
    $("note").value = entry ? entry.note : "";
    fillGrid(entry ? entry.grid : null);
    renderAttachments(entry ? entry.attachments : []);
  } finally {
    busy = false;
  }
}

function renderAttachments(ids) {
  const box = $("attList");
  box.innerHTML = "";
  if (!ids || !ids.length) {
    box.innerText = "（暂无附件）";
    return;
  }
  ids.forEach((id) => {
    const meta = attMeta[id];
    const div = document.createElement("div");
    div.className = "att";
    const a = document.createElement("a");
    a.innerText = meta ? meta.filename || id : id;
    a.title = "点击下载";
    a.addEventListener("click", () => downloadAttachment(id));
    div.appendChild(a);
    const del = document.createElement("span");
    del.className = "del";
    del.innerText = "删除";
    del.addEventListener("click", () => removeAttachment(id));
    div.appendChild(del);
    box.appendChild(div);
  });
}

function base64ToBytes(b64) {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

async function downloadAttachment(id) {
  try {
    const att = await window.metadataApi.readAttachment(id);
    if (!att) {
      log("附件数据缺失：" + id);
      return;
    }
    const blob = new Blob([base64ToBytes(att.base64)], { type: att.mime || "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = att.filename || id;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 10000);
    log("已触发下载：" + a.download);
  } catch (e) {
    log("download error: " + e);
  }
}

function onUpload(ev) {
  const input = ev.target;
  const f = input.files && input.files[0];
  if (!f) return;
  if (f.size > window.metadataApi.MAX_FILE_BYTES) {
    log("附件超过 10MB 上限，已取消：" + f.name);
    input.value = "";
    return;
  }
  const reader = new FileReader();
  reader.onload = async () => {
    try {
      const b64 = String(reader.result).split(",")[1] || "";
      const id = "att-" + Date.now().toString(36) + "-" + Math.floor(Math.random() * 1679616).toString(36);
      await window.metadataApi.writeAttachment(id, { filename: f.name, mime: f.type, size: f.size, base64: b64 });
      const entry = currentEntry || { key: lastKey, type: "cell", source: "", url: "", note: "", attachments: [] };
      entry.note = $("note").value;
      entry.grid = readGridInputs();
      entry.attachments = (entry.attachments || []).concat([id]);
      await window.metadataApi.writeEntry(entry);
      await refresh(true);
      log("已上传 " + f.name + "（" + (f.size / 1024).toFixed(1) + " KB）");
    } catch (e) {
      log("upload error: " + e);
    }
    input.value = "";
  };
  reader.readAsDataURL(f);
}

async function removeAttachment(id) {
  try {
    await window.metadataApi.deleteAttachment(id);
    if (currentEntry) {
      currentEntry.attachments = (currentEntry.attachments || []).filter((x) => x !== id);
      currentEntry.note = $("note").value;
      currentEntry.grid = readGridInputs();
      await window.metadataApi.writeEntry(currentEntry);
    }
    await refresh(true);
    log("已删除附件 " + id);
  } catch (e) {
    log("del attachment error: " + e);
  }
}

async function onSave() {
  try {
    const entry = {
      key: lastKey,
      type: "cell",
      source: currentEntry ? currentEntry.source : "",
      url: currentEntry ? currentEntry.url : "",
      note: $("note").value,
      attachments: currentEntry ? currentEntry.attachments || [] : [],
      grid: readGridInputs()
    };
    const row = await window.metadataApi.writeEntry(entry);
    currentEntry = entry;
    await refresh(true);
    log("已保存 " + entry.key + "（第 " + row + " 行）");
  } catch (e) {
    log("save error: " + e);
  }
}
