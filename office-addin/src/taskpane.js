// taskpane.js - simplified MVP implementation (worksheet-based metadata)

Office.onReady(() => {
  document.getElementById('status').innerText = 'Office ready';
  startPollingSelection();
  document.getElementById('saveBtn').addEventListener('click', onSaveClicked);
  document.getElementById('fileInput').addEventListener('change', onFileSelected);
});

let lastAddress = null;
let currentEntry = null;
let pendingAttachment = null; // {filename,mime,base64,size}

function startPollingSelection() {
  setInterval(async () => {
    try {
      await Office.run(async context => {
        const range = context.workbook.getSelectedRange();
        range.load(['address', 'values']);
        await context.sync();
        const addr = range.address.replace(/\$/g, ''); // Sheet1!A1
        if (addr !== lastAddress) {
          lastAddress = addr;
          document.getElementById('selectionAddr').innerText = addr;
          await loadEntryForAddress(addr);
        }
      });
    } catch (e) {
      console.error('Polling error', e);
    }
  }, 1200);
}

async function loadEntryForAddress(address) {
  currentEntry = null;
  // Read metadata sheet
  try {
    await Excel.run(async context => {
      const wb = context.workbook;
      let metaSheet = null;
      try {
        metaSheet = wb.worksheets.getItem('__metadata');
        metaSheet.load('name');
      } catch (err) {
        metaSheet = null;
      }
      if (!metaSheet) {
        clearForm();
        return;
      }
      const used = metaSheet.getUsedRange();
      used.load('values');
      await context.sync();
      const rows = used.values; // header expected in first row
      // rows is array of arrays
      // Find a matching key in column 0
      for (let i = 1; i < rows.length; i++) {
        const key = rows[i][0];
        if (!key) continue;
        if (key === address) {
          // map columns per schema
          const entry = {
            key: key,
            type: rows[i][1] || 'cell',
            source: rows[i][2] || '',
            url: rows[i][3] || '',
            note: rows[i][4] || '',
            attachments: rows[i][5] ? JSON.parse(rows[i][5]) : [],
            author: rows[i][6] || '',
            last_modified: rows[i][7] || ''
          };
          currentEntry = entry;
          renderEntry(entry);
          return;
        }
      }
      // not found
      clearForm();
    });
  } catch (e) {
    console.error('loadEntryForAddress error', e);
    clearForm();
  }
}

function renderEntry(entry) {
  document.getElementById('inputSource').value = entry.source || '';
  document.getElementById('inputUrl').value = entry.url || '';
  document.getElementById('inputNote').value = entry.note || '';
  renderAttachments(entry.attachments || []);
}

function clearForm() {
  document.getElementById('inputSource').value = '';
  document.getElementById('inputUrl').value = '';
  document.getElementById('inputNote').value = '';
  renderAttachments([]);
}

function renderAttachments(list) {
  const container = document.getElementById('attachmentsList');
  container.innerHTML = '';
  if (!list || list.length === 0) { container.innerText = 'No attachments'; return; }
  list.forEach(id => {
    const div = document.createElement('div');
    div.innerText = id + ' '; // In a full impl we map id to filename via attachments table
    container.appendChild(div);
  });
}

async function onSaveClicked() {
  const addr = lastAddress;
  if (!addr) { alert('No selection'); return; }
  const entry = {
    key: addr,
    type: 'cell',
    source: document.getElementById('inputSource').value,
    url: document.getElementById('inputUrl').value,
    note: document.getElementById('inputNote').value,
    attachments: currentEntry ? currentEntry.attachments || [] : []
  };
  if (pendingAttachment) {
    // append new attachment id
    const attId = 'att-' + Date.now();
    entry.attachments.push(attId);
    // write attachment record into __metadata_attachments
    await writeAttachmentToSheet(attId, pendingAttachment);
    pendingAttachment = null;
    document.getElementById('fileInput').value = '';
  }
  await writeEntryToSheet(entry);
  alert('Saved metadata to workbook');
}

async function writeEntryToSheet(entry) {
  try {
    await Excel.run(async context => {
      const wb = context.workbook;
      let metaSheet = null;
      try { metaSheet = wb.worksheets.getItem('__metadata'); metaSheet.load('name'); }
      catch (e) { metaSheet = wb.worksheets.add('__metadata'); metaSheet.getRange().clear(); }
      await context.sync();
      // Ensure header exists
      const headerRange = metaSheet.getRange('A1:I1');
      headerRange.values = [[ 'key','type','source','url','note','attachments','author','last_modified','version' ]];
      // read used rows
      const used = metaSheet.getUsedRangeOrNullObject();
      used.load('rowCount, values');
      await context.sync();
      let startRow = 2;
      if (!used.isNullObject && used.rowCount >= 2) {
        startRow = used.rowCount + 1;
      }
      // Try to find existing row for key
      const allRange = metaSheet.getRange('A2:A' + (startRow));
      allRange.load('values');
      await context.sync();
      let foundRow = -1;
      for (let i = 0; i < allRange.values.length; i++) {
        if (allRange.values[i][0] === entry.key) { foundRow = i + 2; break; }
      }
      const rowIndex = foundRow === -1 ? startRow : foundRow;
      const rowRange = metaSheet.getRange('A' + rowIndex + ':I' + rowIndex);
      const now = new Date().toISOString();
      const attachmentsJson = JSON.stringify(entry.attachments || []);
      rowRange.values = [[ entry.key, entry.type, entry.source, entry.url, entry.note, attachmentsJson, '', now, '' ]];
      await context.sync();
    });
  } catch (e) {
    console.error('writeEntryToSheet error', e);
    alert('Failed to write metadata: ' + e.message);
  }
}

async function writeAttachmentToSheet(id, att) {
  try {
    await Excel.run(async context => {
      const wb = context.workbook;
      let attSheet = null;
      try { attSheet = wb.worksheets.getItem('__metadata_attachments'); attSheet.load('name'); }
      catch (e) { attSheet = wb.worksheets.add('__metadata_attachments'); }
      await context.sync();
      const used = attSheet.getUsedRangeOrNullObject();
      used.load('rowCount');
      await context.sync();
      let startRow = 2;
      if (!used.isNullObject && used.rowCount >= 2) startRow = used.rowCount + 1;
      const rowRange = attSheet.getRange('A' + startRow + ':G' + startRow);
      // columns: id, filename, mime, size, base64, uploaded_by, uploaded_at
      rowRange.values = [[id, att.filename, att.mime, att.size, att.base64, '', new Date().toISOString()]];
      await context.sync();
    });
  } catch (e) {
    console.error('writeAttachmentToSheet error', e);
    alert('Failed to write attachment: ' + e.message);
  }
}

function onFileSelected(ev) {
  const f = ev.target.files[0];
  if (!f) return;
  if (f.size > 10 * 1024 * 1024) { alert('File exceeds 10MB limit'); ev.target.value = ''; return; }
  const reader = new FileReader();
  reader.onload = function(evt) {
    const b64 = evt.target.result.split(',')[1];
    pendingAttachment = { filename: f.name, mime: f.type || 'application/octet-stream', size: f.size, base64: b64 };
    alert('Attachment ready: ' + f.name);
  };
  reader.readAsDataURL(f);
}
