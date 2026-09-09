// Updated taskpane.js to use event handler and metadata helper

Office.onReady(() => {
  document.getElementById('status').innerText = 'Office ready';
  // try to register proper selection change event; fallback to polling
  try {
    Excel.run(async context => {
      context.workbook.onSelectionChanged.add(async () => {
        try { await loadEntryForCurrentSelection(); } catch (e) { console.error('handler error', e); }
      });
      await context.sync();
    }).catch(e => {
      console.warn('Could not register selectionChanged event, falling back to polling', e);
      startPollingSelection();
    });
  } catch (e) {
    console.warn('Event registration not available, use polling', e);
    startPollingSelection();
  }
  document.getElementById('saveBtn').addEventListener('click', onSaveClicked);
  document.getElementById('fileInput').addEventListener('change', onFileSelected);
});

let lastAddress = null;
let currentEntry = null;
let pendingAttachment = null; // {filename,mime,base64,size}

function startPollingSelection() {
  setInterval(async () => {
    try { await loadEntryForCurrentSelection(); } catch (e) { /* ignore */ }
  }, 1200);
}

async function loadEntryForCurrentSelection() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(['address','values']);
      await context.sync();
      const addr = range.address.replace(/\$/g, '');
      if (addr === lastAddress) return;
      lastAddress = addr;
      document.getElementById('selectionAddr').innerText = addr;

      // read metadata from workbook
      const meta = await window.metadataApi.readMetadataFromWorkbook();
      // lookup exact cell
      let entry = meta.entries.find(e => e.key === addr);
      if (!entry) {
        // try header if first row header exists
        try {
          const header = range.worksheet.getCell(0, range.columnIndex).getText();
        } catch (e) {
          // ignore
        }
      }
      currentEntry = entry || null;
      if (entry) renderEntry(entry);
      else clearForm();
    });
  } catch (e) {
    console.error('loadEntryForCurrentSelection error', e);
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
    div.innerText = id + ' ';
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
    attachments: currentEntry ? (currentEntry.attachments || []) : []
  };
  if (pendingAttachment) {
    const attId = 'att-' + Date.now();
    entry.attachments.push(attId);
    await window.metadataApi.writeAttachment(attId, pendingAttachment);
    pendingAttachment = null;
    document.getElementById('fileInput').value = '';
  }
  await window.metadataApi.writeEntry(entry);
  currentEntry = entry;
  alert('Saved metadata to workbook');
}

async function onFileSelected(ev) {
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
