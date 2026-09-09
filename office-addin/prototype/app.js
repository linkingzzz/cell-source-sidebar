// app.js - simple prototype that loads examples/sample-metadata.json and shows entry for Sheet1!B3

async function loadSample() {
  const res = await fetch('/cell-source-sidebar/examples/sample-metadata.json');
  const json = await res.json();
  return json;
}

document.addEventListener('DOMContentLoaded', async () => {
  const meta = await loadSample();
  const entry = meta.entries.find(e => e.key === 'Sheet1!B3');
  if (entry) {
    document.getElementById('src').value = entry.source || '';
    document.getElementById('url').value = entry.url || '';
    document.getElementById('note').value = entry.note || '';
    document.getElementById('attList').innerText = (entry.attachments || []).join(', ');
  }
  document.getElementById('save').addEventListener('click', () => {
    document.getElementById('log').innerText = 'Simulated save: ' + JSON.stringify({ source: document.getElementById('src').value, url: document.getElementById('url').value, note: document.getElementById('note').value }, null, 2);
  });
});
