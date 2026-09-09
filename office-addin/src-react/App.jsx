import React, { useEffect, useState } from 'react'

export default function App() {
  const [selection, setSelection] = useState('');
  const [entry, setEntry] = useState({ source: '', url: '', note: '', attachments: [] });

  useEffect(() => {
    // Minimal placeholder: real Office integration will be added
    setSelection('Sheet1!B3')
  }, [])

  const onSave = () => {
    alert('Save triggered (React skeleton)')
  }

  return (
    <div style={{ padding: 12, fontFamily: 'Arial, sans-serif' }}>
      <h3>Cell Source Sidebar (React - MVP)</h3>
      <div>Selection: {selection}</div>
      <label style={{ display: 'block', marginTop: 8 }}>Source:
        <input value={entry.source} onChange={e => setEntry({ ...entry, source: e.target.value })} style={{ width: '100%' }} />
      </label>
      <label style={{ display: 'block', marginTop: 8 }}>URL:
        <input value={entry.url} onChange={e => setEntry({ ...entry, url: e.target.value })} style={{ width: '100%' }} />
      </label>
      <label style={{ display: 'block', marginTop: 8 }}>Note:
        <textarea value={entry.note} onChange={e => setEntry({ ...entry, note: e.target.value })} style={{ width: '100%', height: 120 }} />
      </label>
      <div style={{ marginTop: 8 }}>
        <input type="file" />
      </div>
      <div style={{ marginTop: 12 }}>
        <button onClick={onSave}>Save to workbook</button>
      </div>
    </div>
  )
}
