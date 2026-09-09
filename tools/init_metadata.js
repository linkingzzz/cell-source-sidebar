// tools/init_metadata.js
// Usage: node init_metadata.js <inputWorkbook.xlsx> <outputWorkbook.xlsx> <metadata.json>

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

async function main() {
  const args = process.argv.slice(2);
  if (args.length < 3) {
    console.error('Usage: node init_metadata.js <input.xlsx> <output.xlsx> <metadata.json>');
    process.exit(1);
  }
  const [inputPath, outputPath, metadataJsonPath] = args;
  if (!fs.existsSync(inputPath)) { console.error('Input workbook not found:', inputPath); process.exit(1); }
  if (!fs.existsSync(metadataJsonPath)) { console.error('Metadata JSON not found:', metadataJsonPath); process.exit(1); }
  const metaRaw = fs.readFileSync(metadataJsonPath, 'utf8');
  const meta = JSON.parse(metaRaw);
  const workbook = XLSX.readFile(inputPath, { cellStyles: true });

  // Prepare entries sheet
  const entriesHeader = ['key','type','source','url','note','attachments','author','last_modified','version'];
  const entriesData = [entriesHeader];
  (meta.entries || []).forEach(e => {
    entriesData.push([
      e.key || '',
      e.type || '',
      e.source || '',
      e.url || '',
      e.note || '',
      JSON.stringify(e.attachments || []),
      e.author || '',
      e.last_modified || '',
      e.version || ''
    ]);
  });
  const wsEntries = XLSX.utils.aoa_to_sheet(entriesData);

  // Prepare attachments sheet
  const attHeader = ['id','filename','mime','size','base64','uploaded_by','uploaded_at'];
  const attData = [attHeader];
  (meta.attachments || []).forEach(a => {
    attData.push([a.id || '', a.filename || '', a.mime || '', a.size || 0, a.base64 || '', a.uploaded_by || '', a.uploaded_at || '']);
  });
  const wsAtt = XLSX.utils.aoa_to_sheet(attData);

  // Insert or replace sheets
  workbook.SheetNames.push('__metadata');
  workbook.Sheets['__metadata'] = wsEntries;
  workbook.SheetNames.push('__metadata_attachments');
  workbook.Sheets['__metadata_attachments'] = wsAtt;

  // Write out
  XLSX.writeFile(workbook, outputPath, { bookType: 'xlsx', cellStyles: true });
  console.log('Wrote initialized workbook to', outputPath);
}

main();
