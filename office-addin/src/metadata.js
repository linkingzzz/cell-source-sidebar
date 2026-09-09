// office-addin/src/metadata.js
// Helper functions to read/write __metadata sheets in the workbook

async function readMetadataFromWorkbook() {
  return await Excel.run(async context => {
    const wb = context.workbook;
    const result = { entries: [], attachments: [] };

    const metaSheet = wb.worksheets.getItemOrNullObject('__metadata');
    metaSheet.load('name');
    await context.sync();
    if (!metaSheet.isNullObject) {
      const used = metaSheet.getUsedRangeOrNullObject();
      used.load('values,rowCount');
      await context.sync();
      if (!used.isNullObject && used.rowCount >= 2) {
        const rows = used.values; // array of arrays
        for (let i = 1; i < rows.length; i++) {
          const r = rows[i];
          if (!r || !r[0]) continue;
          const entry = {
            key: r[0],
            type: r[1] || 'cell',
            source: r[2] || '',
            url: r[3] || '',
            note: r[4] || '',
            attachments: [] ,
            author: r[6] || '',
            last_modified: r[7] || '',
            version: r[8] || ''
          };
          try { entry.attachments = r[5] ? JSON.parse(r[5]) : []; } catch (e) { entry.attachments = []; }
          result.entries.push(entry);
        }
      }
    }

    const attSheet = wb.worksheets.getItemOrNullObject('__metadata_attachments');
    attSheet.load('name');
    await context.sync();
    if (!attSheet.isNullObject) {
      const used = attSheet.getUsedRangeOrNullObject();
      used.load('values,rowCount');
      await context.sync();
      if (!used.isNullObject && used.rowCount >= 2) {
        const rows = used.values;
        for (let i = 1; i < rows.length; i++) {
          const r = rows[i];
          if (!r || !r[0]) continue;
          result.attachments.push({
            id: r[0],
            filename: r[1] || '',
            mime: r[2] || '',
            size: r[3] || 0,
            base64: r[4] || '',
            uploaded_by: r[5] || '',
            uploaded_at: r[6] || ''
          });
        }
      }
    }

    return result;
  });
}

async function initMetadataSheetFromJson(metaJson) {
  return await Excel.run(async context => {
    const wb = context.workbook;
    // create or clear __metadata
    let metaSheet = wb.worksheets.getItemOrNullObject('__metadata');
    metaSheet.load('name');
    await context.sync();
    if (!metaSheet.isNullObject) {
      // clear sheet
      metaSheet.getRange().clear();
    } else {
      metaSheet = wb.worksheets.add('__metadata');
    }
    const header = ['key','type','source','url','note','attachments','author','last_modified','version'];
    const entries = [header];
    (metaJson.entries || []).forEach(e => {
      entries.push([
        e.key || '', e.type || '', e.source || '', e.url || '', e.note || '', JSON.stringify(e.attachments || []), e.author || '', e.last_modified || '', e.version || ''
      ]);
    });
    const metaRange = metaSheet.getRangeByIndexes(0,0,entries.length, header.length);
    metaRange.values = entries;

    // attachments
    let attSheet = wb.worksheets.getItemOrNullObject('__metadata_attachments');
    attSheet.load('name');
    await context.sync();
    if (!attSheet.isNullObject) {
      attSheet.getRange().clear();
    } else {
      attSheet = wb.worksheets.add('__metadata_attachments');
    }
    const attHeader = ['id','filename','mime','size','base64','uploaded_by','uploaded_at'];
    const attData = [attHeader];
    (metaJson.attachments || []).forEach(a => {
      attData.push([a.id || '', a.filename || '', a.mime || '', a.size || 0, a.base64 || '', a.uploaded_by || '', a.uploaded_at || '']);
    });
    const attRange = attSheet.getRangeByIndexes(0,0,attData.length, attHeader.length);
    attRange.values = attData;

    // optionally hide sheets (VeryHidden not possible via Office.js; mark as hidden)
    metaSheet.visibility = "Hidden";
    attSheet.visibility = "Hidden";

    await context.sync();
    return true;
  });
}

async function writeEntry(entry) {
  return await Excel.run(async context => {
    const wb = context.workbook;
    let metaSheet = wb.worksheets.getItemOrNullObject('__metadata');
    metaSheet.load('name');
    await context.sync();
    if (metaSheet.isNullObject) {
      metaSheet = wb.worksheets.add('__metadata');
      const headerRange = metaSheet.getRange('A1:I1');
      headerRange.values = [[ 'key','type','source','url','note','attachments','author','last_modified','version' ]];
    }
    // find existing
    const used = metaSheet.getUsedRangeOrNullObject();
    used.load('values,rowCount');
    await context.sync();
    const rows = used.isNullObject ? [] : used.values;
    let foundRow = -1;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === entry.key) { foundRow = i + 1; break; }
    }
    const now = new Date().toISOString();
    const attachmentsJson = JSON.stringify(entry.attachments || []);
    if (foundRow === -1) {
      const appendIndex = (rows.length === 0) ? 2 : (rows.length + 1);
      const rowRange = metaSheet.getRange('A' + appendIndex + ':I' + appendIndex);
      rowRange.values = [[ entry.key, entry.type || 'cell', entry.source || '', entry.url || '', entry.note || '', attachmentsJson, entry.author || '', now, entry.version || '' ]];
    } else {
      const rowRange = metaSheet.getRange('A' + foundRow + ':I' + foundRow);
      rowRange.values = [[ entry.key, entry.type || 'cell', entry.source || '', entry.url || '', entry.note || '', attachmentsJson, entry.author || '', now, entry.version || '' ]];
    }
    await context.sync();
    return true;
  });
}

async function writeAttachment(attId, attObj) {
  return await Excel.run(async context => {
    const wb = context.workbook;
    let attSheet = wb.worksheets.getItemOrNullObject('__metadata_attachments');
    attSheet.load('name');
    await context.sync();
    if (attSheet.isNullObject) {
      attSheet = wb.worksheets.add('__metadata_attachments');
      const headerRange = attSheet.getRange('A1:G1');
      headerRange.values = [[ 'id','filename','mime','size','base64','uploaded_by','uploaded_at' ]];
    }
    const used = attSheet.getUsedRangeOrNullObject();
    used.load('values,rowCount');
    await context.sync();
    const rows = used.isNullObject ? [] : used.values;
    const appendIndex = (rows.length === 0) ? 2 : (rows.length + 1);
    const rowRange = attSheet.getRange('A' + appendIndex + ':G' + appendIndex);
    rowRange.values = [[ attId, attObj.filename || '', attObj.mime || '', attObj.size || 0, attObj.base64 || '', attObj.uploaded_by || '', attObj.uploaded_at || new Date().toISOString() ]];
    await context.sync();
    return true;
  });
}

// export functions for use in taskpane
window.metadataApi = {
  readMetadataFromWorkbook,
  initMetadataSheetFromJson,
  writeEntry,
  writeAttachment
};
