// office-addin/src/commands.js
// Ribbon function file. The 删除侧边栏批注 button runs deleteAnnotation (ExecuteFunction).

Office.onReady(() => {
  // Requested functions are dispatched by office.js once this file has loaded.
});

async function deleteAnnotation(event) {
  try {
    const state = await window.metadataApi.readState();
    const entry = state.entries.find((e) => e.key === state.key);
    if (entry) {
      const ids = entry.attachments || [];
      for (let i = 0; i < ids.length; i++) {
        await window.metadataApi.deleteAttachment(ids[i]);
      }
      await window.metadataApi.deleteEntry(state.key);
    }
    await revealPane();
  } catch (e) {
    await revealPane();
  } finally {
    event.completed();
  }
}

// Show the pane so the result is visible; the pane re-reads the workbook on its own poll.
async function revealPane() {
  try {
    await Office.addin.showAsTaskpane();
  } catch (e) {
    // showAsTaskpane is missing in some host versions; the deletion itself already happened.
  }
}