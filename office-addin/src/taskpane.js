Office.onReady(() => {
  document.getElementById('status').innerText = 'Office ready. Waiting for selection...';
  // Office.js selection change events require Excel.run context. The real implementation will add selection change handler.
});

// Placeholder functions for later implementation
async function onSelectionChanged() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'values']);
      await context.sync();
      document.getElementById('content').innerText = 'Selected: ' + range.address;
    });
  } catch (e) {
    console.error(e);
    document.getElementById('content').innerText = 'Error: ' + e.message;
  }
}
