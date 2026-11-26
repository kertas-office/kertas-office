// PDF to Word or to a workbook, in a process of its own (started by writer2.cjs), so a long conversion never
// freezes the window. One message in: the PDF's bytes and where the PDFium engine file is.
// One message out: the new file's bytes (as `docx`, whichever kind it is), or what went wrong. Progress is reported per page.
const path = require('path');

process.parentPort.on('message', async ({ data }) => {
  try {
    const lib = require(path.join(__dirname, 'writer2', 'main', 'pdf2docx.cjs'));
    const convert = data.format === 'xlsx' ? lib.convertPdfBytesToXlsx : lib.convertPdfBytesToDocx;
    const docx = await convert(new Uint8Array(data.pdf), data.wasm, (page, total) => process.parentPort.postMessage({ page, total }));
    process.parentPort.postMessage({ docx });
  } catch (e) {
    process.parentPort.postMessage({ error: String((e && e.message) || e) });
  }
});
