import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { paymentPdfLines, extractLawsonPdfRows, generateLawsonPaymentPdfWorkbook } from './lawsonPaymentPdf.js';
import { createLawsonPaymentWorkbook } from './lawsonPaymentLogic.js';

const header = 'Item No Invoice No. Date Gross Amount WHT Amount';
const first = '1 P3581IAP-2607- 01/09/2026 1,432.07';
const bill = '009195_69070503';

test('reads PDF text using a reader without stream async-iterator support', async () => {
  let released = false;
  let destroyed = false;
  const chunks = [header, first, bill].map((str, index) => ({
    done: false,
    value: { items: [{ str, transform: [1, 0, 0, 1, 30, 300 - index * 20] }] },
  }));
  const page = {
    streamTextContent: () => ({
      getReader: () => ({
        read: async () => chunks.shift() || { done: true },
        releaseLock: () => { released = true; },
      }),
    }),
    cleanup() {},
  };
  const pdfjs = {
    getDocument: () => ({
      promise: Promise.resolve({ numPages: 1, getPage: async () => page }),
      destroy: async () => { destroyed = true; },
    }),
  };
  const result = await generateLawsonPaymentPdfWorkbook(new ArrayBuffer(0), pdfjs);
  assert.equal(result.recordCount, 1);
  assert.equal(result.total, 1432.07);
  assert.equal(released, true);
  assert.equal(destroyed, true);
});

test('reconstructs PDF rows from unsorted text and slightly different baselines', () => {
  const item = (str, x, y) => ({ str, transform: [1, 0, 0, 1, x, y] });
  assert.deepEqual(paymentPdfLines([
    item(bill, 80, 156.91), item('288.04', 317, 171.5),
    item('P3581IAP-2607-', 80, 170.31), item('1', 30, 171.5),
    item('01/09/2026', 205, 171.5),
  ]), ['1 P3581IAP-2607- 01/09/2026 288.04', bill]);
});

test('reads across pages, preserves leading zeros and sorts branches', async () => {
  const { rows } = extractLawsonPdfRows([
    [header, first],
    ['Credit Advice Report', header, bill, '2 P0829IAP-2608- 01/09/2026 216.03 0.00', '000282_69070912'],
  ]);
  assert.deepEqual(rows.map(row => row.branch), ['0829', '3581']);
  assert.equal(rows[1].paidDate.toISOString(), '2026-09-01T00:00:00.000Z');
  const result = await createLawsonPaymentWorkbook(rows, 'PDF');
  assert.equal(result.recordCount, 2);
  assert.equal(result.total, 1648.1);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await result.blob.arrayBuffer());
  assert.equal(workbook.getWorksheet('รายการบิลที่จ่าย').getCell('E5').value, '000282_69070912');
  assert.equal(workbook.getWorksheet('สรุปตามสาขา').getCell('B7').value, 2);
});

test('rejects scanned or unreadable pages instead of silently skipping them', () => {
  assert.throws(() => extractLawsonPdfRows([[header, first, bill], []]), /หน้า 2.*ไฟล์สแกน/);
});

test('rejects missing bill numbers including at the end of the document', () => {
  assert.throws(() => extractLawsonPdfRows([[header, first]]), /ไม่พบเลขบิล/);
  assert.throws(() => extractLawsonPdfRows([[header, first, '2 P0829IAP-2608- 01/09/2026 216.03']]), /ไม่พบเลขบิล/);
});

test('rejects skipped and duplicate item numbers', () => {
  for (const number of [1, 3]) {
    assert.throws(() => extractLawsonPdfRows([[header, first, bill, `${number} P0829IAP-2608- 01/09/2026 216.03`]]), /ไม่ต่อเนื่อง/);
  }
});

test('rejects malformed amounts and invalid dates', () => {
  assert.throws(() => extractLawsonPdfRows([[header, first.replace('1,432.07', ''), bill]]), /อ่านรายการ/);
  assert.throws(() => extractLawsonPdfRows([[header, first.replace('01/09/2026', '31/02/2026'), bill]]), /วันที่/);
});
