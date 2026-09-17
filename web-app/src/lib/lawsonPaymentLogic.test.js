import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { extractLawsonPaymentRows, generateLawsonPaymentWorkbook } from './lawsonPaymentLogic.js';

function makeSourceWorkbook() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('1.1');
  sheet.addRow(['Item No', 'Invoice No.', 'Date', 'Gross Amount']);
  sheet.addRow([2, 'P3581IAP-2607-', new Date(2026, 0, 9), 432.07]);
  sheet.addRow([null, '009195_69070503']);
  sheet.addRow([1, 'P0829IAP-2608-', new Date(2026, 0, 9), 216.03]);
  sheet.addRow([null, '000282_69070912']);
  return workbook;
}

test('extracts the bill from the row below and sorts by branch', () => {
  const { sourceSheetName, rows } = extractLawsonPaymentRows(makeSourceWorkbook());
  assert.equal(sourceSheetName, '1.1');
  assert.equal(rows.length, 2);
  assert.deepEqual(rows.map(row => [row.branch, row.billNumber]), [
    ['0829', '000282_69070912'],
    ['3581', '009195_69070503'],
  ]);
});

test('rejects a payment row when its following bill number is missing', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('1');
  sheet.addRow(['Item No', 'Invoice No.', 'Date', 'Gross Amount']);
  sheet.addRow([1, 'P3581IAP-2607-', new Date(2026, 0, 9), 432.07]);
  assert.throws(() => extractLawsonPaymentRows(workbook), /ไม่พบเลขบิล/);
});

test('rejects a payment row when the following row is another payment', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('1');
  sheet.addRow(['Item No', 'Invoice No.', 'Date', 'Gross Amount']);
  sheet.addRow([1, 'P0829IAP-2607-', new Date(2026, 0, 9), 216.03]);
  sheet.addRow([2, 'P3581IAP-2607-', new Date(2026, 0, 9), 432.07]);
  sheet.addRow([null, '009195_69070503']);
  assert.throws(() => extractLawsonPaymentRows(workbook), /ไม่พบเลขบิล.*1/);
});

test('rejects a payment row when its paid amount is blank', () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('1');
  sheet.addRow(['Item No', 'Invoice No.', 'Date', 'Gross Amount']);
  sheet.addRow([1, 'P3581IAP-2607-', new Date(2026, 0, 9), null]);
  sheet.addRow([null, '009195_69070503']);
  assert.throws(() => extractLawsonPaymentRows(workbook), /ยอดจ่าย.*ไม่ใช่ตัวเลข/);
});

test('generates detail and branch summary sheets', async () => {
  const input = await makeSourceWorkbook().xlsx.writeBuffer();
  const result = await generateLawsonPaymentWorkbook(input);
  assert.equal(result.recordCount, 2);
  assert.equal(result.branchCount, 2);
  assert.equal(result.total, 648.1);

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(await result.blob.arrayBuffer());
  assert.equal(output.getWorksheet('รายการบิลที่จ่าย').getCell('E5').value, '000282_69070912');
  assert.equal(output.getWorksheet('สรุปตามสาขา').getCell('B7').value, 2);
});
