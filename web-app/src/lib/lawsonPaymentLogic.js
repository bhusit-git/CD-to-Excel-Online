import ExcelJS from 'exceljs';

const DETAIL_SHEET_NAME = 'รายการบิลที่จ่าย';
const SUMMARY_SHEET_NAME = 'สรุปตามสาขา';
const PAYMENT_REFERENCE_PATTERN = /^P(\d+)IAP-(\d+)-/i;
const BILL_NUMBER_PATTERN = /^\d{6}_\d{8}$/;

function resolvedCellValue(cell) {
  const value = cell?.value;
  if (value && typeof value === 'object' && 'result' in value) return value.result;
  return value;
}

function cellText(cell) {
  const value = resolvedCellValue(cell);
  if (value == null) return '';
  if (typeof value === 'object' && Array.isArray(value.richText)) {
    return value.richText.map(part => part.text).join('').replaceAll('\u00a0', ' ').trim();
  }
  return String(value).replaceAll('\u00a0', ' ').trim();
}

function cellNumber(cell) {
  const value = resolvedCellValue(cell);
  if (typeof value === 'number') return value;
  const text = String(value ?? '').replaceAll(',', '').trim();
  if (!text) return NaN;
  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : NaN;
}

function findPaymentSheet(workbook) {
  const preferredNames = ['1.1', '1'];
  const candidates = [
    ...preferredNames.map(name => workbook.getWorksheet(name)).filter(Boolean),
    ...workbook.worksheets.filter(sheet => !preferredNames.includes(sheet.name)),
  ];

  return candidates.find(sheet => {
    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
      const itemNumber = cellText(sheet.getCell(rowNumber, 1));
      const reference = cellText(sheet.getCell(rowNumber, 2));
      if (/^\d+$/.test(itemNumber) && PAYMENT_REFERENCE_PATTERN.test(reference)) return true;
    }
    return false;
  });
}

export function extractLawsonPaymentRows(workbook) {
  const sourceSheet = findPaymentSheet(workbook);
  if (!sourceSheet) {
    throw new Error('ไม่พบชีตข้อมูล Lawson ที่มีรหัสสาขารูปแบบ PxxxxIAP-xxxx-');
  }

  const rows = [];
  const missingBills = [];
  for (let rowNumber = 1; rowNumber <= sourceSheet.rowCount; rowNumber += 1) {
    const originalItemText = cellText(sourceSheet.getCell(rowNumber, 1));
    const reference = cellText(sourceSheet.getCell(rowNumber, 2));
    const match = reference.match(PAYMENT_REFERENCE_PATTERN);
    if (!/^\d+$/.test(originalItemText) || !match) continue;

    const billNumber = cellText(sourceSheet.getCell(rowNumber + 1, 2));
    if (!BILL_NUMBER_PATTERN.test(billNumber)) {
      missingBills.push(originalItemText);
      continue;
    }

    const paidAmount = cellNumber(sourceSheet.getCell(rowNumber, 4));
    if (!Number.isFinite(paidAmount)) {
      throw new Error(`ยอดจ่ายของรายการลำดับ ${originalItemText} ไม่ใช่ตัวเลข`);
    }

    rows.push({
      originalItem: Number(originalItemText),
      branch: match[1].padStart(4, '0'),
      cycle: match[2],
      reference,
      billNumber,
      paidDate: resolvedCellValue(sourceSheet.getCell(rowNumber, 3)),
      paidAmount,
    });
  }

  if (missingBills.length > 0) {
    throw new Error(`ไม่พบเลขบิลในบรรทัดถัดไปของรายการ: ${missingBills.join(', ')}`);
  }
  if (rows.length === 0) {
    throw new Error('ไม่พบรายการบิลที่ Lawson แจ้งจ่าย');
  }

  const billCollator = new Intl.Collator('en', { numeric: true });
  rows.sort((left, right) => (
    Number(left.branch) - Number(right.branch)
    || left.cycle.localeCompare(right.cycle, 'en')
    || billCollator.compare(left.billNumber, right.billNumber)
    || left.originalItem - right.originalItem
  ));

  return { sourceSheetName: sourceSheet.name, rows };
}

function styleTitle(sheet, title, note, columnCount) {
  sheet.mergeCells(1, 1, 1, columnCount);
  sheet.getCell(1, 1).value = title;
  sheet.getCell(1, 1).font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FF1F2937' } };
  sheet.getRow(1).height = 26;

  sheet.mergeCells(2, 1, 2, columnCount);
  sheet.getCell(2, 1).value = note;
  sheet.getCell(2, 1).font = { name: 'Arial', size: 10, italic: true, color: { argb: 'FF5B6472' } };
  sheet.getCell(2, 1).alignment = { vertical: 'middle', wrapText: true };
  sheet.getRow(2).height = 30;
}

function styleHeader(row) {
  row.height = 30;
  row.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
    cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });
}

function writeDetailSheet(workbook, rows) {
  const sheet = workbook.addWorksheet(DETAIL_SHEET_NAME, { views: [{ state: 'frozen', ySplit: 4 }] });
  sheet.properties.showGridLines = false;
  styleTitle(
    sheet,
    'รายการบิลที่ Lawson แจ้งจ่าย',
    'เรียงตามสาขา โดยนำเลขบิลจากบรรทัดถัดจากรหัสสาขาในไฟล์ต้นฉบับ',
    8,
  );

  const headers = ['ลำดับ', 'สาขา', 'รอบเอกสาร', 'รหัสอ้างอิง Lawson', 'เลขบิล', 'วันที่จ่าย', 'ยอดจ่าย', 'ลำดับในไฟล์เดิม'];
  sheet.getRow(4).values = headers;
  styleHeader(sheet.getRow(4));

  rows.forEach((payment, index) => {
    const row = sheet.addRow([
      index + 1,
      Number(payment.branch),
      payment.cycle,
      payment.reference,
      payment.billNumber,
      payment.paidDate,
      payment.paidAmount,
      payment.originalItem,
    ]);
    row.font = { name: 'Arial', size: 10 };
    row.getCell(2).numFmt = '0000';
    row.getCell(3).numFmt = '@';
    row.getCell(4).numFmt = '@';
    row.getCell(5).numFmt = '@';
    row.getCell(6).numFmt = 'dd/mm/yyyy';
    row.getCell(7).numFmt = '#,##0.00';
    [1, 2, 3, 6, 8].forEach(column => {
      row.getCell(column).alignment = { horizontal: 'center', vertical: 'middle' };
    });
  });

  sheet.columns = [
    { width: 9 }, { width: 10 }, { width: 13 }, { width: 23 },
    { width: 21 }, { width: 14 }, { width: 14 }, { width: 17 },
  ];
  sheet.autoFilter = { from: 'A4', to: `H${rows.length + 4}` };
  return sheet;
}

function writeSummarySheet(workbook, rows) {
  const sheet = workbook.addWorksheet(SUMMARY_SHEET_NAME, { views: [{ state: 'frozen', ySplit: 4 }] });
  sheet.properties.showGridLines = false;
  styleTitle(
    sheet,
    'สรุปการจ่ายตามสาขา',
    'ไฟล์ Lawson แสดงเฉพาะบิลที่แจ้งว่าจ่ายแล้ว สถานะครบ/ไม่ครบต้องตรวจเทียบกับรายการบิลที่ควรได้รับ',
    5,
  );

  sheet.getRow(4).values = ['สาขา', 'จำนวนบิลที่จ่าย', 'ยอดรวมที่จ่าย', 'สถานะตรวจ', 'หมายเหตุ'];
  styleHeader(sheet.getRow(4));

  const byBranch = new Map();
  rows.forEach(payment => {
    const summary = byBranch.get(payment.branch) || { count: 0, total: 0 };
    summary.count += 1;
    summary.total += payment.paidAmount;
    byBranch.set(payment.branch, summary);
  });

  [...byBranch.entries()]
    .sort(([left], [right]) => Number(left) - Number(right))
    .forEach(([branch, summary]) => {
      const row = sheet.addRow([Number(branch), summary.count, Math.round((summary.total + Number.EPSILON) * 100) / 100, '', '']);
      row.font = { name: 'Arial', size: 10 };
      row.getCell(1).numFmt = '0000';
      row.getCell(1).alignment = { horizontal: 'center' };
      row.getCell(2).alignment = { horizontal: 'center' };
      row.getCell(3).numFmt = '#,##0.00';
      row.getCell(4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
      row.getCell(4).dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['"รอตรวจ,ครบ,ไม่ครบ"'],
      };
    });

  const total = Math.round((rows.reduce((sum, payment) => sum + payment.paidAmount, 0) + Number.EPSILON) * 100) / 100;
  const totalRow = sheet.addRow(['รวม', rows.length, total, '', '']);
  totalRow.font = { name: 'Arial', size: 10, bold: true };
  totalRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9EAF7' } };
  totalRow.getCell(3).numFmt = '#,##0.00';

  sheet.columns = [{ width: 11 }, { width: 18 }, { width: 19 }, { width: 16 }, { width: 28 }];
  sheet.autoFilter = { from: 'A4', to: `E${sheet.rowCount - 1}` };
  return { branchCount: byBranch.size, total };
}

export async function generateLawsonPaymentWorkbook(inputBuffer) {
  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.load(inputBuffer);
  const { sourceSheetName, rows } = extractLawsonPaymentRows(sourceWorkbook);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Lawson Billing Web App';
  workbook.created = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;
  writeDetailSheet(workbook, rows);
  const { branchCount, total } = writeSummarySheet(workbook, rows);

  const outputBuffer = await workbook.xlsx.writeBuffer();
  return {
    blob: new Blob([outputBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    recordCount: rows.length,
    branchCount,
    total,
    sourceSheetName,
  };
}
