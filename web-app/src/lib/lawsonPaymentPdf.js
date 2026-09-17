import ExcelJS from 'exceljs';
import { extractLawsonPaymentRows, createLawsonPaymentWorkbook } from './lawsonPaymentLogic.js';

// PDF text order is not reading order: group baselines before sorting horizontally.
export function paymentPdfLines(items) {
  const lines = [];
  for (const item of items.filter(item => item.str?.trim()).sort((a, b) => b.transform[5] - a.transform[5])) {
    const y = item.transform[5];
    let line = lines.find(line => Math.abs(line.y - y) <= 2);
    if (!line) {
      line = { y, items: [] };
      lines.push(line);
    }
    line.items.push(item);
  }
  return lines.map(line => line.items.sort((a, b) => a.transform[4] - b.transform[4]).map(item => item.str.trim()).join(' '));
}

export function extractLawsonPdfRows(pages) {
  // Adapt the PDF table to the existing Excel extractor so both formats share
  // bill validation, branch sorting and the same output columns.
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('PDF');
  let expectedItem = 1;
  let pending = false;
  const money = '-?(?:\\d{1,3}(?:,\\d{3})+|\\d+)\\.\\d{2}';
  const rowPattern = new RegExp(`^(\\d+)\\s+(P\\d+IAP-\\d+-)\\s+(\\d{2})/(\\d{2})/(\\d{4})\\s+(${money})(?:\\s+${money})?$`, 'i');

  for (const [index, lines] of pages.entries()) {
    if (!lines.some(line => /Item\s+No.*Invoice\s+No.*Gross\s+Amount/i.test(line))) {
      throw new Error(`อ่านตารางใน PDF หน้า ${index + 1} ไม่ได้ รองรับ PDF ข้อความจาก Lawson เท่านั้น (ไม่รองรับไฟล์สแกน)`);
    }
    let inTable = false;
    for (const line of lines) {
      if (/Item\s+No.*Invoice\s+No.*Gross\s+Amount/i.test(line)) {
        inTable = true;
        continue;
      }
      if (!inTable) continue;
      const match = line.match(rowPattern);
      if (match) {
        if (pending) throw new Error(`ไม่พบเลขบิลของรายการ ${expectedItem - 1} ใน PDF`);
        const [, item, reference, day, month, year, amount] = match;
        if (Number(item) !== expectedItem) throw new Error(`ลำดับรายการ PDF ไม่ต่อเนื่อง: ควรเป็น ${expectedItem} แต่พบ ${item}`);
        const date = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
        if (date.getUTCDate() !== Number(day) || date.getUTCMonth() !== Number(month) - 1) {
          throw new Error(`วันที่ในรายการ ${item} ไม่ถูกต้อง`);
        }
        sheet.addRow([Number(item), reference, date, Number(amount.replaceAll(',', ''))]);
        expectedItem += 1;
        pending = true;
      } else if (/^\d{6}_\d{8}$/.test(line)) {
        if (!pending) throw new Error(`พบเลขบิลที่ไม่มีรายการอ้างอิงใน PDF หน้า ${index + 1}`);
        sheet.addRow([null, line]);
        pending = false;
      } else if (/^\d+\s|IAP-|\d+_\d+/i.test(line)) {
        throw new Error(`อ่านรายการใน PDF หน้า ${index + 1} ไม่ครบ กรุณาตรวจรูปแบบรหัสสาขา วันที่ และยอดเงิน`);
      }
    }
  }
  if (pending) throw new Error(`ไม่พบเลขบิลของรายการ ${expectedItem - 1} ใน PDF`);
  return extractLawsonPaymentRows(workbook);
}

export async function generateLawsonPaymentPdfWorkbook(inputBuffer, pdfjs) {
  const task = pdfjs.getDocument({ data: new Uint8Array(inputBuffer), useSystemFonts: true, isEvalSupported: false });
  try {
    const document = await task.promise;
    const pages = [];
    for (let number = 1; number <= document.numPages; number += 1) {
      const page = await document.getPage(number);
      // Safari supports stream readers but not ReadableStream async iteration,
      // which PDF.js getTextContent() uses internally.
      const reader = page.streamTextContent().getReader();
      const items = [];
      try {
        while (true) {
          const { value, done } = await reader.read();
          if (done) break;
          items.push(...value.items);
        }
      } finally {
        reader.releaseLock();
      }
      pages.push(paymentPdfLines(items));
      page.cleanup();
    }
    const { rows, sourceSheetName } = extractLawsonPdfRows(pages);
    return await createLawsonPaymentWorkbook(rows, sourceSheetName);
  } finally {
    await task.destroy();
  }
}
