import ExcelJS from 'exceljs';

const SMALL_PRICE = 60.0;
const SMALL_SUMMARY_BILLS = new Set(["69 040870", "69 040871", "69 040872", "69 040873"]);

const THAI_MONTHS = [
    "",
    "มกราคม",
    "กุมภาพันธ์",
    "มีนาคม",
    "เมษายน",
    "พฤษภาคม",
    "มิถุนายน",
    "กรกฎาคม",
    "สิงหาคม",
    "กันยายน",
    "ตุลาคม",
    "พฤศจิกายน",
    "ธันวาคม",
];

const THAI_DIGITS = ["ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"];
const THAI_POSITIONS = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน"];

const SHEET_CONFIGS = [
    {
        name: "ลอสันหลอดใหญ่",
        kind: "lawson_big",
        bill_no: "6905001",
        customer_line: "ชื่อลูกค้า : บริษัท สห ลอว์สัน จำกัด",
        address_line: "เลขที่ 2170 อาคารกรุงเทพทาวเวอร์ ชั้น 3 ถนนเพชรบุรีตัดใหม่ แขวงบางกะปิ เขตห้วยขวาง กรุงเทพมหานคร 10310",
        tax_line: "เลขประจำตัวผู้เสียภาษี 105-55516-6337",
        product_code: "04",
        branch_ids: [
            "0829", "1089", "1360", "2880", "3089", "3204", "3208", "3255", "3534",
            "3548", "3554", "3555", "3572", "3575", "3581", "3639", "3666", "3667",
            "3674", "3675", "3676", "3709", "3728", "3733", "3735", "3747", "3750",
            "3754", "3758", "3760", "3766", "3767", "3775", "3793", "3794", "3807",
            "3814",
        ],
        title: "รายการ : น้ำแข็งหลอดใหญ่แพ็ค 1.4 KG (ห่อ)",
    },
    {
        name: "แฟรนไชส์(ประเวศ)",
        kind: "franchise_big",
        bill_no: "6905002",
        customer_line: "ชื่อลูกค้า : บริษัท ซัสโก้มาร์เก็ตติ้ง จำกัด (สาขาที่ 00025 ประเวศ)",
        address_line: "ที่อยู่ : 161 ถ.มอเตอร์เวย์ แขวงทับช้าง เขตสะพานสูง กรุงเทพ 10250",
        tax_line: "เลขประจำผู้เสียภาษี : 0105539021915",
        product_code: "04",
        branch_ids: ["3653"],
    },
    {
        name: "แฟรนไชส์หัวหมาก",
        kind: "franchise_big",
        bill_no: "6905003",
        customer_line: "ชื่อลูกค้า : บริษัท ซัสโก้มาร์เก็ตติ้ง จำกัด (สาขาที่ 00032 หัวหมาก)",
        address_line: "",
        tax_line: "",
        product_code: "04",
        branch_ids: ["3658"],
    },
    {
        name: "แฟรนไชส์สวนหลวง",
        kind: "franchise_big",
        bill_no: "6905004",
        customer_line: "ชื่อลูกค้า : บริษัท ซัสโก้มาร์เก็ตติ้ง จำกัด (สาขาที่ 00019 สวนหลวง)",
        address_line: "",
        tax_line: "",
        product_code: "04",
        branch_ids: ["3645"],
    },
    {
        name: "แฟรนไชสัลโก้บางพลี",
        kind: "franchise_big",
        bill_no: "6905005",
        customer_line: "ชื่อลูกค้า : บริษัท ซัสโก้มาร์เก็ตติ้ง จำกัด (สาขาที 00031 บางพลี)",
        address_line: "",
        tax_line: "",
        product_code: "04",
        branch_ids: ["3663"],
    },
    {
        name: "แฟรนไชสัสโก้บางบ่อ",
        kind: "franchise_big",
        bill_no: "6905006",
        customer_line: "ชื่อลูกค้า : บริษัท ซัสโก้มาร์เก็ตติ้ง จำกัด (สาขาที 00038 บางบ่อ)",
        address_line: "",
        tax_line: "",
        product_code: "04",
        branch_ids: ["3719"],
    },
    {
        name: "แฟรนไชสัสโก้สะพานสูง",
        kind: "franchise_big",
        bill_no: "6905007",
        customer_line: "ชื่อลูกค้า : บริษัท ซัสโก้มาร์เก็ตติ้ง จำกัด (สาขาที่ 00048 ซัสโก้สะพานสูง )",
        address_line: "",
        tax_line: "",
        product_code: "04",
        branch_ids: ["3730"],
    },
    {
        name: "แฟรนไชสัสโก้คลองสองต้นนุ่น",
        kind: "franchise_big",
        bill_no: "6905008",
        customer_line: "ชื่อลูกค้า : บริษัท ซัสโก้มาร์เก็ตติ้ง จำกัด (สาขาที 00041 ซัสโก้คลองสองต้นนุ่น)",
        address_line: "",
        tax_line: "",
        product_code: "04",
        branch_ids: ["3703"],
    },
    {
        name: "ลอสันหลอดเล็ก",
        kind: "lawson_small_combined",
        bill_no: "6905009",
        customer_line: "ชื่อลูกค้า : บริษัท สห ลอว์สัน จำกัด",
        address_line: "เลขที่ 2170 อาคารกรุงเทพทาวเวอร์ ชั้น 3 ถนนเพชรบุรีตัดใหม่ แขวงบางกะปิ เขตห้วยขวาง กรุงเทพมหานคร 10310",
        tax_line: "เลขประจำตัวผู้เสียภาษี 105-55516-6337",
        product_code: "06",
        branch_ids: ["3666", "3639", "3760", "3735", "3766"],
        title: "รายการ : น้ำแข็งหลอดเล็ก",
    },
    {
        name: "3666 (ล.)",
        kind: "lawson_small_single",
        bill_no: "6905010",
        customer_line: "ชื่อลูกค้า : บริษัท สห ลอว์สัน จำกัด",
        address_line: "เลขที่ 2170 อาคารกรุงเทพทาวเวอร์ ชั้น 3 ถนนเพชรบุรีตัดใหม่ แขวงบางกะปิ เขตห้วยขวาง กรุงเทพมหานคร 10310",
        tax_line: "เลขประจำตัวผู้เสียภาษี 105-55516-6337",
        product_code: "06",
        branch_ids: ["3666"],
        title: "รายการ : น้ำแข็งหลอดเล็ก",
    },
    {
        name: "3639 (ล.)",
        kind: "lawson_small_single",
        bill_no: "6905011",
        customer_line: "ชื่อลูกค้า : บริษัท สห ลอว์สัน จำกัด",
        address_line: "เลขที่ 2170 อาคารกรุงเทพทาวเวอร์ ชั้น 3 ถนนเพชรบุรีตัดใหม่ แขวงบางกะปิ เขตห้วยขวาง กรุงเทพมหานคร 10310",
        tax_line: "เลขประจำตัวผู้เสียภาษี 105-55516-6337",
        product_code: "06",
        branch_ids: ["3639"],
        title: "รายการ : น้ำแข็งหลอดเล็ก",
    },
    {
        name: "3760 (ล.)",
        kind: "lawson_small_single",
        bill_no: "6905012",
        customer_line: "ชื่อลูกค้า : บริษัท สห ลอว์สัน จำกัด",
        address_line: "เลขที่ 2170 อาคารกรุงเทพทาวเวอร์ ชั้น 3 ถนนเพชรบุรีตัดใหม่ แขวงบางกะปิ เขตห้วยขวาง กรุงเทพมหานคร 10310",
        tax_line: "เลขประจำตัวผู้เสียภาษี 105-55516-6337",
        product_code: "06",
        branch_ids: ["3760"],
        title: "รายการ : น้ำแข็งหลอดเล็ก",
    },
    {
        name: "3735 (ล.)",
        kind: "lawson_small_single",
        bill_no: "6905013",
        customer_line: "ชื่อลูกค้า : บริษัท สห ลอว์สัน จำกัด",
        address_line: "เลขที่ 2170 อาคารกรุงเทพทาวเวอร์ ชั้น 3 ถนนเพชรบุรีตัดใหม่ แขวงบางกะปิ เขตห้วยขวาง กรุงเทพมหานคร 10310",
        tax_line: "เลขประจำตัวผู้เสียภาษี 105-55516-6337",
        product_code: "06",
        branch_ids: ["3735"],
        title: "รายการ : น้ำแข็งหลอดเล็ก",
    },
];

export function getSheetConfigs() {
    return SHEET_CONFIGS;
}

export function formatThaiBillDate(dateObj) {
    if (!dateObj) return "";
    return `วันที่ ${dateObj.getDate()} ${THAI_MONTHS[dateObj.getMonth() + 1]} ${dateObj.getFullYear() + 543}`;
}

export function monthBounds(monthText) {
    const parts = monthText.split("-");
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const startDay = new Date(year, month, 1);
    const startStr = `${year}${String(month + 1).padStart(2, "0")}01`;
    const lastDay = new Date(year, month + 1, 0);
    const endStr = `${year}${String(month + 1).padStart(2, "0")}${String(lastDay.getDate()).padStart(2, "0")}`;
    return [startStr, endStr];
}

export function parseBillNumber(raw) {
    raw = (raw || "").replace(/\s/g, "");
    if (!raw) return "";
    const val = Number(raw);
    if (!isNaN(val)) {
        return Math.round(val).toString();
    }
    return raw;
}

export function numericKey(raw) {
    const digits = String(raw).replace(/\\D/g, "");
    return digits ? parseInt(digits, 10) : 0;
}

export function parseYmd(raw) {
    raw = (raw || "").trim();
    if (raw.length !== 8) return new Date();
    const year = parseInt(raw.substring(0, 4), 10);
    const month = parseInt(raw.substring(4, 6), 10) - 1;
    const day = parseInt(raw.substring(6, 8), 10);
    return new Date(year, month, day);
}

export function cleanBranchName(raw) {
    raw = (raw || "").trim();
    return raw.replace(/^P\\d+\\s*สาขา\\s*/, "");
}

export function thaiBahtText(amount) {
    const integer = Math.round(amount);

    function readUnderMillion(num) {
        if (num === 0) return "";
        const parts = [];
        const digits = String(num).split("").map(Number);
        const size = digits.length;
        for (let idx = 0; idx < size; idx++) {
            const digit = digits[idx];
            const pos = size - idx - 1;
            if (digit === 0) continue;
            if (pos === 0) {
                if (digit === 1 && size > 1) {
                    parts.push("เอ็ด");
                } else {
                    parts.push(THAI_DIGITS[digit]);
                }
            } else if (pos === 1) {
                if (digit === 1) {
                    parts.push("สิบ");
                } else if (digit === 2) {
                    parts.push("ยี่สิบ");
                } else {
                    parts.push(THAI_DIGITS[digit] + "สิบ");
                }
            } else {
                parts.push(THAI_DIGITS[digit] + THAI_POSITIONS[pos]);
            }
        }
        return parts.join("");
    }

    if (integer === 0) return "ศูนย์บาทถ้วน";

    const chunks = [];
    let temp = integer;
    while (temp > 0) {
        chunks.push(temp % 1000000);
        temp = Math.floor(temp / 1000000);
    }

    const words = [];
    for (let idx = chunks.length - 1; idx >= 0; idx--) {
        const chunk = chunks[idx];
        if (chunk === 0) continue;
        const chunkText = readUnderMillion(chunk);
        if (idx > 0) {
            words.push(chunkText + "ล้าน");
        } else {
            words.push(chunkText);
        }
    }
    return words.join("") + "บาทถ้วน";
}

export function buildTransRows(customersList, transRows, branchIds, productCode, periodStart, periodEnd) {
    const customers = {};
    customersList.forEach(c => customers[c.ID] = c);

    const branchSet = new Set(branchIds || []);
    const rows = [];

    for (const row of transRows) {
        const custId = row.CUST_VEND || "";
        if (branchSet.size > 0 && !branchSet.has(custId)) continue;
        if (row.PROD_CODE !== productCode) continue;
        const dateRaw = row.DATE || "";
        if (dateRaw < periodStart || dateRaw > periodEnd) continue;

        const cust = customers[custId] || {};
        const total = parseFloat(row.AMT || 0);
        const amount = Math.round((total / 1.07) * 1e12) / 1e12; // JS floating point workaround
        const vat = Math.round((total - amount) * 1e12) / 1e12;
        const branchName = cleanBranchName(cust.SH_NAME || cust.SHIP_NAME || cust.NAME || "");

        rows.push({
            dateObj: parseYmd(dateRaw),
            date: dateRaw, // For sorting consistency
            branch: `P${custId}`,
            bill_no: parseBillNumber(row.BILL_NO),
            branch_name: branchName,
            qty: parseFloat(row.QTY || 0),
            price: parseFloat(row.PRICE || 0),
            amount: amount,
            vat: vat,
            total: total,
            product_name: productCode === "04" ? "หลอดใหญ่" : branchName,
        });
    }

    // Sort by branch_name, date, branch, bill_no
    rows.sort((a, b) => {
        if (a.branch_name !== b.branch_name) return a.branch_name.localeCompare(b.branch_name);
        if (a.date !== b.date) return a.date.localeCompare(b.date);
        if (a.branch !== b.branch) return a.branch.localeCompare(b.branch);
        return numericKey(a.bill_no) - numericKey(b.bill_no);
    });

    return rows.map((r, idx) => ({ ...r, seq: idx + 1, dateStr: `${String(r.dateObj.getDate()).padStart(2, '0')}/${String(r.dateObj.getMonth()+1).padStart(2, '0')}/${r.dateObj.getFullYear()}` }));
}

export function buildSmallBillRows(customersList, billRows, branchIds, periodStart, periodEnd) {
    const customers = {};
    customersList.forEach(c => customers[c.ID] = c);

    const branchSet = new Set(branchIds || []);
    const rows = [];

    for (const row of billRows) {
        const custId = row.CUST_VEND || "";
        if (branchSet.size > 0 && !branchSet.has(custId)) continue;
        const dateRaw = row.REF_DATE || "";
        if (dateRaw < periodStart || dateRaw > periodEnd) continue;
        if (SMALL_SUMMARY_BILLS.has(row.NO || "")) continue;

        const amount = parseFloat(row.BAL_AMT || 0);
        const vat = parseFloat(row.VAT_AMT || 0);
        const total = Math.round((amount + vat) * 100) / 100;
        if (total <= 0) continue;

        const cust = customers[custId] || {};
        const branchName = cleanBranchName(cust.SH_NAME || cust.SHIP_NAME || cust.NAME || "");
        const qty = Math.round(total / SMALL_PRICE);

        rows.push({
            dateObj: parseYmd(dateRaw),
            date: dateRaw,
            branch: `P${custId}`,
            bill_no: parseBillNumber(row.NO),
            branch_name: branchName,
            qty: qty,
            price: SMALL_PRICE,
            amount: amount,
            vat: vat,
            total: total,
            product_name: branchName,
        });
    }

    // Sort by branch numeric key, date, bill_no numeric key
    rows.sort((a, b) => {
        const branchDiff = numericKey(a.branch) - numericKey(b.branch);
        if (branchDiff !== 0) return branchDiff;
        if (a.date !== b.date) return a.date.localeCompare(b.date);
        return numericKey(a.bill_no) - numericKey(b.bill_no);
    });

    return rows.map((r, idx) => ({ ...r, seq: idx + 1, dateStr: `${String(r.dateObj.getDate()).padStart(2, '0')}/${String(r.dateObj.getMonth()+1).padStart(2, '0')}/${r.dateObj.getFullYear()}` }));
}

function applyBaseFormat(ws) {
    ws.views = [{ showGridLines: false, state: 'frozen', ySplit: 11 }];
    ws.columns = [
        { width: 8 },  // A
        { width: 18 }, // B
        { width: 14 }, // C
        { width: 12 }, // D
        { width: 30 }, // E
        { width: 12 }, // F
        { width: 12 }, // G
        { width: 14 }, // H
        { width: 14 }, // I
        { width: 14 }, // J
        { width: 14 }, // K
        { width: 14 }, // L
    ];
}

function writeCommonHeader(ws, config, billDateText) {
    ws.mergeCells("A1:L1");
    ws.getCell("A1").value = "ใบวางบิล/ใบแจ้งหนี้";
    ws.getCell("A1").font = { name: "Angsana New", size: 20, bold: true };
    ws.getCell("A1").alignment = { horizontal: "center" };

    ws.mergeCells("A2:J2");
    ws.getCell("A2").value = "บริษัท ซูเปอร์ ไอซ์ จำกัด  สำนักงานใหญ่";
    ws.getCell("A2").font = { name: "Angsana New", size: 18, bold: true };
    ws.getCell("K2").value = "ต้นฉบับ";
    ws.getCell("K2").font = { name: "Angsana New", size: 16, bold: true };

    ws.mergeCells("A3:H3");
    ws.getCell("A3").value = "ที่อยู่ 18/39 ซอยนวมินทร์ 111 แยก 15 แขวงนวมินทร์ เขตบึงกุ่ม กรุงเทพมหานคร 10240";
    ws.mergeCells("I3:L3");
    ws.getCell("I3").value = `เลขที่บิล ${config.bill_no}`;

    ws.mergeCells("A4:H4");
    ws.getCell("A4").value = "เลขประจำตัวผู้เสียภาษี 0105542031756";

    ws.mergeCells("I5:L5");
    ws.getCell("I5").value = billDateText;
    ws.getCell("I5").alignment = { horizontal: "center" };
    ws.getCell("I5").font = { name: "Angsana New", size: 16, bold: true };

    ws.mergeCells("A6:L6");
    ws.getCell("A6").value = config.customer_line;
    ws.getCell("A6").font = { name: "Angsana New", size: 16, bold: true };

    ws.mergeCells("A7:L7");
    ws.getCell("A7").value = config.address_line;

    ws.mergeCells("A8:L8");
    ws.getCell("A8").value = config.tax_line;

    if (config.title) {
        ws.mergeCells("A9:L9");
        ws.getCell("A9").value = config.title;
        ws.getCell("A9").font = { name: "Angsana New", size: 16, bold: true };
    }
}

const borderStyle = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

export function writeLawsonBigSheet(ws, config, rows, billDateText) {
    applyBaseFormat(ws);
    writeCommonHeader(ws, config, billDateText);

    const headers = ["ลำดับ", "วันที่", "เลขที่สาขา", "เลขที่บิล", "สาขา", "จำนวน", "ราคา/หน่วย", "จำนวนเงิน", "ภาษีมูลค่าเพิ่ม", "เงินรวมทั้งสิ้น"];
    const rowIdx = 11;

    for (let i = 0; i < headers.length; i++) {
        const cell = ws.getCell(rowIdx, i + 1);
        cell.value = headers[i];
        cell.font = { name: "Angsana New", size: 16, bold: true };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = borderStyle;
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9EAD3" } };
    }

    const dataStart = 12;
    for (let i = 0; i < rows.length; i++) {
        const record = rows[i];
        const row = ws.getRow(dataStart + i);
        
        row.getCell(1).value = record.seq;
        row.getCell(2).value = record.dateStr;
        row.getCell(3).value = record.branch;
        row.getCell(4).value = record.bill_no;
        row.getCell(5).value = record.branch_name;
        row.getCell(6).value = record.qty;
        row.getCell(7).value = record.price;
        row.getCell(8).value = { formula: `ROUND(F${dataStart + i}*G${dataStart + i}/1.07,2)` };
        row.getCell(9).value = { formula: `ROUND(J${dataStart + i}-H${dataStart + i},2)` };
        row.getCell(10).value = { formula: `ROUND(F${dataStart + i}*G${dataStart + i},2)` };

        for (let col = 1; col <= 10; col++) {
            const cell = row.getCell(col);
            cell.border = borderStyle;
            cell.font = { name: "Angsana New", size: 16 };
            if (col === 2) cell.alignment = { horizontal: "center" };
            else if (col === 1 || col === 6) { cell.numFmt = "0"; cell.alignment = { horizontal: "center" }; }
            else if (col === 7) { cell.numFmt = "0.00"; cell.alignment = { horizontal: "center" }; }
            else if (col >= 8) { cell.numFmt = "#,##0.00"; cell.alignment = { horizontal: "right" }; }
            else if (col === 3 || col === 4) cell.alignment = { horizontal: "center" };
        }
    }

    const totalRow = dataStart + rows.length;
    const totalGrand = rows.reduce((sum, r) => sum + r.total, 0);

    ws.mergeCells(totalRow, 1, totalRow, 2);
    ws.getCell(totalRow, 1).value = "รวมทั้งหมด";
    ws.getCell(totalRow, 3).value = "IN64002102";
    ws.mergeCells(totalRow, 4, totalRow, 7);
    ws.getCell(totalRow, 4).value = thaiBahtText(totalGrand);
    
    if (rows.length > 0) {
        ws.getCell(totalRow, 8).value = { formula: `SUM(H${dataStart}:H${totalRow - 1})` };
        ws.getCell(totalRow, 9).value = { formula: `SUM(I${dataStart}:I${totalRow - 1})` };
        ws.getCell(totalRow, 10).value = { formula: `SUM(J${dataStart}:J${totalRow - 1})` };
    }

    for (let col = 1; col <= 10; col++) {
        const cell = ws.getCell(totalRow, col);
        cell.border = borderStyle;
        cell.font = { name: "Angsana New", size: 16, bold: col >= 8 };
        if (col >= 8) cell.numFmt = "#,##0.00";
    }

    const footerRow = totalRow + 3;
    ws.mergeCells(footerRow, 1, footerRow, 4);
    ws.getCell(footerRow, 1).value = "ผู้วางบิล ……………………………………..";
    ws.mergeCells(footerRow, 7, footerRow, 9);
    ws.getCell(footerRow, 7).value = "ผู้รับวางบิล  .................................................";
    ws.mergeCells(footerRow + 1, 7, footerRow + 1, 9);
    ws.getCell(footerRow + 1, 7).value = "วันที่รับวางบิล………………………..";
    ws.getCell(footerRow, 10).value = "วันที่วางบิล.................................";
    ws.getCell(footerRow + 1, 10).value = "วันที่จ่ายชำระ.............................";
    
    // Add missing fonts for footer
    [footerRow, footerRow+1].forEach(r => {
        [1, 7, 10].forEach(c => {
            if(ws.getCell(r, c).value) ws.getCell(r, c).font = { name: "Angsana New", size: 16 };
        });
    });
}

export function writeNarrowSheet(ws, config, rows, combinedSmall, billDateText) {
    applyBaseFormat(ws);
    writeCommonHeader(ws, config, billDateText);

    const headers = ["ลำดับ", "วันที่", "เลขที่บิล", "รหัสสาขา", "ประเภท", "จำนวน", "ราคา", "จำนวนเงิน", "ภาษีมูลค่าเพิ่ม", "เงินรวมทั้งสิ้น"];
    const rowIdx = 10;

    for (let i = 0; i < headers.length; i++) {
        const cell = ws.getCell(rowIdx, i + 1);
        cell.value = headers[i];
        cell.font = { name: "Angsana New", size: 16, bold: true };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = borderStyle;
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9EAD3" } };
    }

    const dataStart = 11;
    for (let i = 0; i < rows.length; i++) {
        const record = rows[i];
        const row = ws.getRow(dataStart + i);
        const label = (combinedSmall || config.kind === "lawson_small_single") ? record.branch_name : "หลอดใหญ่";
        
        row.getCell(1).value = record.seq;
        row.getCell(2).value = record.dateStr;
        row.getCell(3).value = record.bill_no;
        row.getCell(4).value = record.branch;
        row.getCell(5).value = label;
        row.getCell(6).value = record.qty;
        row.getCell(7).value = record.price;
        row.getCell(8).value = { formula: `ROUND(F${dataStart + i}*G${dataStart + i}/1.07,2)` };
        row.getCell(9).value = { formula: `ROUND(J${dataStart + i}-H${dataStart + i},2)` };
        row.getCell(10).value = { formula: `ROUND(F${dataStart + i}*G${dataStart + i},2)` };

        for (let col = 1; col <= 10; col++) {
            const cell = row.getCell(col);
            cell.border = borderStyle;
            cell.font = { name: "Angsana New", size: 16 };
            if (col === 2) cell.alignment = { horizontal: "center" };
            else if (col === 1 || col === 6) { cell.numFmt = "0"; cell.alignment = { horizontal: "center" }; }
            else if (col === 7) { cell.numFmt = "0.00"; cell.alignment = { horizontal: "center" }; }
            else if (col >= 8) { cell.numFmt = "#,##0.00"; cell.alignment = { horizontal: "right" }; }
            else if (col === 3 || col === 4) cell.alignment = { horizontal: "center" };
        }
    }

    const totalRow = dataStart + rows.length;
    const totalGrand = rows.reduce((sum, r) => sum + r.total, 0);

    ws.getCell(totalRow, 2).value = "รวม";
    ws.getCell(totalRow, 3).value = thaiBahtText(totalGrand);
    
    if (rows.length > 0) {
        ws.getCell(totalRow, 6).value = { formula: `SUM(F${dataStart}:F${totalRow - 1})` };
        ws.getCell(totalRow, 8).value = { formula: `SUM(H${dataStart}:H${totalRow - 1})` };
        ws.getCell(totalRow, 9).value = { formula: `SUM(I${dataStart}:I${totalRow - 1})` };
        ws.getCell(totalRow, 10).value = { formula: `SUM(J${dataStart}:J${totalRow - 1})` };
    }

    for (let col = 1; col <= 10; col++) {
        const cell = ws.getCell(totalRow, col);
        cell.border = borderStyle;
        cell.font = { name: "Angsana New", size: 16, bold: true };
        if (col >= 8) cell.numFmt = "#,##0.00";
    }

    const footerRow = totalRow + 4;
    ws.getCell(footerRow, 1).value = "ลงชื่อ";
    ws.getCell(footerRow, 2).value = "............................................... ผู้วางบิล";
    ws.getCell(footerRow, 7).value = "ลงชื่อ  ............................................................";
    ws.getCell(footerRow, 10).value = "ผู้รับวางบิล";
    ws.getCell(footerRow + 2, 1).value = "วันที่วางบิล.................................";
    ws.getCell(footerRow + 2, 7).value = "วันที่รับวางบิล .................................................";
    ws.getCell(footerRow + 4, 1).value = "วันที่จ่ายชำระ.............................";
    
    // Add missing fonts for footer
    [footerRow, footerRow+2, footerRow+4].forEach(r => {
        [1, 2, 7, 10].forEach(c => {
            if(ws.getCell(r, c).value) ws.getCell(r, c).font = { name: "Angsana New", size: 16 };
        });
    });
}

export async function generateBillingWorkbook(bigSourceData, smallSourceData, monthStr, billDateStr) {
    const [periodStart, periodEnd] = monthBounds(monthStr);
    
    let billDateObj;
    if (billDateStr) {
        billDateObj = new Date(billDateStr);
    } else {
        billDateObj = new Date();
    }
    const billDateText = formatThaiBillDate(billDateObj);

    const wb = new ExcelJS.Workbook();
    wb.creator = "Billing Web App";
    wb.created = new Date();

    for (const config of SHEET_CONFIGS) {
        const ws = wb.addWorksheet(config.name);
        let rows = [];
        
        if (config.kind.startsWith("lawson_small")) {
            if (smallSourceData) {
                rows = buildSmallBillRows(smallSourceData.mcust, smallSourceData.abillno, config.branch_ids, periodStart, periodEnd);
            }
        } else {
            if (bigSourceData) {
                rows = buildTransRows(bigSourceData.mcust, bigSourceData.atrans, config.branch_ids, config.product_code, periodStart, periodEnd);
            }
        }

        if (config.kind === "lawson_big") {
            writeLawsonBigSheet(ws, config, rows, billDateText);
        } else if (config.kind === "franchise_big") {
            writeNarrowSheet(ws, config, rows, false, billDateText);
        } else if (config.kind === "lawson_small_combined") {
            writeNarrowSheet(ws, config, rows, true, billDateText);
        } else {
            writeNarrowSheet(ws, config, rows, true, billDateText);
        }
    }

    const buffer = await wb.xlsx.writeBuffer();
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
