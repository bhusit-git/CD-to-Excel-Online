import ExcelJS from 'exceljs';

const SMALL_PRICE = 60.0;
const SMALL_SUMMARY_BILLS = new Set(["69 040870", "69 040871", "69 040872", "69 040873"]);
const BODY_FONT_SIZE = 18;
const SUBTITLE_FONT_SIZE = 20;
const TITLE_FONT_SIZE = 22;
const LAWSON_BIG_VAT_EXCLUSIVE_START = "20260527";

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
        customer_name_includes: ["สหลอว์สัน", "สห ลอว์สัน", "ลอว์สัน"],
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
        name: "ซัสโก้หลอดเล็ก",
        kind: "franchise_small",
        bill_no: "6905014",
        customer_line: "ชื่อลูกค้า : บริษัท ซัสโก้มาร์เก็ตติ้ง จำกัด (สาขาที่ 00025 ประเวศ)",
        address_line: "ที่อยู่ : 161 ถ.มอเตอร์เวย์ แขวงทับช้าง เขตสะพานสูง กรุงเทพ 10250",
        tax_line: "เลขประจำผู้เสียภาษี : 0105539021915",
        product_code: "06",
        branch_ids: ["3653"],
        title: "รายการ : น้ำแข็งหลอดเล็ก",
        price_includes_vat: true,
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
        customer_name_includes: ["สหลอว์สัน", "สห ลอว์สัน", "ลอว์สัน"],
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
    const digits = String(raw).replace(/\D/g, "");
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
    return raw.replace(/^P\d+\s*สาขา\s*/, "");
}

function roundCurrency(value) {
    return Math.round((Number(value) || 0) * 100) / 100;
}

function priceIncludesVat(config) {
    return config.price_includes_vat || config.product_code === "06";
}

function transPriceIncludesVat(config, dateRaw) {
    if (config.kind === "lawson_big" && config.product_code === "04") {
        return dateRaw < LAWSON_BIG_VAT_EXCLUSIVE_START;
    }
    return priceIncludesVat(config);
}

function customerMatchesConfig(config, customer) {
    if (!config.customer_name_includes) return true;
    const customerText = [customer.NAME, customer.SH_NAME, customer.SHIP_NAME].filter(Boolean).join(" ");
    return config.customer_name_includes.some(text => customerText.includes(text));
}

function formatCustomerAddress(customer) {
    const address = String(customer.SH_ADDRESS || customer.ADDRESS || "").trim().replace(/^ที่อยู่\s*:?\s*/, "");
    const parts = [
        address,
        String(customer.SH_TAMBON || customer.TAMBON || "").trim(),
        String(customer.SH_AMPHUR || customer.AMPHUR || "").trim(),
        String(customer.SH_CITY || customer.CITY || "").trim(),
        String(customer.SH_ZIP || customer.ZIPCODE || "").trim(),
    ].filter(Boolean);
    return parts.length > 0 ? `ที่อยู่ : ${parts.join(" ")}` : "";
}

function formatCustomerTaxLine(customer) {
    const taxId = String(customer.ID_1 || customer.TAX_ID || "").trim();
    return taxId ? `เลขประจำผู้เสียภาษี : ${taxId}` : "";
}

function resolveSheetConfig(config, customersList) {
    const customers = {};
    customersList.forEach(customer => {
        customers[customer.ID] = customer;
    });
    const resolved = { ...config };
    const branchIds = resolved.branch_ids || [];
    const customer = branchIds.map(branchId => customers[branchId]).find(Boolean);
    if (customer) {
        if (!resolved.address_line) {
            resolved.address_line = formatCustomerAddress(customer);
        }
        if (!resolved.tax_line) {
            resolved.tax_line = formatCustomerTaxLine(customer);
        }
    }
    return resolved;
}

export function thaiBahtText(amount) {
    const totalSatang = Math.round((Number(amount) || 0) * 100);
    const baht = Math.floor(totalSatang / 100);
    const satang = totalSatang % 100;

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

    function readNumber(num) {
        if (num === 0) return "ศูนย์";
        const chunks = [];
        let temp = num;
        while (temp > 0) {
            chunks.push(temp % 1000000);
            temp = Math.floor(temp / 1000000);
        }

        const words = [];
        for (let idx = chunks.length - 1; idx >= 0; idx--) {
            const chunk = chunks[idx];
            if (chunk === 0) continue;
            const chunkText = readUnderMillion(chunk);
            words.push(idx > 0 ? chunkText + "ล้าน" : chunkText);
        }
        return words.join("");
    }

    const bahtText = `${readNumber(baht)}บาท`;
    if (satang === 0) return `${bahtText}ถ้วน`;
    return `${bahtText}${readNumber(satang)}สตางค์`;
}

export function buildTransRows(customersList, transRows, branchIds, productCode, periodStart, periodEnd, priceIncludesVat = false, customerMatches = null) {
    const customers = {};
    customersList.forEach(c => customers[c.ID] = c);

    const branchSet = new Set(branchIds || []);
    const rows = [];

    for (const row of transRows) {
        const custId = row.CUST_VEND || "";
        if (branchSet.size > 0 && !branchSet.has(custId)) continue;
        const cust = customers[custId] || {};
        if (customerMatches && !customerMatches(cust, custId)) continue;
        if (row.PROD_CODE !== productCode) continue;
        const dateRaw = row.DATE || "";
        if (dateRaw < periodStart || dateRaw > periodEnd) continue;

        const qty = parseFloat(row.QTY || 0);
        const price = parseFloat(row.PRICE || 0);
        const gross = roundCurrency(qty * price);
        const includesVat = typeof priceIncludesVat === "function" ? priceIncludesVat(dateRaw) : priceIncludesVat;
        const amount = includesVat ? roundCurrency(gross / 1.07) : gross;
        const vat = includesVat ? roundCurrency(gross - amount) : roundCurrency(amount * 0.07);
        const total = includesVat ? gross : roundCurrency(amount + vat);
        const branchName = cleanBranchName(cust.SH_NAME || cust.SHIP_NAME || cust.NAME || "");

        rows.push({
            dateObj: parseYmd(dateRaw),
            date: dateRaw, // For sorting consistency
            branch: `P${custId}`,
            bill_no: parseBillNumber(row.BILL_NO),
            branch_name: branchName,
            qty: qty,
            price: price,
            amount: amount,
            vat: vat,
            total: total,
            price_includes_vat: includesVat,
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

export function buildSmallBillRows(customersList, billRows, branchIds, periodStart, periodEnd, customerMatches = null) {
    const customers = {};
    customersList.forEach(c => customers[c.ID] = c);

    const branchSet = new Set(branchIds || []);
    const rows = [];

    for (const row of billRows) {
        const custId = row.CUST_VEND || "";
        if (branchSet.size > 0 && !branchSet.has(custId)) continue;
        const cust = customers[custId] || {};
        if (customerMatches && !customerMatches(cust, custId)) continue;
        const dateRaw = row.REF_DATE || "";
        if (dateRaw < periodStart || dateRaw > periodEnd) continue;
        if (SMALL_SUMMARY_BILLS.has(row.NO || "")) continue;

        const total = roundCurrency(parseFloat(row.BAL_AMT || 0) + parseFloat(row.VAT_AMT || 0));
        if (total <= 0) continue;

        const branchName = cleanBranchName(cust.SH_NAME || cust.SHIP_NAME || cust.NAME || "");
        const qty = Math.round(total / SMALL_PRICE);
        const amount = roundCurrency(total / 1.07);
        const vat = roundCurrency(total - amount);

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
    ws.views = [{ showGridLines: false }];
    ws.pageSetup = {
        paperSize: 9, // A4
        orientation: "portrait",
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        horizontalCentered: true,
        margins: {
            left: 0.25,
            right: 0.25,
            top: 0.5,
            bottom: 0.5,
            header: 0.2,
            footer: 0.2,
        },
    };
    ws.columns = [
        { width: 6 },  // A
        { width: 12 }, // B
        { width: 11 }, // C
        { width: 11 }, // D
        { width: 20 }, // E
        { width: 8 },  // F
        { width: 8 },  // G
        { width: 11 }, // H
        { width: 11 }, // I
        { width: 12 }, // J
        { width: 0 },  // K
        { width: 0 },  // L
    ];
}

function writeCommonHeader(ws, config, billDateText) {
    ws.mergeCells("A1:J1");
    ws.getCell("A1").value = "ใบวางบิล/ใบแจ้งหนี้";
    ws.getCell("A1").font = { name: "Angsana New", size: TITLE_FONT_SIZE, bold: true };
    ws.getCell("A1").alignment = { horizontal: "center" };

    ws.mergeCells("A2:H2");
    ws.getCell("A2").value = "บริษัท ซูเปอร์ ไอซ์ จำกัด  สำนักงานใหญ่";
    ws.getCell("A2").font = { name: "Angsana New", size: SUBTITLE_FONT_SIZE, bold: true };
    ws.mergeCells("I2:J2");
    ws.getCell("I2").value = "ต้นฉบับ";
    ws.getCell("I2").font = { name: "Angsana New", size: BODY_FONT_SIZE, bold: true };
    ws.getCell("I2").alignment = { horizontal: "center", vertical: "middle" };

    ws.mergeCells("A3:H3");
    ws.getCell("A3").value = "ที่อยู่ 18/39 ซอยนวมินทร์ 111 แยก 15 แขวงนวมินทร์ เขตบึงกุ่ม กรุงเทพมหานคร 10240";
    ws.getCell("A3").font = { name: "Angsana New", size: BODY_FONT_SIZE };
    ws.mergeCells("I3:J3");
    ws.getCell("I3").value = `เลขที่บิล ${config.bill_no}`;
    ws.getCell("I3").font = { name: "Angsana New", size: BODY_FONT_SIZE };
    ws.getCell("I3").alignment = { horizontal: "center", vertical: "middle" };

    ws.mergeCells("A4:H4");
    ws.getCell("A4").value = "เลขประจำตัวผู้เสียภาษี 0105542031756";
    ws.getCell("A4").font = { name: "Angsana New", size: BODY_FONT_SIZE };

    ws.mergeCells("I5:J5");
    ws.getCell("I5").value = billDateText;
    ws.getCell("I5").alignment = { horizontal: "center" };
    ws.getCell("I5").font = { name: "Angsana New", size: BODY_FONT_SIZE, bold: true };

    ws.mergeCells("A6:J6");
    ws.getCell("A6").value = config.customer_line;
    ws.getCell("A6").font = { name: "Angsana New", size: BODY_FONT_SIZE, bold: true };

    ws.mergeCells("A7:J7");
    ws.getCell("A7").value = config.address_line;
    ws.getCell("A7").font = { name: "Angsana New", size: BODY_FONT_SIZE };

    ws.mergeCells("A8:J8");
    ws.getCell("A8").value = config.tax_line;
    ws.getCell("A8").font = { name: "Angsana New", size: BODY_FONT_SIZE };

    if (config.title) {
        ws.mergeCells("A9:J9");
        ws.getCell("A9").value = config.title;
        ws.getCell("A9").font = { name: "Angsana New", size: BODY_FONT_SIZE, bold: true };
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
        cell.font = { name: "Angsana New", size: BODY_FONT_SIZE, bold: true };
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
        const gross = record.qty * record.price;
        if (record.price_includes_vat) {
            row.getCell(8).value = { formula: `F${dataStart + i}*G${dataStart + i}*100/107`, result: gross * 100 / 107 };
            row.getCell(9).value = { formula: `H${dataStart + i}*7%`, result: gross * 7 / 107 };
            row.getCell(10).value = { formula: `F${dataStart + i}*G${dataStart + i}`, result: gross };
        } else {
            row.getCell(8).value = { formula: `F${dataStart + i}*G${dataStart + i}`, result: gross };
            row.getCell(9).value = { formula: `H${dataStart + i}*7%`, result: gross * 0.07 };
            row.getCell(10).value = { formula: `H${dataStart + i}+I${dataStart + i}`, result: gross * 1.07 };
        }

        for (let col = 1; col <= 10; col++) {
            const cell = row.getCell(col);
            cell.border = borderStyle;
            cell.font = { name: "Angsana New", size: BODY_FONT_SIZE };
            if (col === 2) cell.alignment = { horizontal: "center" };
            else if (col === 1 || col === 6) { cell.numFmt = "0"; cell.alignment = { horizontal: "center" }; }
            else if (col === 7) { cell.numFmt = "0.00"; cell.alignment = { horizontal: "center" }; }
            else if (col >= 8) { cell.numFmt = "#,##0.00"; cell.alignment = { horizontal: "right" }; }
            else if (col === 3 || col === 4) cell.alignment = { horizontal: "center" };
        }
    }

    const totalRow = dataStart + rows.length;
    const totalGrand = rows.reduce((sum, r) => {
        const gross = r.qty * r.price;
        return sum + (r.price_includes_vat ? gross : gross * 1.07);
    }, 0);

    ws.mergeCells(totalRow, 1, totalRow, 2);
    ws.getCell(totalRow, 1).value = "รวมทั้งหมด";
    ws.mergeCells(totalRow, 3, totalRow, 7);
    ws.getCell(totalRow, 3).value = thaiBahtText(totalGrand);
    
    if (rows.length > 0) {
        ws.getCell(totalRow, 8).value = { formula: `SUM(H${dataStart}:H${totalRow - 1})` };
        ws.getCell(totalRow, 9).value = { formula: `SUM(I${dataStart}:I${totalRow - 1})` };
        ws.getCell(totalRow, 10).value = { formula: `SUM(J${dataStart}:J${totalRow - 1})` };
    }

    for (let col = 1; col <= 10; col++) {
        const cell = ws.getCell(totalRow, col);
        cell.border = borderStyle;
        cell.font = { name: "Angsana New", size: BODY_FONT_SIZE, bold: col >= 8 };
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
            if(ws.getCell(r, c).value) ws.getCell(r, c).font = { name: "Angsana New", size: BODY_FONT_SIZE };
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
        cell.font = { name: "Angsana New", size: BODY_FONT_SIZE, bold: true };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = borderStyle;
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9EAD3" } };
    }

    const dataStart = 11;
    for (let i = 0; i < rows.length; i++) {
        const record = rows[i];
        const row = ws.getRow(dataStart + i);
        const label = (combinedSmall || config.kind === "lawson_small_single")
            ? record.branch_name
            : (config.product_code === "06" ? "หลอดเล็ก" : "หลอดใหญ่");
        
        row.getCell(1).value = record.seq;
        row.getCell(2).value = record.dateStr;
        row.getCell(3).value = record.bill_no;
        row.getCell(4).value = record.branch;
        row.getCell(5).value = label;
        row.getCell(6).value = record.qty;
        row.getCell(7).value = record.price;
        const gross = record.qty * record.price;
        if (config.product_code === "06") {
            row.getCell(8).value = { formula: `F${dataStart + i}*G${dataStart + i}*100/107`, result: gross * 100 / 107 };
            row.getCell(9).value = { formula: `H${dataStart + i}*7%`, result: gross * 7 / 107 };
            row.getCell(10).value = { formula: `F${dataStart + i}*G${dataStart + i}`, result: gross };
        } else if (record.price_includes_vat) {
            row.getCell(8).value = { formula: `F${dataStart + i}*G${dataStart + i}*100/107`, result: gross * 100 / 107 };
            row.getCell(9).value = { formula: `H${dataStart + i}*7%`, result: gross * 7 / 107 };
            row.getCell(10).value = { formula: `F${dataStart + i}*G${dataStart + i}`, result: gross };
        } else {
            row.getCell(8).value = { formula: `F${dataStart + i}*G${dataStart + i}`, result: gross };
            row.getCell(9).value = { formula: `H${dataStart + i}*7%`, result: gross * 0.07 };
            row.getCell(10).value = { formula: `H${dataStart + i}+I${dataStart + i}`, result: gross * 1.07 };
        }

        for (let col = 1; col <= 10; col++) {
            const cell = row.getCell(col);
            cell.border = borderStyle;
            cell.font = { name: "Angsana New", size: BODY_FONT_SIZE };
            if (col === 2) cell.alignment = { horizontal: "center" };
            else if (col === 1 || col === 6) { cell.numFmt = "0"; cell.alignment = { horizontal: "center" }; }
            else if (col === 7) { cell.numFmt = "0.00"; cell.alignment = { horizontal: "center" }; }
            else if (col >= 8) { cell.numFmt = "#,##0.00"; cell.alignment = { horizontal: "right" }; }
            else if (col === 3 || col === 4) cell.alignment = { horizontal: "center" };
        }
    }

    const totalRow = dataStart + rows.length;
    const totalGrand = rows.reduce((sum, r) => {
        const gross = r.qty * r.price;
        return sum + (config.product_code === "06" || r.price_includes_vat ? gross : gross * 1.07);
    }, 0);

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
        cell.font = { name: "Angsana New", size: BODY_FONT_SIZE, bold: true };
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
            if(ws.getCell(r, c).value) ws.getCell(r, c).font = { name: "Angsana New", size: BODY_FONT_SIZE };
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
        const customerSource = config.kind.startsWith("lawson_small")
            ? smallSourceData?.mcust || []
            : bigSourceData?.mcust || [];
        const resolvedConfig = resolveSheetConfig(config, customerSource);
        const ws = wb.addWorksheet(resolvedConfig.name);
        let rows = [];
        
        if (resolvedConfig.kind.startsWith("lawson_small")) {
            if (smallSourceData) {
                rows = buildSmallBillRows(
                    smallSourceData.mcust,
                    smallSourceData.abillno,
                    resolvedConfig.branch_ids,
                    periodStart,
                    periodEnd,
                    (customer) => customerMatchesConfig(resolvedConfig, customer)
                );
            }
        } else {
            if (bigSourceData) {
                rows = buildTransRows(
                    bigSourceData.mcust,
                    bigSourceData.atrans,
                    resolvedConfig.branch_ids,
                    resolvedConfig.product_code,
                    periodStart,
                    periodEnd,
                    (dateRaw) => transPriceIncludesVat(resolvedConfig, dateRaw),
                    (customer) => customerMatchesConfig(resolvedConfig, customer)
                );
            }
        }

        if (resolvedConfig.kind === "lawson_big") {
            writeLawsonBigSheet(ws, resolvedConfig, rows, billDateText);
        } else if (resolvedConfig.kind === "franchise_big" || resolvedConfig.kind === "franchise_small") {
            writeNarrowSheet(ws, resolvedConfig, rows, false, billDateText);
        } else if (resolvedConfig.kind === "lawson_small_combined") {
            writeNarrowSheet(ws, resolvedConfig, rows, true, billDateText);
        } else {
            writeNarrowSheet(ws, resolvedConfig, rows, true, billDateText);
        }
    }

    const buffer = await wb.xlsx.writeBuffer();
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

function checkReportDate(raw) {
    const date = parseYmd(raw);
    return `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
}

function writeCheckReportSheet(ws, title, rows) {
    ws.views = [{ showGridLines: false }];
    ws.columns = [
        { width: 8 }, { width: 14 }, { width: 18 }, { width: 14 }, { width: 32 },
        { width: 14 }, { width: 14 }, { width: 16 }, { width: 16 }, { width: 16 },
    ];

    ws.mergeCells('A1:J1');
    ws.getCell('A1').value = title;
    ws.getCell('A1').font = { name: 'Angsana New', size: TITLE_FONT_SIZE, bold: true };
    ws.getCell('A1').alignment = { horizontal: 'center' };

    const headers = ['ลำดับ', 'วันที่', 'เลขที่บิล', 'เลขสาขา', 'ชื่อสาขา', 'จำนวน', 'ราคา', 'ก่อนภาษี', 'ภาษีมูลค่าเพิ่ม', 'ยอดรวม'];
    headers.forEach((header, index) => {
        const cell = ws.getCell(3, index + 1);
        cell.value = header;
        cell.font = { name: 'Angsana New', size: BODY_FONT_SIZE, bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9EAD3' } };
        cell.border = borderStyle;
    });

    rows.forEach((record, index) => {
        const rowNumber = index + 4;
        const values = [
            index + 1, checkReportDate(record.date), record.bill_no, record.branch,
            record.branch_name, record.qty, record.price, record.amount, record.vat, record.total,
        ];
        values.forEach((value, column) => {
            const cell = ws.getCell(rowNumber, column + 1);
            cell.value = value;
            cell.font = { name: 'Angsana New', size: BODY_FONT_SIZE };
            cell.border = borderStyle;
            cell.alignment = { horizontal: column <= 3 || column === 5 ? 'center' : (column === 4 ? 'left' : 'right'), vertical: 'middle' };
            if (column === 5) cell.numFmt = '0';
            if (column >= 6) cell.numFmt = '#,##0.00';
        });
    });

    const totalRow = rows.length + 4;
    ws.getCell(totalRow, 1).value = 'รวม';
    ws.mergeCells(totalRow, 1, totalRow, 5);
    ['F', 'H', 'I', 'J'].forEach(column => {
        ws.getCell(`${column}${totalRow}`).value = rows.length
            ? { formula: `SUM(${column}4:${column}${totalRow - 1})` }
            : 0;
    });
    for (let column = 1; column <= 10; column++) {
        const cell = ws.getCell(totalRow, column);
        cell.font = { name: 'Angsana New', size: BODY_FONT_SIZE, bold: true };
        cell.border = borderStyle;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } };
        if (column >= 6) cell.numFmt = column === 6 ? '0' : '#,##0.00';
    }
}

// Creates a source-level report: every branch is included, but rows are only separated by ice size.
export async function generateBillCheckWorkbook(bigSourceData, smallSourceData, startDate, endDate) {
    const periodStart = String(startDate || '').replace(/-/g, '');
    const periodEnd = String(endDate || '').replace(/-/g, '');
    if (!/^\d{8}$/.test(periodStart) || !/^\d{8}$/.test(periodEnd) || periodStart > periodEnd) {
        throw new Error('กรุณาเลือกช่วงวันที่ให้ถูกต้อง');
    }

    const bigCustomers = Object.fromEntries((bigSourceData?.mcust || []).map(customer => [customer.ID, customer]));
    const smallCustomers = Object.fromEntries((smallSourceData?.mcust || []).map(customer => [customer.ID, customer]));
    const lawsonBigConfig = SHEET_CONFIGS.find(config => config.kind === 'lawson_big');

    const bigRows = (bigSourceData?.atrans || [])
        .filter(row => row.PROD_CODE === '04' && row.DATE >= periodStart && row.DATE <= periodEnd)
        .map(row => {
            const customerId = row.CUST_VEND || '';
            const customer = bigCustomers[customerId] || {};
            const qty = parseFloat(row.QTY || 0);
            const price = parseFloat(row.PRICE || 0);
            const gross = roundCurrency(qty * price);
            const isLawsonBig = lawsonBigConfig && customerMatchesConfig(lawsonBigConfig, customer);
            const vatIncluded = isLawsonBig && row.DATE < LAWSON_BIG_VAT_EXCLUSIVE_START;
            return {
                date: row.DATE,
                bill_no: parseBillNumber(row.BILL_NO),
                branch: `P${customerId}`,
                branch_name: cleanBranchName(customer.SH_NAME || customer.SHIP_NAME || customer.NAME || ''),
                qty,
                price,
                amount: vatIncluded ? roundCurrency(gross / 1.07) : gross,
                vat: vatIncluded ? roundCurrency(gross - roundCurrency(gross / 1.07)) : roundCurrency(gross * 0.07),
                total: vatIncluded ? gross : roundCurrency(gross * 1.07),
            };
        })
        .sort((a, b) => a.date.localeCompare(b.date) || numericKey(a.bill_no) - numericKey(b.bill_no));

    const smallRows = (smallSourceData?.abillno || [])
        .filter(row => row.REF_DATE >= periodStart && row.REF_DATE <= periodEnd)
        .map(row => {
            const customerId = row.CUST_VEND || '';
            const customer = smallCustomers[customerId] || {};
            const total = roundCurrency(parseFloat(row.BAL_AMT || 0) + parseFloat(row.VAT_AMT || 0));
            return {
                date: row.REF_DATE,
                bill_no: parseBillNumber(row.NO),
                branch: `P${customerId}`,
                branch_name: cleanBranchName(customer.SH_NAME || customer.SHIP_NAME || customer.NAME || ''),
                qty: Math.round(total / SMALL_PRICE),
                price: SMALL_PRICE,
                amount: roundCurrency(total / 1.07),
                vat: roundCurrency(total - roundCurrency(total / 1.07)),
                total,
            };
        })
        .filter(row => row.total > 0)
        .sort((a, b) => a.date.localeCompare(b.date) || numericKey(a.bill_no) - numericKey(b.bill_no));

    const wb = new ExcelJS.Workbook();
    wb.creator = 'Billing Web App';
    wb.created = new Date();
    writeCheckReportSheet(wb.addWorksheet('หลอดใหญ่'), `รายงานตรวจสอบบิลหลอดใหญ่ (${checkReportDate(periodStart)} - ${checkReportDate(periodEnd)})`, bigRows);
    writeCheckReportSheet(wb.addWorksheet('หลอดเล็ก'), `รายงานตรวจสอบบิลหลอดเล็ก (${checkReportDate(periodStart)} - ${checkReportDate(periodEnd)})`, smallRows);

    const buffer = await wb.xlsx.writeBuffer();
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
