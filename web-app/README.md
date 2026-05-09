# Lawson Billing Web App

ระบบสร้างใบวางบิลและใบแจ้งหนี้ (Lawson & Franchise) รูปแบบ Web Application พัฒนาด้วย React และทำงานแบบ **Client-Side Processing 100%** เพื่อความปลอดภัยและเป็นส่วนตัวของข้อมูล โดยไม่ต้องพึ่งพา Backend Server ข้อมูลไฟล์ DBF ทั้งหมดจะถูกประมวลผลบนเบราว์เซอร์ของคุณ

---

## ฟีเจอร์หลัก (Features)
- 🚀 **Client-Side Processing:** ปลอดภัย ข้อมูลทั้งหมด (ยอดเงิน, รายการลูกค้า) ไม่ถูกอัปโหลดขึ้นเซิร์ฟเวอร์
- 📂 **Dual Folder Upload:** รองรับการอัปโหลดโฟลเดอร์ 2 แหล่งพร้อมกัน (Big Source และ Small Source) 
- 🗄️ **DBF Parser:** มีตัวอ่านไฟล์ `.DBF` ในตัวที่รองรับการถอดรหัสภาษาไทย (TIS-620 / CP874)
- 📊 **Excel Generation:** สรุปยอด ภาษีมูลค่าเพิ่ม และสร้างไฟล์ Excel แบบหลายชีท (Multi-sheet) พร้อมจัดหน้าตารางพร้อมพิมพ์ให้อัตโนมัติด้วย `exceljs`
- 🇹🇭 **Thai Text Conversion:** รองรับการแปลงตัวเลขยอดเงินรวมเป็นคำอ่านภาษาไทย (เช่น "หนึ่งร้อยบาทถ้วน")

---

## 📖 วิธีใช้งานระบบ (User Guide)

1. เปิดหน้าเว็บแอปพลิเคชัน
2. **ตั้งค่าการวางบิล (Report Settings):**
   - **Billing Month (เดือนที่วางบิล):** เลือกเดือนและปีที่ต้องการสรุปยอด
   - **Billing Date (วันที่วางบิล):** ระบุวันที่สำหรับการวางบิล (จะถูกนำไปแสดงในหัวบิล Excel)
3. **อัปโหลดข้อมูล (Upload Data):**
   - **Big Source Folder:** คลิกที่กล่องนี้ แล้วเลือกโฟลเดอร์ **SI-xxxx** (ระบบจะค้นหาไฟล์ `ATRANS.DBF` และ `MCUST.DBF` ในนี้อัตโนมัติ)
   - **Small Source Folder:** คลิกที่กล่องนี้ แล้วเลือกโฟลเดอร์ **LAWSON** (ระบบจะค้นหาไฟล์ `ABILLNO.DBF` และ `MCUST.DBF` ในนี้อัตโนมัติ)
4. **สร้างไฟล์ Excel:**
   - เมื่อเลือกข้อมูลครบถ้วนแล้ว ปุ่ม **Generate Excel** จะสามารถกดได้
   - กดปุ่ม **Generate Excel** ระบบจะทำการประมวลผล และดาวน์โหลดไฟล์นามสกุล `.xlsx` ลงเครื่องของคุณทันที!

---

## สำหรับนักพัฒนา (For Developers)

### เทคโนโลยีที่ใช้ (Tech Stack)
- **Framework:** React + Vite
- **Styling:** Tailwind CSS + Framer Motion
- **Icons:** Lucide React
- **Data Processing:** `exceljs` (สร้าง Excel), `iconv-lite` (ถอดรหัสภาษาไทยจากไฟล์ DBF)

### การติดตั้งและรันระบบแบบ Local (Local Development)

1. ต้องมี **Node.js** (แนะนำเวอร์ชัน 18 ขึ้นไป) ติดตั้งอยู่ในเครื่อง
2. เปิด Terminal แล้วเข้าไปที่โฟลเดอร์โปรเจกต์:
   \`\`\`bash
   cd path/to/web-app
   \`\`\`
3. ติดตั้ง Dependencies:
   \`\`\`bash
   npm install
   \`\`\`
4. เริ่มเซิร์ฟเวอร์จำลอง:
   \`\`\`bash
   npm run dev
   \`\`\`
5. เปิดเบราว์เซอร์ไปที่ `http://localhost:5173/`

### การนำไปใช้งานจริง (Deployment - Cloudflare Pages)

ระบบนี้ถูกออกแบบมาเป็น Static Site ซึ่งเหมาะอย่างยิ่งสำหรับการนำไปฝากไว้บน **Cloudflare Pages**, **Vercel** หรือ **GitHub Pages** ได้ฟรี!

**วิธีการตั้งค่าบน Cloudflare Pages:**
- **Framework Preset:** Vite (หรือ None)
- **Build Command:** `npm run build`
- **Build Output Directory:** `dist`

หลังจากตั้งค่าและ Deploy สำเร็จ คุณจะสามารถใช้งานเว็บแอปพลิเคชันผ่าน URL ที่ Cloudflare กำหนดได้ทันที และใช้งานได้ตลอดเวลาโดยไม่ต้องคอยเปิดคอมพิวเตอร์ทิ้งไว้
