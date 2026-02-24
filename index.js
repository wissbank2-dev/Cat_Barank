const express = require('express');
const path = require('path');
const ExcelJS = require('exceljs');
const app = express();
const PORT = process.env.PORT || 3000;

// Set EJS as the template engine
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Middleware to parse URL-encoded bodies and JSON
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Serve static files
app.use(express.static(path.join(__dirname, 'public')));

// Main Route
app.get('/', (req, res) => {
    res.render('index');
});

// Template Download Endpoint
app.get('/api/template', async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();

        // Helper: style a data sheet
        function styleDataSheet(ws, data) {
            ws.columns = [
                { header: 'No.', key: 'no', width: 8 },
                { header: 'Key', key: 'key', width: 30 },
                { header: 'Value', key: 'value', width: 40 },
                { header: 'Description', key: 'description', width: 40 }
            ];
            ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
            ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDA4AF' } };
            ws.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
            data.forEach(row => {
                const r = ws.addRow(row);
                r.alignment = { vertical: 'top' };
            });
        }

        // ---- Payload 1 ----
        const ws1 = workbook.addWorksheet('Payload 1');
        styleDataSheet(ws1, [
            { no: 1, key: 'job_no', value: 'JF18022600000001', description: 'เลขที่งาน' },
            { no: 2, key: 'forms[0].form_type', value: 'A', description: 'ประเภทฟอร์ม ชุดที่ 1' },
            { no: 3, key: 'forms[0].form_receive_date', value: '2026-02-19T17:10:00.000Z', description: 'วันที่รับ ชุดที่ 1' },
            { no: 4, key: 'forms[0].form_remark', value: 'Axxx', description: 'หมายเหตุ ชุดที่ 1' },
            { no: 5, key: 'forms[0].form_sts', value: 'S', description: 'สถานะ ชุดที่ 1' },
            { no: 6, key: 'forms[1].form_type', value: 'B', description: 'ประเภทฟอร์ม ชุดที่ 2' },
            { no: 7, key: 'forms[1].form_receive_date', value: '2026-02-19T17:10:00.000Z', description: 'วันที่รับ ชุดที่ 2' },
            { no: 8, key: 'forms[1].form_remark', value: 'Bxxx', description: 'หมายเหตุ ชุดที่ 2' },
            { no: 9, key: 'forms[1].form_sts', value: 'J', description: 'สถานะ ชุดที่ 2' }
        ]);

        // ---- Payload 2 (ตัวอย่างที่ 2) ----
        const ws2 = workbook.addWorksheet('Payload 2');
        styleDataSheet(ws2, [
            { no: 1, key: 'job_no', value: 'JF18022600000002', description: 'เลขที่งาน' },
            { no: 2, key: 'forms[0].form_type', value: 'C', description: 'ประเภทฟอร์ม' },
            { no: 3, key: 'forms[0].form_receive_date', value: '2026-03-01T09:00:00.000Z', description: 'วันที่รับ' },
            { no: 4, key: 'forms[0].form_remark', value: 'Cxxx', description: 'หมายเหตุ' },
            { no: 5, key: 'forms[0].form_sts', value: 'P', description: 'สถานะ' }
        ]);

        // ---- คำแนะนำ (Instructions) ----
        const instructionSheet = workbook.addWorksheet('คำแนะนำ');
        instructionSheet.columns = [{ header: '', key: 'text', width: 80 }];
        const instructions = [
            '📋 คำแนะนำการใช้งาน (Instructions)',
            '',
            '📌 แต่ละชีท (Sheet) = 1 Payload',
            '   → เพิ่มชีทใหม่เพื่อสร้าง Payload เพิ่ม',
            '   → ตั้งชื่อชีทอะไรก็ได้ (ห้ามชื่อ "คำแนะนำ")',
            '',
            '🔑 รูปแบบ Key รองรับ Nested JSON (dot-notation):',
            '',
            '  ✅ key ธรรมดา          → "job_no"',
            '  ✅ object ซ้อน         → "address.city"',
            '  ✅ array of objects    → "forms[0].form_type"',
            '  ✅ array ซ้อนหลายชั้น  → "data[0].items[1].name"',
            '',
            '1. ใส่ข้อมูลในแต่ละชีท',
            '2. คอลัมน์ No. = ลำดับ (ไม่จำเป็นต้องกรอก)',
            '3. คอลัมน์ Key = ชื่อ key / path ของ JSON (ห้ามเว้นว่าง)',
            '4. คอลัมน์ Value = ค่าที่ต้องการ',
            '5. คอลัมน์ Description = คำอธิบาย (ไม่จำเป็น)',
            '',
            '🐱 ฅ^•ﻌ•^ฅ Cat Test Case Builder'
        ];
        instructions.forEach(text => {
            instructionSheet.addRow({ text });
        });

        // Send file
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=JSON_Template_Cat.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error('Template generation failed:', error);
        res.status(500).json({ error: 'Failed to generate template' });
    }
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
