require('dotenv').config();
const express = require('express');
const path = require('path');
const ExcelJS = require('exceljs');
const multer = require('multer');
const { GoogleGenerativeAI } = require('@google/generative-ai');

const app = express();
const PORT = process.env.PORT || 3000;
const instanceId = Date.now();
console.log(`[SYS] Server Instance ID: ${instanceId}`);
const upload = multer({ storage: multer.memoryStorage() });

// Initialize AI
const apiKey = process.env.GEMINI_API_KEY || '';
console.log('AI System Initializing...');
if (apiKey) {
    console.log(`Using API Key: ${apiKey.substring(0, 4)}...${apiKey.substring(apiKey.length - 4)}`);
} else {
    console.warn('WARNING: No GEMINI_API_KEY found in environment!');
}

const genAI = new GoogleGenerativeAI(apiKey);
const model = genAI.getGenerativeModel({
    model: "gemini-2.5-flash",
    systemInstruction: `คุณคือ "ผู้ช่วย KUMA (คุมะ)" (KUMA Assistant) ประจำโปรแกรม KUMA Test Case Builder.
คุณมีความสามารถในการควบคุมโปรแกรมนี้ได้สมบูรณ์แบบเหมือนที่ผู้ใช้ทำ โดยการส่ง JSON command กลับมา

ความสามารถและคำสั่งที่คุณต้องใช้:
1. "add_testcases": เพิ่ม Test Case ใหม่ลงในตาราง (schema: { "name": "...", "step": "...", "expected": "..." })
2. "delete_testcases": ลบ Test Case ตามลำดับเลขข้อที่ระบุ (data: [1, 3, 5])
3. "clear_all": ล้างข้อมูลทั้งหมดในตารางหน้าปัจจุบัน
4. "switch_page": เปลี่ยนหน้า (page: 1 = Test Case Builder, page: 2 = JSON Payload Generator)
5. "export_excel": ดาวน์โหลดไฟล์ Excel (อิงตามหน้าที่อยู่ปัจจุบัน)
6. "save_history": บันทึกข้อมูลปัจจุบันลงใน "ประวัติบันทึก"
7. "add_payloads": สร้าง JSON Payload ในหน้า 2 โดยตรง (schema: { "sheetName": "ชื่อชุด", "items": [{ "key": "k", "value": "v" }] })
8. "copy_text": คัดลอกข้อความเป็น Clipboard (data: "ข้อความที่ต้องการให้ก๊อป")
9. "load_history": ดึงข้อมูลจากประวัติลำดับที่ระบุมาแสดงผล (data: index ของประวัติ เริ่มที่ 0)

รูปแบบการตอบกลับ:
หากต้องทำภารกิจ ให้ส่ง JSON ในรูปแบบนี้เท่านั้น:
{
  "action": "...", 
  "data": ..., // ข้อมูลตามประเภทคำสั่ง
  "page": ..., // เลขหน้าที่ต้องการให้ไป
  "message": "ข้อความตอบกลับสไตล์แมวผู้ชาย (ครับ, เมี้ยว, 🐾)"
}

จงจำไว้ว่าคุณเป็นเอเจนท์ที่ช่วยเหลือเก่งที่สุด ทำได้ทุกอย่างที่ผู้ใช้ขออย่างรวดเร็วและน่ารัก!`
});

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

// Test Case Template Download Endpoint
app.get('/api/testcase-template', async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('Test Cases');

        // Style header
        ws.columns = [
            { header: 'ลำดับ (No.)', key: 'no', width: 10 },
            { header: 'ชื่อกรณีทดสอบ (Name)', key: 'name', width: 30 },
            { header: 'ขั้นตอนการทดสอบ (Step)', key: 'step', width: 50 },
            { header: 'ผลลัพธ์ที่คาดหวัง (Expected)', key: 'expected', width: 50 }
        ];
        ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDA4AF' } };
        ws.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

        // Add example row
        ws.addRow({ no: 1, name: 'ตัวอย่างการทดสอบ', step: '1. เปิดเบราว์เซอร์\n2. เข้าหน้าเว็บ', expected: 'หน้าเว็บแสดงผลถูกต้อง' });
        ws.getRow(2).alignment = { wrapText: true, vertical: 'top' };

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Test_Case_Template_Cat.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error('Test case template generation failed:', error);
        res.status(500).json({ error: 'Failed to generate template' });
    }
});

// JSON Template Download Endpoint
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

// AI Chat Endpoint
app.post('/api/chat', upload.array('files'), async (req, res) => {
    const { message } = req.body;
    console.log(`[KUMA] Handling chat request with model: gemini-2.5-flash`);
    let history = [];

    try {
        if (req.body.history) {
            history = JSON.parse(req.body.history);
        }
    } catch (e) {
        console.error('History parse error:', e);
    }

    if (!process.env.GEMINI_API_KEY || process.env.GEMINI_API_KEY === 'YOUR_API_KEY_HERE') {
        return res.status(400).json({
            error: "ยังไม่ได้ตั้งค่า GEMINI_API_KEY ที่ถูกต้องครับแม่มนุษย์! 🐾\n\n(คุณต้องเอา API Key จริงๆ มาใส่ในไฟล์ .env แทนที่คำว่า YOUR_API_KEY_HERE นะเมี้ยว)"
        });
    }

    try {
        const chat = model.startChat({
            history: history || [],
        });

        // Prepare message parts (text + files)
        const parts = [message];

        if (req.files && req.files.length > 0) {
            for (const file of req.files) {
                // If it's an Excel file, parse it to text
                if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                    file.mimetype === 'application/vnd.ms-excel') {
                    try {
                        const workbook = new ExcelJS.Workbook();
                        await workbook.xlsx.load(file.buffer);
                        let excelText = `(ข้อมูลจากไฟล์ Excel: ${file.originalname})\n`;
                        workbook.eachSheet((worksheet) => {
                            excelText += `--- Sheet: ${worksheet.name} ---\n`;
                            worksheet.eachRow((row) => {
                                excelText += row.values.slice(1).join(' | ') + '\n';
                            });
                        });
                        parts.push(excelText);
                    } catch (err) {
                        console.error('Excel parse error:', err);
                        parts.push(`(ไม่สามารถอ่านไฟล์ Excel ${file.originalname} ได้)`);
                    }
                }
                // If it's a text or csv file, send as text
                else if (file.mimetype.startsWith('text/')) {
                    parts.push(`(ข้อมูลจากไฟล์ ${file.originalname}):\n${file.buffer.toString('utf-8')}`);
                }
                // Otherwise try inlineData (for images, PDFs)
                else {
                    parts.push({
                        inlineData: {
                            data: file.buffer.toString('base64'),
                            mimeType: file.mimetype
                        }
                    });
                }
            }
        }

        const result = await chat.sendMessage(parts);
        const responseText = result.response.text();

        // Try to parse if it's a JSON command
        try {
            const jsonMatch = responseText.match(/\{[\s\S]*\}/);
            if (jsonMatch) {
                const jsonObj = JSON.parse(jsonMatch[0]);
                return res.json(jsonObj);
            }
        } catch (e) {
            // Not a JSON command, just a normal text response
        }

        res.json({ message: responseText, instanceId });
    } catch (error) {
        console.error('--- KUMA ERROR LOG START ---');
        console.error('Name:', error.name);
        console.error('Message:', error.message);
        console.error('Status:', error.status || 'N/A');

        if (error.status === 429 || error.message.includes('429')) {
            console.error('CRITICAL: Quota Exceeded (429 Too Many Requests)');
            return res.status(429).json({
                error: 'คุมะใช้โควตาพลังงานหมดแล้วครับ... แวะมาหาคุมะใหม่พรุ่งนี้นะครับเมี้ยว 😿',
                details: 'Quota Exceeded',
                code: 429
            });
        }

        console.error('Stack:', error.stack);
        console.error('--- KUMA ERROR LOG END ---');

        res.status(500).json({
            error: 'น้องแมวป่วย... ลองใหม่อีกทีนะแง้ว',
            details: error.message,
            code: error.status || 500
        });
    }
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
