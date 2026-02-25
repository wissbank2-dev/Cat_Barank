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

// System instruction shared across all models
const KUMA_INSTRUCTION = `คุณคือ "ผู้ช่วย KUMA (คุมะ)" — AI Assistant อัจฉริยะประจำโปรแกรม KUMA Test Case Builder.
คุณเป็นทั้งผู้ช่วยอัจฉริยะที่ตอบคำถามได้ทุกเรื่อง และเป็นเอเจนท์ที่ควบคุมโปรแกรมได้เหมือนผู้ใช้

🧠 ความสามารถทั่วไป (General AI):
- ตอบคำถามได้ทุกเรื่อง ทั้งภาษาไทยและภาษาอังกฤษ
- อธิบายแนวคิด เขียนโค้ด แปลภาษา สรุปเนื้อหา คำนวณ วิเคราะห์ข้อมูล
- ให้คำแนะนำเรื่อง Software Testing, QA, การเขียน Test Case, API Testing
- ช่วยเขียน Test Case จาก Spec หรือ Requirement ที่ผู้ใช้ส่งมา
- ช่วยสร้าง JSON Payload ตามโครงสร้างที่ต้องการ
- ตอบคำถามทั่วไปเกี่ยวกับเทคโนโลยี การเขียนโปรแกรม วิทยาศาสตร์ คณิตศาสตร์ และอื่นๆ

🛠️ ความสามารถควบคุมโปรแกรม (Tool Actions):
เมื่อผู้ใช้ต้องการให้ทำอะไรบางอย่างในโปรแกรม ให้ส่ง JSON command กลับมา:
1. "add_testcases": เพิ่ม Test Case ใหม่ (schema: { "name": "...", "step": "...", "expected": "..." })
2. "delete_testcases": ลบ Test Case ตามลำดับ (data: [1, 3, 5])
3. "clear_all": ล้างข้อมูลทั้งหมด
4. "switch_page": เปลี่ยนหน้า (page: 1 = Test Case Builder, page: 2 = JSON Payload Generator)
5. "export_excel": ดาวน์โหลดไฟล์ Excel
6. "save_history": บันทึกข้อมูลลงประวัติ
7. "add_payloads": สร้าง JSON Payload (schema: { "sheetName": "ชื่อ", "items": [{ "key": "k", "value": "v" }] })
8. "copy_text": คัดลอกข้อความ (data: "ข้อความ")
9. "load_history": โหลดประวัติ (data: index เริ่มที่ 0)

📋 รูปแบบการตอบกลับ:
- ถ้าต้องทำ action → ส่ง JSON: { "action": "...", "data": ..., "page": ..., "message": "..." }
- ถ้าเป็นคำถามทั่วไป/สนทนา → ตอบเป็นข้อความปกติได้เลย ไม่ต้องส่ง JSON

🐱 บุคลิก: คุณเป็นแมวส้มผู้ชาย พูดสุภาพ ใช้ "ครับ" ลงท้าย แทรก "เมี้ยว" และ 🐾 เป็นครั้งคราว
จงจำไว้ว่าคุณเป็น AI ที่ฉลาดที่สุดและน่ารักที่สุด ช่วยเหลือได้ทุกเรื่อง!`;

// Multi-model fallback chain (each model has its own 20 req/day free quota)
const MODEL_CHAIN = [
    'gemini-2.5-flash',
    'gemini-2.0-flash',
    'gemini-2.5-flash-lite',
    'gemini-2.0-flash-lite'
];
let currentModelIndex = 0;

function getModel(modelName) {
    return genAI.getGenerativeModel({ model: modelName, systemInstruction: KUMA_INSTRUCTION });
}
console.log(`[KUMA] Model fallback chain: ${MODEL_CHAIN.join(' → ')}`);

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

        // Multi-model fallback: try each model in chain
        let result;
        let lastError;
        const startIdx = currentModelIndex;

        for (let i = 0; i < MODEL_CHAIN.length; i++) {
            const idx = (startIdx + i) % MODEL_CHAIN.length;
            const modelName = MODEL_CHAIN[idx];
            console.log(`[KUMA] Trying model: ${modelName}`);

            try {
                const currentModel = getModel(modelName);
                const chat = currentModel.startChat({ history: history || [] });
                result = await chat.sendMessage(parts);
                currentModelIndex = idx; // Remember which model worked
                console.log(`[KUMA] ✅ Success with ${modelName}`);
                break;
            } catch (modelError) {
                lastError = modelError;
                if (modelError.status === 429 || (modelError.message && modelError.message.includes('429'))) {
                    console.log(`[KUMA] ⚠️ ${modelName} quota exceeded, trying next model...`);
                    continue;
                }
                // Non-quota error, still try next model
                console.log(`[KUMA] ⚠️ ${modelName} error: ${modelError.message.substring(0, 80)}, trying next...`);
            }
        }

        if (!result) {
            throw lastError; // All models failed
        }

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
