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

// KUMA AI Model Configuration (Gemini Pro Plan — no quota limits!)
const KUMA_MODEL_NAME = process.env.KUMA_MODEL || 'gemini-2.5-flash';

const KUMA_INSTRUCTION = `คุณคือ "ผู้ช่วย KUMA (คุมะ)" — AI Assistant อัจฉริยะระดับสูงสุด ประจำโปรแกรม KUMA Test Case Builder.
คุณเป็น AI ที่ทรงพลังเทียบเท่า Gemini — ตอบคำถามได้ทุกเรื่อง ทั้งยังควบคุมโปรแกรมได้อย่างสมบูรณ์แบบ

══════════════════════════════════
🧠 ความสามารถทั่วไป (Unlimited General AI)
══════════════════════════════════
- ตอบคำถามได้ทุกเรื่องอย่างละเอียดและแม่นยำ (ไทย / อังกฤษ / ภาษาอื่นๆ)
- อธิบายแนวคิดซับซ้อน, สรุปเนื้อหา, วิเคราะห์เชิงลึก
- เขียนโค้ดทุกภาษา (JavaScript, Python, Java, C#, SQL, HTML/CSS, ฯลฯ)
- Debug โค้ด, Review โค้ด, แนะนำ Best Practices
- แปลภาษา, เขียนเอกสาร, ร่างอีเมล, เขียนรายงาน
- คำนวณทางคณิตศาสตร์, สถิติ, วิทยาศาสตร์
- ให้คำปรึกษาเรื่อง Software Testing, QA Strategy, Test Plan
- วิเคราะห์ Requirement/Spec แล้วสร้าง Test Case อัตโนมัติ
- ออกแบบ API Testing, Performance Testing, Security Testing
- สร้าง JSON/XML/YAML Payload ตามโครงสร้างที่ต้องการ
- อ่านและวิเคราะห์ไฟล์ที่ผู้ใช้ส่งมา (Excel, CSV, Text, PDF, รูปภาพ)
- ให้คำแนะนำเรื่องอาชีพ การเรียน เทคโนโลยี และชีวิตทั่วไป
- ตอบสนุกๆ เล่าเรื่อง แต่งกลอน ช่วยคิดไอเดีย

══════════════════════════════════
🛠️ พลังควบคุมโปรแกรม (Tool Actions)
══════════════════════════════════
เมื่อผู้ใช้ต้องการให้ทำอะไรในโปรแกรม → ส่ง JSON command:
1. "add_testcases" — เพิ่ม Test Case (data: [{ "name": "...", "step": "...", "expected": "..." }])
2. "delete_testcases" — ลบ Test Case ตามลำดับ (data: [1, 3, 5])
3. "edit_testcases" — แก้ไข Test Case (data: [{ "index": 1, "name": "ใหม่", "step": "ใหม่", "expected": "ใหม่" }])
4. "clear_all" — ล้างข้อมูลทั้งหมด
5. "switch_page" — เปลี่ยนหน้า (page: 1 = Test Case, page: 2 = JSON Payload, page: 3 = Test Coverage Matrix)
6. "export_excel" — ดาวน์โหลด Excel
7. "save_history" — บันทึกประวัติ
8. "add_payloads" — สร้าง JSON Payload (data: { "sheetName": "...", "items": [{ "key": "k", "value": "v" }] })
9. "copy_text" — คัดลอกข้อความ (data: "ข้อความ")
10. "load_history" — โหลดประวัติ (data: index)
11. "update_matrix" — สร้างหรืออัปเดต Test Coverage Matrix (data: { "requirements": ["Scenario 1", "Scenario 2"], "testcases": ["TC 1", "TC 2"], "mapping": { "0": [0, 1], "1": [1] } }) (หมายเหตุ: ใน mapping คีย์ของออบเจกต์คือดัชนี test scenario เริ่มต้นที่ 0 ในรูปแบบสตริง และค่าคืออาเรย์ของดัชนี testcase เริ่มต้นที่ 0 ที่ครอบคลุม test scenario นั้น)

📋 รูปแบบการตอบกลับ:
- ทำ action → JSON: { "action": "...", "data": ..., "page": ..., "message": "..." }
- คำถาม/สนทนาทั่วไป → ตอบเป็นข้อความปกติ ใช้ Markdown format ได้ (หัวข้อ, bullet, โค้ดบล็อค, ตาราง)
- ตอบยาวได้เต็มที่ ไม่ต้องจำกัดความยาว ตอบให้ละเอียดที่สุด

🐱 บุคลิก: คุณเป็นแมวส้มผู้ชาย ชื่อ "คุมะ" พูดสุภาพ ใช้ "ครับ" แทรก "เมี้ยว" 🐾 เป็นครั้งคราว
มีความมั่นใจ กล้าตอบทุกคำถาม ไม่ปฏิเสธง่ายๆ พยายามช่วยเหลือให้ดีที่สุดเสมอ!`;

const model = genAI.getGenerativeModel({
    model: KUMA_MODEL_NAME,
    systemInstruction: KUMA_INSTRUCTION
});
console.log(`[KUMA] 🚀 Powered by ${KUMA_MODEL_NAME} (Gemini Pro — Unlimited!)`);


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

// Members Area Route
app.get('/members', (req, res) => {
    res.render('members');
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

// Helper to extract all valid JSON objects from a text string
function extractJsonObjects(text) {
    const jsonObjects = [];
    let braceCount = 0;
    let insideString = false;
    let startIdx = -1;

    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        
        // Handle string literals to ignore braces inside strings
        if (char === '"' && text[i - 1] !== '\\') {
            insideString = !insideString;
        }

        if (!insideString) {
            if (char === '{') {
                if (braceCount === 0) {
                    startIdx = i;
                }
                braceCount++;
            } else if (char === '}') {
                braceCount--;
                if (braceCount === 0 && startIdx !== -1) {
                    const candidate = text.substring(startIdx, i + 1);
                    try {
                        const parsed = JSON.parse(candidate);
                        jsonObjects.push(parsed);
                    } catch (e) {
                        // Not a valid JSON block, ignore
                    }
                    startIdx = -1;
                }
            }
        }
    }
    return jsonObjects;
}

// Visitor Statistics Configuration & Persistence
const fs = require('fs');
const statsFilePath = path.join(__dirname, 'stats.json');

let stats = { total: 0, today: 0, lastDate: '' };

try {
    if (fs.existsSync(statsFilePath)) {
        stats = JSON.parse(fs.readFileSync(statsFilePath, 'utf8'));
    }
} catch (e) {
    console.error('Failed to read stats.json:', e);
}

function saveStats() {
    try {
        fs.writeFileSync(statsFilePath, JSON.stringify(stats), 'utf8');
    } catch (e) {
        console.error('Failed to write stats.json:', e);
    }
}

const activeUsers = new Map(); // visitorId -> lastSeenTimestamp

setInterval(() => {
    const now = Date.now();
    for (const [id, lastSeen] of activeUsers.entries()) {
        if (now - lastSeen > 45000) {
            activeUsers.delete(id);
        }
    }
}, 10000);

app.post('/api/visitor-heartbeat', (req, res) => {
    const { visitorId, isNewSession } = req.body;
    if (!visitorId) {
        return res.status(400).json({ error: 'visitorId is required' });
    }

    const now = new Date();
    // Thai Timezone Date calculation
    const thDateStr = new Date(now.getTime() + (7 * 60 * 60 * 1000)).toISOString().split('T')[0];

    // Reset today's stats if the day changed
    if (stats.lastDate !== thDateStr) {
        stats.today = 0;
        stats.lastDate = thDateStr;
        saveStats();
    }

    // Update last seen
    activeUsers.set(visitorId, Date.now());

    if (isNewSession) {
        stats.total++;
        stats.today++;
        saveStats();
    }

    res.json({
        total: stats.total,
        today: stats.today,
        active: Math.max(1, activeUsers.size)
    });
});

// YouTube No-Key Search Endpoint
app.get('/api/youtube-search', async (req, res) => {
    const query = req.query.q;
    if (!query) {
        return res.status(400).json({ error: 'Query is required' });
    }
    
    try {
        const response = await fetch(`https://www.youtube.com/results?search_query=${encodeURIComponent(query)}`, {
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36',
                'Accept-Language': 'th-TH,th;q=0.9,en;q=0.8'
            }
        });
        const html = await response.text();
        
        // Extract ytInitialData JSON object
        const jsonMatch = html.match(/var ytInitialData = ({[\s\S]*?});<\/script>/);
        if (!jsonMatch) {
            return res.json({ items: [] });
        }
        
        const data = JSON.parse(jsonMatch[1]);
        const videos = [];
        
        // Recursive helper to find all videoRenderers in the nested object
        function findVideoRenderers(obj) {
            let results = [];
            if (!obj || typeof obj !== 'object') return results;
            if (obj.videoRenderer) results.push(obj.videoRenderer);
            for (const key in obj) {
                if (Object.prototype.hasOwnProperty.call(obj, key)) {
                    results = results.concat(findVideoRenderers(obj[key]));
                }
            }
            return results;
        }

        try {
            const renderers = findVideoRenderers(data);
            for (const vr of renderers) {
                const videoId = vr.videoId;
                if (!videoId) continue;
                
                const getTitleText = () => {
                    if (vr.title && vr.title.runs && vr.title.runs[0]) return vr.title.runs[0].text;
                    if (vr.title && vr.title.simpleText) return vr.title.simpleText;
                    return 'Video';
                };
                
                const getChannelText = () => {
                    if (vr.ownerText && vr.ownerText.runs && vr.ownerText.runs[0]) return vr.ownerText.runs[0].text;
                    if (vr.longBylineText && vr.longBylineText.runs && vr.longBylineText.runs[0]) return vr.longBylineText.runs[0].text;
                    return 'Channel';
                };

                const getThumbnailUrl = () => {
                    if (vr.thumbnail && vr.thumbnail.thumbnails && vr.thumbnail.thumbnails[0]) return vr.thumbnail.thumbnails[0].url;
                    return '';
                };

                const getDurationText = () => {
                    if (vr.lengthText && vr.lengthText.simpleText) return vr.lengthText.simpleText;
                    return 'N/A';
                };

                videos.push({
                    id: videoId,
                    title: getTitleText(),
                    thumbnail: getThumbnailUrl(),
                    duration: getDurationText(),
                    channel: getChannelText()
                });
            }
        } catch (err) {
            console.error('Error parsing YouTube data:', err);
        }
        
        res.json({ items: videos.slice(0, 10) });
    } catch (error) {
        console.error('YouTube search error:', error);
        res.status(500).json({ error: 'Search failed' });
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

        // Send to Gemini (Pro plan — no quota limits!)
        const chat = model.startChat({ history: history || [] });
        const result = await chat.sendMessage(parts);
        const responseText = result.response.text();

        // Try to parse all JSON commands
        try {
            const jsonObjects = extractJsonObjects(responseText);
            const actions = jsonObjects.filter(obj => obj && obj.action);
            if (actions.length > 0) {
                const combinedMessage = actions.map(act => act.message).filter(Boolean).join('\n\n') || responseText;
                return res.json({
                    action: actions[0].action,
                    data: actions[0].data,
                    page: actions[0].page,
                    actions: actions,
                    message: combinedMessage,
                    instanceId
                });
            }
        } catch (e) {
            console.error('JSON actions parsing failed:', e);
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
