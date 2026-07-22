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
const todayVisitors = new Set(); // Keep track of unique visitors counted today (resets on day change/server reboot)

setInterval(() => {
    const now = Date.now();
    for (const [id, lastSeen] of activeUsers.entries()) {
        if (now - lastSeen > 45000) {
            activeUsers.delete(id);
        }
    }
}, 10000);

app.post('/api/visitor-heartbeat', (req, res) => {
    const { visitorId } = req.body;
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
        todayVisitors.clear(); // Reset today's tracking set
        saveStats();
    }

    // Update last seen
    activeUsers.set(visitorId, Date.now());

    // Deduplicate unique visitors on backend (in case client is already active or reloads)
    if (!todayVisitors.has(visitorId)) {
        todayVisitors.add(visitorId);
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

// ==================== GISX Automate Endpoints ====================

// 1. Download GISX Register Case Excel Template
app.get('/api/gisx/template', async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('Register Case');

        // Fields config (all required and optional fields from GISX screen layout)
        const fields = [
            { header: 'Quotation No. *', key: 'quotationNo', width: 25, required: true },
            { header: 'Policy Holder Title *', key: 'title', width: 25, required: true },
            { header: 'Policy Holder Name (Thai) *', key: 'nameTh', width: 35, required: true },
            { header: 'Policy Holder Name (English) *', key: 'nameEn', width: 35, required: true },
            { header: 'Line of Business *', key: 'lineOfBusiness', width: 22, required: true },
            { header: 'Risk Level *', key: 'riskLevel', width: 18, required: true },
            { header: 'Occupational Classification (Priority) *', key: 'occupationClass', width: 40, required: true },
            { header: 'Policy Effective Date *', key: 'effDate', width: 25, required: true },
            { header: 'Policy Effective Time *', key: 'effTime', width: 25, required: true },
            { header: 'Policy End Date *', key: 'endDate', width: 25, required: true },
            { header: 'Policy End Time *', key: 'endTime', width: 25, required: true },
            { header: 'Policy Language *', key: 'language', width: 22, required: true },
            { header: 'Copy of Policy *', key: 'copyCount', width: 20, required: true },
            
            // Policy Holder Address / Contact info
            { header: 'Address 1 *', key: 'address1', width: 30, required: true },
            { header: 'Address 2', key: 'address2', width: 30, required: false },
            { header: 'Country *', key: 'country', width: 20, required: true },
            { header: 'Province *', key: 'province', width: 20, required: true },
            { header: 'District *', key: 'district', width: 20, required: true },
            { header: 'Sub District *', key: 'subDistrict', width: 25, required: true },
            { header: 'Zip Code *', key: 'zipCode', width: 15, required: true },
            { header: 'Contact Name *', key: 'contactName', width: 30, required: true },
            { header: 'Contact Position *', key: 'contactPosition', width: 25, required: true },
            { header: 'Contact Mobile', key: 'contactMobile', width: 20, required: false },
            { header: 'Contact Phone', key: 'contactPhone', width: 20, required: false },
            { header: 'Contact Email', key: 'contactEmail', width: 25, required: false },
            
            // Coverage section
            { header: 'Product Type *', key: 'productType', width: 20, required: true },
            { header: 'Sub Product *', key: 'subProduct', width: 20, required: true },
            { header: 'Age Average *', key: 'ageAverage', width: 18, required: true },
            { header: 'Min Age *', key: 'minAge', width: 15, required: true },
            { header: 'Max Age *', key: 'maxAge', width: 15, required: true },
            { header: 'Plan Number *', key: 'planNumber', width: 15, required: true },
            { header: 'Plan Type *', key: 'planType', width: 45, required: true },
            { header: 'Mode of Payment *', key: 'modeOfPayment', width: 25, required: true },
            
            // Agent/Broker
            { header: 'Channel *', key: 'channel', width: 25, required: true },
            { header: 'Agent/Broker Code *', key: 'agentBrokerCode', width: 22, required: true },
            { header: 'Sales Team *', key: 'salesTeam', width: 22, required: true },
            { header: 'Sales Name *', key: 'salesName', width: 25, required: true },
            
            // Experience Refund
            { header: 'Experience Refund (ER) *', key: 'erType', width: 25, required: true },
            { header: 'Loss Ratio', key: 'lossRatio', width: 15, required: false },
            { header: 'Refund Rate', key: 'refundRate', width: 15, required: false },

            // Account Detail modal fields
            { header: 'Account Title *', key: 'accTitle', width: 25, required: true },
            { header: 'Account Name (Thai) *', key: 'accNameTh', width: 35, required: true },
            { header: 'Account Name (English) *', key: 'accNameEn', width: 35, required: true },
            { header: 'Account Tax ID *', key: 'accTaxId', width: 25, required: true },
            { header: 'Account Type *', key: 'accType', width: 25, required: true },
            { header: 'Account Head Count Type *', key: 'accHeadCountType', width: 28, required: true },
            { header: 'Account Head Count Desc', key: 'accHeadCountDesc', width: 30, required: false },
            { header: 'Account Line of Business *', key: 'accLineOfBusiness', width: 28, required: true },
            { header: 'Account Risk Level *', key: 'accRiskLevel', width: 22, required: true },
            { header: 'Account Occupation Class *', key: 'accOccupationClass', width: 35, required: true },

            // Commission Rates
            { header: 'Commission Plan Type', key: 'commPlanType1', width: 25, required: false },
            { header: 'Commission Rate (%)', key: 'commRate1', width: 25, required: false },
            { header: 'Additional Commission (%)', key: 'addCommRate1', width: 25, required: false }
        ];

        ws.columns = fields.map(f => ({ header: f.header, key: f.key, width: f.width }));

        // Style the headers (Red for Required, Blue for Optional)
        const headerRow = ws.getRow(1);
        headerRow.height = 30;
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

        fields.forEach((f, idx) => {
            const cell = headerRow.getCell(idx + 1);
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: f.required ? 'FFEF4444' : 'FF3B82F6' } // Red for required, Blue for optional
            };
        });

        // Add validations for rows 2 to 100
        const validations = {
            title: {
                type: 'list',
                allowBlank: true,
                formulae: ['"นาย,นาง,นางสาว,เด็กชาย,เด็กหญิง,บจก.,บมจ."']
            },
            lineOfBusiness: {
                type: 'list',
                allowBlank: true,
                formulae: ['"Ordinary,Group,Credit Life,Accident"']
            },
            riskLevel: {
                type: 'list',
                allowBlank: true,
                formulae: ['"Low,Medium,High"']
            },
            occupationClass: {
                type: 'list',
                allowBlank: true,
                formulae: ['"Class 1,Class 2,Class 3,Class 4"']
            },
            language: {
                type: 'list',
                allowBlank: true,
                formulae: ['"Thai,English"']
            },
            country: {
                type: 'list',
                allowBlank: true,
                formulae: ['"Thailand"']
            },
            province: {
                type: 'list',
                allowBlank: true,
                formulae: ['"กทม."']
            },
            productType: {
                type: 'list',
                allowBlank: true,
                formulae: ['"01"']
            },
            planType: {
                type: 'list',
                allowBlank: true,
                formulae: ['"1 : ชีวิต,2 : อุบัติเหตุ,3 : ทุพพลภาพ,4 : สุขภาพ,5 : โรคร้ายแรง,6 : รพก.,7 : อุบัติเหตุกลุ่มส่วนบุคคล,9 : QA"']
            },
            commPlanType1: {
                type: 'list',
                allowBlank: true,
                formulae: ['"Select All,เลือกทั้งหมด,1 : ชีวิต,2 : อุบัติเหตุ,3 : ทุพพลภาพ,4 : สุขภาพ,5 : โรคร้ายแรง,6 : รพก.,7 : อุบัติเหตุกลุ่มส่วนบุคคล,9 : QA"']
            },
            modeOfPayment: {
                type: 'list',
                allowBlank: true,
                formulae: ['"Monthly, รายเดือน"']
            },
            channel: {
                type: 'list',
                allowBlank: true,
                formulae: ['"Agent (บุคคลธรรมดา)"']
            },
            erType: {
                type: 'list',
                allowBlank: true,
                formulae: ['"ER,NON_ER"']
            },
            accTitle: {
                type: 'list',
                allowBlank: true,
                formulae: ['"นาย,นาง,นางสาว,เด็กชาย,เด็กหญิง,บจก.,บมจ."']
            },
            accHeadCountType: {
                type: 'list',
                allowBlank: true,
                formulae: ['"Non Head Count,Head Count"']
            },
            accLineOfBusiness: {
                type: 'list',
                allowBlank: true,
                formulae: ['"Ordinary,Group,Credit Life,Accident"']
            },
            accRiskLevel: {
                type: 'list',
                allowBlank: true,
                formulae: ['"Low,Medium,High"']
            },
            accOccupationClass: {
                type: 'list',
                allowBlank: true,
                formulae: ['"Class 1,Class 2,Class 3,Class 4"']
            }
        };

        for (let r = 2; r <= 100; r++) {
            fields.forEach((f, colIdx) => {
                const cell = ws.getCell(r, colIdx + 1);
                if (validations[f.key]) {
                    cell.dataValidation = validations[f.key];
                }
                
                // Format numeric styling
                if (f.key === 'copyCount' || f.key === 'planNumber') {
                    cell.numFmt = '0';
                }
            });
        }

        // Add a demo/example row on row 2 (first row under headers)
        const demoRow = ws.getRow(2);
        demoRow.values = {
            quotationNo: 'QT20260721-DEMO',
            title: 'นาย',
            nameTh: 'สมชาย มั่งคั่งคุมะ',
            nameEn: 'Somchai Mangkang',
            lineOfBusiness: 'Ordinary',
            riskLevel: 'Low',
            occupationClass: 'Class 1',
            effDate: '21/07/2026',
            effTime: '00:00:00',
            endDate: '20/07/2027',
            endTime: '23:59:59',
            language: 'Thai',
            copyCount: 1,
            address1: '123/45 Kuma Tower',
            address2: '',
            country: 'Thailand',
            province: 'กทม.',
            district: 'วัฒนา',
            subDistrict: 'คลองเตย',
            zipCode: '10310',
            contactName: 'สมชาย มั่งคั่งคุมะ',
            contactPosition: 'ผู้จัดการ',
            contactMobile: '',
            contactPhone: '',
            contactEmail: '',
            productType: '01',
            subProduct: '',
            ageAverage: '40',
            minAge: '2',
            maxAge: '80',
            planNumber: 4,
            planType: '1 : ชีวิต',
            modeOfPayment: 'Monthly, รายเดือน',
            channel: 'Agent (บุคคลธรรมดา)',
            agentBrokerCode: '144660',
            salesTeam: '',
            salesName: '',
            erType: 'NON_ER',
            lossRatio: '',
            refundRate: '',
            accTitle: 'นาย',
            accNameTh: 'สมชาย มั่งคั่งคุมะ',
            accNameEn: 'Somchai Mangkang',
            accTaxId: '1101800262649',
            accType: 'Compulsory (Base Account)',
            accHeadCountType: 'Non Head Count',
            accHeadCountDesc: '',
            accLineOfBusiness: 'Ordinary',
            accRiskLevel: 'Low',
            accOccupationClass: 'Class 1',
            
            // Commission Rates Demo Values
            commPlanType1: '1 : ชีวิต',
            commRate1: 10,
            addCommRate1: 2
        };

        // Set response headers and transmit the template file
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=GISX_Register_Case_Template.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error('GISX Template generation failed:', error);
        res.status(500).json({ error: 'Failed to generate GISX Excel template' });
    }
});

// 2. Upload and Parse GISX Excel file
app.post('/api/gisx/upload', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'ไม่พบไฟล์อัปโหลด' });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);

        const ws = workbook.worksheets[0];
        if (!ws) {
            return res.status(400).json({ error: 'ไม่พบตารางงานในไฟล์ Excel' });
        }

        const cases = [];
        const requiredFields = [
            'quotationNo', 'title', 'nameTh', 'nameEn', 'lineOfBusiness', 'riskLevel', 'occupationClass', 'effDate', 'effTime', 'endDate', 'endTime', 'language', 'copyCount',
            'address1', 'country', 'province', 'district', 'subDistrict', 'zipCode', 'contactName', 'contactPosition',
            'productType', 'subProduct', 'ageAverage', 'minAge', 'maxAge', 'planNumber', 'planType', 'modeOfPayment',
            'channel', 'agentBrokerCode', 'salesTeam', 'salesName', 'erType',
            'accTitle', 'accNameTh', 'accNameEn', 'accTaxId', 'accType', 'accHeadCountType', 'accLineOfBusiness', 'accRiskLevel', 'accOccupationClass'
        ];

        // Map column indices dynamically based on headers in row 1
        const headerRow = ws.getRow(1);
        const colMap = {};
        
        const fieldMapping = {
            'Quotation No. *': 'quotationNo',
            'Policy Holder Title *': 'title',
            'Policy Holder Name (Thai) *': 'nameTh',
            'Policy Holder Name (English) *': 'nameEn',
            'Line of Business *': 'lineOfBusiness',
            'Risk Level *': 'riskLevel',
            'Occupational Classification (Priority) *': 'occupationClass',
            'Policy Effective Date *': 'effDate',
            'Policy Effective Time *': 'effTime',
            'Policy End Date *': 'endDate',
            'Policy End Time *': 'endTime',
            'Policy Language *': 'language',
            'Copy of Policy *': 'copyCount',
            'Address 1 *': 'address1',
            'Address 2': 'address2',
            'Country *': 'country',
            'Province *': 'province',
            'District *': 'district',
            'Sub District *': 'subDistrict',
            'Zip Code *': 'zipCode',
            'Contact Name *': 'contactName',
            'Contact Position *': 'contactPosition',
            'Contact Mobile': 'contactMobile',
            'Contact Phone': 'contactPhone',
            'Contact Email': 'contactEmail',
            'Product Type *': 'productType',
            'Sub Product *': 'subProduct',
            'Age Average *': 'ageAverage',
            'Min Age *': 'minAge',
            'Max Age *': 'maxAge',
            'Plan Number *': 'planNumber',
            'Plan Type *': 'planType',
            'Mode of Payment *': 'modeOfPayment',
            'Channel *': 'channel',
            'Agent/Broker Code *': 'agentBrokerCode',
            'Sales Team *': 'salesTeam',
            'Sales Name *': 'salesName',
            'Experience Refund (ER) *': 'erType',
            'Loss Ratio': 'lossRatio',
            'Refund Rate': 'refundRate',
            'Account Title *': 'accTitle',
            'Account Name (Thai) *': 'accNameTh',
            'Account Name (English) *': 'accNameEn',
            'Account Tax ID *': 'accTaxId',
            'Account Type *': 'accType',
            'Account Head Count Type *': 'accHeadCountType',
            'Account Head Count Desc': 'accHeadCountDesc',
            'Account Line of Business *': 'accLineOfBusiness',
            'Account Risk Level *': 'accRiskLevel',
            'Account Occupation Class *': 'accOccupationClass',
            'Commission Plan Type': 'commPlanType1',
            'Commission Rate (%)': 'commRate1',
            'Additional Commission (%)': 'addCommRate1',
            'Commission Plan Type 1': 'commPlanType1',
            'Commission Rate 1 (%)': 'commRate1',
            'Additional Commission 1 (%)': 'addCommRate1',
            'Commission Plan Type 2': 'commPlanType2',
            'Commission Rate 2 (%)': 'commRate2',
            'Additional Commission 2 (%)': 'addCommRate2',
            'Commission Plan Type 3': 'commPlanType3',
            'Commission Rate 3 (%)': 'commRate3',
            'Additional Commission 3 (%)': 'addCommRate3',
            'Commission Plan Type 4': 'commPlanType4',
            'Commission Rate 4 (%)': 'commRate4',
            'Additional Commission 4 (%)': 'addCommRate4'
        };

        headerRow.eachCell((cell, colNumber) => {
            const text = cell.text ? cell.text.trim() : '';
            if (fieldMapping[text]) {
                colMap[fieldMapping[text]] = colNumber;
            }
        });

        // Fallback to absolute index positioning
        const keysList = [
            'quotationNo', 'title', 'nameTh', 'nameEn', 'lineOfBusiness', 'riskLevel', 'occupationClass', 'effDate', 'effTime', 'endDate', 'endTime', 'language', 'copyCount',
            'address1', 'address2', 'country', 'province', 'district', 'subDistrict', 'zipCode', 'contactName', 'contactPosition', 'contactMobile', 'contactPhone', 'contactEmail',
            'productType', 'subProduct', 'ageAverage', 'minAge', 'maxAge', 'planNumber', 'planType', 'modeOfPayment',
            'channel', 'agentBrokerCode', 'salesTeam', 'salesName', 'erType', 'lossRatio', 'refundRate',
            'accTitle', 'accNameTh', 'accNameEn', 'accTaxId', 'accType', 'accHeadCountType', 'accHeadCountDesc', 'accLineOfBusiness', 'accRiskLevel', 'accOccupationClass',
            'commPlanType1', 'commRate1', 'addCommRate1',
            'commPlanType2', 'commRate2', 'addCommRate2',
            'commPlanType3', 'commRate3', 'addCommRate3',
            'commPlanType4', 'commRate4', 'addCommRate4'
        ];

        keysList.forEach((key, index) => {
            if (!colMap[key]) {
                colMap[key] = index + 1;
            }
        });

        // Parse data rows starting from row 2
        ws.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Skip headers

            const rowData = {};
            keysList.forEach(key => {
                const colIdx = colMap[key];
                const cell = row.getCell(colIdx);
                let val = cell.value;
                
                // Handle complex ExcelJS cell structures
                if (val && typeof val === 'object') {
                    if (val.result !== undefined) val = val.result;
                    else if (val.richText) val = val.richText.map(t => t.text).join('');
                    else if (val.text) val = val.text;
                }
                
                // Trim string values
                if (typeof val === 'string') {
                    val = val.trim();
                }
                
                // Ensure date formatting is clean (DD/MM/YYYY)
                if ((key === 'effDate' || key === 'endDate') && val instanceof Date) {
                    const d = val;
                    const dd = String(d.getDate()).padStart(2, '0');
                    const mm = String(d.getMonth() + 1).padStart(2, '0');
                    const yyyy = d.getFullYear();
                    val = `${dd}/${mm}/${yyyy}`;
                }

                rowData[key] = val !== null && val !== undefined ? val : '';
            });

            // Skip empty rows
            const isRowEmpty = Object.values(rowData).every(v => v === '');
            if (isRowEmpty) return;

            // Validate fields
            const missing = [];
            requiredFields.forEach(f => {
                if (rowData[f] === undefined || rowData[f] === null || rowData[f] === '') {
                    missing.push(f);
                }
            });

            rowData.isValid = missing.length === 0;
            rowData.missingFields = missing;
            rowData.rowNum = rowNumber;

            cases.push(rowData);
        });

        res.json({ cases });
    } catch (error) {
        console.error('GISX File Upload Parse Error:', error);
        res.status(500).json({ error: 'ไม่สามารถประมวลผลไฟล์ Excel ได้: ' + error.message });
    }
});

// ==================== GISX Run Automation (SSE Streaming) ====================
const { spawn } = require('child_process');

// Active run jobs: jobId -> { proc, clients }
const gisxJobs = new Map();

// 3. POST /api/gisx/run — start a new automation job
app.post('/api/gisx/run', async (req, res) => {
    try {
        const { cases, headless } = req.body;

        if (!cases || !Array.isArray(cases) || cases.length === 0) {
            return res.status(400).json({ error: 'ไม่พบข้อมูล cases ที่ต้องการรัน' });
        }

        // Write cases to a temp JSON file
        const jobId = 'gisx_job_' + Date.now();
        const tmpDir = path.join(__dirname, 'gisx_tmp');
        if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true });

        const inputFile = path.join(tmpDir, jobId + '_cases.json');
        fs.writeFileSync(inputFile, JSON.stringify(cases, null, 2), 'utf8');

        const screenshotDir = path.join(tmpDir, jobId + '_screenshots');

        // Build CLI args
        const scriptArgs = [
            path.join(__dirname, 'run_create_gisx.js'),
            '--input', inputFile,
            '--screenshotDir', screenshotDir
        ];
        if (headless) scriptArgs.push('--headless');

        console.log(`[GISX RUN] Starting job ${jobId} with ${cases.length} case(s)...`);

        const proc = spawn('node', scriptArgs, {
            cwd: __dirname,
            env: { ...process.env }
        });

        gisxJobs.set(jobId, { proc, clients: new Set(), inputFile, screenshotDir });

        proc.stdout.on('data', (data) => {
            const text = data.toString();
            process.stdout.write('[GISX PROC] ' + text);
            broadcastJobLog(jobId, 'stdout', text);
        });

        proc.stderr.on('data', (data) => {
            const text = data.toString();
            process.stderr.write('[GISX PROC ERR] ' + text);
            broadcastJobLog(jobId, 'stderr', text);
        });

        proc.on('close', (code) => {
            console.log(`[GISX RUN] Job ${jobId} finished with code ${code}`);
            broadcastJobLog(jobId, 'done', `\n[KUMA AUTO] ✅ กระบวนการเสร็จสิ้นแล้วครับเมี้ยว (exit code: ${code})\n`);

            // Read and broadcast results if available
            try {
                const resultsFile = path.join(screenshotDir, 'batch_results.json');
                if (fs.existsSync(resultsFile)) {
                    const results = JSON.parse(fs.readFileSync(resultsFile, 'utf8'));
                    broadcastJobLog(jobId, 'results', JSON.stringify(results));
                }
            } catch (e) {}

            // Cleanup temp input file after 5 minutes
            setTimeout(() => {
                try { fs.unlinkSync(inputFile); } catch (e) {}
                gisxJobs.delete(jobId);
            }, 5 * 60 * 1000);
        });

        proc.on('error', (err) => {
            console.error(`[GISX RUN] Job ${jobId} spawn error:`, err);
            broadcastJobLog(jobId, 'stderr', `\n[ERROR] ไม่สามารถเริ่ม Playwright ได้: ${err.message}\n`);
        });

        res.json({ jobId, message: `เริ่มรัน Automation แล้วครับเมี้ยว! Job ID: ${jobId}` });

    } catch (error) {
        console.error('GISX Run Error:', error);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการเริ่มรัน: ' + error.message });
    }
});

// 4. GET /api/gisx/run/:jobId/stream — SSE log stream
app.get('/api/gisx/run/:jobId/stream', (req, res) => {
    const { jobId } = req.params;
    const job = gisxJobs.get(jobId);

    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.setHeader('X-Accel-Buffering', 'no');
    res.flushHeaders();

    if (!job) {
        res.write(`data: ${JSON.stringify({ type: 'stderr', text: 'ไม่พบ Job นี้ในระบบ หรือรันเสร็จไปแล้วครับ' })}\n\n`);
        res.end();
        return;
    }

    job.clients.add(res);

    req.on('close', () => {
        job.clients.delete(res);
    });
});

// 5. POST /api/gisx/run/:jobId/stop — stop a running job
app.post('/api/gisx/run/:jobId/stop', (req, res) => {
    const { jobId } = req.params;
    const job = gisxJobs.get(jobId);
    if (!job) {
        return res.status(404).json({ error: 'ไม่พบ Job นี้' });
    }
    try {
        job.proc.kill('SIGTERM');
        broadcastJobLog(jobId, 'stderr', '\n[KUMA AUTO] ⛔ ผู้ใช้หยุด Automation แล้วครับ\n');
        res.json({ message: 'หยุดการรัน Automation แล้ว' });
    } catch (e) {
        res.status(500).json({ error: 'ไม่สามารถหยุดได้: ' + e.message });
    }
});

function broadcastJobLog(jobId, type, text) {
    const job = gisxJobs.get(jobId);
    if (!job) return;
    const payload = JSON.stringify({ type, text });
    for (const client of job.clients) {
        try {
            client.write(`data: ${payload}\n\n`);
        } catch (e) {
            job.clients.delete(client);
        }
    }
}

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});

