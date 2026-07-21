/**
 * GISX Register Case Automation Script (Batch Mode)
 * 🐾 Powered by KUMA Test Case Builder
 *
 * Usage:
 *   node run_create_gisx.js --input cases.json [--headless]
 *
 * Input file: JSON array of case objects (exported from KUMA Web)
 */

const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

// ---- Parse CLI args ----
const args = process.argv.slice(2);
const inputFlag = args.indexOf('--input');
const inputFile = inputFlag !== -1 ? args[inputFlag + 1] : null;
const headless = args.includes('--headless');

// Screenshot output dir (can be overridden by --screenshotDir)
const sdFlag = args.indexOf('--screenshotDir');
const screenshotDir = sdFlag !== -1
    ? args[sdFlag + 1]
    : path.join(__dirname, 'gisx_screenshots_' + Date.now());

// Credentials (can be overridden by env)
const GISX_USERNAME = process.env.GISX_USERNAME || 'mtl807032';
const GISX_PASSWORD = process.env.GISX_PASSWORD || 'Love35795wissbank@2';
const GISX_BASE_URL = process.env.GISX_BASE_URL || 'https://gisx-qa.muangthai.co.th';

// ---- Load cases ----
let cases = [];
if (inputFile && fs.existsSync(inputFile)) {
    try {
        cases = JSON.parse(fs.readFileSync(inputFile, 'utf8'));
        console.log(`[KUMA AUTO] Loaded ${cases.length} case(s) from: ${inputFile}`);
    } catch (e) {
        console.error('[KUMA AUTO] Failed to parse input JSON:', e.message);
        process.exit(1);
    }
} else {
    // Fallback demo case
    console.log('[KUMA AUTO] No --input file provided. Running with demo case...');
    cases = [{
        quotationNo: 'QT20260721-DEMO',
        title: 'นาย',
        nameTh: 'สมชาย มั่งคั่งคุมะ',
        nameEn: 'Somchai MangkangKuma',
        lineOfBusiness: 'Ordinary',
        riskLevel: 'Low',
        occupationClass: 'Class 1',
        effDate: '01/01/2026',
        effTime: '00:00:00',
        endDate: '31/12/2026',
        endTime: '23:59:59',
        language: 'Thai',
        copyCount: 1,
        planNumber: 4
    }];
}

// ---- Create screenshot dir ----
if (!fs.existsSync(screenshotDir)) {
    fs.mkdirSync(screenshotDir, { recursive: true });
}
console.log(`[KUMA AUTO] Screenshots will be saved to: ${screenshotDir}`);

// ---- Results tracking ----
const results = [];

(async () => {
    console.log('[KUMA AUTO] Launching browser...');
    const browser = await chromium.launch({
        headless: headless,
        args: ['--start-maximized']
    });

    const context = await browser.newContext({ viewport: null });
    const page = await context.newPage();

    // ---------- Helpers ----------

    async function takeScreenshot(label) {
        try {
            const filename = `${label}_${Date.now()}.png`;
            const filepath = path.join(screenshotDir, filename);
            await page.screenshot({ path: filepath, fullPage: false });
            console.log(`[KUMA AUTO] 📸 Screenshot: ${filename}`);
        } catch (e) {
            console.log(`[KUMA AUTO] ⚠️  Screenshot failed: ${e.message}`);
        }
    }

    async function fillDropdown(dataQaName, valueText = null, index = 0) {
        try {
            // Apply QA environment mappings
            if (valueText) {
                const valClean = valueText.trim().toLowerCase();
                if (dataQaName.includes('line_of_business')) {
                    if (valClean.startsWith('o')) valueText = 'O'; // e.g. Ordinary -> O
                    else if (valClean.startsWith('g')) valueText = 'G'; // e.g. Group -> G
                    else if (valClean.startsWith('c')) valueText = 'P'; // e.g. Credit Life -> P (since codes start with P)
                    else if (valClean.startsWith('a')) valueText = 'F'; // e.g. Accident -> F (since codes start with F)
                } else if (dataQaName.includes('risk_level')) {
                    if (valClean === 'low') valueText = 'ความเสี่ยงต่ำ';
                    else if (valClean === 'medium') valueText = 'ความเสี่ยงปานกลาง';
                    else if (valClean === 'high') valueText = 'ความเสี่ยงสูง';
                } else if (dataQaName.includes('occup_classified') || dataQaName.includes('occupational_classification')) {
                    if (valClean.includes('1') || valClean.includes('ชั้น 1')) valueText = 'ประเภทอาชีพ ชั้น 1';
                    else if (valClean.includes('2') || valClean.includes('ชั้น 2')) valueText = 'ประเภทอาชีพ ชั้น 2';
                    else if (valClean.includes('3') || valClean.includes('ชั้น 3')) valueText = 'ประเภทอาชีพ ชั้น 3';
                    else if (valClean.includes('4') || valClean.includes('ชั้น 4')) valueText = 'ประเภทอาชีพ ชั้น 4';
                } else if (dataQaName.includes('title')) {
                    if (valClean === 'บริษัท' || valClean.includes('บริษัท') || valClean.includes('บจก') || valClean.includes('บมจ')) {
                        valueText = 'บริษัท,บจก.,บมจ.';
                    }
                }
                if (valClean === 'select all' || valClean.includes('select all') || valClean === 'เลือกทั้งหมด' || valClean.includes('เลือกทั้งหมด')) {
                    valueText = 'Select All,เลือกทั้งหมด';
                }
            }

            console.log(`[KUMA AUTO]   → Dropdown "${dataQaName}" [index: ${index}] = "${valueText ?? '(first option)'}"`);
            await page.keyboard.press('Escape').catch(() => {});
            await page.waitForTimeout(300);

            const selector = `div[data-qa="${dataQaName}"] [data-qa="btn_dropdown_toggle_ddl"]`;
            const trigger = page.locator(selector).nth(index);
            await trigger.scrollIntoViewIfNeeded().catch(() => {});
            await trigger.click({ force: true });
            await page.waitForTimeout(800);

            // Find the active visible dropdown overlay in DOM
            let activeOverlay = null;
            for (let attempt = 0; attempt < 30; attempt++) {
                const overlays = page.locator('[data-qa="dropdown_overlay"]');
                const count = await overlays.count().catch(() => 0);
                for (let i = 0; i < count; i++) {
                    const ov = overlays.nth(i);
                    const isVisible = await ov.isVisible().catch(() => false);
                    if (isVisible) {
                        activeOverlay = ov;
                        break;
                    }
                }
                if (activeOverlay) break;
                await page.waitForTimeout(100);
            }
            if (!activeOverlay) {
                console.log('[KUMA AUTO] No visible dropdown overlay found via visibility check, using last() fallback');
                activeOverlay = page.locator('[data-qa="dropdown_overlay"]').last();
            }

            // Wait for at least one dropdown item inside this active overlay to render and become visible
            await activeOverlay.locator('[data-qa^="dropdown_item"], [id^="dropdown-overlay-item-"]').first()
                .waitFor({ state: 'visible', timeout: 5000 })
                .catch(() => {});

            // Dump dropdown options to console!
            try {
                const optionsText = await page.evaluate(() => {
                    const overlays = Array.from(document.querySelectorAll('[data-qa="dropdown_overlay"]'));
                    const activeOverlay = overlays.find(el => el.getBoundingClientRect().height > 0) || overlays[overlays.length - 1] || document;
                    const items = Array.from(activeOverlay.querySelectorAll('[data-qa^="dropdown_item"], [id^="dropdown-overlay-item-"]'));
                    return items.map(el => el.textContent.trim());
                });
                console.log(`[KUMA DUMP] Dropdown "${dataQaName}" options:`, JSON.stringify(optionsText));
            } catch (e) {
                console.log(`[KUMA DUMP] Dropdown "${dataQaName}" options evaluation failed:`, e.message);
            }
            
            // Check for Apply button presence in DOM (multi-select)
            const hasApplyBtn = (await activeOverlay
                .locator('[data-qa="btn_dropdown_confirm"], button:has-text("Apply"), button:has-text("ตกลง"), button:has-text("นำไปใช้"), button:has-text("OK")')
                .count().catch(() => 0)) > 0;

            if (valueText) {
                if (valueText === 'Select All,เลือกทั้งหมด') {
                    let selectAllItem = activeOverlay.getByText('Select all', { exact: false }).first();
                    if (!(await selectAllItem.isVisible().catch(() => false))) {
                        selectAllItem = activeOverlay.getByText('เลือกทั้งหมด', { exact: false }).first();
                    }

                    const isSelectAllVisible = await selectAllItem.isVisible().catch(() => false);
                    console.log(`[KUMA AUTO] "Select all" element visibility: ${isSelectAllVisible}`);

                    if (isSelectAllVisible) {
                        console.log(`[KUMA AUTO] Clicking "Select all" item in dropdown overlay for "${dataQaName}"...`);
                        await selectAllItem.scrollIntoViewIfNeeded().catch(() => {});
                        await selectAllItem.click().catch(() => selectAllItem.click({ force: true }));
                        await page.waitForTimeout(500);
                    } else {
                        const items = activeOverlay.locator('[data-qa^="dropdown_item"], [id^="dropdown-overlay-item-"]');
                        const count = await items.count().catch(() => 0);
                        console.log(`[KUMA AUTO] "Select all" item not found. Total items in overlay: ${count}`);
                        for (let i = 0; i < count; i++) {
                            const text = await items.nth(i).textContent().catch(() => '');
                            console.log(`[KUMA AUTO] Item ${i} text: "${text.trim()}"`);
                        }

                        console.log(`[KUMA AUTO] Selecting all ${count} items one by one in "${dataQaName}"...`);
                        for (let i = 0; i < count; i++) {
                            const text = await items.nth(i).textContent().catch(() => '');
                            if (text.toLowerCase().includes('select all') || text.includes('เลือกทั้งหมด')) {
                                continue;
                            }
                            await items.nth(i).scrollIntoViewIfNeeded().catch(() => {});
                            await items.nth(i).click().catch(() => items.nth(i).click({ force: true }));
                            await page.waitForTimeout(200);
                        }
                    }

                    if (hasApplyBtn) {
                        const applyBtn = activeOverlay.locator('[data-qa="btn_dropdown_confirm"], button:has-text("Apply"), button:has-text("ตกลง"), button:has-text("นำไปใช้"), button:has-text("OK")').first();
                        if (await applyBtn.isVisible().catch(() => false)) {
                            console.log(`[KUMA AUTO] Clicking Apply button in dropdown overlay...`);
                            await applyBtn.scrollIntoViewIfNeeded().catch(() => {});
                            await applyBtn.click().catch(() => applyBtn.click({ force: true }));
                            await page.waitForTimeout(1000);
                        }
                    }
                    return;
                }

                const alternatives = valueText.split(',').map(v => v.trim());

                if (hasApplyBtn) {
                    for (const alt of alternatives) {
                        const matched = await page.evaluate(({ value }) => {
                            const overlays = Array.from(document.querySelectorAll('[data-qa="dropdown_overlay"]'));
                            const activeOverlay = overlays.find(el => el.getBoundingClientRect().height > 0) || overlays[overlays.length - 1] || document;
                            const items = Array.from(activeOverlay.querySelectorAll('[data-qa^="dropdown_item"], [id^="dropdown-overlay-item-"]'));
                            let match = items.find(el => el.textContent.trim().toLowerCase() === value.toLowerCase());
                            if (!match) {
                                match = items.find(el => el.textContent.trim().toLowerCase().includes(value.toLowerCase()));
                            }
                            if (!match && value.length === 1) {
                                match = items.find(el => el.textContent.trim().toUpperCase().startsWith(value.toUpperCase()));
                            }
                            if (match) {
                                if (match.getAttribute('data-qa')) {
                                    return { type: 'data-qa', value: match.getAttribute('data-qa') };
                                }
                                if (match.getAttribute('id')) {
                                    return { type: 'id', value: match.getAttribute('id') };
                                }
                            }
                            return null;
                        }, { value: alt });

                        if (matched) {
                            let item;
                            if (matched.type === 'data-qa') {
                                item = activeOverlay.locator(`[data-qa="${matched.value}"]`).first();
                            } else {
                                item = activeOverlay.locator(`#${matched.value}`).first();
                            }
                            await item.click({ force: true });
                            await page.waitForTimeout(300);
                        }
                    }
                    const applyBtn = activeOverlay.locator('[data-qa="btn_dropdown_confirm"], button:has-text("Apply"), button:has-text("ตกลง"), button:has-text("นำไปใช้"), button:has-text("OK")').first();
                    if (await applyBtn.isVisible().catch(() => false)) {
                        await applyBtn.click({ force: true });
                        await page.waitForTimeout(600);
                    }
                } else {
                    let matched = null;
                    for (const alt of alternatives) {
                        matched = await page.evaluate(({ value }) => {
                            const overlays = Array.from(document.querySelectorAll('[data-qa="dropdown_overlay"]'));
                            const activeOverlay = overlays.find(el => el.getBoundingClientRect().height > 0) || overlays[overlays.length - 1] || document;
                            const items = Array.from(activeOverlay.querySelectorAll('[data-qa^="dropdown_item"], [id^="dropdown-overlay-item-"]'));
                            let match = items.find(el => el.textContent.trim().toLowerCase() === value.toLowerCase());
                            if (!match) {
                                match = items.find(el => el.textContent.trim().toLowerCase().includes(value.toLowerCase()));
                            }
                            if (!match && value.length === 1) {
                                match = items.find(el => el.textContent.trim().toUpperCase().startsWith(value.toUpperCase()));
                            }
                            if (match) {
                                if (match.getAttribute('data-qa')) {
                                    return { type: 'data-qa', value: match.getAttribute('data-qa') };
                                }
                                if (match.getAttribute('id')) {
                                    return { type: 'id', value: match.getAttribute('id') };
                                }
                            }
                            return null;
                        }, { value: alt });
                        if (matched) break;
                    }

                    if (matched) {
                        let option;
                        if (matched.type === 'data-qa') {
                            option = activeOverlay.locator(`[data-qa="${matched.value}"]`).first();
                        } else {
                            option = activeOverlay.locator(`#${matched.value}`).first();
                        }
                        await option.click({ force: true });
                        await page.waitForTimeout(600);
                    } else {
                        // Fallback: select first option in list
                        const firstOpt = activeOverlay.locator('[data-qa^="dropdown_item"], [id^="dropdown-overlay-item-"]').first();
                        if (await firstOpt.isVisible().catch(() => false)) {
                            await firstOpt.click({ force: true });
                            await page.waitForTimeout(600);
                        } else {
                            console.log(`[KUMA AUTO]   ⚠️  Option "${valueText}" not found in "${dataQaName}"`);
                            await page.keyboard.press('Escape').catch(() => {});
                        }
                    }
                }
            } else {
                // Pick first option
                const firstItem = activeOverlay.locator('[data-qa^="dropdown_item"], [id^="dropdown-overlay-item-"]').first();
                await firstItem.waitFor({ state: 'visible', timeout: 4000 }).catch(() => {});
                if (await firstItem.isVisible().catch(() => false)) {
                    await firstItem.click({ force: true });
                    await page.waitForTimeout(600);
                }
                if (hasApplyBtn) {
                    const applyBtn = activeOverlay.locator('[data-qa="btn_dropdown_confirm"], button:has-text("Apply"), button:has-text("ตกลง"), button:has-text("นำไปใช้"), button:has-text("OK")').first();
                    if (await applyBtn.isVisible().catch(() => false)) {
                        await applyBtn.click({ force: true });
                        await page.waitForTimeout(600);
                    }
                }
            }
        } catch (err) {
            console.log(`[KUMA AUTO]   ⚠️  Dropdown "${dataQaName}" error: ${err.message}`);
            await page.keyboard.press('Escape').catch(() => {});
        }
    }

    async function fillText(selector, value) {
        if (!value && value !== 0) return;
        try {
            const el = page.locator(selector).first();
            await el.scrollIntoViewIfNeeded().catch(() => {});
            await el.fill(String(value));
            await page.waitForTimeout(300);
        } catch (err) {
            console.log(`[KUMA AUTO]   ⚠️  fillText "${selector}" error: ${err.message}`);
        }
    }

    async function clickEl(selector) {
        try {
            const el = page.locator(selector).first();
            await el.scrollIntoViewIfNeeded().catch(() => {});
            await el.click({ force: true });
            await page.waitForTimeout(400);
        } catch (err) {
            console.log(`[KUMA AUTO]   ⚠️  click "${selector}" error: ${err.message}`);
        }
    }

    // ---------- Login once ----------
    console.log(`[KUMA AUTO] Navigating to ${GISX_BASE_URL}/new-business/register-case ...`);
    await page.goto(`${GISX_BASE_URL}/new-business/register-case`);

    try {
        console.log('[KUMA AUTO] Checking for login form...');
        // Wait up to 30 seconds for either the login form (username input) or the dashboard (e.g. Create button or logout)
        await Promise.race([
            page.waitForSelector('input[type="text"], input[name="username"]', { timeout: 30000 }),
            page.waitForSelector('button:has-text("Create"), button:has-text("สร้าง"), a:has-text("Create")', { timeout: 30000 })
        ]);

        if (await page.locator('input[type="text"], input[name="username"]').first().isVisible()) {
            console.log('[KUMA AUTO] Login page detected. Logging in...');
            await page.locator('input[type="text"], input[name="username"]').first().fill(GISX_USERNAME);
            await page.locator('input[type="password"]').first().fill(GISX_PASSWORD);

            const submitBtn = page.locator('input[type="submit"], #kc-login, button[type="submit"], button:has-text("Login")').first();
            await submitBtn.click();
            console.log('[KUMA AUTO] Clicked login. Waiting for navigation...');
            await page.waitForNavigation({ timeout: 30000 }).catch(() => {});
        } else {
            console.log('[KUMA AUTO] Already logged in (found dashboard controls).');
        }
    } catch (e) {
        console.log('[KUMA AUTO] Login check or step error:', e.message);
    }

    await page.waitForTimeout(4000);
    console.log(`[KUMA AUTO] Current URL after login: ${page.url()}`);
    await takeScreenshot('00_after_login');

    // ---------- Process each case ----------
    for (let caseIdx = 0; caseIdx < cases.length; caseIdx++) {
        const item = cases[caseIdx];
        const caseLabel = `case${String(caseIdx + 1).padStart(3, '0')}_${(item.quotationNo || 'unknown').replace(/[^a-z0-9]/gi, '_')}`;

        console.log(`\n[KUMA AUTO] ========================================`);
        console.log(`[KUMA AUTO] 🚀 Processing Case ${caseIdx + 1}/${cases.length}`);
        console.log(`[KUMA AUTO]    Quotation No : ${item.quotationNo}`);
        console.log(`[KUMA AUTO]    Name (TH)    : ${item.nameTh}`);
        console.log(`[KUMA AUTO]    Line of Biz  : ${item.lineOfBusiness}`);
        console.log(`[KUMA AUTO] ========================================`);

        let caseStatus = 'success';
        let caseError = null;

        try {
            // Navigate to Create Case page
            console.log('[KUMA AUTO] Navigating to Register Case create page...');
            await page.goto(`${GISX_BASE_URL}/new-business/register-case/create`).catch(async () => {
                // fallback: click Create button
                const btn = page.locator('button:has-text("Create"), button:has-text("สร้าง"), a:has-text("Create")').first();
                if (await btn.isVisible().catch(() => false)) {
                    await btn.click({ force: true });
                }
            });

            await page.waitForTimeout(6000);
            const currentUrl = page.url();
            console.log(`[KUMA AUTO] Create page URL: ${currentUrl}`);
            await takeScreenshot(`${caseLabel}_01_create_page`);

            if (!currentUrl.includes('/new-business/register-case/create')) {
                throw new Error(`ไม่สามารถเปิดหน้าสร้างข้อมูลเคสได้ (URL ปัจจุบัน: ${currentUrl})`);
            }

            // ====================================================
            // STEP 1 — Policy Detail Tab
            // ====================================================
            console.log('[KUMA AUTO] --- SUB-TAB: Policy Detail ---');
            await clickEl('a:has-text("Policy Detail"), button:has-text("Policy Detail")');
            await page.waitForTimeout(1000);

            // 1. Quotation No.
            await fillText(
                'input[placeholder*="Quotation" i], input[placeholder*="ใบเสนอราคา" i]',
                item.quotationNo
            );

            // 2. Policy Holder Title
            await fillDropdown(
                'field_type_dropdown_name_detail_policy.policy_info.policy_holder_title',
                item.title
            );

            // 3. Policy Holder Name (Thai)
            await fillText(
                'input[placeholder*="Holder Name" i], input[placeholder*="ชื่อผู้เอาประกัน" i]',
                item.nameTh
            );

            // 4. Policy Holder Name (English) — second input with same placeholder
            try {
                const nameEnInput = page.locator(
                    'input[placeholder*="Holder Name" i], input[placeholder*="ชื่อผู้เอาประกัน" i]'
                ).nth(1);
                await nameEnInput.fill(item.nameEn || '');
            } catch (e) {}

            // 5. Line of Business
            await fillDropdown(
                'field_type_dropdown_name_detail_policy.policy_info.line_of_business',
                item.lineOfBusiness
            );

            // 6. Risk Level
            await fillDropdown(
                'field_type_dropdown_name_detail_policy.policy_info.risk_level',
                item.riskLevel
            );

            // 7. Occupational Classification
            await fillDropdown(
                'field_type_dropdown_name_detail_policy.policy_info.occup_classified',
                item.occupationClass
            );

            // 8-9. Effective Date / End Date
            try {
                const dateInputs = page.locator('input[placeholder*="DD/MM/YYYY" i]');
                await dateInputs.nth(0).fill(item.effDate || '');
                await page.waitForTimeout(400);
                await dateInputs.nth(1).fill(item.endDate || '');
                await page.keyboard.press('Escape').catch(() => {});
                await page.waitForTimeout(500);
            } catch (e) {}

            // 10. Policy Language (checkbox/radio)
            if (item.language === 'Thai' || item.language === 'TH') {
                await clickEl('label:has-text("Thai"), label:has-text("ภาษาไทย"), input[value="Thai"], input[value="TH"]');
            } else {
                await clickEl('label:has-text("English"), input[value="English"], input[value="EN"]');
            }

            // 11. Copy of Policy
            await fillText('input[placeholder*="Copy" i], input[type="number"]', String(item.copyCount || 1));

            // 12. Address lines
            await fillText('input[name="detail_policy.policy_holder_address.address_1"]', item.address1 || '123/45 Kuma Tower');
            if (item.address2) {
                await fillText('input[name="detail_policy.policy_holder_address.address_2"]', item.address2);
            }

            // 13-16. Country > Province > District > Sub District
            await fillDropdown('field_type_dropdown_name_detail_policy.policy_holder_address.country', item.country);
            await fillDropdown('field_type_dropdown_name_detail_policy.policy_holder_address.province', item.province);
            await page.waitForTimeout(2000);
            await fillDropdown('field_type_dropdown_name_detail_policy.policy_holder_address.district', item.district);
            await page.waitForTimeout(2000);
            await fillDropdown('field_type_dropdown_name_detail_policy.policy_holder_address.sub_district', item.subDistrict);
            await page.waitForTimeout(1500);

            // 17. Zip Code (check if it has value first, do not modify if it does)
            try {
                const zipCodeInput = page.locator('input[name="detail_policy.policy_holder_address.zip_code"]').first();
                await zipCodeInput.waitFor({ state: 'visible', timeout: 5000 });
                const currentZip = await zipCodeInput.inputValue().catch(() => '');
                if (!currentZip) {
                    await zipCodeInput.fill(item.zipCode || '10310');
                } else {
                    console.log(`[KUMA AUTO] Zip Code already has value: "${currentZip}", skipping fill.`);
                }
            } catch (e) {
                console.log('[KUMA AUTO]   ⚠️  Zip Code input error:', e.message);
                await fillText('input[name="detail_policy.policy_holder_address.zip_code"]', item.zipCode || '10310').catch(() => {});
            }

            // 18-19. Contact
            await fillText('input[name="detail_policy.contact.name"]', item.contactName || item.nameTh || '');
            await fillText('input[name="detail_policy.contact.position"]', item.contactPosition || 'ผู้จัดการ');
            if (item.contactMobile) {
                await fillText('input[name="detail_policy.contact.mobile"]', item.contactMobile);
            }
            if (item.contactPhone) {
                await fillText('input[name="detail_policy.contact.phone"]', item.contactPhone);
            }
            if (item.contactEmail) {
                await fillText('input[name="detail_policy.contact.email"]', item.contactEmail);
            }

            await takeScreenshot(`${caseLabel}_02_policy_detail`);

            // ====================================================
            // STEP 1 — Coverage Tab
            // ====================================================
            console.log('[KUMA AUTO] --- SUB-TAB: Coverage ---');
            await clickEl('a:has-text("Coverage"), button:has-text("Coverage")');
            await page.waitForTimeout(1000);

            await fillDropdown('field_type_dropdown_name_coverage.coverage_info.product_type', item.productType || '01');
            await page.waitForTimeout(2000);
            await fillDropdown('field_type_dropdown_name_coverage.coverage_info.sub_product_type', item.subProduct);

            await fillText('div[data-qa="field_type_inputSelection_name_coverage.coverage_info.age_average"] input', item.ageAverage || '40');
            await fillDropdown('field_type_inputSelection_name_coverage.coverage_info.age_average');

            await fillText('div[data-qa="field_type_inputSelection_name_coverage.coverage_info.min_age"] input', item.minAge || '2');
            await fillDropdown('field_type_inputSelection_name_coverage.coverage_info.min_age');

            await fillText('div[data-qa="field_type_inputSelection_name_coverage.coverage_info.max_age"] input', item.maxAge || '80');
            await fillDropdown('field_type_inputSelection_name_coverage.coverage_info.max_age');

            try {
                await page.locator('input[name="coverage.coverage_table.plan_number"]').first().fill(String(item.planNumber || '1'));
            } catch (e) {
                await fillText('input[placeholder*="จำนวนแผน" i], input[placeholder*="plan" i]', String(item.planNumber || '1'));
            }

            try {
                const enterBtn = page.locator('div[data-qa="field_type_inputBtn_name_coverage.coverage_table.plan_number"] button').first();
                await enterBtn.click().catch(() => enterBtn.click({ force: true }));
                await page.waitForTimeout(2000);
            } catch (e) {}

            const ptSelector = 'div[data-qa="field_type_dropdown_name_coverage.coverage_table.plan_type"] [data-qa="btn_dropdown_toggle_ddl"]';
            const countPlanType = await page.locator(ptSelector).count().catch(() => 0);
            console.log(`[KUMA AUTO] Filling ${countPlanType} plan rows in Coverage table...`);
            for (let i = 0; i < countPlanType; i++) {
                await fillDropdown('field_type_dropdown_name_coverage.coverage_table.plan_type', item.planType || '1 :, 2 :, 3 :, 4 :, 5 :, 6 :', i);
                await fillDropdown('field_type_dropdown_name_coverage.coverage_table.mode_of_payment', item.modeOfPayment || 'Monthly, รายเดือน', i);
            }

            await clickEl('input[id$="-WAIVED"]');

            await takeScreenshot(`${caseLabel}_03_coverage`);

            // ====================================================
            // STEP 1 — Agent/Broker Tab
            // ====================================================
            console.log('[KUMA AUTO] --- SUB-TAB: Agent/Broker ---');
            await clickEl('a:has-text("Agent/Broker"), button:has-text("Agent/Broker")');
            await page.waitForTimeout(1000);

            await fillDropdown(
                'field_type_dropdown_name_agent_broker.agent_broker_info.channel',
                item.channel || 'Agent (บุคคลธรรมดา)'
            );

            try {
                const agentCodeInput = page.locator('input[name="agent_broker.agent_broker_info.agent_broker_code"]').first();
                await agentCodeInput.waitFor({ state: 'visible', timeout: 5000 });
                await agentCodeInput.fill(item.agentBrokerCode || '144660');
                await page.keyboard.press('Tab');
                await page.waitForTimeout(5000); // Wait for API call
            } catch (e) {
                console.log('[KUMA AUTO]   ⚠️  Agent Code input not found');
            }

            await fillDropdown('field_type_dropdown_name_agent_broker.mtl_sales_info.sales_team', item.salesTeam);
            await fillDropdown('field_type_dropdown_name_agent_broker.mtl_sales_info.sales_name', item.salesName);

            // Click "+ Add Commission Rate" and fill commission fields
            try {
                const addCommBtn = page.locator('button:has-text("Add Commission Rate"), button:has-text("Commission Rate"), button:has-text("คอมมิชชั่น")').first();
                const hasCustomComm = item.commPlanType1 || item.commRate1 || item.addCommRate1;
                
                if (hasCustomComm) {
                    const planCount = parseInt(item.planNumber || 1);
                    console.log(`[KUMA AUTO] Need to add custom commission rates for ${planCount} plans...`);
                    for (let i = 0; i < planCount; i++) {
                        if (await addCommBtn.isVisible().catch(() => false)) {
                            console.log(`[KUMA AUTO] Clicking + Add Commission Rate button (Index ${i})...`);
                            await addCommBtn.click({ force: true });
                            await page.waitForTimeout(1000);
                            
                            const planTypeDdl = `field_type_dropdown_name_agent_broker.commission_rate.${i}.plan_type`;
                            const userPlanType = item[`commPlanType${i+1}`];
                            await fillDropdown(planTypeDdl, userPlanType || null, 0);

                            const commRateVal = item[`commRate${i+1}`] !== undefined && item[`commRate${i+1}`] !== '' ? item[`commRate${i+1}`] : '10';
                            await fillText(`input[name="agent_broker.commission_rate.${i}.commission_rate"]`, String(commRateVal));

                            const addCommVal = item[`addCommRate${i+1}`] !== undefined && item[`addCommRate${i+1}`] !== '' ? item[`addCommRate${i+1}`] : '0';
                            await fillText(`input[name="agent_broker.commission_rate.${i}.additional_commission"]`, String(addCommVal));
                        }
                    }
                } else {
                    // Default behavior: Click once, select "Select All" and fill "10" for commission rate
                    if (await addCommBtn.isVisible().catch(() => false)) {
                        console.log('[KUMA AUTO] Clicking + Add Commission Rate button (Default Select All)...');
                        await addCommBtn.click({ force: true });
                        await page.waitForTimeout(1000);
                        
                        const planTypeDdl = `field_type_dropdown_name_agent_broker.commission_rate.0.plan_type`;
                        await fillDropdown(planTypeDdl, 'Select All,เลือกทั้งหมด', 0);

                        await fillText('input[name="agent_broker.commission_rate.0.commission_rate"]', '10');
                        await fillText('input[name="agent_broker.commission_rate.0.additional_commission"]', '0');
                    }
                }
                await takeScreenshot(`${caseLabel}_04_agent_broker_commission_filled`);
            } catch (e) {
                console.log('[KUMA AUTO]   ⚠️  Commission Rate section error:', e.message);
            }

            await takeScreenshot(`${caseLabel}_04_agent_broker`);

            // ====================================================
            // STEP 1 — Experience Refund Tab
            // ====================================================
            console.log('[KUMA AUTO] --- SUB-TAB: Experience Refund ---');
            await clickEl('a:has-text("Experience Refund"), button:has-text("Experience Refund")');
            await page.waitForTimeout(800);
            
            const isER = String(item.erType || '').toUpperCase() === 'ER';
            if (isER) {
                await clickEl('input[id$="-ER"]');
                await page.waitForTimeout(500);
                if (item.lossRatio) {
                    await fillText('input[name="reinsurance_detail.experience_refund.loss_ratio"]', item.lossRatio);
                }
                if (item.refundRate) {
                    await fillText('input[name="reinsurance_detail.experience_refund.refund_rate"]', item.refundRate);
                }
            } else {
                await clickEl('input[id$="-NON_ER"]');
            }

            await takeScreenshot(`${caseLabel}_05_step1_complete`);

            // ====================================================
            // STEP 1 → Click NEXT
            // ====================================================
            console.log('[KUMA AUTO] Clicking Next (Step 1 → Step 2)...');
            const nextBtn = page.locator('button:has-text("Next"), button:has-text("ถัดไป")').last();
            await nextBtn.scrollIntoViewIfNeeded().catch(() => {});
            await nextBtn.click({ force: true });
            await page.waitForTimeout(8000);
            await takeScreenshot(`${caseLabel}_06_step2_account`);

            // ====================================================
            // STEP 2 — Account Detail
            // ====================================================
            console.log('[KUMA AUTO] --- STEP 2: Account Detail ---');

            try {
                const accountSummaryTab = page.locator(
                    'a:has-text("Account Summary"), button:has-text("Account Summary")'
                ).first();
                await accountSummaryTab.scrollIntoViewIfNeeded().catch(() => {});
                await accountSummaryTab.click({ force: true });
                await page.waitForTimeout(2000);
            } catch (e) {}

            // Scroll down to find + Create Account button
            await page.evaluate(() => {
                const el = document.querySelector('main, .content, [class*="content"], [class*="main-content"]') || document.documentElement;
                el.scrollTop = el.scrollHeight;
                window.scrollTo(0, document.body.scrollHeight);
            });
            await page.waitForTimeout(1000);

            try {
                await page.evaluate(() => {
                    const btn = Array.from(document.querySelectorAll('button')).find(
                        b => b.textContent.includes('+ Create Account') || b.textContent.includes('Create Account')
                    );
                    if (btn) { btn.scrollIntoView({ behavior: 'smooth', block: 'center' }); btn.click(); }
                });
                await page.waitForTimeout(4000);
                await takeScreenshot(`${caseLabel}_07_step2_create_account`);

                // FILL ACCOUNT DETAIL MODAL
                let modal = page.locator('[role="dialog"], div.fixed, div.absolute').filter({ hasText: 'Account Detail' }).last();
                if (!(await modal.isVisible().catch(() => false))) {
                    console.log('[KUMA AUTO] Preferred modal locator not visible, trying fallback...');
                    modal = page.locator('[role="dialog"], div.fixed, div.absolute').first();
                    if (!(await modal.isVisible().catch(() => false))) {
                        console.log('[KUMA AUTO] Fallback modal locator not visible, using page-wide locators.');
                        modal = page;
                    }
                }

                // Dump all inputs inside modal to console so we can see them in logs
                try {
                    const modalInputs = await modal.evaluate((el) => {
                        const inputs = Array.from(el.querySelectorAll('input, select, textarea'));
                        return inputs.map((inp, idx) => ({
                            idx,
                            tagName: inp.tagName,
                            placeholder: inp.getAttribute('placeholder') || '',
                            name: inp.getAttribute('name') || '',
                            id: inp.getAttribute('id') || '',
                            dataQa: inp.getAttribute('data-qa') || '',
                            outerHTML: inp.outerHTML.substring(0, 200)
                        }));
                    });
                    console.log('[KUMA DUMP] Modal Inputs:', JSON.stringify(modalInputs, null, 2));
                } catch (e) {
                    console.log('[KUMA DUMP] Modal evaluation failed:', e.message);
                }
                
                // Fill Account Name (Thai)
                await page.locator('input[name="account_detail.account_information.account_name.th_TH"]').first().fill(item.accNameTh || item.nameTh || '');
                await page.waitForTimeout(300);

                // Fill Account Name (English)
                await page.locator('input[name="account_detail.account_information.account_name.en_US"]').first().fill(item.accNameEn || item.nameEn || '');
                await page.waitForTimeout(300);

                // Fill Tax ID
                await page.locator('input[name="account_detail.account_information.tax_id"]').first().fill(item.accTaxId || '1101800262649');
                await page.waitForTimeout(300);

                // Head count
                const hcType = String(item.accHeadCountType || 'Non Head Count').toLowerCase();
                if (hcType.includes('non')) {
                    await page.locator('input[value="NON_HEAD_COUNT"]').first().click({ force: true });
                } else {
                    await page.locator('input[value="HEAD_COUNT"]').first().click({ force: true });
                }
                await page.waitForTimeout(300);

                if (item.accHeadCountDesc) {
                    await page.locator('input[name="account_detail.account_information.head_count_desc"]').first().fill(item.accHeadCountDesc);
                    await page.waitForTimeout(300);
                }

                // Account Address (ตาม Policy Holder)
                await page.locator('input[value="POLICY_HOLDER"]').first().click({ force: true });
                await page.waitForTimeout(300);

                // ผู้รับ Invoice & Receipt (A : Account)
                await page.locator('input[value="A"]').first().click({ force: true });
                await page.waitForTimeout(300);

                // ชุดเอกสาร Invoice & Receipt (INCLUDE : รวมชุด)
                await page.locator('input[value="INCLUDE"]').first().click({ force: true });
                await page.waitForTimeout(500);

                // Dropdowns
                // 1. Account Title
                await fillDropdown('field_type_dropdown_name_account_detail.account_information.account_title', item.accTitle || 'บริษัท');
                await page.waitForTimeout(300);

                // 2. Account Type
                const titleVal = item.accTitle || item.title || 'นาย';
                let defaultAccType = 'Compulsory (Base Account)';
                await fillDropdown('field_type_dropdown_name_account_detail.account_information.account_type', item.accType || defaultAccType);
                await page.waitForTimeout(300);

                // 3. Line of Business (only if custom value specified)
                if (item.accLineOfBusiness && item.accLineOfBusiness !== 'Ordinary' && item.accLineOfBusiness !== 'O') {
                    await fillDropdown('field_type_dropdown_name_account_detail.account_information.line_of_business', item.accLineOfBusiness);
                    await page.waitForTimeout(300);
                }

                // 4. Risk Level (only if custom value specified)
                if (item.accRiskLevel && item.accRiskLevel !== 'Low') {
                    await fillDropdown('field_type_dropdown_name_account_detail.account_information.risk_level', item.accRiskLevel);
                    await page.waitForTimeout(300);
                }

                // 5. Occupational Classification (only if custom value specified)
                if (item.accOccupationClass && item.accOccupationClass !== 'Class 1') {
                    await fillDropdown('field_type_dropdown_name_account_detail.account_information.occupational_classification', item.accOccupationClass);
                    await page.waitForTimeout(1000);
                }

                // Fill Claim Payment Tab
                console.log('[KUMA AUTO] Switching to Claim Payment tab in modal...');
                let claimTab = page.locator('xpath=/html/body/div[2]/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div[2]/div[2]').first();
                if (!(await claimTab.isVisible().catch(() => false))) {
                    claimTab = page.locator('[role="dialog"] :text-is("Claim Payment")').first();
                    if (!(await claimTab.isVisible().catch(() => false))) {
                        console.log('[KUMA AUTO] Preferred claimTab locator not visible, trying fallback...');
                        claimTab = page.locator('[role="dialog"] p, [role="dialog"] span, [role="dialog"] button, [role="dialog"] [role="tab"]')
                            .filter({ hasText: /^Claim Payment$/ })
                            .first();
                    }
                }
                await claimTab.scrollIntoViewIfNeeded().catch(() => {});
                await claimTab.click({ force: true });
                await page.waitForTimeout(1500);

                // Fill Claim Payment fields
                console.log('[KUMA AUTO] Filling Claim Payment fields...');
                // 1. Plan Type
                await fillDropdown('field_type_dropdown_name_claim_payment_object.claim_payment.0.plan_type', 'Select All,เลือกทั้งหมด');
                await page.waitForTimeout(300);

                // 2. Payment Type (Select first option)
                await fillDropdown('field_type_dropdown_name_claim_payment_object.claim_payment.0.payment_type', '(first option)');
                await page.waitForTimeout(300);

                // 3. Paid To (Select first option)
                await fillDropdown('field_type_dropdown_name_claim_payment_object.claim_payment.0.paid_type', '(first option)');
                await page.waitForTimeout(500);

                // Submit modal
                const modalSubmitBtn = page.locator('#account-detail-modal-content_Account\\ Detail button:has-text("Submit"), button:has-text("Submit"), button:has-text("ตกลง"), button:has-text("บันทึก")').first();
                await modalSubmitBtn.click().catch(() => modalSubmitBtn.click({ force: true }));
                await page.waitForTimeout(4000);

                // Verify modal is closed
                if (await page.locator('#account-detail-modal-content_Account\\ Detail').isVisible().catch(() => false)) {
                    const errors = await page.locator('#account-detail-modal-content_Account\\ Detail #helper-text').allTextContents().catch(() => []);
                    throw new Error(`Modal submission failed. Validation errors: ${errors.join(', ')}`);
                }

                await takeScreenshot(`${caseLabel}_07_step2_account_saved`);

            } catch (e) {
                console.log('[KUMA AUTO]   ⚠️  Create Account error:', e.message);
                throw e; // rethrow to abort remaining steps on failure!
            }

            // ====================================================
            // STEP 3 — Document Upload (click NEXT to proceed)
            // ====================================================
            console.log('[KUMA AUTO] --- STEP 3: Upload Document → Click Next ---');
            const nextBtn3 = page.locator('button:has-text("Next"), button:has-text("ถัดไป")').last();
            await nextBtn3.scrollIntoViewIfNeeded().catch(() => {});
            await nextBtn3.click({ force: true });
            await page.waitForTimeout(6000);
            await takeScreenshot(`${caseLabel}_08_step4_summary`);

            // ====================================================
            // STEP 4 — Case Summary → Submit
            // ====================================================
            console.log('[KUMA AUTO] --- STEP 4: Case Summary → Submit ---');
            const submitBtn = page.locator(
                'button:has-text("Submit"), button:has-text("Confirm"), button:has-text("บันทึก"), button:has-text("ตกลง"), button:has-text("สร้าง")'
            ).first();

            if (await submitBtn.isVisible().catch(() => false)) {
                console.log('[KUMA AUTO] Clicking final Submit/Confirm button...');
                await submitBtn.click({ force: true });
                await page.waitForTimeout(8000);
                await takeScreenshot(`${caseLabel}_09_DONE`);
                console.log(`[KUMA AUTO] ✅ Case ${caseIdx + 1} submitted successfully!`);
            } else {
                console.log('[KUMA AUTO]   ℹ️  Submit button not visible — may require manual review.');
                await takeScreenshot(`${caseLabel}_09_submit_not_found`);
            }

        } catch (err) {
            caseStatus = 'error';
            caseError = err.message;
            console.error(`[KUMA AUTO] ❌ Case ${caseIdx + 1} FAILED: ${err.message}`);
            await takeScreenshot(`${caseLabel}_ERROR`);
        }

        results.push({
            caseIndex: caseIdx + 1,
            quotationNo: item.quotationNo,
            nameTh: item.nameTh,
            status: caseStatus,
            error: caseError
        });

        // Brief pause between cases
        if (caseIdx < cases.length - 1) {
            console.log('[KUMA AUTO] Waiting 3s before next case...');
            await page.waitForTimeout(3000);
        }
    }

    // ---- Print Summary ----
    console.log('\n[KUMA AUTO] ============ BATCH SUMMARY ============');
    results.forEach(r => {
        const icon = r.status === 'success' ? '✅' : '❌';
        console.log(`${icon}  Case ${r.caseIndex}: ${r.quotationNo} | ${r.nameTh} → ${r.status}${r.error ? ' — ' + r.error : ''}`);
    });
    console.log(`[KUMA AUTO] Total: ${results.length} | Success: ${results.filter(r => r.status === 'success').length} | Failed: ${results.filter(r => r.status !== 'success').length}`);
    console.log(`[KUMA AUTO] Screenshots saved to: ${screenshotDir}`);

    // Save results JSON
    const resultsFile = path.join(screenshotDir, 'batch_results.json');
    fs.writeFileSync(resultsFile, JSON.stringify(results, null, 2), 'utf8');
    console.log(`[KUMA AUTO] Results JSON saved: ${resultsFile}`);

    console.log('\n[KUMA AUTO] Browser will remain open for 30 seconds for review...');
    await page.waitForTimeout(30000);
    await browser.close();
    console.log('[KUMA AUTO] Done. Browser closed.');
})();
