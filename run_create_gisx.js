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
const defaultDatasetPath = path.join(__dirname, 'dataset', 'cases.json');
const targetInputFile = inputFile || (fs.existsSync(defaultDatasetPath) ? defaultDatasetPath : null);

if (targetInputFile && fs.existsSync(targetInputFile)) {
    try {
        cases = JSON.parse(fs.readFileSync(targetInputFile, 'utf8'));
        console.log(`[KUMA AUTO] Loaded ${cases.length} case(s) from: ${targetInputFile}`);
    } catch (e) {
        console.error('[KUMA AUTO] Failed to parse input JSON:', e.message);
        process.exit(1);
    }
} else {
    // Fallback demo case
    console.log('[KUMA AUTO] No dataset/cases.json or --input file provided. Running with demo case...');
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
        planNumber: 4,
        accTitle: 'บริษัท',
        accNameTh: 'บริษัท สมชาย มั่งคั่งคุมะ จำกัด',
        accNameEn: 'Somchai MangkangKuma Co., Ltd.',
        accTaxId: '0105561000123',
        accType: 'Compulsory (Base Account)',
        accHeadCountType: 'Non Head Count',
        accLineOfBusiness: 'Ordinary',
        accRiskLevel: 'Low',
        accOccupationClass: 'Class 1'
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
    let browser;
    try {
        console.log('[KUMA AUTO] Launching browser...');
        browser = await chromium.launch({
            headless: headless,
            args: ['--start-maximized']
        });
    } catch (launchErr) {
        if (launchErr.message.includes("Executable doesn't exist") || launchErr.message.includes("Please run the following command")) {
            console.log('[KUMA AUTO] Playwright browser not found. Automatically running "npx playwright install chromium" to install on the fly...');
            try {
                const { execSync } = require('child_process');
                execSync('npx playwright install chromium', { stdio: 'inherit' });
                console.log('[KUMA AUTO] Browser installed successfully! Retrying launch...');
                browser = await chromium.launch({
                    headless: headless,
                    args: ['--start-maximized']
                });
            } catch (installErr) {
                console.error('[KUMA AUTO] Failed to install Playwright browser automatically:', installErr.message);
                throw launchErr;
            }
        } else {
            throw launchErr;
        }
    }

    const context = await browser.newContext({ viewport: null });
    const page = await context.newPage();

    // Listen to network request failures and console errors
    page.on('console', msg => {
        if (msg.type() === 'error' || msg.text().includes('failed') || msg.text().includes('Error')) {
            console.log(`[BROWSER CONSOLE] [${msg.type()}] ${msg.text()}`);
        }
    });
    page.on('requestfailed', request => {
        console.log(`[BROWSER REQUEST FAILED] ${request.url()} | Error: ${request.failure()?.errorText || 'unknown'}`);
    });
    page.on('response', response => {
        const status = response.status();
        if (status >= 400) {
            console.log(`[BROWSER RESPONSE ERROR] ${response.url()} | Status: ${status}`);
        }
    });

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
        if (valueText === undefined || valueText === null || valueText === '') {
            return;
        }
        const safeClick = async (loc) => {
            await loc.scrollIntoViewIfNeeded().catch(() => {});
            try {
                await loc.click();
            } catch (err) {
                console.log(`[KUMA AUTO]     safeClick primary failed: ${err.message}. Retrying with force: true...`);
                await loc.click({ force: true });
            }
        };
        try {
            // Apply QA environment mappings
            if (valueText) {
                const valClean = valueText.trim().toLowerCase();
                if (dataQaName.includes('line_of_business')) {
                    if (valClean.startsWith('o')) valueText = 'TINSU'; // Map Ordinary to TINSU to get standard occupations
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
                        valueText = 'บริษัทจำกัด,บริษัท,บจก.,บมจ.';
                    }
                }
                if (valClean === 'select all' || valClean.includes('select all') || valClean === 'เลือกทั้งหมด' || valClean.includes('เลือกทั้งหมด')) {
                    valueText = 'Select All,เลือกทั้งหมด';
                }
            }

            let targetOptIndex = 0;
            if (valueText && valueText.startsWith('(nth option:')) {
                const match = valueText.match(/\d+/);
                if (match) {
                    targetOptIndex = parseInt(match[0], 10);
                }
                valueText = null;
            } else if (valueText === '(first option)') {
                valueText = null;
            }

            console.log(`[KUMA AUTO]   → Dropdown "${dataQaName}" [index: ${index}] = "${valueText ?? ('(option index: ' + targetOptIndex + ')')}"`);
            
            // Wait for any loading spinner to disappear
            await page.waitForTimeout(500);
            try {
                await page.locator('.ant-spin-spinning, .ant-spin, [class*="loading"], [class*="spinner"]').first().waitFor({ state: 'hidden', timeout: 10000 });
            } catch (e) {}

            // Wait for Ant Design select loading/disabled class to disappear
            try {
                const selectEl = page.locator(`div[data-qa="${dataQaName}"] .ant-select`).first();
                if (await selectEl.count().catch(() => 0) > 0) {
                    for (let i = 0; i < 80; i++) {
                        const classes = await selectEl.getAttribute('class').catch(() => '');
                        if (!classes.includes('loading') && !classes.includes('disabled')) {
                            break;
                        }
                        await page.waitForTimeout(100);
                    }
                }
            } catch (e) {}

            await page.keyboard.press('Escape').catch(() => {});
            await page.waitForTimeout(300);

            const selector = `div[data-qa="${dataQaName}"] [data-qa="btn_dropdown_toggle_ddl"]`;
            const triggerCount = await page.locator(selector).count().catch(() => 0);
            console.log(`[KUMA AUTO]   Trigger count for "${dataQaName}": ${triggerCount}`);
            const trigger = page.locator(selector).nth(index);

            for (let attempt = 1; attempt <= 3; attempt++) {
                try {
                    await trigger.scrollIntoViewIfNeeded().catch(() => {});
                    await trigger.click({ force: true });
                    await page.waitForTimeout(800);

                    // Scroll the modal and page to the bottom ONLY if we are inside a modal dialog (Account Detail modal)
                    try {
                        const isInModal = dataQaName.includes('account_detail') || dataQaName.includes('claim_payment_object') || dataQaName.includes('account_information');
                        if (isInModal) {
                            console.log('[KUMA AUTO]   Scrolling modal/page to bottom inside fillDropdown...');
                            await page.evaluate(() => {
                                const els = Array.from(document.querySelectorAll('*'));
                                els.forEach(el => {
                                    if (el.scrollHeight > el.clientHeight) {
                                        el.scrollTop = el.scrollHeight;
                                    }
                                });
                                window.scrollTo(0, document.body.scrollHeight);
                            });
                            await page.waitForTimeout(800);
                        }
                    } catch (e) {}

                // Find the active visible dropdown overlay in DOM
                const activeOverlay = page.locator('[data-qa="dropdown_overlay"]:visible, .ant-select-dropdown:not(.ant-select-dropdown-hidden)').last();
                await activeOverlay.waitFor({ state: 'visible', timeout: 5000 }).catch(() => {});

                // Wait for at least one dropdown item inside this active overlay to render and become visible
                await activeOverlay.locator('[data-qa^="dropdown_item"], [id^="dropdown-overlay-item-"], .ant-select-item-option').first()
                    .waitFor({ state: 'visible', timeout: 5000 })
                    .catch(() => {});

                // Scroll dropdown overlay to load all paginated items (Ant Design virtual list / select dropdown)
                try {
                    console.log('[KUMA AUTO]   Attempting to scroll overlay to load all options...');
                    for (let s = 0; s < 10; s++) {
                        await page.evaluate(() => {
                            const overlays = Array.from(document.querySelectorAll('[data-qa="dropdown_overlay"]'));
                            const activeOverlay = overlays.find(el => {
                                const rect = el.getBoundingClientRect();
                                const style = window.getComputedStyle(el);
                                return rect.height > 0 && rect.width > 0 && style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0' && !el.className.includes('hidden') && !el.classList.contains('ant-select-dropdown-hidden');
                            }) || overlays[overlays.length - 1];
                            if (activeOverlay) {
                                const scrollable = Array.from(activeOverlay.querySelectorAll('*')).find(
                                    el => el.scrollHeight > el.clientHeight && 
                                          (window.getComputedStyle(el).overflowY === 'auto' || 
                                           window.getComputedStyle(el).overflowY === 'scroll' || 
                                           el.classList.contains('rc-virtual-list-holder') || 
                                           el.tagName === 'UL')
                                ) || activeOverlay;
                                if (scrollable) {
                                    scrollable.scrollTop = scrollable.scrollHeight;
                                }
                            }
                        });
                        await page.waitForTimeout(150);
                    }
                } catch (e) {
                    console.log('[KUMA AUTO]   Scroll error:', e.message);
                }

                // Dump dropdown options to console!
                try {
                    const optionsText = await page.evaluate(() => {
                        const overlays = Array.from(document.querySelectorAll('[data-qa="dropdown_overlay"]'));
                        const activeOverlay = overlays.find(el => {
                            const rect = el.getBoundingClientRect();
                            const style = window.getComputedStyle(el);
                            return rect.height > 0 && rect.width > 0 && style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0' && !el.className.includes('hidden') && !el.classList.contains('ant-select-dropdown-hidden');
                        }) || overlays[overlays.length - 1] || document;
                        const items = Array.from(activeOverlay.querySelectorAll('[data-qa^="dropdown_item"], [id^="dropdown-overlay-item-"], .ant-select-item-option'));
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
                        console.log(`[KUMA AUTO] Selecting all items one-by-one and clicking Apply browser-side...`);
                        const result = await page.evaluate(() => {
                            const overlays = Array.from(document.querySelectorAll('[data-qa="dropdown_overlay"]'));
                            const activeOverlay = overlays.find(el => {
                                const rect = el.getBoundingClientRect();
                                const style = window.getComputedStyle(el);
                                return rect.height > 0 && rect.width > 0 && style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0' && !el.className.includes('hidden') && !el.classList.contains('ant-select-dropdown-hidden');
                            }) || overlays[overlays.length - 1];
                            if (!activeOverlay) return { clicked: 0, applied: false };

                            const items = Array.from(activeOverlay.querySelectorAll('[data-qa^="dropdown_item"], [id^="dropdown-overlay-item-"], .ant-select-item-option'));
                            let clicked = 0;
                            items.forEach(el => {
                                const text = el.textContent.trim().toLowerCase();
                                if (text.includes('select all') || text.includes('เลือกทั้งหมด')) return;

                                const container = el.closest('.ant-select-item-option, .ant-select-item, [class*="option"], [class*="item"]') || el;
                                const checkbox = container.querySelector('input[type="checkbox"]');
                                const isSelected = el.classList.contains('ant-select-item-option-selected') || 
                                                   el.getAttribute('aria-selected') === 'true' || 
                                                   el.classList.contains('selected') ||
                                                   container.classList.contains('ant-select-item-option-selected') || 
                                                   container.getAttribute('aria-selected') === 'true' || 
                                                   container.classList.contains('selected') ||
                                                   container.classList.contains('checked') ||
                                                   (checkbox && checkbox.checked) ||
                                                   !!container.querySelector('.ant-select-item-option-state-icon, [class*="selected-icon"], [class*="checked-icon"]');

                                if (!isSelected) {
                                    // Target the text content span or checkbox container for maximum click compatibility
                                    const clickTarget = container.querySelector('.ant-select-item-option-content') || checkbox || el;
                                    clickTarget.click();
                                    clicked++;
                                }
                            });

                            // Find and click the Apply button browser-side immediately
                            const applyBtn = activeOverlay.querySelector('[data-qa="btn_dropdown_confirm"]') ||
                                             Array.from(activeOverlay.querySelectorAll('button, div, span')).find(btn => {
                                                 const txt = btn.textContent.trim().toLowerCase();
                                                 return txt === 'apply' || txt === 'ตกลง' || txt === 'นำไปใช้' || txt === 'ok';
                                             });
                            let applied = false;
                            if (applyBtn) {
                                applyBtn.click();
                                applied = true;
                            }
                            return { clicked, applied };
                        });
                        console.log(`[KUMA AUTO] Browser-clicked ${result.clicked} options. Apply button clicked browser-side: ${result.applied}`);
                        await page.waitForTimeout(1000);
                        return;
                    }

                    const alternatives = valueText.split(',').map(v => v.trim());
                    console.log(`[KUMA AUTO]   → alternatives to select: ${JSON.stringify(alternatives)} | hasApplyBtn: ${hasApplyBtn}`);

                    if (hasApplyBtn) {
                        for (const alt of alternatives) {
                            const matched = await page.evaluate(({ value }) => {
                                const overlays = Array.from(document.querySelectorAll('[data-qa="dropdown_overlay"]'));
                                const activeOverlay = overlays.find(el => {
                                    const rect = el.getBoundingClientRect();
                                    const style = window.getComputedStyle(el);
                                    return rect.height > 0 && rect.width > 0 && style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0' && !el.className.includes('hidden') && !el.classList.contains('ant-select-dropdown-hidden');
                                }) || overlays[overlays.length - 1] || document;
                                const items = Array.from(activeOverlay.querySelectorAll('[data-qa^="dropdown_item"], [id^="dropdown-overlay-item-"], .ant-select-item-option'));
                                let match = items.find(el => el.textContent.trim().toLowerCase() === value.toLowerCase());
                                if (!match) {
                                    match = items.find(el => el.textContent.trim().toLowerCase().startsWith(value.toLowerCase()));
                                }
                                if (!match && value.length > 1) {
                                    match = items.find(el => el.textContent.trim().toLowerCase().includes(value.toLowerCase()));
                                }
                                if (!match && /^\d+\s*:/.test(value)) {
                                    const prefix = value.match(/^\d+\s*:/)[0].trim().toLowerCase();
                                    match = items.find(el => el.textContent.trim().toLowerCase().startsWith(prefix));
                                }
                                if (match) {
                                    const container = match.closest('.ant-select-item-option, .ant-select-item, [class*="option"], [class*="item"]') || match;
                                    const checkbox = container.querySelector('input[type="checkbox"]');
                                    const isSelected = match.classList.contains('ant-select-item-option-selected') || 
                                                       match.getAttribute('aria-selected') === 'true' || 
                                                       match.classList.contains('selected') ||
                                                       container.classList.contains('ant-select-item-option-selected') || 
                                                       container.getAttribute('aria-selected') === 'true' || 
                                                       container.classList.contains('selected') ||
                                                       container.classList.contains('checked') ||
                                                       (checkbox && checkbox.checked) ||
                                                       !!container.querySelector('.ant-select-item-option-state-icon, [class*="selected-icon"], [class*="checked-icon"]');
                                    if (match.getAttribute('data-qa')) {
                                        return { type: 'data-qa', value: match.getAttribute('data-qa'), text: match.textContent.trim(), isSelected };
                                    }
                                    if (match.getAttribute('id')) {
                                        return { type: 'id', value: match.getAttribute('id'), text: match.textContent.trim(), isSelected };
                                    }
                                }
                                return null;
                            }, { value: alt });

                            if (matched) {
                                if (matched.isSelected) {
                                    console.log(`[KUMA AUTO]     Option "${alt}" ("${matched.text}") is already selected, skipping click.`);
                                } else {
                                    console.log(`[KUMA AUTO]     Matched "${alt}" to overlay item: "${matched.text}" (${matched.type}: ${matched.value})`);
                                    let item;
                                    if (matched.type === 'data-qa') {
                                        item = activeOverlay.locator(`[data-qa="${matched.value}"]`).first();
                                    } else {
                                        item = activeOverlay.locator(`#${matched.value}`).first();
                                    }
                                    await safeClick(item);
                                    await page.waitForTimeout(300);
                                }
                            } else {
                                console.log(`[KUMA AUTO]     ⚠️ Could not match alternative "${alt}"`);
                            }
                        }
                        // Scroll the modal and page to the bottom if we are in modal to make sure Apply is visible!
                        try {
                            const isInModal = dataQaName.includes('account_detail') || dataQaName.includes('claim_payment_object') || dataQaName.includes('account_information');
                            if (isInModal) {
                                console.log('[KUMA AUTO] Scrolling modal/page to bottom before clicking Apply (Alternatives)...');
                                await page.evaluate(() => {
                                    const els = Array.from(document.querySelectorAll('*'));
                                    els.forEach(el => {
                                        if (el.scrollHeight > el.clientHeight) {
                                            el.scrollTop = el.scrollHeight;
                                        }
                                    });
                                    window.scrollTo(0, document.body.scrollHeight);
                                });
                                await page.waitForTimeout(600);
                            }
                        } catch (e) {}

                        const applyBtn = activeOverlay.locator('[data-qa="btn_dropdown_confirm"], button:has-text("Apply"), button:has-text("ตกลง"), button:has-text("นำไปใช้"), button:has-text("OK")').first();
                        if (await applyBtn.isVisible().catch(() => false)) {
                            console.log(`[KUMA AUTO]     Clicking Apply button...`);
                            await safeClick(applyBtn);
                            await page.waitForTimeout(600);
                        } else {
                            console.log(`[KUMA AUTO]     ⚠️ Apply button not visible!`);
                        }
                    } else {
                        let matched = null;
                        for (const alt of alternatives) {
                            matched = await page.evaluate(({ value }) => {
                                const overlays = Array.from(document.querySelectorAll('[data-qa="dropdown_overlay"]'));
                                const activeOverlay = overlays.find(el => {
                                    const rect = el.getBoundingClientRect();
                                    const style = window.getComputedStyle(el);
                                    return rect.height > 0 && rect.width > 0 && style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0' && !el.className.includes('hidden') && !el.classList.contains('ant-select-dropdown-hidden');
                                }) || overlays[overlays.length - 1] || document;
                                const items = Array.from(activeOverlay.querySelectorAll('[data-qa^="dropdown_item"], [id^="dropdown-overlay-item-"], .ant-select-item-option'));
                                let match = items.find(el => el.textContent.trim().toLowerCase() === value.toLowerCase());
                                if (!match) {
                                    match = items.find(el => el.textContent.trim().toLowerCase().startsWith(value.toLowerCase()));
                                }
                                if (!match && value.length > 1) {
                                    match = items.find(el => el.textContent.trim().toLowerCase().includes(value.toLowerCase()));
                                }
                                if (!match && /^\d+\s*:/.test(value)) {
                                    const prefix = value.match(/^\d+\s*:/)[0].trim().toLowerCase();
                                    match = items.find(el => el.textContent.trim().toLowerCase().startsWith(prefix));
                                }
                                if (match) {
                                    if (match.getAttribute('data-qa')) {
                                        return { type: 'data-qa', value: match.getAttribute('data-qa'), text: match.textContent.trim() };
                                    }
                                    if (match.getAttribute('id')) {
                                        return { type: 'id', value: match.getAttribute('id'), text: match.textContent.trim() };
                                    }
                                }
                                return null;
                            }, { value: alt });
                            if (matched) break;
                        }

                        if (matched) {
                            console.log(`[KUMA AUTO]     Matched "${valueText}" to overlay item: "${matched.text}" (${matched.type}: ${matched.value})`);
                            let option;
                            if (matched.type === 'data-qa') {
                                option = activeOverlay.locator(`[data-qa="${matched.value}"]`).first();
                            } else {
                                option = activeOverlay.locator(`#${matched.value}`).first();
                            }
                            await safeClick(option);
                            await page.waitForTimeout(600);
                        } else {
                            // Fallback: select first option in list
                            const firstOpt = activeOverlay.locator('[data-qa^="dropdown_item"], [id^="dropdown-overlay-item-"], .ant-select-item-option').first();
                            if (await firstOpt.isVisible().catch(() => false)) {
                                await safeClick(firstOpt);
                                await page.waitForTimeout(600);
                            } else {
                                console.log(`[KUMA AUTO]   ⚠️  Option "${valueText}" not found in "${dataQaName}"`);
                                await page.keyboard.press('Escape').catch(() => {});
                            }
                        }
                    }
                } else {
                    // Pick Nth option (0-indexed)
                    const items = activeOverlay.locator('[data-qa^="dropdown_item"], [id^="dropdown-overlay-item-"], .ant-select-item-option');
                    const targetItem = items.nth(targetOptIndex);
                    await targetItem.waitFor({ state: 'visible', timeout: 4000 }).catch(() => {});
                    if (await targetItem.isVisible().catch(() => false)) {
                        await safeClick(targetItem);
                        await page.waitForTimeout(600);
                    }
                    if (hasApplyBtn) {
                        const applyBtn = activeOverlay.locator('[data-qa="btn_dropdown_confirm"], button:has-text("Apply"), button:has-text("ตกลง"), button:has-text("นำไปใช้"), button:has-text("OK")').first();
                        if (await applyBtn.isVisible().catch(() => false)) {
                            await safeClick(applyBtn);
                            await page.waitForTimeout(600);
                        }
                    }
                }

                // Check verification
                if (valueText) {
                    await page.waitForTimeout(500);
                    const triggerText = (await trigger.textContent().catch(() => '')).trim();
                    const isStillEmpty = !triggerText || 
                                         triggerText.toLowerCase().includes('select') || 
                                         triggerText.includes('เลือก') || 
                                         triggerText.toLowerCase().includes('empty') ||
                                         triggerText.includes('cannot be empty');
                    if (!isStillEmpty) {
                        console.log(`[KUMA AUTO]   Successfully selected value for "${dataQaName}": "${triggerText}"`);
                        break;
                    }
                    console.log(`[KUMA AUTO]   ⚠️ Selection attempt ${attempt} for "${dataQaName}" appears empty. Retrying...`);
                } else {
                    break;
                }
            } catch (err) {
                console.log(`[KUMA AUTO]   ⚠️ Dropdown "${dataQaName}" attempt ${attempt} error: ${err.message}`);
                await page.keyboard.press('Escape').catch(() => {});
            }
        }
        } catch (outerErr) {
            console.log(`[KUMA AUTO]   ⚠️ Outer Dropdown "${dataQaName}" error: ${outerErr.message}`);
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
            // Wait for any loading spinner to disappear
            await page.waitForTimeout(500);
            try {
                await page.locator('.ant-spin-spinning, .ant-spin, [class*="loading"], [class*="spinner"]').first().waitFor({ state: 'hidden', timeout: 10000 });
            } catch (e) {}

            const el = page.locator(selector).first();
            await el.scrollIntoViewIfNeeded().catch(() => {});
            await el.click({ force: true });
            await page.waitForTimeout(400);
        } catch (err) {
            console.log(`[KUMA AUTO]   ⚠️  click "${selector}" error: ${err.message}`);
        }
    }

    async function approveCase(caseLabel, caseIdx) {
        try {
            console.log('[KUMA AUTO] --- AUTOMATIC APPROVAL PROCESS ---');
            await page.waitForTimeout(5000); // Wait for redirection back to list page

            // Wait for register case list table to load
            console.log('[KUMA AUTO] Waiting for Register Case list table...');
            const tableLocator = page.locator('table tbody tr').first();
            await tableLocator.waitFor({ state: 'visible', timeout: 15000 }).catch(() => {});

            // Find and click the first row Case No cell (column index 1) to enter detail view
            const caseCell = page.locator('table tbody tr').first().locator('td').nth(1);
            const linkText = await caseCell.textContent().catch(() => '');
            console.log(`[KUMA AUTO] Clicking Case No cell text: "${linkText.trim()}" to enter detail page...`);
            
            const cellClickable = caseCell.locator('a, span, button, div').first();
            if (await cellClickable.count() > 0) {
                await cellClickable.click({ force: true });
            } else {
                await caseCell.click({ force: true });
            }
            await page.waitForTimeout(4000); // Wait for detail page to load

            // Wait for Approve button to be visible
            const approveBtn = page.locator('button:has-text("Approve"), button:has-text("อนุมัติ")').first();
            await approveBtn.waitFor({ state: 'visible', timeout: 10000 });

            console.log('[KUMA AUTO] Clicking Approve button...');
            await approveBtn.click({ force: true });
            await page.waitForTimeout(3000);

            // Handle Approve Confirmation dialog if it appears
            try {
                const confirmBtn = page.locator('button:has-text("Confirm"), button:has-text("Approve"), button:has-text("ตกลง"), button:has-text("ยืนยัน")').filter({ state: 'visible' }).last();
                if (await confirmBtn.isVisible().catch(() => false)) {
                    console.log('[KUMA AUTO] Approve confirmation pop-up detected. Clicking Confirm...');
                    await confirmBtn.click({ force: true });
                    await page.waitForTimeout(6000); // Wait for approval to save
                }
            } catch (e) {
                console.log('[KUMA AUTO] No approve confirmation pop-up handled:', e.message);
            }

            // Verify that the status changed to Approve (button becomes hidden and/or "Approve" text is found in status area)
            console.log('[KUMA AUTO] Verifying case status has changed to Approve...');
            try {
                await Promise.race([
                    approveBtn.waitFor({ state: 'hidden', timeout: 15000 }),
                    page.locator('span:text-is("Approve"), div:text-is("Approve"), span:text-is("อนุมัติ"), div:text-is("อนุมัติ")').first().waitFor({ state: 'visible', timeout: 15000 })
                ]);
                console.log('[KUMA AUTO] Case status verified as APPROVED.');
            } catch (verifErr) {
                throw new Error(`Approval verification timed out: ${verifErr.message}`);
            }

            await takeScreenshot(`${caseLabel}_10_APPROVED`);
            console.log(`[KUMA AUTO] ✅ Case ${caseIdx + 1} approved successfully!`);
        } catch (approveErr) {
            console.log('[KUMA AUTO] ⚠️ Automatic Approval failed:', approveErr.message);
            await takeScreenshot(`${caseLabel}_APPROVE_FAILED`);
            throw approveErr; // Rethrow to let the main case runner catch it and mark case as error!
        }
    }

    // ---------- Login once ----------
    console.log(`[KUMA AUTO] Navigating to ${GISX_BASE_URL}/new-business/register-case ...`);
    
    let loggedIn = false;
    for (let loginAttempt = 1; loginAttempt <= 5; loginAttempt++) {
        try {
            await page.goto(`${GISX_BASE_URL}/new-business/register-case`, { waitUntil: 'load', timeout: 30000 }).catch(() => {});
            
            const title = await page.title().catch(() => '');
            if (title.includes('unknown error') || title.includes('520') || title.includes('Error') || (await page.locator('text="Error code 520"').count().catch(() => 0)) > 0) {
                console.log(`[KUMA AUTO]   ⚠️  Cloudflare / Server error detected (Attempt ${loginAttempt}/5). Retrying in 10s...`);
                await page.waitForTimeout(10000);
                continue;
            }

            console.log('[KUMA AUTO] Checking for login form...');
            await Promise.race([
                page.waitForSelector('input[type="text"], input[name="username"]', { timeout: 20000 }),
                page.waitForSelector('button:has-text("Create"), button:has-text("สร้าง"), a:has-text("Create")', { timeout: 20000 })
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
            loggedIn = true;
            break;
        } catch (e) {
            console.log(`[KUMA AUTO]   ⚠️  Login attempt ${loginAttempt} error: ${e.message}`);
            await page.waitForTimeout(5000);
        }
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

            // Wait for key Quotation No. input element to render instead of hardcoded 6s
            await page.locator('input[placeholder*="Quotation" i], input[placeholder*="ใบเสนอราคา" i]').first().waitFor({ state: 'visible', timeout: 15000 }).catch(() => {});
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
            await page.waitForTimeout(200);

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
            await page.waitForTimeout(300);

            // 6. Risk Level
            await fillDropdown(
                'field_type_dropdown_name_detail_policy.policy_info.risk_level',
                item.riskLevel
            );
            await page.waitForTimeout(300);

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
            await page.waitForTimeout(300);
            // Fill District (with Province re-select retry if options are blank/empty)
            for (let retryDist = 0; retryDist < 3; retryDist++) {
                await fillDropdown('field_type_dropdown_name_detail_policy.policy_holder_address.district', item.district);
                await page.waitForTimeout(300);

                // Verify if selected successfully (check for any placeholder keywords)
                const labelText = await page.locator('div[data-qa="field_type_dropdown_name_detail_policy.policy_holder_address.district"] selection-item, div[data-qa="field_type_dropdown_name_detail_policy.policy_holder_address.district"] [class*="selection-item"], div[data-qa="field_type_dropdown_name_detail_policy.policy_holder_address.district"] #dropdown-label-ddl').first().textContent().catch(() => '');
                const isPlaceholder = !labelText || 
                                      labelText.toLowerCase().includes('select') || 
                                      labelText.includes('เลือก') || 
                                      labelText.trim() === '';
                if (!isPlaceholder) {
                    break;
                }
                
                // If it fails, select a DIFFERENT province to force the API to load!
                const fallbackProv = `(nth option: ${retryDist + 1})`;
                console.log(`[KUMA AUTO]   ⚠️  District dropdown selection failed or empty. Selecting a DIFFERENT Province "${fallbackProv}" to trigger new API fetch (Attempt ${retryDist + 1}/3)...`);
                await fillDropdown('field_type_dropdown_name_detail_policy.policy_holder_address.province', fallbackProv);
                await page.waitForTimeout(3000);
            }

            // Fill Sub District (with District re-select retry if options are blank/empty)
            for (let retrySub = 0; retrySub < 3; retrySub++) {
                await fillDropdown('field_type_dropdown_name_detail_policy.policy_holder_address.sub_district', item.subDistrict);
                await page.waitForTimeout(300);

                // Verify if selected successfully (check for any placeholder keywords)
                const labelText = await page.locator('div[data-qa="field_type_dropdown_name_detail_policy.policy_holder_address.sub_district"] selection-item, div[data-qa="field_type_dropdown_name_detail_policy.policy_holder_address.sub_district"] [class*="selection-item"], div[data-qa="field_type_dropdown_name_detail_policy.policy_holder_address.sub_district"] #dropdown-label-ddl').first().textContent().catch(() => '');
                const isPlaceholder = !labelText || 
                                      labelText.toLowerCase().includes('select') || 
                                      labelText.includes('เลือก') || 
                                      labelText.trim() === '';
                if (!isPlaceholder) {
                    break;
                }
                
                const fallbackDist = `(nth option: ${retrySub + 1})`;
                console.log(`[KUMA AUTO]   ⚠️  Sub District dropdown selection failed or empty. Selecting a DIFFERENT District "${fallbackDist}" to trigger new API fetch (Attempt ${retrySub + 1}/3)...`);
                await fillDropdown('field_type_dropdown_name_detail_policy.policy_holder_address.district', fallbackDist);
                await page.waitForTimeout(3000);
            }

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
            await page.waitForTimeout(200);

            await fillDropdown('field_type_dropdown_name_coverage.coverage_info.product_type', item.productType || '01');
            await page.waitForTimeout(300);
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
                await page.waitForTimeout(300);
            } catch (e) {}

            const ptSelector = 'div[data-qa="field_type_dropdown_name_coverage.coverage_table.plan_type"] [data-qa="btn_dropdown_toggle_ddl"]';
            const countPlanType = await page.locator(ptSelector).count().catch(() => 0);
            console.log(`[KUMA AUTO] Filling ${countPlanType} plan rows in Coverage table...`);
            for (let i = 0; i < countPlanType; i++) {
                await fillDropdown('field_type_dropdown_name_coverage.coverage_table.plan_type', item.planType || '1 :, 2 :, 3 :, 4 :, 5 :, 6 :', i);
                
                // Wait 1.5 seconds for Mode of Payment to load dynamically
                await page.waitForTimeout(1500);

                // Retry mode of payment dropdown up to 3 times
                let mopSuccess = false;
                for (let retry = 0; retry < 3; retry++) {
                    await fillDropdown('field_type_dropdown_name_coverage.coverage_table.mode_of_payment', item.modeOfPayment || 'Monthly, รายเดือน', i);
                    
                    // Verify if the option is actually selected and not empty
                    const selector = `div[data-qa="field_type_dropdown_name_coverage.coverage_table.mode_of_payment"] [data-qa="btn_dropdown_toggle_ddl"]`;
                    const triggerText = (await page.locator(selector).nth(i).textContent().catch(() => '')).trim();
                    const isEmpty = !triggerText || triggerText.toLowerCase().includes('select') || triggerText.includes('เลือก');
                    if (!isEmpty) {
                        mopSuccess = true;
                        break;
                    }
                    console.log(`[KUMA AUTO]   ⚠️  mode_of_payment index ${i} is still empty. Retry ${retry + 1}/3...`);
                    await page.waitForTimeout(1000);
                }
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
                await agentCodeInput.click({ force: true });
                await agentCodeInput.fill(item.agentBrokerCode || '144660');
                await page.keyboard.press('Enter');
                await page.keyboard.press('Tab');
                await page.waitForTimeout(5000); // Wait for API call
            } catch (e) {
                console.log('[KUMA AUTO]   ⚠️  Agent Code input error:', e.message);
            }

            if (item.salesTeam) {
                await fillDropdown('field_type_dropdown_name_agent_broker.mtl_sales_info.sales_team', item.salesTeam);
                await page.waitForTimeout(2000); // Wait for sales names to load based on team selection
            }
            if (item.salesName) {
                await fillDropdown('field_type_dropdown_name_agent_broker.mtl_sales_info.sales_name', item.salesName);
            }

            // Click "+ Add Commission Rate" and fill commission fields
            try {
                const addCommBtn = page.locator('button:has-text("Add Commission Rate"), button:has-text("Commission Rate"), button:has-text("คอมมิชชั่น")').first();
                const hasCustomComm = item.commPlanType1 || item.commRate1 || item.addCommRate1;
                
                if (hasCustomComm) {
                    // Collect all populated commission rows (supporting both old 1-4 format and new single dropdown format)
                    const commRows = [];
                    if (item.commPlanType1 || item.commRate1 || item.addCommRate1) {
                        commRows.push({
                            planType: item.commPlanType1,
                            commRate: item.commRate1,
                            addCommRate: item.addCommRate1
                        });
                    }
                    if (item.commPlanType2 || item.commRate2 || item.addCommRate2) {
                        commRows.push({
                            planType: item.commPlanType2,
                            commRate: item.commRate2,
                            addCommRate: item.addCommRate2
                        });
                    }
                    if (item.commPlanType3 || item.commRate3 || item.addCommRate3) {
                        commRows.push({
                            planType: item.commPlanType3,
                            commRate: item.commRate3,
                            addCommRate: item.addCommRate3
                        });
                    }
                    if (item.commPlanType4 || item.commRate4 || item.addCommRate4) {
                        commRows.push({
                            planType: item.commPlanType4,
                            commRate: item.commRate4,
                            addCommRate: item.addCommRate4
                        });
                    }

                    console.log(`[KUMA AUTO] Need to add custom commission rates for ${commRows.length} plans...`);
                    for (let i = 0; i < commRows.length; i++) {
                        if (await addCommBtn.isVisible().catch(() => false)) {
                            console.log(`[KUMA AUTO] Clicking + Add Commission Rate button (Index ${i})...`);
                            await addCommBtn.click({ force: true });
                            await page.waitForTimeout(1000);
                            
                            const planTypeDdl = `field_type_dropdown_name_agent_broker.commission_rate.${i}.plan_type`;
                            const row = commRows[i];
                            await fillDropdown(planTypeDdl, row.planType || null, 0);

                            const commRateVal = row.commRate !== undefined && row.commRate !== '' ? row.commRate : '10';
                            await fillText(`input[name="agent_broker.commission_rate.${i}.commission_rate"]`, String(commRateVal));

                            const addCommVal = row.addCommRate !== undefined && row.addCommRate !== '' ? row.addCommRate : '0';
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
            // Wait up to 15s for the Account Detail Step 2 tab or content to appear
            await page.locator('a:has-text("Account Summary"), button:has-text("Account Summary"), button:has-text("Create Account")').first().waitFor({ state: 'visible', timeout: 15000 }).catch(() => {});
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
                await page.waitForTimeout(300);
            } catch (e) {}

            // Scroll down to find + Create Account button
            await page.evaluate(() => {
                const el = document.querySelector('main, .content, [class*="content"], [class*="main-content"]') || document.documentElement;
                el.scrollTop = el.scrollHeight;
                window.scrollTo(0, document.body.scrollHeight);
            });
            await page.waitForTimeout(200);

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
                if (item.accRiskLevel && item.accRiskLevel !== 'Low' && item.accRiskLevel !== 'ความเสี่ยงต่ำ') {
                    await fillDropdown('field_type_dropdown_name_account_detail.account_information.risk_level', item.accRiskLevel);
                    await page.waitForTimeout(300);
                }

                // 5. Occupational Classification (only if custom value specified)
                if (item.accOccupationClass && item.accOccupationClass !== 'Class 1' && item.accOccupationClass !== 'ประเภทอาชีพ ชั้น 1') {
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
                // Scroll the modal body/main page to the bottom to make dropdown and apply button visible!
                try {
                    console.log('[KUMA AUTO] Scrolling modal and page to the bottom...');
                    await page.evaluate(() => {
                        const scrollables = Array.from(document.querySelectorAll('*')).filter(
                            el => el.scrollHeight > el.clientHeight && 
                                  (window.getComputedStyle(el).overflowY === 'auto' || 
                                   window.getComputedStyle(el).overflowY === 'scroll' ||
                                   el.className.includes('modal') ||
                                   el.className.includes('dialog'))
                        );
                        scrollables.forEach(el => {
                            el.scrollTop = el.scrollHeight;
                        });
                        window.scrollTo(0, document.body.scrollHeight);
                    });
                    await page.waitForTimeout(1000);
                } catch (e) {
                    console.log('[KUMA AUTO] Scroll error:', e.message);
                }

                // 1. Plan Type
                await fillDropdown('field_type_dropdown_name_claim_payment_object.claim_payment.0.plan_type', item.accPlanType || 'Select All,เลือกทั้งหมด');
                await page.waitForTimeout(300);

                // 2. Payment Type
                await fillDropdown('field_type_dropdown_name_claim_payment_object.claim_payment.0.payment_type', item.accPaymentType || 'Bank Transfer');
                await page.waitForTimeout(300);

                // 3. Paid To
                await fillDropdown('field_type_dropdown_name_claim_payment_object.claim_payment.0.paid_type', item.accPaidTo || 'ผู้ถือกรมธรรม์ (Account)');
                await page.waitForTimeout(300);
 
                // 4. Payment Description
                if (item.accPaymentDesc) {
                    await fillText('input[name="claim_payment_object.claim_payment.0.payment_description"], input[name$="payment_description"]', item.accPaymentDesc);
                    await page.waitForTimeout(300);
                }

                // Submit modal
                const modalSubmitBtn = page.locator('#account-detail-modal-content_Account\\ Detail button:has-text("Submit"), button:has-text("Submit"), button:has-text("ตกลง"), button:has-text("บันทึก")').first();
                await modalSubmitBtn.click().catch(() => modalSubmitBtn.click({ force: true }));
                await page.waitForTimeout(2000);

                // Handle Saving Confirmation dialog if it appears
                try {
                    const confirmBtn = page.locator('[role="dialog"] button:has-text("Confirm"), button:has-text("Confirm"), button:has-text("ตกลง"), button:has-text("ยืนยัน")').last();
                    if (await confirmBtn.isVisible().catch(() => false)) {
                        console.log('[KUMA AUTO] Confirmation pop-up detected. Clicking Confirm...');
                        await confirmBtn.click({ force: true });
                        await page.waitForTimeout(4000);
                    }
                } catch (e) {
                    console.log('[KUMA AUTO] No confirmation pop-up handled:', e.message);
                }

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
            // STEP 3 — Document Upload
            // ====================================================
            console.log('[KUMA AUTO] --- STEP 3: Upload Documents ---');
            
            // Click Next (Step 2 → Step 3)
            console.log('[KUMA AUTO] Clicking Next (Step 2 → Step 3)...');
            const nextBtn2 = page.locator('button:has-text("Next"), button:has-text("ถัดไป")').last();
            await nextBtn2.scrollIntoViewIfNeeded().catch(() => {});
            await nextBtn2.click({ force: true });
            // Wait for document upload elements to appear
            await page.locator('button:has-text("Upload File"), span:has-text("Upload File"), div:has-text("Upload File")').first().waitFor({ state: 'visible', timeout: 15000 }).catch(() => {});
            
            // Use the user's provided PNG file for uploading
            const dummyPath = path.join(__dirname, 'dummy.png');
            if (!fs.existsSync(dummyPath)) {
                fs.writeFileSync(dummyPath, 'Dummy image content');
            }

            // Wait for the document table (specifically "Upload File" buttons) to be visible in the DOM
            try {
                console.log('[KUMA AUTO] Waiting for document upload table to render...');
                const uploadBtn = page.locator('button:has-text("Upload File"), span:has-text("Upload File"), div:has-text("Upload File")').first();
                await uploadBtn.waitFor({ state: 'visible', timeout: 15000 });
                console.log('[KUMA AUTO] Document upload table loaded successfully.');
            } catch (err) {
                console.log('[KUMA AUTO]   ⚠️ Upload File button did not become visible within 15s:', err.message);
            }

            // Make sure all required rows have ATTACHED_FILE selected
            console.log('[KUMA AUTO] Selecting ATTACHED_FILE radio buttons for document rows...');
            try {
                const radioLocator = page.locator('input[type="radio"][value="ATTACHED_FILE"]');
                await radioLocator.first().waitFor({ state: 'attached', timeout: 5000 }).catch(() => {});
                
                const radios = await radioLocator.all();
                console.log(`[KUMA AUTO] Found ${radios.length} ATTACHED_FILE radio targets.`);
                for (let r of radios) {
                    try {
                        await r.click({ force: true });
                        await page.waitForTimeout(500); // Wait for UI transition
                    } catch (e) {
                        console.log('[KUMA AUTO]   ⚠️  Failed clicking radio:', e.message);
                    }
                }
                await page.waitForTimeout(1000); // Let all file inputs mount
            } catch (e) {
                console.log('[KUMA AUTO] Failed to click ATTACHED_FILE radios:', e.message);
            }

            // Find all file inputs on the Step 3 page
            let fileInputs = await page.locator('input[type="file"]').all();
            console.log(`[KUMA AUTO] Found ${fileInputs.length} file inputs for upload.`);

            if (fileInputs.length > 0) {
                console.log(`[KUMA AUTO] Uploading dummy.png directly to file inputs...`);
                for (let idx = 0; idx < fileInputs.length; idx++) {
                    try {
                        await fileInputs[idx].setInputFiles(dummyPath);
                        await page.waitForTimeout(800); // Small delay to let UI process the file
                    } catch (e) {
                        console.log(`[KUMA AUTO]   ⚠️ Failed direct upload for index ${idx}:`, e.message);
                    }
                }
            } else {
                console.log('[KUMA AUTO] Direct file inputs not found. Attempting filechooser intercept on Upload File buttons...');
                const uploadButtons = await page.locator('button:has-text("Upload File"), span:has-text("Upload File"), div:has-text("Upload File")').all();
                console.log(`[KUMA AUTO] Found ${uploadButtons.length} Upload File click targets.`);
                for (let idx = 0; idx < uploadButtons.length; idx++) {
                    try {
                        console.log(`[KUMA AUTO] Uploading file via click intercept for index ${idx}...`);
                        const fileChooserPromise = page.waitForEvent('filechooser', { timeout: 4000 });
                        await uploadButtons[idx].scrollIntoViewIfNeeded().catch(() => {});
                        await uploadButtons[idx].click({ force: true });
                        const fileChooser = await fileChooserPromise;
                        await fileChooser.setFiles(dummyPath);
                        await page.waitForTimeout(1000); // Wait for upload completion
                    } catch (e) {
                        console.log(`[KUMA AUTO]   ⚠️ Failed upload via click intercept for index ${idx}:`, e.message);
                    }
                }
            }

            await page.waitForTimeout(2000);
            await takeScreenshot(`${caseLabel}_08_step3_uploaded`);

            console.log('[KUMA AUTO] Clicking Next button on Step 3...');
            const nextBtn3 = page.locator('button:has-text("Next"), button:has-text("ถัดไป")').last();
            await nextBtn3.scrollIntoViewIfNeeded().catch(() => {});
            await nextBtn3.click({ force: true });
            // Wait for summary elements to appear
            await page.locator('button:text-is("Submit"), button:text-is("Confirm"), button:text-is("สร้าง"), button:text-is("ส่งใบสมัคร"), button:has-text("Submit Case")').first().waitFor({ state: 'visible', timeout: 15000 }).catch(() => {});
            await takeScreenshot(`${caseLabel}_08_step4_summary`);

            // ====================================================
            // STEP 4 — Case Summary → Submit
            // ====================================================
            console.log('[KUMA AUTO] --- STEP 4: Case Summary → Submit ---');

            const handleConfirmModal = async () => {
                try {
                    await page.waitForTimeout(500);
                    console.log('[KUMA AUTO] Searching for Submit Case confirmation modal...');
                    const modalSubmit = page.locator('button:has-text("Submit"), button:has-text("ตกลง"), button:has-text("ยืนยัน")').filter({ state: 'visible' }).last();
                    
                    if (await modalSubmit.isVisible().catch(() => false)) {
                        console.log('[KUMA AUTO] Clicking Submit button inside the visible confirmation modal...');
                        await modalSubmit.click({ force: true });
                        // Wait for submit request to finish
                        await page.waitForLoadState('networkidle').catch(() => {});
                        await page.waitForTimeout(1000);
                    } else {
                        console.log('[KUMA AUTO] Confirm Submit Case modal button not found via visible locator.');
                    }
                } catch (e) {
                    console.log('[KUMA AUTO] Failed to handle confirm submit inside modal:', e.message);
                }
            };

            // Select the actual Submit/Confirm/สร้าง button (avoiding Draft/บันทึกร่าง/Save Draft)
            const submitBtn = page.locator(
                'button:text-is("Submit"), button:text-is("Confirm"), button:text-is("สร้าง"), button:text-is("ส่งใบสมัคร"), button:has-text("Submit Case")'
            ).first();

            if (await submitBtn.isVisible().catch(() => false)) {
                console.log('[KUMA AUTO] Clicking final Submit button...');
                await submitBtn.click({ force: true });
                await handleConfirmModal();
                await takeScreenshot(`${caseLabel}_09_DONE`);
                console.log(`[KUMA AUTO] ✅ Case ${caseIdx + 1} submitted successfully!`);
                await approveCase(caseLabel, caseIdx);
            } else {
                console.log('[KUMA AUTO]   ℹ️  Submit button not visible — attempting backup selector...');
                // Fallback to general Submit but make sure it isn't "Save Draft" or "บันทึกร่าง"
                const fallbackSubmit = page.locator('button:has-text("Submit"), button:has-text("สร้าง")').filter({ hasNotText: 'Draft' }).filter({ hasNotText: 'ร่าง' }).first();
                if (await fallbackSubmit.isVisible().catch(() => false)) {
                    await fallbackSubmit.click({ force: true });
                    await handleConfirmModal();
                    await takeScreenshot(`${caseLabel}_09_DONE`);
                    console.log(`[KUMA AUTO] ✅ Case ${caseIdx + 1} submitted successfully (via fallback selector)!`);
                    await approveCase(caseLabel, caseIdx);
                } else {
                    console.log('[KUMA AUTO]   ℹ️  Submit button not found — may require manual review.');
                    await takeScreenshot(`${caseLabel}_09_submit_not_found`);
                }
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
            await page.waitForTimeout(1000);
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
