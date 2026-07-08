const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer-core');
const ExcelJS = require('exceljs');
const prompts = require('prompts');
const pc = require('picocolors');

// Default Config
const DEFAULT_CONFIG = {
  chromePath: "C:\\Program Files\\Google\Chrome\\Application\\chrome.exe",
  excelFile: "data.xlsx",
  startRow: 2,
  mode: "attach", // "attach" or "launch"
  delayBetweenFields: 500,
  autoSave: false
};

// Fallback Web Form Selectors (Matches mock form + eGramSwaraj potential forms)
const SELECTORS = {
  theme: ['select#theme', 'select[name="themeId"]', 'select[name="theme"]', 'select[id*="theme"]'],
  activity: ['select#activity', 'select[name="activityId"]', 'select[name="activity"]', 'select[id*="activity"]'],
  focus_area: ['select#focus_area', 'select[name="focusAreaId"]', 'select[name="focusArea"]', 'select[id*="focusArea"]', 'select[id*="focus_area"]'],
  activity_type: ['select#activity_type', 'select[name="activityType"]', 'select[id*="activityType"]', 'select[id*="activity_type"]'],
  activity_nature: ['select#activity_nature', 'select[name="activityNature"]', 'select[id*="activityNature"]', 'select[id*="activity_nature"]'],
  description: ['textarea#description', 'textarea[name="activityDescription"]', 'textarea[name="description"]', 'textarea[id*="description"]'],
  indicators: ['select#indicators', 'select[name="indicators"]', 'select[name="indicatorId"]', 'select[id*="indicator"]'],
  remarks_for: ['select#remarks_for', 'select[name="remarksFor"]', 'select[name="remarks_for"]', 'select[id*="remarks"]'],
  targeted_population: ['select#targeted_population', 'select[name="targetedPopulation"]', 'select[name="targeted_population"]', 'select[id*="target"]'],
  funded_by_panchayat: ['input[type="radio"][name="funded_by_panchayat"]', 'input[type="radio"][name="directlyFunded"]', 'input[type="radio"][name="isFunded"]'],
  completion_year: ['input#completion_year', 'input[name="completionYear"]', 'input[name="estimatedCompletionYear"]', 'input[id*="completion_year"]', 'input[placeholder="Year"]'],
  completion_month: ['input#completion_month', 'input[name="completionMonth"]', 'input[name="estimatedCompletionMonth"]', 'input[id*="completion_month"]', 'input[placeholder="Month"]'],
  completion_days: ['input#completion_days', 'input[name="completionDays"]', 'input[name="estimatedCompletionDays"]', 'input[id*="completion_days"]', 'input[placeholder="Days"]'],
  start_year: ['select#start_year', 'select[name="startYear"]', 'select[id*="startYear"]', 'select[id*="start_year"]'],
  start_month: ['select#start_month', 'select[name="startMonth"]', 'select[id*="startMonth"]', 'select[id*="start_month"]'],
  beneficiaries_general: ['input#beneficiaries_general', 'input[name="beneficiaryGeneral"]', 'input[id*="general"]', 'input[id*="General"]'],
  beneficiaries_sc: ['input#beneficiaries_sc', 'input[name="beneficiarySc"]', 'input[id*="sc"]', 'input[id*="SC"]'],
  beneficiaries_st: ['input#beneficiaries_st', 'input[name="beneficiarySt"]', 'input[id*="st"]', 'input[id*="ST"]'],
  estimated_cost: ['input#estimated_cost', 'input[name="estimatedCost"]', 'input[name="totalCost"]', 'input[id*="cost"]', 'input[id*="Cost"]']
};

/**
 * Loads config.json if it exists, otherwise creates one with default values
 */
function loadConfig() {
  const configPath = path.join(process.cwd(), 'config.json');
  if (fs.existsSync(configPath)) {
    try {
      const userConfig = JSON.parse(fs.readFileSync(configPath, 'utf8'));
      return { ...DEFAULT_CONFIG, ...userConfig };
    } catch (e) {
      console.log(pc.red(`❌ Error reading config.json: ${e.message}. Using default settings.`));
    }
  } else {
    try {
      fs.writeFileSync(configPath, JSON.stringify(DEFAULT_CONFIG, null, 2), 'utf8');
      console.log(pc.green(`💡 Created default config.json`));
    } catch (e) {
      console.log(pc.yellow(`⚠️ Could not create config.json: ${e.message}`));
    }
  }
  return DEFAULT_CONFIG;
}

/**
 * Generates a sample Excel template with columns and demo rows
 */
async function generateTemplate(filePath) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('GPDP Activities');

  // Set columns
  sheet.columns = [
    { header: 'theme', key: 'theme', width: 45 },
    { header: 'activity', key: 'activity', width: 40 },
    { header: 'focus_area', key: 'focus_area', width: 35 },
    { header: 'activity_type', key: 'activity_type', width: 25 },
    { header: 'activity_nature', key: 'activity_nature', width: 25 },
    { header: 'description', key: 'description', width: 50 },
    { header: 'indicators', key: 'indicators', width: 45 },
    { header: 'remarks_for', key: 'remarks_for', width: 25 },
    { header: 'targeted_population', key: 'targeted_population', width: 25 },
    { header: 'funded_by_panchayat', key: 'funded_by_panchayat', width: 25 },
    { header: 'completion_year', key: 'completion_year', width: 18 },
    { header: 'completion_month', key: 'completion_month', width: 18 },
    { header: 'completion_days', key: 'completion_days', width: 18 },
    { header: 'start_year', key: 'start_year', width: 15 },
    { header: 'start_month', key: 'start_month', width: 15 },
    { header: 'beneficiaries_general', key: 'beneficiaries_general', width: 22 },
    { header: 'beneficiaries_sc', key: 'beneficiaries_sc', width: 18 },
    { header: 'beneficiaries_st', key: 'beneficiaries_st', width: 18 },
    { header: 'estimated_cost', key: 'estimated_cost', width: 22 },
    { header: 'upload_status', key: 'upload_status', width: 30 }
  ];

  // Add demo row matching screenshot
  sheet.addRow({
    theme: 'Theme 1 - Poverty Free and Enhanced Livelihoods Village',
    activity: 'Facilitate formation of bank sakhi',
    focus_area: 'Poverty alleviation programme',
    activity_type: 'Community Works',
    activity_nature: 'New Asset',
    description: 'Facilitate formation of bank sakhi and self-help groups in the village.',
    indicators: 'Percentage of households with bank accounts',
    remarks_for: 'Women',
    targeted_population: 'SC and ST',
    funded_by_panchayat: 'Yes',
    completion_year: 0,
    completion_month: 6,
    completion_days: 0,
    start_year: '2026-2027',
    start_month: 'July',
    beneficiaries_general: 10,
    beneficiaries_sc: 5,
    beneficiaries_st: 5,
    estimated_cost: 50000
  });

  // Second demo row
  sheet.addRow({
    theme: 'Theme 2 - Healthy Village',
    activity: 'Provide drinking water facilities',
    focus_area: 'Drinking Water',
    activity_type: 'Community Works',
    activity_nature: 'New Asset',
    description: 'Installation of handpumps and piping for clean drinking water.',
    indicators: 'Percentage of households with tap water',
    remarks_for: 'General Public',
    targeted_population: 'All Villagers',
    funded_by_panchayat: 'Yes',
    completion_year: 1,
    completion_month: 2,
    completion_days: 15,
    start_year: '2026-2027',
    start_month: 'August',
    beneficiaries_general: 150,
    beneficiaries_sc: 40,
    beneficiaries_st: 25,
    estimated_cost: 250000
  });

  // Format header row
  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
  sheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '2C3E50' }
  };

  await workbook.xlsx.writeFile(filePath);
}

/**
 * Tries to find an element using fallback selector list
 */
async function findSelector(page, selectorList) {
  for (const selector of selectorList) {
    try {
      const el = await page.$(selector);
      if (el) return selector;
    } catch (e) {
      // Selector might be syntax-invalid or not found, proceed to next
    }
  }
  return null;
}

/**
 * Fills out standard text inputs or textareas
 */
async function fillInput(page, selectorList, value, delayMs, labelName) {
  if (value === undefined || value === null || value === '') return;
  const selector = await findSelector(page, selectorList);
  if (!selector) {
    console.log(pc.yellow(`⚠️ Field not found for: ${labelName} (Checked: ${selectorList.join(', ')})`));
    return;
  }

  // Clear and prepare field
  await page.evaluate((sel) => {
    const el = document.querySelector(sel);
    if (el) {
      el.value = '';
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
    }
  }, selector);

  // Type the value using the top-level page keyboard (frames don't have their own keyboard interface)
  const topPage = typeof page.page === 'function' ? page.page() : page;
  await page.focus(selector);
  await topPage.keyboard.type(value.toString());
  
  // Fire blur and change events to let validation / dynamic updates execute
  await page.evaluate((sel) => {
    const el = document.querySelector(sel);
    if (el) {
      el.dispatchEvent(new Event('change', { bubbles: true }));
      el.dispatchEvent(new Event('blur', { bubbles: true }));
    }
  }, selector);

  await new Promise(r => setTimeout(r, delayMs));
}

/**
 * Selects an option from a drop-down based on text matching
 */
async function selectDropdown(page, selectorList, text, delayMs, labelName) {
  if (!text) return;
  const selector = await findSelector(page, selectorList);
  if (!selector) {
    console.log(pc.yellow(`⚠️ Dropdown not found for: ${labelName} (Checked: ${selectorList.join(', ')})`));
    return;
  }

  const selectedValue = await page.evaluate((sel, txt) => {
    const select = document.querySelector(sel);
    if (!select) return null;
    
    const options = Array.from(select.options);
    const cleanTxt = txt.toString().trim().toLowerCase();
    
    // Find option where text contains the Excel value, or matches exactly
    const match = options.find(o => 
      o.text.trim().toLowerCase().includes(cleanTxt) || 
      o.value.trim().toLowerCase() === cleanTxt
    );

    if (match) {
      select.value = match.value;
      select.dispatchEvent(new Event('change', { bubbles: true }));
      return match.text;
    }
    return null;
  }, selector, text);

  if (!selectedValue) {
    console.log(pc.red(`❌ Option "${text}" not found in dropdown for: ${labelName}`));
  } else {
    // Wait for cascading dropdown animations/API loads
    await new Promise(r => setTimeout(r, delayMs));
  }
}

/**
 * Selects a radio button matching the text/value
 */
async function selectRadio(page, selectorList, value, delayMs, labelName) {
  if (value === undefined || value === null || value === '') return;
  const selector = await findSelector(page, selectorList);
  if (!selector) {
    console.log(pc.yellow(`⚠️ Radio buttons not found for: ${labelName}`));
    return;
  }

  const nameAttr = await page.evaluate((sel) => {
    const el = document.querySelector(sel);
    return el ? el.getAttribute('name') : null;
  }, selector);

  if (!nameAttr) return;

  const clicked = await page.evaluate((name, val) => {
    const radios = Array.from(document.querySelectorAll(`input[type="radio"][name="${name}"]`));
    const targetVal = val.toString().trim().toLowerCase();
    
    const target = radios.find(r => 
      r.value.toLowerCase() === targetVal || 
      r.nextSibling?.textContent?.trim().toLowerCase().includes(targetVal) ||
      r.parentElement?.textContent?.trim().toLowerCase().includes(targetVal)
    );

    if (target) {
      target.click();
      target.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    }
    return false;
  }, nameAttr, value);

  if (!clicked) {
    console.log(pc.yellow(`⚠️ Radio option "${value}" not found for: ${labelName}`));
  } else {
    await new Promise(r => setTimeout(r, delayMs));
  }
}

/**
 * Helper to save progress to Excel file safely
 */
async function saveWorkbook(workbook, filePath) {
  try {
    await workbook.xlsx.writeFile(filePath);
  } catch (err) {
    console.log(pc.yellow(`\n⚠️ Warning: Could not write status to Excel (is "${path.basename(filePath)}" open in Excel?).`));
    console.log(pc.yellow(`Please close Microsoft Excel and press Enter to save your progress.`));
    await prompts({
      type: 'text',
      name: 'wait',
      message: 'Press [Enter] when ready...'
    });
    try {
      await workbook.xlsx.writeFile(filePath);
      console.log(pc.green(`✔ Progress saved successfully.`));
    } catch (retryErr) {
      console.log(pc.red(`❌ Failed to save progress: ${retryErr.message}`));
    }
  }
}

/**
 * Core application runner
 */
async function run() {
  let browser;
  try {
    console.log(pc.cyan('=================================================='));
  console.log(pc.bold(pc.cyan('         GPDP Form Autofill Automator             ')));
  console.log(pc.cyan('==================================================\n'));

  const config = loadConfig();
  const excelFilePath = path.join(process.cwd(), config.excelFile);

  // Check Excel File Presence
  if (!fs.existsSync(excelFilePath)) {
    console.log(pc.yellow(`Excel file "${config.excelFile}" not found in current directory.`));
    const response = await prompts({
      type: 'confirm',
      name: 'generate',
      message: 'Would you like to generate a sample Excel template?',
      initial: true
    });

    if (response.generate) {
      await generateTemplate(excelFilePath);
      console.log(pc.green(`\n✔ Generated sample Excel: "${config.excelFile}"`));
      console.log(pc.cyan(`Please populate your data in this sheet and rerun the tool.\n`));
    } else {
      console.log(pc.red('Please create an Excel sheet with activity data to proceed. Exiting.'));
    }
    return;
  }

  // Read Excel File
  const workbook = new ExcelJS.Workbook();
  console.log(pc.blue(`Reading "${config.excelFile}"...`));
  await workbook.xlsx.readFile(excelFilePath);
  const sheet = workbook.worksheets[0];

  // Map rows
  const activities = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber < config.startRow) return;

    // Build data object matching cells
    const getVal = (colNum) => {
      const cell = row.getCell(colNum);
      if (cell.value && typeof cell.value === 'object' && cell.value.result !== undefined) {
        return cell.value.result;
      }
      return cell.value;
    };

    activities.push({
      rowNum: rowNumber,
      rowObj: row,
      theme: getVal(1),
      activity: getVal(2),
      focus_area: getVal(3),
      activity_type: getVal(4),
      activity_nature: getVal(5),
      description: getVal(6),
      indicators: getVal(7),
      remarks_for: getVal(8),
      targeted_population: getVal(9),
      funded_by_panchayat: getVal(10),
      completion_year: getVal(11),
      completion_month: getVal(12),
      completion_days: getVal(13),
      start_year: getVal(14),
      start_month: getVal(15),
      beneficiaries_general: getVal(16),
      beneficiaries_sc: getVal(17),
      beneficiaries_st: getVal(18),
      estimated_cost: getVal(19),
      upload_status: getVal(20)
    });
  });

  if (activities.length === 0) {
    console.log(pc.yellow(`No rows to process starting from row ${config.startRow}. Exiting.`));
    return;
  }

  console.log(pc.green(`✔ Loaded ${activities.length} activity rows from Excel.`));

  // Choose launch/attach mode
  let page;

  if (config.mode === 'attach') {
    console.log('\n--------------------------------------------------');
    console.log(pc.bold(pc.yellow('Connection Mode: ATTACH TO RUNNING CHROME')));
    console.log('Instructions:');
    console.log('1. Close all active Google Chrome windows completely.');
    console.log('2. Open Chrome using the command prompt or terminal with debugging:');
    console.log(pc.cyan(`   chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\\chrome-profile"`));
    console.log('3. In that debugging Chrome window, log in and open the GPDP Activity Form.');
    console.log('--------------------------------------------------\n');

    const connectPrompt = await prompts({
      type: 'text',
      name: 'url',
      message: 'Enter testing URL or press [Enter] to scan open Chrome tabs:',
      initial: ''
    });

      browser = await puppeteer.connect({
        browserURL: 'http://localhost:9222',
        defaultViewport: null
      });

      const pages = await browser.pages();
      
      // Filter pages safely (skip system target pages and background services that throw errors)
      let activePages = [];
      for (const p of pages) {
        try {
          const url = p.url();
          if (url && !url.startsWith('chrome://') && !url.startsWith('chrome-extension://')) {
            activePages.push(p);
          }
        } catch (e) {
          // ignore closed or system targets
        }
      }

      if (activePages.length === 0) {
        page = pages[0] || await browser.newPage();
      } else {
        if (connectPrompt.url) {
          const cleanUrl = connectPrompt.url.trim().toLowerCase();
          page = activePages.find(p => {
            try {
              return p.url().toLowerCase().includes(cleanUrl);
            } catch (e) {
              return false;
            }
          });
        } else {
          page = activePages.find(p => {
            try {
              const url = p.url().toLowerCase();
              return url.includes('egramswaraj') || url.includes('mock') || url.includes('.html');
            } catch (e) {
              return false;
            }
          });
        }
        if (!page) {
          page = activePages[0];
        }
      }

      let pageTitle = 'Unknown Tab';
      let pageUrl = 'Unknown URL';
      try {
        pageTitle = await page.title();
        pageUrl = page.url();
      } catch (err) {
        // ignore title/url errors
      }

      console.log(pc.green(`✔ Connected successfully to page: ${pageTitle} (${pageUrl})`));
  } else {
    console.log('\n--------------------------------------------------');
    console.log(pc.bold(pc.yellow('Connection Mode: LAUNCH NEW CHROME')));
    console.log('--------------------------------------------------\n');

    try {
      browser = await puppeteer.launch({
        executablePath: config.chromePath,
        headless: false,
        defaultViewport: null,
        args: ['--start-maximized']
      });
      const pages = await browser.pages();
      page = pages[0] || await browser.newPage();

      // Ask for target URL or use mock
      const targetResponse = await prompts({
        type: 'text',
        name: 'url',
        message: 'Enter GPDP Form URL (leave blank to test local mock form):',
        initial: ''
      });

      let targetUrl = targetResponse.url;
      if (!targetUrl) {
        const localMockPath = path.join(process.cwd(), 'mock_form', 'index.html');
        if (fs.existsSync(localMockPath)) {
          targetUrl = `file://${localMockPath}`;
        } else {
          console.log(pc.red('❌ Local mock form not found and no URL entered. Exiting.'));
          await browser.close();
          return;
        }
      }

      console.log(pc.blue(`Navigating to: ${targetUrl}...`));
      await page.goto(targetUrl, { waitUntil: 'networkidle2' });
      console.log(pc.green(`✔ Form loaded.`));

    } catch (e) {
      console.log(pc.red(`❌ Launch failed: ${e.message}`));
      console.log(pc.yellow(`Make sure the chromePath in config.json is correct: "${config.chromePath}"`));
      return;
    }
  }

  // Loop through rows
  for (let i = 0; i < activities.length; i++) {
    const act = activities[i];

    // Check if already uploaded
    if (act.upload_status && act.upload_status.toString().toLowerCase().includes('success')) {
      console.log(pc.green(`Row ${act.rowNum}: Already successfully uploaded (SUCCESS). Skipping.`));
      continue;
    }

    console.log('\n' + pc.cyan('--------------------------------------------------'));
    console.log(pc.bold(pc.green(`Filling Activity [Row ${act.rowNum}]`)));
    console.log(pc.cyan('--------------------------------------------------'));
    console.log(`${pc.bold('Theme:')} ${act.theme || 'N/A'}`);
    console.log(`${pc.bold('Activity:')} ${act.activity || 'N/A'}`);
    console.log(`${pc.bold('Cost:')} Rs. ${act.estimated_cost || 0}`);
    console.log(pc.cyan('--------------------------------------------------\n'));

    // Dynamic context helper (resolves forms inside iframes/framesets)
    let context = page;
    try {
      const frames = page.frames();
      for (const frame of frames) {
        try {
          const hasTheme = await frame.$('select[name="themeId"], select#theme, select[name="theme"], select[id*="theme"]');
          if (hasTheme) {
            context = frame;
            break;
          }
        } catch (e) {
          // ignore cross-origin access blocks
        }
      }
    } catch (e) {
      // ignore frame check errors
    }

    if (context !== page) {
      console.log(pc.green(`✔ Form detected inside iframe context: ${context.url()}`));
    }

    try {
      const delay = config.delayBetweenFields;

      // 1. Choose Theme
      await selectDropdown(context, SELECTORS.theme, act.theme, delay, 'Theme');

      // 2. Select Activity
      await selectDropdown(context, SELECTORS.activity, act.activity, delay, 'Activity');

      // 3. Focus Area
      await selectDropdown(context, SELECTORS.focus_area, act.focus_area, delay, 'Focus Area');

      // 4. Activity Type
      await selectDropdown(context, SELECTORS.activity_type, act.activity_type, delay, 'Activity Type');

      // 5. Activity Nature
      await selectDropdown(context, SELECTORS.activity_nature, act.activity_nature, delay, 'Activity Nature');

      // 6. Enter Description
      await fillInput(context, SELECTORS.description, act.description, delay, 'Description');

      // 7. Indicators
      await selectDropdown(context, SELECTORS.indicators, act.indicators, delay, 'Indicators');

      // 8. Remarks For
      await selectDropdown(context, SELECTORS.remarks_for, act.remarks_for, delay, 'Remarks For');

      // 9. Targeted Population
      await selectDropdown(context, SELECTORS.targeted_population, act.targeted_population, delay, 'Targeted Population');

      // 10. Funded directly by Panchayat
      await selectRadio(context, SELECTORS.funded_by_panchayat, act.funded_by_panchayat, delay, 'Funded by Panchayat');

      // 11. Completion Time
      await fillInput(context, SELECTORS.completion_year, act.completion_year, delay, 'Completion Year');
      await fillInput(context, SELECTORS.completion_month, act.completion_month, delay, 'Completion Month');
      await fillInput(context, SELECTORS.completion_days, act.completion_days, delay, 'Completion Days');

      // 12. Start Timeline
      await selectDropdown(context, SELECTORS.start_year, act.start_year, delay, 'Start Year');
      await selectDropdown(context, SELECTORS.start_month, act.start_month, delay, 'Start Month');

      // 13. Beneficiaries
      await fillInput(context, SELECTORS.beneficiaries_general, act.beneficiaries_general, delay, 'General Beneficiaries');
      await fillInput(context, SELECTORS.beneficiaries_sc, act.beneficiaries_sc, delay, 'SC Beneficiaries');
      await fillInput(context, SELECTORS.beneficiaries_st, act.beneficiaries_st, delay, 'ST Beneficiaries');

      // 14. Total Cost
      await fillInput(context, SELECTORS.estimated_cost, act.estimated_cost, delay, 'Estimated Cost');

      console.log(pc.green(`\n✔ Row ${act.rowNum} successfully filled in browser.`));

      if (config.autoSave) {
        console.log(pc.blue('Auto-saving form...'));
        const saveBtn = await findSelector(context, ['button#btn_save', 'input[type="submit"]', '#btn_save', '#saveBtn']);
        if (saveBtn) {
          await context.click(saveBtn);
          await new Promise(r => setTimeout(r, 2000)); // wait for alert/submission
          console.log(pc.green('✔ Save clicked.'));
          
          // Write success status
          act.rowObj.getCell(20).value = 'SUCCESS - ' + new Date().toLocaleString();
          await saveWorkbook(workbook, excelFilePath);
        } else {
          console.log(pc.yellow('⚠️ Save button not found. Skipping auto-save.'));
        }
      } else {
        // Wait for user to review and press Enter in the CLI
        const confirm = await prompts({
          type: 'select',
          name: 'action',
          message: 'Review form on screen. Choose action to continue:',
          choices: [
            { title: 'Proceed to next row', value: 'next' },
            { title: 'Retry current row', value: 'retry' },
            { title: 'Exit program', value: 'exit' }
          ]
        });

        if (confirm.action === 'next') {
          // Write success status
          act.rowObj.getCell(20).value = 'SUCCESS - ' + new Date().toLocaleString();
          await saveWorkbook(workbook, excelFilePath);
        } else if (confirm.action === 'retry') {
          i--; // decrement loop counter to repeat current index
          console.log(pc.yellow('Refilling current row...'));
        } else if (confirm.action === 'exit') {
          console.log(pc.yellow('Exiting data fill loop.'));
          break;
        }
      }

    } catch (err) {
      console.log(pc.red(`❌ Error filling Row ${act.rowNum}: ${err.message}`));
      
      const errorAction = await prompts({
        type: 'select',
        name: 'action',
        message: 'An error occurred. Choose action:',
        choices: [
          { title: 'Retry current row', value: 'retry' },
          { title: 'Skip row', value: 'skip' },
          { title: 'Exit', value: 'exit' }
        ]
      });

      if (errorAction.action === 'retry') {
        i--;
      } else if (errorAction.action === 'skip') {
        // Write skipped/error status
        act.rowObj.getCell(20).value = 'SKIPPED - ' + err.message;
        await saveWorkbook(workbook, excelFilePath);
      } else if (errorAction.action === 'exit') {
        break;
      }
    }
  }

  console.log(pc.cyan('==================================================\n'));

  // Close browser prompts
  const closeResponse = await prompts({
    type: 'confirm',
    name: 'close',
    message: 'Do you want to close the automated browser window?',
    initial: false
  });

  if (browser) {
    if (closeResponse.close) {
      await browser.close();
      browser = null;
    } else {
      await browser.disconnect();
      browser = null;
    }
  }

  } catch (err) {
    console.error(pc.red(`\n❌ Error: ${err.message}`));
  } finally {
    if (browser) {
      try {
        await browser.disconnect();
      } catch (e) {}
    }
    // Keep terminal open for double-clicked EXE runs
    await prompts({
      type: 'text',
      name: 'exit',
      message: pc.bold(pc.cyan('\nPress [Enter] to exit the program...'))
    });
  }
}

// Start execution
run().catch(console.error);
