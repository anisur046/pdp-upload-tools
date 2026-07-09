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
  activity: ['select#activity', 'select[name="activityId"]', 'select[name="activityCode"]', 'select[name="activity"]', 'select[id*="activity"]'],
  focus_area: ['select#focus_area', 'select[name="focusAreaId"]', 'select[name="focusAreaCode"]', 'select[name="focusArea"]', 'select[id*="focusArea"]', 'select[id*="focus_area"]'],
  activity_type: ['select#activity_type', 'select[name="activityType"]', 'select[name="activityTypeCode"]', 'select[name="activityTypeId"]', 'select[id*="activityType"]', 'select[id*="activity_type"]'],
  activity_nature: ['select#activity_nature', 'select[name="activityNature"]', 'select[name="natureOfActivity"]', 'select[name="activityNatureId"]', 'select[id*="natureOfActivity"]', 'select[id*="activityNature"]', 'select[id*="activity_nature"]'],
  description: ['textarea#description', 'textarea[name="activityDescription"]', 'textarea[name="description"]', 'textarea[id*="description"]'],
  indicators: ['select#indicators', 'select[name="indicators"]', 'select[name="indicatorId"]', 'select[id*="indicator"]'],
  remarks_for: ['select#remarks_for', 'select[name="remarksFor"]', 'select[name="remarks_for"]', 'select[id*="remarks"]'],
  targeted_population: ['select#targeted_population', 'select[name="targetedPopulation"]', 'select[name="targetedPopulace"]', 'select[name="targeted_population"]', 'select[id*="target"]', 'select[id*="populace"]'],
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
 * Finds the Chrome executable path by checking common Windows install directories
 */
function findChromePath() {
  const standardPaths = [
    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
  ];
  if (process.env.LOCALAPPDATA) {
    standardPaths.push(path.join(process.env.LOCALAPPDATA, 'Google', 'Chrome', 'Application', 'chrome.exe'));
  }
  for (const p of standardPaths) {
    if (fs.existsSync(p)) {
      return p;
    }
  }
  return DEFAULT_CONFIG.chromePath;
}

/**
 * Loads config.json if it exists, otherwise creates one with default values.
 * Re-creates config.json if it is empty or corrupted.
 */
function loadConfig() {
  const configPath = path.join(process.cwd(), 'config.json');
  if (fs.existsSync(configPath)) {
    try {
      const content = fs.readFileSync(configPath, 'utf8').trim();
      if (!content) {
        throw new Error('Config file is empty');
      }
      const userConfig = JSON.parse(content);
      return { ...DEFAULT_CONFIG, ...userConfig };
    } catch (e) {
      console.log(pc.red(`❌ Error reading config.json: ${e.message}. Re-creating default settings.`));
      try {
        fs.writeFileSync(configPath, JSON.stringify(DEFAULT_CONFIG, null, 2), 'utf8');
      } catch (writeErr) {
        // ignore
      }
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
 * Searches for Excel file in multiple directories (CWD, Desktop, OneDrive Desktop)
 */
function findExcelFile(fileName) {
  const os = require('os');
  const pathsToCheck = [
    path.join(process.cwd(), fileName),
    path.join(os.homedir(), 'Desktop', fileName),
    path.join(os.homedir(), 'OneDrive', 'Desktop', fileName)
  ];
  for (const p of pathsToCheck) {
    if (fs.existsSync(p)) {
      return p;
    }
  }
  return null;
}

/**
 * Checks if a TCP port is active (listening)
 */
function checkPortActive(port) {
  return new Promise((resolve) => {
    const net = require('net');
    const socket = new net.Socket();
    socket.setTimeout(300);
    socket.once('connect', () => {
      socket.destroy();
      resolve(true);
    });
    socket.once('error', () => {
      resolve(false);
    });
    socket.once('timeout', () => {
      resolve(false);
    });
    socket.connect(port, '127.0.0.1');
  });
}

/**
 * Spawns a browser process in remote debugging mode and waits for it to listen
 */
function spawnBrowser(executablePath, profileDir) {
  return new Promise((resolve) => {
    const { spawn } = require('child_process');
    const child = spawn(executablePath, [
      '--remote-debugging-port=9222',
      `--user-data-dir=${profileDir}`,
      '--no-first-run'
    ], {
      detached: true,
      stdio: 'ignore'
    });
    child.unref();

    let attempts = 0;
    const interval = setInterval(async () => {
      attempts++;
      const inUse = await checkPortActive(9222);
      if (inUse) {
        clearInterval(interval);
        resolve(true);
      } else if (attempts >= 10) {
        clearInterval(interval);
        resolve(false);
      }
    }, 1000);
  });
}

/**
 * Fetches the browser WebSocket debugger URL from local debugger JSON endpoint using
 * low-level Node.js http API to bypass global proxy/experimental fetch issues.
 */
function getWSEndpoint(browserURL) {
  return new Promise((resolve, reject) => {
    const http = require('http');
    const url = new URL(browserURL + '/json/version');
    
    const options = {
      hostname: url.hostname,
      port: url.port,
      path: url.pathname,
      method: 'GET',
      headers: {
        'host': `127.0.0.1:${url.port}`
      },
      agent: false
    };

    const req = http.request(options, (res) => {
      let data = '';
      res.on('data', (chunk) => {
        data += chunk;
      });
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (json.webSocketDebuggerUrl) {
            resolve(json.webSocketDebuggerUrl);
          } else {
            reject(new Error('webSocketDebuggerUrl not found in response'));
          }
        } catch (e) {
          reject(new Error(`Failed to parse json: ${e.message}`));
        }
      });
    });

    req.on('error', (err) => {
      reject(err);
    });

    req.setTimeout(2000, () => {
      req.destroy();
      reject(new Error('Request timed out'));
    });

    req.end();
  });
}

/**
 * Launches Chrome or Microsoft Edge in debugging mode
 */
async function launchBrowser(config) {
  let browserURL = 'http://127.0.0.1:9222';
  let browserInstance = null;
  const os = require('os');

  const connectWithWS = async () => {
    const wsEndpoint = await getWSEndpoint(browserURL);
    return await puppeteer.connect({
      browserWSEndpoint: wsEndpoint,
      defaultViewport: null
    });
  };

  // 1. Try to connect to an already active debugging browser on port 9222
  try {
    browserInstance = await connectWithWS();
    console.log(pc.green('✔ Connected to active browser debugging session on port 9222.'));
    return browserInstance;
  } catch (e) {
    // Debugging port is closed
  }

  // 2. Try to launch Chrome in debugging mode
  const chromePath = findChromePath();
  if (fs.existsSync(chromePath)) {
    console.log(pc.blue('Attempting to launch Google Chrome in debugging mode...'));
    try {
      const userProfileDir = path.join(os.homedir(), '.pdp-chrome-profile');
      const success = await spawnBrowser(chromePath, userProfileDir);
      if (success) {
        browserInstance = await connectWithWS();
        console.log(pc.green('✔ Connected to Google Chrome.'));
        return browserInstance;
      }
    } catch (err) {
      console.log(pc.yellow(`⚠️ Chrome launch failed: ${err.message}`));
    }
  }

  // 3. Try to launch Microsoft Edge in debugging mode (Fallback)
  const edgePath = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
  if (fs.existsSync(edgePath)) {
    console.log(pc.blue('Attempting to launch Microsoft Edge in debugging mode (fallback)...'));
    try {
      const userProfileDir = path.join(os.homedir(), '.pdp-edge-profile');
      const success = await spawnBrowser(edgePath, userProfileDir);
      if (success) {
        browserInstance = await connectWithWS();
        console.log(pc.green('✔ Connected to Microsoft Edge.'));
        return browserInstance;
      }
    } catch (err) {
      console.log(pc.red(`❌ Edge launch failed: ${err.message}`));
    }
  }

  throw new Error('Could not launch Google Chrome or Microsoft Edge in remote debugging mode. Please close all active browser windows and try running the tool again.');
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
    throw new Error(`Field not found for: ${labelName} (Checked: ${selectorList.join(', ')})`);
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
    throw new Error(`Dropdown not found for: ${labelName} (Checked: ${selectorList.join(', ')})`);
  }

  // Wait for the dropdown to contain the desired option (up to 8 seconds)
  let optionValue = null;
  const maxWaitMs = 8000;
  const startTime = Date.now();

  while (Date.now() - startTime < maxWaitMs) {
    optionValue = await page.evaluate((sel, txt) => {
      const select = document.querySelector(sel);
      if (!select) return null;
      const options = Array.from(select.options);
      
      const normalize = (s) => {
        if (!s) return '';
        return s.toString()
          .toLowerCase()
          .replace(/&/g, 'and')
          .replace(/[^a-z0-9]/g, '');
      };

      const cleanSearch = normalize(txt);
      if (!cleanSearch) return null;

      // Tier 1: Try exact normalized match
      let match = options.find(o => normalize(o.text) === cleanSearch || normalize(o.value) === cleanSearch);
      if (match) return match.value;

      // Tier 2: Try substring normalized match
      match = options.find(o => {
        const cleanOptText = normalize(o.text);
        return cleanOptText.includes(cleanSearch) || cleanSearch.includes(cleanOptText);
      });
      if (match) return match.value;

      // Tier 3: Token-based overlap similarity (fuzzy)
      const tokenize = (s) => {
        if (!s) return [];
        return s.toString()
          .toLowerCase()
          .replace(/&/g, 'and')
          .replace(/[^a-z0-9]/g, ' ')
          .split(/\s+/)
          .filter(Boolean);
      };

      const searchTokens = tokenize(txt);
      const searchSet = new Set(searchTokens);
      if (searchSet.size === 0) return null;

      let bestMatch = null;
      let highestScore = 0;

      for (const o of options) {
        const optTokens = tokenize(o.text);
        const optSet = new Set(optTokens);
        if (optSet.size === 0) continue;

        let intersection = 0;
        for (const t of searchSet) {
          if (optSet.has(t)) {
            intersection++;
          }
        }

        const unionSize = searchSet.size + optSet.size - intersection;
        const jaccard = unionSize > 0 ? intersection / unionSize : 0;
        const coverage = intersection / searchSet.size;
        const score = (jaccard * 0.4) + (coverage * 0.6);

        if (score > highestScore && score >= 0.70) {
          highestScore = score;
          bestMatch = o;
        }
      }

      return bestMatch ? bestMatch.value : null;
    }, selector, text);

    if (optionValue) {
      break;
    }
    // Desired option not loaded yet, wait 300ms and check again
    await new Promise(r => setTimeout(r, 300));
  }

  if (!optionValue) {
    let availableText = '';
    try {
      const availableOptions = await page.evaluate((sel) => {
        const select = document.querySelector(sel);
        if (!select) return [];
        return Array.from(select.options).map(o => o.text.trim()).filter(Boolean);
      }, selector);
      if (availableOptions.length > 0) {
        availableText = `\n   Available options in "${labelName}" dropdown:\n` + 
          availableOptions.map(opt => `     - "${opt}"`).join('\n');
      } else {
        availableText = `\n   The "${labelName}" dropdown is currently empty.`;
      }
    } catch (e) {
      availableText = `\n   Could not retrieve available options: ${e.message}`;
    }
    throw new Error(`Option "${text}" not found in dropdown for: ${labelName} (timed out waiting for options to load).${availableText}`);
  }

  // Select the option natively in Puppeteer/browser frame context
  try {
    await page.select(selector, optionValue);
    // Force extra events to trigger any dynamic page updates
    await page.evaluate((sel) => {
      const el = document.querySelector(sel);
      if (el) {
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('blur', { bubbles: true }));
      }
    }, selector);
  } catch (err) {
    throw new Error(`Failed to select "${text}" in dropdown for ${labelName}: ${err.message}`);
  }

  // Wait for dynamic options or loaders to settle
  await new Promise(r => setTimeout(r, delayMs));
}

/**
 * Selects a radio button matching the text/value
 */
async function selectRadio(page, selectorList, value, delayMs, labelName) {
  if (value === undefined || value === null || value === '') return;
  const selector = await findSelector(page, selectorList);
  if (!selector) {
    throw new Error(`Radio buttons not found for: ${labelName} (Checked: ${selectorList.join(', ')})`);
  }

  const nameAttr = await page.evaluate((sel) => {
    const el = document.querySelector(sel);
    return el ? el.getAttribute('name') : null;
  }, selector);

  if (!nameAttr) {
    throw new Error(`Could not read name attribute of radio buttons for: ${labelName}`);
  }

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
    throw new Error(`Radio option "${value}" not found for: ${labelName}`);
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
  const excelFilePath = findExcelFile(config.excelFile);

  // Check Excel File Presence
  if (!excelFilePath) {
    const desktopPath = path.join(require('os').homedir(), 'Desktop', config.excelFile);
    console.log(pc.yellow(`Excel file "${config.excelFile}" not found.`));
    const response = await prompts({
      type: 'confirm',
      name: 'generate',
      message: `Would you like to generate a sample Excel template on your Desktop ("${desktopPath}")?`,
      initial: true
    });

    if (response.generate) {
      await generateTemplate(desktopPath);
      console.log(pc.green(`\n✔ Generated sample Excel: "${desktopPath}"`));
      console.log(pc.cyan(`Please populate your data in this sheet and rerun the tool.\n`));
    } else {
      console.log(pc.red('Please create an Excel sheet with activity data to proceed. Exiting.'));
    }
    return;
  }

  // Read Excel File
  const workbook = new ExcelJS.Workbook();
  console.log(pc.blue(`Reading "${path.basename(excelFilePath)}"...`));
  await workbook.xlsx.readFile(excelFilePath);
  const sheet = workbook.worksheets[0];

  // Parse headers to map columns dynamically
  const headerRow = sheet.getRow(config.startRow - 1 || 1);
  const colMap = {};
  headerRow.eachCell((cell, colNum) => {
    const val = cell.value ? cell.value.toString().trim().toLowerCase().replace(/[^a-z0-9]/g, '_') : '';
    if (val) {
      colMap[val] = colNum;
    }
  });

  const uploadStatusCol = colMap['upload_status'] || 20;

  // Map rows
  const activities = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber < config.startRow) return;

    // Build data object matching cells
    const getValByName = (name, defaultCol) => {
      const colNum = colMap[name] || defaultCol;
      const cell = row.getCell(colNum);
      if (cell.value && typeof cell.value === 'object' && cell.value.result !== undefined) {
        return cell.value.result;
      }
      return cell.value;
    };

    activities.push({
      rowNum: rowNumber,
      rowObj: row,
      theme: getValByName('theme', 1),
      activity: getValByName('activity', 2),
      focus_area: getValByName('focus_area', 3),
      activity_type: getValByName('activity_type', 4),
      activity_nature: getValByName('activity_nature', 5),
      description: getValByName('description', 6),
      indicators: getValByName('indicators', 7),
      remarks_for: getValByName('remarks_for', 8),
      targeted_population: getValByName('targeted_population', 9),
      funded_by_panchayat: getValByName('funded_by_panchayat', 10),
      completion_year: getValByName('completion_year', 11),
      completion_month: getValByName('completion_month', 12),
      completion_days: getValByName('completion_days', 13),
      start_year: getValByName('start_year', 14),
      start_month: getValByName('start_month', 15),
      beneficiaries_general: getValByName('beneficiaries_general', 16),
      beneficiaries_sc: getValByName('beneficiaries_sc', 17),
      beneficiaries_st: getValByName('beneficiaries_st', 18),
      estimated_cost: getValByName('estimated_cost', 19),
      upload_status: getValByName('upload_status', 20)
    });
  });

  if (activities.length === 0) {
    console.log(pc.yellow(`No rows to process starting from row ${config.startRow}. Exiting.`));
    return;
  }

  console.log(pc.green(`✔ Loaded ${activities.length} activity rows from Excel.`));

  // Choose launch/attach mode
  let page;

  console.log(pc.blue('Detecting browser session...'));
  try {
    browser = await launchBrowser(config);
  } catch (err) {
    console.log(pc.red(`\n❌ Browser Error: ${err.message}`));
    return;
  }

  const pages = await browser.pages();
  let activePages = [];
  for (const p of pages) {
    try {
      const url = p.url();
      if (url && !url.startsWith('chrome://') && !url.startsWith('chrome-extension://') && !url.startsWith('edge://')) {
        activePages.push(p);
      }
    } catch (e) {}
  }

  const targetUrl = 'https://egramswaraj.gov.in/addactivity.htm';
  page = activePages.find(p => {
    try {
      return p.url().toLowerCase().includes('egramswaraj.gov.in');
    } catch (e) {
      return false;
    }
  });

  if (!page) {
    page = activePages[0] || await browser.newPage();
    console.log(pc.blue(`Navigating to: ${targetUrl}...`));
    try {
      await page.goto(targetUrl, { waitUntil: 'load' });
    } catch (gotoErr) {
      console.log(pc.yellow(`⚠️ Navigation took too long or failed: ${gotoErr.message}`));
    }
  } else {
    console.log(pc.green(`✔ Found active eGramSwaraj portal tab.`));
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

    // Dynamic context helper & form visibility polling loop (waits automatically until form is found)
    let context = page;
    let hasForm = false;
    let printedWaiting = false;

    while (!hasForm) {
      try {
        // Resolve frame context in case form is inside an iframe
        context = page;
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

        const themeSelector = await findSelector(context, SELECTORS.theme);
        if (themeSelector) {
          hasForm = true;
          break;
        }
      } catch (e) {
        // page might be closed or navigating
      }

      if (!printedWaiting) {
        console.log(pc.yellow(`⏳ Waiting for GPDP Activity Form to open on: ${targetUrl}`));
        console.log(pc.cyan(`(Please log in if required and navigate to the form. Script will detect it automatically)`));
        printedWaiting = true;
      }

      await new Promise(r => setTimeout(r, 1500));
    }

    if (context !== page) {
      console.log(pc.green(`✔ Form detected inside iframe context: ${context.url()}`));
    } else {
      console.log(pc.green(`✔ Form detected on page.`));
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
          act.rowObj.getCell(uploadStatusCol).value = 'SUCCESS - ' + new Date().toLocaleString();
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
          act.rowObj.getCell(uploadStatusCol).value = 'SUCCESS - ' + new Date().toLocaleString();
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
        act.rowObj.getCell(uploadStatusCol).value = 'SKIPPED - ' + err.message;
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
