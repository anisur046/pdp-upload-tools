# PDP Excel-to-Form Autofill Tool

This tool automates the process of copying data from an Excel spreadsheet and pasting it into the **GPDP (Gram Panchayat Development Plan) Activity Details Form** in Google Chrome.

---

## Features

1. **Dual Modes**:
   - **Attach Mode (Recommended)**: Connects to a Chrome browser that is already open. This allows you to log in manually, bypass CAPTCHAs/OTPs, navigate to the GPDP form, and run the script to fill it out.
   - **Launch Mode**: Opens a brand new Chrome browser and loads the form.
2. **Template Generator**: If no Excel file is found, the tool automatically generates `data.xlsx` filled with correct headers and sample data from the user screenshot.
3. **Interactive Step-by-Step Filling**: The tool prints the row data in your command prompt, fills it in the browser, and pauses so you can verify the information, click submit, and then hit Enter to load the next row. This prevents mistakes and handles slow portal load times.
4. **Local Mock Sandbox**: A built-in mock form (`mock_form/index.html`) is provided so you can test the script locally.
5. **Robust Field Mappings**: Matches inputs using fallback lists, making it resilient to slight changes in web form names/IDs.

---

## Installation & Setup

1. Make sure you have Google Chrome installed on your machine.
2. Download or copy the tool files into a folder (e.g., `C:\Users\HP\WebstormProjects\pdp-upload-tools`).
3. If running via Node.js, run:
   ```bash
   npm install
   ```

---

## How to Use

### Step 1: Generate & Prepare your Excel Data
1. Run the script:
   ```bash
   node index.js
   ```
   *(Or double click `pdp-autofill.exe` once built)*
2. If `data.xlsx` is missing, select **Yes** to generate a template spreadsheet.
3. Open the generated `data.xlsx` and fill it with your real activity data. Keep the header names exactly as they are.

### Step 2: Open Chrome in Debugging Mode (for Attach Mode)
To let the script connect to your active Chrome browser:
1. Close all currently open Chrome windows.
2. Press `Win + R` to open the Run dialog.
3. Copy and paste the following command, then click **OK**:
   ```cmd
   chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\chrome-profile"
   ```
4. In this new Chrome browser window, go to the GPDP / eGramSwaraj portal, log in, and navigate to the **Activity Details** form.

### Step 3: Run the Auto-Filler
1. Start the tool:
   ```bash
   node index.js
   ```
   *(Or run `pdp-autofill.exe`)*
2. The tool will read the Excel file and ask if you want to scan open Chrome tabs. Press **Enter**.
3. It will detect the active GPDP form and begin filling row-by-row.
4. After filling each row, review the screen, submit/save the form on the web page, then press **Enter** in the console window to fill the next row.

---

## Configuration Settings (`config.json`)

You can edit `config.json` next to the executable to customize settings:

```json
{
  "chromePath": "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
  "excelFile": "data.xlsx",
  "startRow": 2,
  "mode": "attach",
  "delayBetweenFields": 500,
  "autoSave": false
}
```

- `mode`: Set to `"attach"` to connect to Chrome on port 9222, or `"launch"` to start a new Chrome window automatically.
- `startRow`: The spreadsheet row to start reading from (default is 2, since row 1 is the header).
- `delayBetweenFields`: Speed of data entry (in milliseconds). Increase if the web page scripts are slow to react.
- `autoSave`: If set to `true`, the tool will automatically click the Save button. (Use `false` to verify manually before submitting).

---

## Compiling to Standalone Executable (`.exe`)

To compile the application into a standalone executable file that runs on any 64-bit Windows machine without needing Node.js installed:

1. Build the binary using `pkg`:
   ```bash
   npm run build
   ```
2. The executable will be generated inside the `dist/` directory as `pdp-autofill.exe`.
3. To distribute, copy the `dist/pdp-autofill.exe` file, the `config.json` file, and the `mock_form/` folder (if using mock tests) together.
