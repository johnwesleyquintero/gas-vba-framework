# Amazon Consolidated Dashboard Script (Refactored)

This documentation provides a comprehensive guide to the refactored Google Apps Script (GAS) solution designed to consolidate Amazon FBA data into a single, dynamic, and user-friendly dashboard within Google Sheets. This script automates the process of combining information from your Amazon Listing and Inventory reports, along with external Keepa data, to provide critical insights into your product catalog.

The solution has been streamlined to generate **only one `Dashboard` sheet**, using a specific header sequence for clarity and efficiency. All previous functionalities for creating separate `Active Listings`, `Inactive Listings`, and `General Analysis` tabs have been removed.

## Data Visualization (Looker Studio Dashboard Links)

These links remain for external reference and are not affected by the script changes.

**Speedtech Mobile**
```
https://lookerstudio.google.com/reporting/e1111771-1a7d-4bc0-9c6b-d19c30fe3b76
```

**Seculife Inc.**
```
https://lookerstudio.google.com/reporting/9a779d88-66fa-49cc-bd75-82fd51c314bc
```

## Report Endpoints Used to Download/Pull Data Manually

To keep the dashboard updated, you will need to periodically download the following reports from Amazon Seller Central and Keepa and paste their contents into the corresponding tabs in your Google Sheet.

**Required Reports:**

*   **All_Listing_Report**
    ```
    https://sellercentral.amazon.com/listing/reports
    ```
*   **FBA_Inventory_Snapshot** (This report provides inventory levels, inbound quantities, and aggregated sales data)
    ```
    https://sellercentral.amazon.com/reportcentral/MANAGE_INVENTORY_HEALTH/1
    ```
*   **Keepa Data** (Used for retrieving product image addresses)
    ```
    https://keepa.com/#!viewer
    ```

---

## 1. Features

*   **Single Dashboard View:** The script generates one consolidated `Dashboard` sheet, providing a comprehensive overview of your entire product catalog in one place.
*   **Data Consolidation:** Seamlessly combines data from `All_Listing_Report`, `FBA_Inventory_Snapshot`, and `Keepa` sheets.
*   **Optimized Header Sequence:** The dashboard is generated with the following clear and logical column order:
    1.  Image
    2.  Name
    3.  Asin
    4.  SKU
    5.  Amazon Link
    6.  Listing status
    7.  Customer Orders
    8.  Available
    9.  FC Transfers
    10. FC Processing
    11. **Inbound** (New)
    12. Days of Supply
    13. Unfulfillable
    14. Sales Last 7 Days
    15. Past 31-60 Day Sales
    16. Current 30 Day Sales
    17. Monthly Difference
*   **Product Visuals:** Automatically pulls product images via Keepa data, displaying them directly in the dashboard using `=IMAGE()` formulas.
*   **Direct Amazon Links:** Generates clickable links to the Amazon product page for each ASIN.
*   **Sales Performance Metrics:** Calculates and displays key sales trends: "Sales Last 7 Days," "Current 30 Day Sales," "Past 31-60 Day Sales," and the "Monthly Difference."
*   **Complete Inventory Insights:** Includes 'Available', 'Inbound', 'Days of Supply', 'Unfulfillable', and detailed 'Reserved' inventory breakdowns (Customer Orders, FC Transfers, FC Processing).
*   **Simplified User Interface:** Provides a custom sidebar within Google Sheets with a single "Generate Dashboard" button for straightforward operation.
*   **Configurable & Resilient:** A centralized `CONFIG` object allows for easy adjustment of sheet names and column headers, providing resilience against future changes in Amazon's report formats.

## 2. Setup Instructions

1.  **Create a New Google Sheet:** Open Google Sheets and create a new blank spreadsheet.
2.  **Open Apps Script Editor:** Go to `Extensions > Apps Script`.
3.  **Create `Code.gs`:** In the Apps Script editor, replace the default content with the JavaScript code provided in **Section 7.1**.
4.  **Create `Sidebar.html`:**
    *   In the Apps Script editor, click the `+` icon next to "Files" and select `HTML`.
    *   Name the new file `Sidebar`.
    *   Paste the HTML code provided in **Section 7.2** into this new `Sidebar.html` file.
5.  **Save the Project:** Click the save icon to save both files.
6.  **Create the Dashboard Sheet:** In your Google Sheet, ensure you have a sheet named `Dashboard`. This is where the output will be generated.
7.  **Create Source Data Sheets:** Create the following sheets. **Their names must precisely match the `CONFIG.sheets` values in the `Code.gs` script.**
    *   `All_Listing_Report`
    *   `FBA_Inventory_Snapshot`
    *   `Keepa`
8.  **Populate Source Data:** Paste your downloaded reports into their respective sheets. Ensure the column headers in your reports match the `CONFIG.headers` defined in the `Code.gs` script.

## 3. Usage

1.  **Open the Spreadsheet:** When you open your Google Sheet, a custom menu "Dashboard Tools" will appear.
2.  **Open Controls:** Click `Dashboard Tools > Open Dashboard Controls` to show the sidebar.
3.  **Enter Optional Context:** In the sidebar, you can enter a "Dashboard Context" string (e.g., "July Inventory Sync").
4.  **Generate the Report:** Click the **Generate Dashboard** button.
5.  **Monitor Progress:** A loading spinner will appear. Status messages (success or error) will be displayed at the top of the sidebar.
6.  **Review Dashboard:** The script will clear and update the `Dashboard` sheet with the newly generated data.

## 4. Configuration

The `CONFIG` object at the top of `Code.gs` is the central hub for all settings.

```javascript
const CONFIG = {
  sheets: {
    listing: 'All_Listing_Report',
    inventory: 'FBA_Inventory_Snapshot',
    keepa: 'Keepa',
    dashboard: 'Dashboard' // The only target sheet
  },
  headers: {
    // All_Listing_Report Headers
    listingSku: 'seller-sku',
    asin: 'asin1',
    itemName: 'item-name',
    status: 'status',

    // FBA_Inventory_Snapshot Headers
    invSku: 'sku',
    available: 'available',
    inbound: 'inbound-quantity', // Added for the new column
    daysOfSupply: 'days-of-supply',
    unfulfillable: 'unfulfillable-quantity',
    reservedCustomer: 'Reserved Customer Order',
    reservedTransfer: 'Reserved FC Transfer',
    reservedProcessing: 'Reserved FC Processing',
    salesT7: 'sales-shipped-last-7-days',
    salesT30: 'sales-shipped-last-30-days',
    salesT60: 'sales-shipped-last-60-days',

    // Keepa Headers
    keepaAsin: 'ASIN',
    keepaImage: 'Image'
  },
  image: {
    width: 100
  }
};
```

*   **`CONFIG.sheets`**: Define the exact names of your source sheets and the single `Dashboard` target sheet.
*   **`CONFIG.headers`**: This is the most important section to maintain. If Amazon changes a column name in a report (e.g., `inbound-quantity` becomes `inbound_qty`), you only need to update the corresponding value here.
*   **`CONFIG.image.width`**: Adjust the width of the image column in the dashboard.

## 5. Code Structure and Explanation

### Core Functions (`Code.gs`)

*   **`onOpen()`**: Creates the "Dashboard Tools" menu when the spreadsheet is opened.
*   **`showSidebar()`**: Displays the `Sidebar.html` content.
*   **`createConsolidatedDashboard(...)`**: The main function called by the sidebar. It orchestrates the entire process by calling the core report generator.
*   **`_generateDashboardReport(...)`**: This is the core logic function. It fetches data from all source sheets, processes it into usable maps, constructs the final dataset according to the specified header sequence, and writes everything to the `Dashboard` sheet.

### Sidebar HTML (`Sidebar.html`)

This file defines the simplified UI in the sidebar. It contains an optional context input and a single "Generate Dashboard" button. Its client-side JavaScript captures the user's click and calls the `createConsolidatedDashboard` function in `Code.gs`.

## 6. Troubleshooting and Common Issues

*   **"Missing required columns in..." Error:** This is the most common error. Double-check that:
    *   Your source sheets are named exactly as defined in `CONFIG.sheets`.
    *   The column headers in your reports (e.g., `seller-sku`, `inbound-quantity`, `sales-shipped-last-30-days`) exactly match the values defined in `CONFIG.headers`.
*   **Incorrect Sales or Inventory Data:**
    *   Confirm that the `FBA_Inventory_Snapshot` sheet is correctly populated with the latest report from Amazon.
    *   Verify all relevant headers exist in the sheet and match the `CONFIG.headers` object.
*   **Images Not Loading:**
    *   Ensure the `Keepa` sheet exists and contains `ASIN` and `Image` columns with valid public image URLs.
    *   Verify the `asin1` values in your `All_Listing_Report` match the `ASIN` values in your `Keepa` sheet.

## 7. Source Code

Below are the complete and final source code files for the project.

### 7.1. `Code.gs`

```javascript
/**
 * @OnlyCurrentDoc
 * This script creates a single, consolidated Amazon dashboard.
 * VERSION 16.0: Single Dashboard Refactor.
 * - Refactored by Gemini to remove Active, Inactive, and Analysis report generation.
 * - The script now creates only one sheet: 'Dashboard'.
 * - The header sequence has been updated to match the user's specification.
 * - Added 'Inbound' quantity to the dashboard, pulled from the inventory report.
 * - Simplified the sidebar and main function router.
 */

// --- CONFIGURATION --- //
const CONFIG = {
  sheets: {
    listing: 'All_Listing_Report',
    inventory: 'FBA_Inventory_Snapshot',
    keepa: 'Keepa',
    dashboard: 'Dashboard' // The only target sheet
  },
  headers: {
    // All_Listing_Report Headers
    listingSku: 'seller-sku',
    asin: 'asin1',
    itemName: 'item-name',
    status: 'status',

    // FBA_Inventory_Snapshot Headers
    invSku: 'sku',
    available: 'available',
    inbound: 'inbound-quantity', // Added for the new column
    daysOfSupply: 'days-of-supply',
    unfulfillable: 'unfulfillable-quantity',
    reservedCustomer: 'Reserved Customer Order',
    reservedTransfer: 'Reserved FC Transfer',
    reservedProcessing: 'Reserved FC Processing',
    salesT7: 'sales-shipped-last-7-days',
    salesT30: 'sales-shipped-last-30-days',
    salesT60: 'sales-shipped-last-60-days',

    // Keepa Headers
    keepaAsin: 'ASIN',
    keepaImage: 'Image'
  },
  image: {
    width: 100
  }
};
// --- END CONFIGURATION ---

// --- UI & INITIALIZATION --- //

/**
 * Adds a custom menu to the spreadsheet UI when the file is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.createMenu('Dashboard Tools')
      .addItem('Open Dashboard Controls', 'showSidebar')
      .addToUi();
  } catch (e) {
    Logger.log(`Error creating custom menu: ${e.message}`);
  }
}

/**
 * Displays the HTML sidebar for user interaction.
 */
function showSidebar() {
  const ui = SpreadsheetApp.getUi();
  try {
    const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Dashboard Controls')
      .setWidth(300);
    ui.showSidebar(html);
  } catch (e) {
    _handleError(e, ui, 'showSidebar');
  }
}

// --- MAIN FUNCTION --- //

/**
 * Main function called from the sidebar to generate the dashboard.
 * @param {string} contextString User-provided context for the report.
 * @returns {string} A success message to be displayed in the sidebar alert.
 */
function createConsolidatedDashboard(contextString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  try {
    return _generateDashboardReport(ss, ui, contextString);
  } catch (e) {
    _handleError(e, ui, 'createConsolidatedDashboard');
    throw e; // Re-throw to trigger failure handler on the client-side
  }
}


// --- CORE DASHBOARD & HELPER FUNCTIONS --- //

/**
 * Generates the consolidated dashboard report.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @param {GoogleAppsScript.Spreadsheet.Ui} ui The spreadsheet UI instance.
 * @param {string} contextString User-provided context for the report.
 * @returns {string} Success message.
 */
function _generateDashboardReport(ss, ui, contextString = '') {
  const targetSheetName = CONFIG.sheets.dashboard;

  // Fetch all necessary data.
  const listingDataObject = _getValidatedData(ss, CONFIG.sheets.listing, [CONFIG.headers.listingSku, CONFIG.headers.asin, CONFIG.headers.itemName, CONFIG.headers.status]);
  const inventoryDataObject = _getValidatedData(ss, CONFIG.sheets.inventory, [
    CONFIG.headers.invSku, CONFIG.headers.available, CONFIG.headers.inbound, CONFIG.headers.daysOfSupply, CONFIG.headers.unfulfillable,
    CONFIG.headers.reservedCustomer, CONFIG.headers.reservedTransfer, CONFIG.headers.reservedProcessing,
    CONFIG.headers.salesT7, CONFIG.headers.salesT30, CONFIG.headers.salesT60
  ]);
  const keepaDataObject = _getValidatedData(ss, CONFIG.sheets.keepa, [CONFIG.headers.keepaAsin, CONFIG.headers.keepaImage]);

  // Process data into maps for easy lookup
  const inventoryMap = processInventory(inventoryDataObject);
  const salesMap = processSales(inventoryDataObject);
  const keepaMap = processKeepa(keepaDataObject);

  const {
    headers: listingHeaders,
    data: listingData
  } = listingDataObject;
  const skuListingIndex = listingHeaders.indexOf(CONFIG.headers.listingSku);
  const asinListingIndex = listingHeaders.indexOf(CONFIG.headers.asin);
  const nameListingIndex = listingHeaders.indexOf(CONFIG.headers.itemName);
  const statusListingIndex = listingHeaders.indexOf(CONFIG.headers.status);

  const dashboardData = [];
  listingData.forEach(row => {
    const sku = row[skuListingIndex];
    if (!sku) return;

    const inventory = inventoryMap[sku] || {};
    const sales = salesMap[sku] || {};
    const asin = row[asinListingIndex];
    let imageFormula = `=IMAGE("https://via.placeholder.com/${CONFIG.image.width}x${CONFIG.image.width / 1.2}.png?text=No+Image", 4, ${CONFIG.image.width}, ${CONFIG.image.width / 1.2})`;
    let amazonLink = "";
    
    if (asin) {
      const keepaImageUrl = keepaMap[asin];
      if (keepaImageUrl) {
        imageFormula = `=IMAGE("${keepaImageUrl}", 4, ${CONFIG.image.width}, ${CONFIG.image.width / 1.2})`;
      }
      amazonLink = `https://www.amazon.com/dp/${asin}`;
    }

    // Push data in the new, specified order
    dashboardData.push([
      imageFormula,
      row[nameListingIndex] || 'N/A',
      asin || 'N/A',
      sku,
      amazonLink,
      row[statusListingIndex] || 'N/A',
      inventory.reservedCustomerOrder || 0,
      inventory.available || 0,
      inventory.reservedFcTransfer || 0,
      inventory.reservedFcProcessing || 0,
      inventory.inbound || 0,
      inventory.daysOfSupply || 0,
      inventory.unfulfillable || 0,
      sales.salesLast7Days || 0,
      sales.pastMonthSales || 0,
      sales.currentMonthSales || 0,
      (sales.currentMonthSales || 0) - (sales.pastMonthSales || 0)
    ]);
  });

  let targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) {
    targetSheet.clear();
  } else {
    targetSheet = ss.insertSheet(targetSheetName);
  }

  // Headers in the new, specified order
  const dashboardHeaders = [
    'Image', 'Name', 'Asin', 'SKU', 'Amazon Link', 'Listing status', 'Customer Orders',
    'Available', 'FC Transfers', 'FC Processing', 'Inbound', 'Days of Supply', 'Unfulfillable',
    'Sales Last 7 Days', 'Past 31-60 Day Sales', 'Current 30 Day Sales', 'Monthly Difference'
  ];
  
  let startRow = 1;
  const reportTitle = "All Listings";

  if (contextString) {
    targetSheet.getRange(startRow, 1).setValue(`Dashboard Context: ${contextString}`).setFontWeight('bold');
    targetSheet.getRange(startRow, 1, 1, dashboardHeaders.length).merge();
    startRow++;
  }

  targetSheet.getRange(startRow, 1).setValue(`Report Type: ${reportTitle}`).setFontWeight('bold');
  targetSheet.getRange(startRow, 1, 1, dashboardHeaders.length).merge();
  startRow++;

  targetSheet.getRange(startRow, 1, 1, dashboardHeaders.length).setValues([dashboardHeaders]).setFontWeight('bold');

  if (dashboardData.length > 0) {
    const imageFormulas = dashboardData.map(row => [row[0]]);
    const dataWithoutImageFormula = dashboardData.map(row => row.slice(1));
    targetSheet.getRange(startRow + 1, 2, dataWithoutImageFormula.length, dataWithoutImageFormula[0].length).setValues(dataWithoutImageFormula);
    targetSheet.getRange(startRow + 1, 1, imageFormulas.length, 1).setFormulas(imageFormulas);
  }

  targetSheet.setColumnWidth(1, CONFIG.image.width);
  targetSheet.autoResizeColumns(2, dashboardHeaders.length - 1);

  return `"${targetSheetName}" has been updated with ${dashboardData.length} products.`;
}

/**
 * Processes inventory data into a SKU-based map.
 * @param {{headers: string[], data: any[][]}} inventoryDataObject The validated inventory data object.
 * @returns {object} A map with SKU as key and inventory details as value.
 */
function processInventory(inventoryDataObject) {
  const map = {};
  const {
    headers,
    data
  } = inventoryDataObject;
  const skuIndex = headers.indexOf(CONFIG.headers.invSku);
  const availableIndex = headers.indexOf(CONFIG.headers.available);
  const inboundIndex = headers.indexOf(CONFIG.headers.inbound);
  const daysOfSupplyIndex = headers.indexOf(CONFIG.headers.daysOfSupply);
  const unfulfillableIndex = headers.indexOf(CONFIG.headers.unfulfillable);
  const reservedCustomerIndex = headers.indexOf(CONFIG.headers.reservedCustomer);
  const reservedTransferIndex = headers.indexOf(CONFIG.headers.reservedTransfer);
  const reservedProcessingIndex = headers.indexOf(CONFIG.headers.reservedProcessing);

  data.forEach(row => {
    const sku = row[skuIndex];
    if (sku) {
      map[sku] = {
        available: parseFloat(row[availableIndex]) || 0,
        inbound: parseFloat(row[inboundIndex]) || 0,
        daysOfSupply: parseFloat(row[daysOfSupplyIndex]) || 0,
        unfulfillable: parseFloat(row[unfulfillableIndex]) || 0,
        reservedCustomerOrder: parseFloat(row[reservedCustomerIndex]) || 0,
        reservedFcTransfer: parseFloat(row[reservedTransferIndex]) || 0,
        reservedFcProcessing: parseFloat(row[reservedProcessingIndex]) || 0
      };
    }
  });
  return map;
}

/**
 * Processes aggregated sales data from the inventory report into a SKU-based map.
 * @param {{headers: string[], data: any[][]}} inventoryDataObject The validated inventory data object.
 * @returns {object} A map with SKU as key and sales details as value.
 */
function processSales(inventoryDataObject) {
  const map = {};
  const {
    headers,
    data
  } = inventoryDataObject;
  const skuIndex = headers.indexOf(CONFIG.headers.invSku);
  const salesT7Index = headers.indexOf(CONFIG.headers.salesT7);
  const salesT30Index = headers.indexOf(CONFIG.headers.salesT30);
  const salesT60Index = headers.indexOf(CONFIG.headers.salesT60);

  data.forEach(row => {
    const sku = String(row[skuIndex]).trim();
    if (sku) {
      const salesT7 = parseInt(row[salesT7Index], 10) || 0;
      const salesT30 = parseInt(row[salesT30Index], 10) || 0;
      const salesT60 = parseInt(row[salesT60Index], 10) || 0;
      const pastMonthSales = salesT60 - salesT30;
      map[sku] = {
        salesLast7Days: salesT7,
        currentMonthSales: salesT30,
        pastMonthSales: pastMonthSales < 0 ? 0 : pastMonthSales
      };
    }
  });
  return map;
}

/**
 * Processes Keepa data into an ASIN-based map for image URLs.
 * @param {{headers: string[], data: any[][]}} keepaDataObject The validated Keepa data object.
 * @returns {object} A map with ASIN as key and image URL as value.
 */
function processKeepa(keepaDataObject) {
  const map = {};
  const {
    headers,
    data
  } = keepaDataObject;
  const asinIndex = headers.indexOf(CONFIG.headers.keepaAsin);
  const imageIndex = headers.indexOf(CONFIG.headers.keepaImage);

  data.forEach(row => {
    const asin = String(row[asinIndex]).trim();
    const imageUrl = String(row[imageIndex]).trim();
    if (asin && imageUrl) map[asin] = imageUrl;
  });
  return map;
}

// --- UTILITY FUNCTIONS --- //

/**
 * A centralized error handler. Logs the error and shows a user-friendly alert.
 * @param {Error} e The error object that was caught.
 * @param {GoogleAppsScript.Spreadsheet.Ui | null} ui The spreadsheet UI instance.
 * @param {string} functionName The name of the function where the error occurred.
 */
function _handleError(e, ui, functionName = 'unknown') {
  const errorMessage = `An error occurred in function '${functionName}': ${e.message}`;
  const errorStack = e.stack ? `\nStack Trace: ${e.stack}` : '';
  Logger.log(`${errorMessage}${errorStack}`);
  if (ui && typeof ui.alert === 'function') {
    ui.alert('Script Error', errorMessage, ui.ButtonSet.OK);
  }
}

/**
 * Validates a sheet and its headers, returning the data if valid.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @param {string} sheetName The name of the sheet to validate.
 * @param {string[]} requiredHeaders An array of header names that must be present.
 * @throws {Error} If the sheet is not found, is empty, or is missing required headers.
 * @returns {{headers: string[], data: any[][]}} An object containing the header row and the data rows.
 */
function _getValidatedData(ss, sheetName, requiredHeaders) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Required sheet "${sheetName}" not found.`);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log(`Warning: Sheet "${sheetName}" is empty or contains only headers.`);
    return {
      headers: (data[0] || []).map(h => String(h).trim()),
      data: []
    };
  }
  const headers = data[0].map(h => String(h).trim());
  const headerSet = new Set(headers);
  const missingHeaders = requiredHeaders.filter(h => !headerSet.has(h));
  if (missingHeaders.length > 0) {
    throw new Error(`Missing required columns in "${sheetName}": ${missingHeaders.join(', ')}`);
  }
  return {
    headers: headers,
    data: data.slice(1)
  };
}
```

### 7.2. `Sidebar.html`

```html
<!DOCTYPE html>
<html>
<head>
    <base target="_top">
    <!-- Using Google's standard CSS package for a clean, modern look -->
    <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">
    <style>
        /* General Body and Container Refinements */
        body {
            padding: 0;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            color: #333;
            margin: 0;
            background-color: #f0f0f0;
        }
        .container {
            display: flex;
            flex-direction: column;
            gap: 20px;
            padding: 15px;
            background-color: #ffffff;
        }

        /* Brand Name Header */
        .brand-header {
            text-align: center;
            padding: 20px 15px;
            margin-bottom: 20px;
            background-color: #ffffff;
            border-bottom: 1px solid #e0e0e0;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            font-size: 1.8em;
            font-weight: bold;
            letter-spacing: 0.5px;
            display: flex;
            justify-content: center;
            align-items: baseline;
        }

        .brand-header span {
            display: inline-block;
        }

        .brand-va {
            color: #000000; /* Black for "VA" */
        }

        .brand-xtreme {
            background: linear-gradient(to right, #FFD700, #DAA520); /* Richer Gold Gradient */
            -webkit-background-clip: text;
            background-clip: text;
            -webkit-text-fill-color: transparent;
            text-fill-color: transparent;
            margin: 0 4px;
            font-size: 1.1em;
        }

        .brand-ph {
            color: #000000; /* Black for "PH" */
        }
        
        /* Input & Select Group Styles */
        .form-group {
            margin-bottom: 15px;
        }

        label {
            font-weight: 600;
            display: block;
            margin-bottom: 8px;
            color: #333;
        }

        input[type="text"] {
            width: 100%;
            padding: 8px;
            box-sizing: border-box;
            border: 1px solid #ccc;
            border-radius: 4px;
            font-size: 0.95em;
            background-color: #fff;
        }
        
        input[type="text"]:focus {
            border-color: #DAA520; /* Gold focus highlight */
            outline: none;
            box-shadow: 0 0 0 2px rgba(218, 165, 32, 0.2);
        }

        /* --- REFINED BUTTON STYLES --- */
        #generateBtn {
            width: 100%;
            padding: 12px 15px;
            border: 1px solid #DAA520;
            border-radius: 5px;
            font-size: 1.1em;
            font-weight: bold;
            cursor: pointer;
            transition: all 0.2s ease-in-out;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            display: flex;
            justify-content: center;
            align-items: center;
            background: linear-gradient(to right, #FFD700, #FFA500);
            color: #000000; /* Black text for high contrast */
        }

        #generateBtn:hover:not(:disabled) {
            background: linear-gradient(to right, #e6c200, #e69500);
            box-shadow: 0 3px 6px rgba(0,0,0,0.2);
            transform: translateY(-1px);
        }
        
        #generateBtn:disabled {
            background: #e0e0e0;
            color: #9e9e9e;
            border-color: #e0e0e0;
            cursor: not-allowed;
            box-shadow: none;
        }

        /* Spinner and Alert styles */
        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #DAA520;
            border-radius: 50%;
            width: 25px;
            height: 25px;
            animation: spin 1s linear infinite;
            margin: 20px auto;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        .alert {
            padding: 12px;
            border-radius: 5px;
            color: white;
            margin-top: 15px;
            font-weight: 500;
        }
        .alert-success { background-color: #28a745; }
        .alert-error { background-color: #dc3545; }
    </style>
</head>
<body>
    <div class="brand-header">
        <span class="brand-va">VA</span><span class="brand-xtreme">Xtreme</span><span class="brand-ph">PH</span>
    </div>

    <div class="container">
        <div id="statusAlert" class="alert" style="display:none;"></div>

        <div class="form-group">
            <label for="contextInput">Dashboard Context (Optional)</label>
            <input type="text" id="contextInput" placeholder="e.g., Q2 Inventory Review">
        </div>
        
        <!-- The button now directly calls the single dashboard generation function -->
        <button id="generateBtn" onclick="runDashboardGeneration()">Generate Dashboard</button>

        <div id="loadingSpinner" class="spinner" style="display:none;"></div>
    </div>

    <script>
        // Renamed function for clarity
        function runDashboardGeneration() {
            const contextString = document.getElementById('contextInput').value;
            
            const generateButton = document.getElementById('generateBtn');
            const spinner = document.getElementById('loadingSpinner');
            const alertBox = document.getElementById('statusAlert');

            generateButton.disabled = true;
            spinner.style.display = 'block';
            alertBox.style.display = 'none';

            google.script.run
                .withSuccessHandler(message => {
                    generateButton.disabled = false;
                    spinner.style.display = 'none';
                    alertBox.textContent = message || 'Operation completed successfully!';
                    alertBox.className = 'alert alert-success';
                    alertBox.style.display = 'block';
                })
                .withFailureHandler(error => {
                    generateButton.disabled = false;
                    spinner.style.display = 'none';
                    alertBox.textContent = 'Error: ' + error.message;
                    alertBox.className = 'alert alert-error';
                    alertBox.style.display = 'block';
                })
                // UPDATED: Call the new, refactored backend function
                .createConsolidatedDashboard(contextString);
        }
    </script>
</body>
</html>
```
