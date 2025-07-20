# Amazon Consolidated Dashboard Script for Google Sheets

This documentation provides a comprehensive guide to the Google Apps Script (GAS) solution designed to consolidate Amazon FBA data into dynamic, user-friendly dashboards within Google Sheets. This script automates the process of combining information from your Amazon Listing and Inventory reports, along with external Keepa data, to provide critical insights into your product catalog.

The solution supports generating distinct dashboards for 'All Listings', 'Active Listings', and 'Inactive Listings', with intelligent filtering and sorting capabilities. A previous feature for a detailed 'Analysis' tab has been disabled due to changes in Amazon's reporting structure, which no longer provides the necessary transactional data.

## Data Visualization (Looker Studio Dashboard Links)

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
*   **FBA_Inventory_Snapshot** (This report provides both inventory levels and aggregated sales data)
    ```
    https://sellercentral.amazon.com/reportcentral/MANAGE_INVENTORY_HEALTH/1
    ```
*   **Keepa Data** (Used for retrieving product image addresses)
    ```
    https://keepa.com/#!viewer
    ```

**Deprecated Reports (No longer used by the script):**
The following reports were used in previous versions but are not required for the current script to function: `Reserved_Inventory_Snapshot` and `FBA_Shipments_60_Days`.

---

## 1. Features

*   **Data Consolidation:** Seamlessly combines data from `All_Listing_Report`, `FBA_Inventory_Snapshot`, and `Keepa` sheets.
*   **Multiple Dashboard Views:** Generates three distinct dashboards:
    *   **All Listings:** A comprehensive view of all your products.
    *   **Active Listings:** Filters for active listings and sorts them by 'Days of Supply' (lowest first) to help identify items needing attention.
    *   **Inactive Listings:** Filters for inactive listings and specifically targets FBA-SKU items for potential archiving or relisting.
*   **Product Visuals:** Automatically pulls product images via Keepa data, displaying them directly in the dashboard using `=IMAGE()` formulas.
*   **Direct Amazon Links:** Generates clickable links to the Amazon product page for each ASIN.
*   **Sales Performance Metrics:** Calculates and displays key sales trends using aggregated data: "Sales Last 7 Days," "Current 30 Day Sales," "Past 31-60 Day Sales," and the "Monthly Difference."
*   **Inventory Insights:** Includes 'Available' quantity, 'Days of Supply', 'Unfulfillable', and detailed 'Reserved' inventory breakdowns (Customer Orders, FC Transfers, FC Processing).
*   **User-Friendly Interface:** Provides a custom sidebar within Google Sheets for easy generation of different report types via a simple dropdown menu.
*   **Configurable & Resilient:** A centralized `CONFIG` object allows for easy adjustment of sheet names and column headers, providing resilience against future changes in Amazon's report formats.

## 2. Setup Instructions

To implement this solution, you will need a Google Sheet and access to Google Apps Script.

### Step-by-Step Installation:

1.  **Create a New Google Sheet:** Open Google Sheets and create a new blank spreadsheet.
2.  **Open Apps Script Editor:** Go to `Extensions > Apps Script`. This will open a new browser tab with the Apps Script editor.
3.  **Create `Code.gs`:** In the Apps Script editor, you'll see a `Code.gs` file by default. Replace its content with the JavaScript code provided in **Section 7.1**.
4.  **Create `Sidebar.html`:**
    *   In the Apps Script editor, click the `+` icon next to "Files" in the left sidebar, then select `HTML`.
    *   Name the new file `Sidebar`.
    *   Paste the HTML code provided in **Section 7.2** into this new `Sidebar.html` file.
5.  **Save the Project:** Click the save icon (floppy disk) or press `Ctrl + S` (`Cmd + S` on Mac) to save both files.
6.  **Rename the Default Sheet:** In your Google Sheet, ensure you have a sheet named `Dashboard`. If your default sheet is `Sheet1`, rename it to `Dashboard`. This is the default target for the 'All Listings' report.
7.  **Create Source Data Sheets:** Create the following sheets in your Google Spreadsheet. **Their names must precisely match the `CONFIG.sheets` values in the `Code.gs` script.**
    *   `All_Listing_Report`
    *   `FBA_Inventory_Snapshot`
    *   `Keepa`
8.  **Populate Source Data:** Import or paste your downloaded reports into their respective sheets. Ensure the column headers in your reports match the `CONFIG.headers` defined in the `Code.gs` script. **This is critical for the script to locate the correct data.**

## 3. Usage

Once set up, the script integrates directly into your Google Sheet interface.

1.  **Open the Spreadsheet:** When you open your Google Sheet, a custom menu "Dashboard Tools" will appear, and the "Dashboard Controls" sidebar will automatically open on the right.
2.  **Enter Optional Context:** In the sidebar, you can enter an optional "Dashboard Context" string (e.g., "Q2 Inventory Review"). This text will be added to the top of the generated dashboard for reference.
3.  **Select Analysis Type:** Use the dropdown menu to choose which report you want to generate.
    *   **All Listings Dashboard:** Creates a consolidated report of all listings on the `Dashboard` sheet.
    *   **Inactive Listings (FBA-SKU):** Creates a filtered report of inactive FBA listings on the `Inactive_Listing` sheet.
    *   **Active Listings (Sorted by DoS):** Creates a filtered report of active listings, sorted by 'Days of Supply', on the `Active_Listing` sheet.
    *   **General Analysis Tab:** This option is disabled. Selecting it will produce an error message explaining that the feature is unavailable due to the new data schema.
4.  **Generate the Report:** Click the **Generate** button.
5.  **Monitor Progress:** A loading spinner will appear while the script runs. Status messages (success or error) will be displayed at the top of the sidebar.
6.  **Review Dashboards:** The script will create or clear and update the relevant target sheet (`Dashboard`, `Active_Listing`, `Inactive_Listing`) with the generated data.

## 4. Configuration

The `CONFIG` object at the top of `Code.gs` is the central hub for all user-adjustable settings.

```javascript
const CONFIG = {
  sheets: {
    listing: 'All_Listing_Report',
    inventory: 'FBA_Inventory_Snapshot',
    shipments: 'FBA_Shipments_60_Days', // Not used in current version
    reserved: 'Reserved_Inventory_Snapshot', // Not used in current version
    keepa: 'Keepa',
    dashboard: 'Dashboard',
    activeListings: 'Active_Listing',
    inactiveListings: 'Inactive_Listing',
    analysis: 'Analysis',
    chartData: '_AnalysisChartData'
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
    width: 100 // Width of the image column in pixels
  },
  listingStatus: {
    active: 'Active',
    inactive: 'Inactive'
  }
};
```

*   **`CONFIG.sheets`**: Define the exact names of your source and target sheets. **Ensure these match your Google Sheet tab names perfectly.**
*   **`CONFIG.headers`**: This is the most important section to maintain. If Amazon changes the column names in their reports, you *only* need to update the corresponding values here. For instance, if `seller-sku` becomes `merchant-sku`, you would update `listingSku: 'merchant-sku'`.
*   **`CONFIG.image.width`**: Adjust the width of the image column in the generated dashboards.
*   **`CONFIG.listingStatus`**: Confirm the exact strings used in your Amazon Listing Report's 'status' column for 'Active' and 'Inactive' listings.

## 5. Code Structure and Explanation

The script is organized into core functions and helper functions to ensure modularity and readability.

### Core Functions (`Code.gs`)

*   **`onOpen()`**: Runs automatically when the spreadsheet opens. It creates the "Dashboard Tools" menu and opens the sidebar.
*   **`showSidebar()`**: Displays the `Sidebar.html` content.
*   **`createConsolidatedDashboardFromSidebar(...)`**: The main router called by the sidebar. It validates the `reportType` and calls the appropriate generator function. It specifically blocks the `'analysis'` type, throwing an error.
*   **`generate...Dashboard(...)`**: Wrapper functions that call the main `_generateDashboardReport` function with the correct `reportType` and target sheet name.
*   **`_generateDashboardReport(...)`**: This is the core logic function.
    *   **Data Retrieval & Validation:** Fetches data from the `listing`, `inventory`, and `keepa` sheets, throwing an error if required sheets or headers are missing.
    *   **Data Processing:** Uses helper functions (`processInventory`, `processSales`, `processKeepa`) to transform raw data into SKU- and ASIN-based maps for efficient lookup.
    *   **Dashboard Construction:** Iterates through each listing, enriching it with corresponding inventory, sales data, and Keepa image formulas.
    *   **Filtering and Sorting:** Applies filtering (for active/inactive) and sorting (by Days of Supply for active) based on the `reportType`.
    *   **Sheet Output:** Clears or creates the target sheet, writes headers, and populates the data.
    *   **Formatting:** Sets column widths for better readability.

### Helper Functions (`Code.gs`)

*   **`processInventory(inventoryDataObject)`**: Parses the `FBA_Inventory_Snapshot` data. It extracts 'available', 'days of supply', 'unfulfillable', and 'reserved' quantities, mapping them by SKU.
*   **`processSales(inventoryDataObject)`**: **This function now processes the `FBA_Inventory_Snapshot` data, not a separate sales report.** It extracts aggregated sales for the last 7, 30, and 60 days. It then calculates 'current 30-day sales' and 'past 31-60 day sales' for each SKU.
*   **`processKeepa(keepaDataObject)`**: Extracts ASINs and their corresponding image URLs from the `Keepa` sheet.
*   **`_getValidatedData(...)`**: A robust utility function that retrieves data from a sheet, validates that it's not empty, and confirms that all required headers are present.
*   **`_handleError(...)`**: A centralized function for logging detailed errors and displaying user-friendly alerts in the UI.

### Sidebar HTML (`Sidebar.html`)

This file defines the UI in the sidebar. It uses a dropdown menu and a single "Generate" button. Its client-side JavaScript captures the user's selections and calls the main `createConsolidatedDashboardFromSidebar` function in `Code.gs` using `google.script.run`.

## 6. Troubleshooting and Common Issues

*   **"Missing required columns in..." Error:** This is the most common error. Double-check that:
    *   Your source sheets (`All_Listing_Report`, `FBA_Inventory_Snapshot`, `Keepa`) are named exactly as defined in `CONFIG.sheets`.
    *   The column headers in your uploaded reports (e.g., `seller-sku`, `asin1`, `available`, `sales-shipped-last-30-days`) exactly match the values defined in `CONFIG.headers`. Verify against your latest downloaded reports.
*   **"Analysis Tab cannot be generated..." Error:** This is expected behavior. If you select "General Analysis Tab" from the dropdown, this error will appear. The feature is disabled because the script no longer has access to the detailed transactional sales data it required.
*   **Incorrect Sales Data:**
    *   Confirm that the `FBA_Inventory_Snapshot` sheet is correctly populated.
    *   Verify that the headers `sales-shipped-last-7-days`, `sales-shipped-last-30-days`, and `sales-shipped-last-60-days` exist in the sheet and in the `CONFIG.headers` object. The script relies entirely on these columns for all sales metrics.
*   **Images Not Loading:**
    *   Ensure the `Keepa` sheet exists and contains `ASIN` and `Image` columns with valid public image URLs.
    *   Verify the `asin1` values in your `All_Listing_Report` match the `ASIN` values in your `Keepa` sheet.
*   **Script Authorization:** The first time you run the script, Google will ask for authorization to access your spreadsheet. You must grant these permissions for the script to function.

## 7. Source Code

Below are the complete source code files for the Google Apps Script project.

### 7.1. `Code.gs`

```javascript
/**
 * @OnlyCurrentDoc
 * This script creates consolidated Amazon dashboards.
 * VERSION 15.2: Authorization and UI Refinement.
 * - Fixed a permissions error by modifying the onOpen() trigger. The sidebar now opens only when the user clicks the menu item, which is a more robust authorization pattern.
 * - Removed unused 'analysis' and 'chartData' references from the CONFIG object for better clarity.
 * - Script logic remains focused on the new schema using 'FBA_Inventory_Snapshot' for both inventory and sales data.
 */

// --- CONFIGURATION --- //
const CONFIG = {
  sheets: {
    listing: 'All_Listing_Report',
    inventory: 'FBA_Inventory_Snapshot',
    keepa: 'Keepa',
    dashboard: 'Dashboard',
    activeListings: 'Active_Listing',
    inactiveListings: 'Inactive_Listing'
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
  },
  listingStatus: {
    active: 'Active',
    inactive: 'Inactive'
  }
};
// --- END CONFIGURATION ---

// --- UI & INITIALIZATION --- //

/**
 * Adds a custom menu to the spreadsheet UI when the file is opened.
 * Note: A simple onOpen() trigger runs in a restricted mode and cannot perform actions
 * that require authorization, like showing a sidebar. The sidebar is now opened
 * by the user via the menu item to ensure proper permissions are handled.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.createMenu('Dashboard Tools')
      .addItem('Open Dashboard Controls', 'showSidebar')
      .addToUi();
  } catch (e) {
    // Use a simple log here as UI alerts may not work in all onOpen contexts.
    Logger.log(`Error creating custom menu: ${e.message}`);
  }
}

/**
 * Displays the HTML sidebar for user interaction.
 * This function is called by the user from the custom menu.
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

// --- MAIN ROUTER FUNCTION --- //

/**
 * Main router function called from the sidebar. It directs traffic to the correct report generator.
 * @param {string} contextString User-provided context for the report.
 * @param {string} reportType The type of report to generate ('all', 'inactive', 'active', 'analysis').
 * @returns {string} A success message to be displayed in the sidebar alert.
 */
function createConsolidatedDashboardFromSidebar(contextString, reportType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  try {
    // The switch statement is a clean way to route execution.
    switch (reportType) {
      case 'all':
        return generateAllListingsDashboard(ss, ui, contextString);
      case 'inactive':
        return generateInactiveListingsDashboard(ss, ui, contextString);
      case 'active':
        return generateActiveListingsDashboard(ss, ui, contextString);
      case 'analysis':
        // This feature is disabled due to schema changes.
        throw new Error('The Analysis Tab cannot be generated. The new data schema does not contain the required transactional sales data.');
      default:
        throw new Error('Invalid report type specified.');
    }
  } catch (e) {
    // All errors from child functions will be caught here.
    _handleError(e, ui, `createConsolidatedDashboardFromSidebar`);
    // Re-throw the error to ensure the .withFailureHandler() on the client-side catches it.
    throw e;
  }
}

// --- DASHBOARD GENERATORS --- //

/**
 * Generates the 'All Listings' dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @param {GoogleAppsScript.Spreadsheet.Ui} ui The spreadsheet UI instance.
 * @param {string} c User-provided context string.
 * @returns {string} Success message.
 */
function generateAllListingsDashboard(ss, ui, c) {
  return _generateDashboardReport(ss, ui, c, 'all', CONFIG.sheets.dashboard);
}

/**
 * Generates the 'Inactive Listings' dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @param {GoogleAppsScript.Spreadsheet.Ui} ui The spreadsheet UI instance.
 * @param {string} c User-provided context string.
 * @returns {string} Success message.
 */
function generateInactiveListingsDashboard(ss, ui, c) {
  return _generateDashboardReport(ss, ui, c, 'inactive', CONFIG.sheets.inactiveListings);
}

/**
 * Generates the 'Active Listings' dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @param {GoogleAppsScript.Spreadsheet.Ui} ui The spreadsheet UI instance.
 * @param {string} c User-provided context string.
 * @returns {string} Success message.
 */
function generateActiveListingsDashboard(ss, ui, c) {
  return _generateDashboardReport(ss, ui, c, 'active', CONFIG.sheets.activeListings);
}

// --- CORE DASHBOARD & HELPER FUNCTIONS --- //

/**
 * Generic function to generate a listings-based dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @param {GoogleAppsScript.Spreadsheet.Ui} ui The spreadsheet UI instance.
 * @param {string} contextString User-provided context for the report.
 * @param {string} reportType The type of report ('all', 'inactive', 'active').
 * @param {string} targetSheetName The name of the sheet where the dashboard will be created.
 * @returns {string} Success message.
 */
function _generateDashboardReport(ss, ui, contextString = '', reportType = 'all', targetSheetName) {
  // Fetch all necessary data.
  const listingDataObject = _getValidatedData(ss, CONFIG.sheets.listing, [CONFIG.headers.listingSku, CONFIG.headers.asin, CONFIG.headers.itemName, CONFIG.headers.status]);
  const inventoryDataObject = _getValidatedData(ss, CONFIG.sheets.inventory, [
    CONFIG.headers.invSku, CONFIG.headers.available, CONFIG.headers.daysOfSupply, CONFIG.headers.unfulfillable,
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

  const allDashboardData = [];
  listingData.forEach(row => {
    const sku = row[skuListingIndex];
    if (!sku) return;

    const inventory = inventoryMap[sku] || {};
    const sales = salesMap[sku] || {};
    const asin = row[asinListingIndex];
    let imageFormula = "";
    let amazonLink = "";
    if (asin) {
      const keepaImageUrl = keepaMap[asin];
      if (keepaImageUrl) {
        imageFormula = `=IMAGE("${keepaImageUrl}", 4, ${CONFIG.image.width}, ${CONFIG.image.width / 1.2})`;
      } else {
        imageFormula = `=IMAGE("https://via.placeholder.com/${CONFIG.image.width}x${CONFIG.image.width / 1.2}.png?text=No+Image", 4, ${CONFIG.image.width}, ${CONFIG.image.width / 1.2})`;
      }
      amazonLink = `https://www.amazon.com/dp/${asin}`;
    }

    allDashboardData.push([
      imageFormula, row[nameListingIndex] || 'N/A', asin || 'N/A', sku, amazonLink,
      row[statusListingIndex] || 'N/A', sales.salesLast7Days || 0, inventory.available || 0,
      inventory.daysOfSupply || 0, inventory.reservedCustomerOrder || 0,
      inventory.reservedFcTransfer || 0, inventory.reservedFcProcessing || 0,
      inventory.unfulfillable || 0, sales.pastMonthSales || 0, sales.currentMonthSales || 0,
      (sales.currentMonthSales || 0) - (sales.pastMonthSales || 0)
    ]);
  });

  let finalDashboardData = [...allDashboardData];
  let reportTitle = "All Listings";

  if (reportType === 'inactive') {
    reportTitle = "Inactive Listings (FBA-SKU Filtered)";
    finalDashboardData = finalDashboardData.filter(row => {
      const status = row[5];
      const sku = String(row[3]);
      return status === CONFIG.listingStatus.inactive && sku.toUpperCase().startsWith('FBA-');
    });
  } else if (reportType === 'active') {
    reportTitle = "Active Listings (Sorted by Days of Supply)";
    finalDashboardData = finalDashboardData.filter(row => row[5] === CONFIG.listingStatus.active);
    finalDashboardData.sort((a, b) => (parseFloat(a[8]) || Infinity) - (parseFloat(b[8]) || Infinity));
  }

  let targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) {
    targetSheet.clear();
  } else {
    targetSheet = ss.insertSheet(targetSheetName);
  }

  const dashboardHeaders = ['Image', 'Name', 'Asin', 'SKU', 'Amazon Link', 'Listing status', "Sales Last 7 Days", 'Available', 'Days of Supply', 'Customer Orders', 'FC Transfers', 'FC Processing', 'Unfulfillable', 'Past 31-60 Day Sales', 'Current 30 Day Sales', 'Monthly Difference'];
  let startRow = 1;

  if (contextString) {
    targetSheet.getRange(startRow, 1).setValue(`Dashboard Context: ${contextString}`).setFontWeight('bold');
    targetSheet.getRange(startRow, 1, 1, dashboardHeaders.length).merge();
    startRow++;
  }

  targetSheet.getRange(startRow, 1).setValue(`Report Type: ${reportTitle}`).setFontWeight('bold');
  targetSheet.getRange(startRow, 1, 1, dashboardHeaders.length).merge();
  startRow++;

  targetSheet.getRange(startRow, 1, 1, dashboardHeaders.length).setValues([dashboardHeaders]).setFontWeight('bold');

  if (finalDashboardData.length > 0) {
    const imageFormulas = finalDashboardData.map(row => [row[0]]);
    const dataWithoutImageFormula = finalDashboardData.map(row => row.slice(1));
    targetSheet.getRange(startRow + 1, 2, dataWithoutImageFormula.length, dataWithoutImageFormula[0].length).setValues(dataWithoutImageFormula);
    targetSheet.getRange(startRow + 1, 1, imageFormulas.length, 1).setFormulas(imageFormulas);
  }

  targetSheet.setColumnWidth(1, CONFIG.image.width);
  targetSheet.autoResizeColumns(2, dashboardHeaders.length - 1);

  return `"${targetSheetName}" has been updated with ${finalDashboardData.length} products for "${reportType}".`;
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

      // Calculate sales for the 31-60 day period.
      const pastMonthSales = salesT60 - salesT30;

      map[sku] = {
        salesLast7Days: salesT7,
        currentMonthSales: salesT30,
        pastMonthSales: pastMonthSales < 0 ? 0 : pastMonthSales // Ensure non-negative
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
 * A centralized error handler. Logs the error and shows a user-friendly alert if in a UI context.
 * @param {Error} e The error object that was caught.
 * @param {GoogleAppsScript.Spreadsheet.Ui | null} ui The spreadsheet UI instance, which may not exist.
 * @param {string} functionName The name of the function where the error occurred.
 */
function _handleError(e, ui, functionName = 'unknown') {
  const errorMessage = `An error occurred in function '${functionName}': ${e.message}`;
  const errorStack = e.stack ? `\nStack Trace: ${e.stack}` : '';

  // Always log the detailed error for developer debugging.
  Logger.log(`${errorMessage}${errorStack}`);

  // Only show a popup alert if the script is running in an interactive UI context.
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
  if (data.length < 2) { // Needs at least a header row and one data row
      Logger.log(`Warning: Sheet "${sheetName}" is empty or contains only headers.`);
      // Return empty data structure but with valid headers to prevent downstream errors
      return { headers: (data[0] || []).map(h => String(h).trim()), data: [] };
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

        input[type="text"], select {
            width: 100%; /* Changed to 100% for better responsiveness */
            padding: 8px;
            box-sizing: border-box;
            border: 1px solid #ccc;
            border-radius: 4px;
            font-size: 0.95em;
            background-color: #fff; /* Ensure select has a background */
        }
        
        input[type="text"]:focus, select:focus {
            border-color: #DAA520; /* Gold focus highlight */
            outline: none;
            box-shadow: 0 0 0 2px rgba(218, 165, 32, 0.2);
        }

        /* --- REFINED BUTTON STYLES --- */
        #generateBtn {
            width: 100%;
            padding: 12px 15px; /* Increased padding for a better feel */
            border: 1px solid #DAA520;
            border-radius: 5px;
            font-size: 1.1em; /* Made text slightly larger */
            font-weight: bold;
            cursor: pointer;
            transition: all 0.2s ease-in-out;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            
            /* FIX: Centering the text */
            display: flex;
            justify-content: center;
            align-items: center;

            /* Primary Action Button - Gradient Gold */
            background: linear-gradient(to right, #FFD700, #FFA500);
            color: #000000; /* Black text for high contrast */
        }

        #generateBtn:hover:not(:disabled) {
            background: linear-gradient(to right, #e6c200, #e69500);
            box-shadow: 0 3px 6px rgba(0,0,0,0.2);
            transform: translateY(-1px); /* Add subtle lift */
        }
        
        /* Disabled State for the button */
        #generateBtn:disabled {
            background: #e0e0e0; /* Muted Grey */
            color: #9e9e9e; /* Lighter grey text */
            border-color: #e0e0e0;
            cursor: not-allowed;
            box-shadow: none;
        }

        /* Spinner and Alert styles */
        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #DAA520; /* Gold to match brand */
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

        <!-- Grouping all form elements together -->
        <div class="form-group">
            <label for="contextInput">Dashboard Context (Optional)</label>
            <input type="text" id="contextInput" placeholder="e.g., Q2 Inventory Review">
        </div>

        <!-- NEW: Dropdown for selecting report type -->
        <div class="form-group">
            <label for="reportTypeSelect">Analysis Type</label>
            <select id="reportTypeSelect">
                <!-- 'analysis' is the value for the original primary button -->
                <option value="analysis">General Analysis Tab</option>
                <option value="all">All Listings Dashboard</option>
                <option value="inactive">Inactive Listings (FBA-SKU)</option>
                <option value="active">Active Listings (Sorted by DoS)</option>
            </select>
        </div>

        <!-- NEW: Single, consolidated button -->
        <button id="generateBtn" onclick="runReport()">Generate</button>

        <div id="loadingSpinner" class="spinner" style="display:none;"></div>
    </div>

    <script>
        document.addEventListener('DOMContentLoaded', function() {
            // No changes needed here
        });

        // MODIFIED: The function no longer needs a parameter
        function runReport() {
            // GETTING VALUES: Now we get the report type from the dropdown
            const contextString = document.getElementById('contextInput').value;
            const reportType = document.getElementById('reportTypeSelect').value; // Get value from the select menu
            
            const generateButton = document.getElementById('generateBtn');
            const spinner = document.getElementById('loadingSpinner');
            const alertBox = document.getElementById('statusAlert');

            generateButton.disabled = true; // Only one button to disable now
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
                // The backend function is called with the selected report type
                .createConsolidatedDashboardFromSidebar(contextString, reportType);
        }
    </script>
</body>
</html>
```
