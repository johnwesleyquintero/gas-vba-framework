# Amazon Consolidated Dashboard Script for Google Sheets

This documentation provides a comprehensive guide to the Google Apps Script (GAS) solution designed to consolidate Amazon FBA data into dynamic, user-friendly dashboards within Google Sheets. This script automates the process of combining information from your Amazon Listing, Inventory, and Sales reports, along with external Keepa data, to provide critical insights into your product catalog.

The solution supports generating distinct dashboards for 'All Listings', 'Active Listings', and 'Inactive Listings', with intelligent filtering and sorting capabilities.

## 1. Features

*   **Data Consolidation:** Seamlessly combines data from `All_Listing_Report`, `FBA_Inventory_Snapshot`, `FBA_Shipments_60_Days`, and `Keepa` sheets.
*   **Multiple Dashboard Views:** Generates three distinct dashboards:
    *   **All Listings:** A comprehensive view of all products.
    *   **Active Listings:** Filters for active listings and sorts them by 'Days of Supply' (lowest first).
    *   **Inactive Listings:** Filters for inactive listings and specifically targets FBA-SKU items.
*   **Product Visuals:** Automatically pulls product images via Keepa data, displaying them directly in the dashboard using `IMAGE()` formulas.
*   **Direct Amazon Links:** Generates clickable links to the Amazon product page for each ASIN.
*   **Sales Performance Metrics:** Calculates and displays 'Yesterday\'s Sales', 'Current 30 Day Sales', 'Past 31-60 Day Sales', and 'Monthly Difference'.
*   **Inventory Insights:** Includes 'Available' quantity, 'Days of Supply', 'Unfulfillable', and detailed 'Reserved' inventory breakdowns.
*   **User-Friendly Interface:** Provides a custom sidebar within Google Sheets for easy generation of different report types.
*   **Configurable Headers:** Centralized configuration allows easy adjustment of column headers, providing resilience against potential changes in Amazon report formats.

## 2. Setup Instructions

To implement this solution, you will need a Google Sheet and access to Google Apps Script.

### Step-by-Step Installation:

1.  **Create a New Google Sheet:** Open Google Sheets and create a new blank spreadsheet.
2.  **Open Apps Script Editor:** Go to `Extensions > Apps Script`. This will open a new browser tab with the Apps Script editor.
3.  **Create `Code.gs`:** In the Apps Script editor, you'll typically see a `Code.gs` file by default. Replace its content with the JavaScript code provided in Section 7.1.
4.  **Create `Sidebar.html`:**
    *   In the Apps Script editor, click `+` next to "Files" in the left sidebar, then select `HTML`.
    *   Name the new file `Sidebar`.
    *   Paste the HTML code provided in Section 7.2 into this new `Sidebar.html` file.
5.  **Save the Project:** Click the save icon (floppy disk) or press `Ctrl + S` (`Cmd + S` on Mac) to save both files.
6.  **Rename the Default Sheet:** In your Google Sheet, ensure you have a sheet named `Dashboard`. If your default sheet is `Sheet1`, rename it to `Dashboard`. This is the default target for the 'All Listings' report.
7.  **Create Source Data Sheets:** Create the following sheets in your Google Spreadsheet. **Their names must precisely match the `CONFIG.sheets` values in the `Code.gs` script.**
    *   `All_Listing_Report`
    *   `FBA_Inventory_Snapshot`
    *   `FBA_Shipments_60_Days`
    *   `Keepa`
8.  **Populate Source Data:** Import or paste your Amazon reports into their respective sheets. Ensure the column headers in your reports match the `CONFIG.headers` defined in the `Code.gs` script. **This is critical for the script to correctly locate data.**

## 3. Usage

Once set up, the script integrates directly into your Google Sheet interface.

1.  **Open the Spreadsheet:** When you open your Google Sheet, a custom menu item "Dashboard Tools" will appear in the top menu bar, and the "Dashboard Controls" sidebar will automatically open on the right.
2.  **Enter Optional Context:** In the sidebar, you can enter an optional "Dashboard Context" string (e.g., "Q2 Inventory Review"). This text will be added to the top of the generated dashboard.
3.  **Select Report Type:** Choose one of the three buttons to generate your desired dashboard:
    *   **Generate All Listings Dashboard:** Creates a consolidated report of all listings on the `Dashboard` sheet.
    *   **Generate Inactive Listings (FBA-SKU) Dashboard:** Creates a filtered report of inactive FBA listings on the `Inactive_Listing` sheet.
    *   **Generate Active Listings (Sorted by DoS) Dashboard:** Creates a filtered report of active listings, sorted by 'Days of Supply', on the `Active_Listing` sheet.
4.  **Monitor Progress:** A loading spinner will appear while the script runs. Status messages (success or error) will be displayed at the top of the sidebar.
5.  **Review Dashboards:** The script will create or clear and update the relevant target sheet (`Dashboard`, `Active_Listing`, `Inactive_Listing`) with the generated data.

## 4. Configuration

The `CONFIG` object at the top of `Code.gs` is the central place for all user-adjustable settings.

```javascript
const CONFIG = {
  sheets: {
    listing: 'All_Listing_Report',
    inventory: 'FBA_Inventory_Snapshot',
    sales: 'FBA_Shipments_60_Days',
    keepa: 'Keepa',
    dashboard: 'Dashboard',
    activeListings: 'Active_Listing',
    inactiveListings: 'Inactive_Listing'
  },
  headers: {
    listingSku: 'seller-sku',
    salesSku: 'sku',
    invSku: 'sku',
    asin: 'asin1',
    itemName: 'item-name',
    status: 'status',
    available: 'available',
    daysOfSupply: 'days-of-supply',
    unfulfillable: 'unfulfillable-quantity',
    reservedCustomer: 'Reserved Customer Order',
    reservedTransfer: 'Reserved FC Transfer',
    reservedProcessing: 'Reserved FC Processing',
    purchaseDate: 'purchase-date',
    quantity: 'quantity',
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

*   **`CONFIG.sheets`**: Define the exact names of your source and target sheets. **Crucially, ensure these match your Google Sheet tab names.**
*   **`CONFIG.headers`**: This is the most important section to maintain. If Amazon changes the column names in their reports, you *only* need to update the corresponding values here. For instance, if `seller-sku` becomes `merchant-sku` in your listing report, you would update `listingSku: 'seller-sku'` to `listingSku: 'merchant-sku'`.
*   **`CONFIG.image.width`**: Adjust the width of the image column in pixels on the dashboard.
*   **`CONFIG.listingStatus`**: Confirm the exact strings used in your Amazon Listing Report's 'status' column for 'Active' and 'Inactive' listings.

## 5. Code Structure and Explanation

The script is organized into core functions and helper functions to ensure modularity and readability.

### Core Functions (`Code.gs`)

*   **`onOpen()`**: This special Google Apps Script function runs automatically when the spreadsheet is opened. It creates the "Dashboard Tools" custom menu and immediately opens the sidebar for user convenience.
*   **`showSidebar()`**: Displays the `Sidebar.html` content as a custom sidebar within the Google Sheet interface.
*   **`createConsolidatedDashboardFromSidebar(contextString, reportType)`**: This function acts as the primary entry point for interactions from the sidebar. It dispatches the request to the appropriate dashboard generation function based on the `reportType` parameter (`'all'`, `'inactive'`, `'active'`).
*   **`generateAllListingsDashboard(contextString)`**, **`generateInactiveListingsDashboard(contextString)`**, **`generateActiveListingsDashboard(contextString)`**: These are wrapper functions that call the main `_generateDashboardReport` function with the specific `reportType` and target sheet name.
*   **`_generateDashboardReport(contextString, reportType, targetSheetName)`**: This is the core logic function.
    *   **Data Retrieval:** Fetches all data from the configured source sheets (`listing`, `inventory`, `sales`, `keepa`).
    *   **Header Validation:** Performs initial checks to ensure critical headers are present in the source sheets, alerting the user if any are missing.
    *   **Data Processing:** Utilizes helper functions (`processInventory`, `processSales`, `processKeepa`) to transform raw sheet data into easily queryable JavaScript objects (maps).
    *   **Dashboard Data Construction:** Iterates through the listing data, enriching each row with corresponding inventory, sales, and Keepa image information. It dynamically generates `IMAGE()` and direct Amazon product `HYPERLINK()` formulas.
    *   **Filtering and Sorting:** Applies the appropriate filtering (e.g., inactive FBA-SKUs) and sorting (e.g., active listings by Days of Supply) based on the `reportType`.
    *   **Sheet Output:** Clears or creates the designated `targetSheetName`, writes the report title and headers, then populates the sheet with the processed dashboard data. It applies the image formulas specifically to the image column.
    *   **Formatting:** Sets the image column width and auto-resizes other columns for optimal display.
    *   **User Feedback:** Provides a success or error alert to the user.

### Helper Functions (`Code.gs`)

*   **`processInventory(data, ui)`**: Parses the `FBA_Inventory_Snapshot` data. It extracts 'available', 'days of supply', 'unfulfillable', and various 'reserved' quantities, mapping them by SKU for quick lookup. Includes robust header validation.
*   **`processSales(data, ui)`**: Processes `FBA_Shipments_60_Days` data. It calculates 'yesterday\'s sales', 'current 30-day sales', and 'past 31-60 day sales' per SKU. It includes robust date parsing to handle various date formats from the report and performs header validation.
*   **`processKeepa(data, ui)`**: Extracts ASINs and their corresponding image URLs from the `Keepa` sheet, mapping ASINs to image URLs. Includes header validation.

### Sidebar HTML (`Sidebar.html`)

This file defines the user interface that appears as a sidebar in Google Sheets. It includes:

*   **Styling:** Embedded CSS for a clean, responsive, and modern look.
*   **Input Field:** An optional text input for adding context to the generated dashboards.
*   **Action Buttons:** Three distinct buttons to trigger the generation of 'All', 'Inactive', or 'Active' listings dashboards.
*   **Loading Spinner:** A visual indicator shown while the script is processing.
*   **Status Alert:** A dynamic message area to display success or error messages to the user.
*   **Client-side JavaScript:** Handles user interactions (button clicks), retrieves input values, manages button states, displays the loading spinner, and communicates with the Google Apps Script functions on the server side using `google.script.run`.

## 6. Troubleshooting and Common Issues

*   **"Missing columns in..." Error:** This is the most common error. Double-check that:
    *   Your source sheets (`All_Listing_Report`, `FBA_Inventory_Snapshot`, `FBA_Shipments_60_Days`, `Keepa`) exist and are spelled exactly as defined in `CONFIG.sheets`.
    *   The column headers in your uploaded reports (e.g., `seller-sku`, `asin1`, `available`) exactly match the values defined in `CONFIG.headers`. Amazon sometimes changes these slightly, so verify against your latest downloaded reports.
*   **"Generation Error: Invalid report type specified..."**: This indicates an internal issue with how the sidebar is calling the script or a mismatch in expected parameters. Ensure you haven't modified the `reportType` strings in `Sidebar.html` or `Code.gs`.
*   **Script Authorization:** The first time you run the script, Google will ask for authorization to access your spreadsheet. Grant these permissions.
*   **Slow Performance:** For very large datasets (tens of thousands of rows), the script might take some time to process. Google Apps Script has execution limits. If you encounter timeouts, consider processing smaller batches or optimizing your reports.
*   **Images Not Loading:**
    *   Ensure the `Keepa` sheet exists and contains `ASIN` and `Image` columns with valid public image URLs.
    *   Verify the `ASIN` values in your `All_Listing_Report` match those in your `Keepa` sheet.
*   **Incorrect Sales Data:**
    *   Confirm the `FBA_Shipments_60_Days` sheet has correct `purchase-date` and `quantity` columns.
    *   The script uses robust date parsing, but inconsistent date formats from Amazon reports can occasionally be an issue. Ensure your dates are in a standard format recognized by JavaScript's `Date` object.

## 7. Source Code

Below are the complete source code files for the Google Apps Script project.

### 7.1. `Code.gs`

```javascript
/**
 * @OnlyCurrentDoc
 * This script creates consolidated Amazon dashboards from three source reports,
 * now supporting distinct sheets for All, Active, and Inactive listings.
 * VERSION 10: Supports separate output tabs for different report types.
 */

// --- CONFIGURATION --- //
// Central configuration for all sheet names and column headers.
// If Amazon changes a report column name, you only need to update it here.
const CONFIG = {
  sheets: {
    listing: 'All_Listing_Report',
    inventory: 'FBA_Inventory_Snapshot',
    sales: 'FBA_Shipments_60_Days', // This is your All Orders Report
    keepa: 'Keepa',                  // New: Keepa sheet
    dashboard: 'Dashboard',          // Main dashboard for 'All' listings
    activeListings: 'Active_Listing',  // New: Sheet for Active listings
    inactiveListings: 'Inactive_Listing' // New: Sheet for Inactive listings
  },
  headers: {
    listingSku: 'seller-sku', // SKU header used in All_Listing_Report
    salesSku: 'sku',         // SKU header used in FBA_Shipments_60_Days (as per your feedback)
    invSku: 'sku',           // Specific SKU header for FBA_Inventory_Snapshot (already 'sku')
    asin: 'asin1',
    itemName: 'item-name',
    status: 'status',
    available: 'available',
    daysOfSupply: 'days-of-supply',
    unfulfillable: 'unfulfillable-quantity',
    reservedCustomer: 'Reserved Customer Order',
    reservedTransfer: 'Reserved FC Transfer',
    reservedProcessing: 'Reserved FC Processing',
    purchaseDate: 'purchase-date',
    quantity: 'quantity',
    // Keepa specific headers
    keepaAsin: 'ASIN',
    keepaImage: 'Image'
  },
  image: {
    width: 100 // Width of the image column in pixels
  },
  // Define exact strings for listing statuses from your reports
  listingStatus: {
    active: 'Active',   // Confirm this exact string from your 'status' column
    inactive: 'Inactive' // Confirm this exact string from your 'status' column
  }
};
// --- END CONFIGURATION ---

/**
 * Creates a custom menu and opens the sidebar when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Dashboard Tools')
    .addItem('Open Dashboard Controls', 'showSidebar') // Emphasize sidebar as primary control
    .addToUi();
  showSidebar(); // Automatically open the sidebar for the user
}

/**
 * Shows the custom HTML sidebar.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Dashboard Controls')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Wrapper function called by the sidebar to dispatch to the correct report generator.
 * This acts as the entry point for sidebar interactions.
 * @param {string} contextString - The optional context string provided by the user.
 * @param {string} reportType - The type of report to generate ('all', 'inactive', 'active').
 */
function createConsolidatedDashboardFromSidebar(contextString, reportType) {
  const ui = SpreadsheetApp.getUi();
  try {
    if (reportType === 'all') {
      generateAllListingsDashboard(contextString);
    } else if (reportType === 'inactive') {
      generateInactiveListingsDashboard(contextString);
    } else if (reportType === 'active') {
      generateActiveListingsDashboard(contextString);
    } else {
      throw new Error('Invalid report type specified. Please select "all", "inactive", or "active".');
    }
  } catch (e) {
    ui.alert('Generation Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Generates the 'All Listings' dashboard on the 'Dashboard' tab.
 * @param {string} contextString - An optional string to add as context.
 */
function generateAllListingsDashboard(contextString = '') {
  _generateDashboardReport(contextString, 'all', CONFIG.sheets.dashboard);
}

/**
 * Generates the 'Inactive Listings' dashboard on the 'Inactive_Listing' tab.
 * @param {string} contextString - An optional string to add as context.
 */
function generateInactiveListingsDashboard(contextString = '') {
  _generateDashboardReport(contextString, 'inactive', CONFIG.sheets.inactiveListings);
}

/**
 * Generates the 'Active Listings' dashboard on the 'Active_Listing' tab.
 * @param {string} contextString - An optional string to add as context.
 */
function generateActiveListingsDashboard(contextString = '') {
  _generateDashboardReport(contextString, 'active', CONFIG.sheets.activeListings);
}


/**
 * Core function to create a consolidated dashboard based on specified report type and target sheet.
 * This version correctly generates image formulas using Keepa data and adds a direct link to the Amazon page,
 * and now supports filtering and sorting based on report type, writing to a specified sheet.
 * @param {string} [contextString=''] - An optional string to add as context to the dashboard.
 * @param {string} [reportType='all'] - The type of report to generate ('all', 'inactive', 'active').
 * @param {string} targetSheetName - The name of the sheet where the report will be written.
 */
function _generateDashboardReport(contextString = '', reportType = 'all', targetSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // --- 1. Get Source Sheets ---
  const listingSheet = ss.getSheetByName(CONFIG.sheets.listing);
  const inventorySheet = ss.getSheetByName(CONFIG.sheets.inventory);
  const salesSheet = ss.getSheetByName(CONFIG.sheets.sales);
  const keepaSheet = ss.getSheetByName(CONFIG.sheets.keepa);

  if (!listingSheet || !inventorySheet || !salesSheet || !keepaSheet) {
    ui.alert('Error', `Please make sure all source sheets exist: ${CONFIG.sheets.listing}, ${CONFIG.sheets.inventory}, ${CONFIG.sheets.sales}, ${CONFIG.sheets.keepa}`, ui.ButtonSet.OK);
    return;
  }

  // --- 2. Get and Validate Data ---
  const listingData = listingSheet.getDataRange().getValues();
  const inventoryData = inventorySheet.getDataRange().getValues();
  const salesData = salesSheet.getDataRange().getValues();
  const keepaData = keepaSheet.getDataRange().getValues();

  // Helper function to find missing headers
  const findMissingHeaders = (actualHeaders, requiredHeaders) => {
    const headerSet = new Set(actualHeaders);
    return requiredHeaders.filter(h => !headerSet.has(h));
  };

  // Validate Listing Sheet Headers
  const listingHeaders = listingData[0];
  let missingListingHeaders = findMissingHeaders(listingHeaders, [CONFIG.headers.listingSku, CONFIG.headers.asin, CONFIG.headers.itemName, CONFIG.headers.status]);
  if (missingListingHeaders.length > 0) {
    ui.alert('Error', `Missing columns in "${CONFIG.sheets.listing}": ${missingListingHeaders.join(', ')}`, ui.ButtonSet.OK);
    return;
  }

  // --- 3. Process Data into Maps ---
  const inventoryMap = processInventory(inventoryData, ui);
  if (!inventoryMap) return;
  const salesMap = processSales(salesData, ui);
  if (!salesMap) return;
  const keepaMap = processKeepa(keepaData, ui);
  if (!keepaMap) return;

  // Get column indexes from Listing Report
  const skuListingIndex = listingHeaders.indexOf(CONFIG.headers.listingSku);
  const asinListingIndex = listingHeaders.indexOf(CONFIG.headers.asin);
  const nameListingIndex = listingHeaders.indexOf(CONFIG.headers.itemName);
  const statusListingIndex = listingHeaders.indexOf(CONFIG.headers.status);

  // --- 4. Build ALL Dashboard Data (before filtering/sorting) ---
  const allDashboardData = [];
  listingData.slice(1).forEach(row => { // Use slice(1) to skip header row
    const sku = row[skuListingIndex];
    if (!sku) return; // Skip rows with no SKU

    const inventory = inventoryMap[sku] || {};
    const sales = salesMap[sku] || {};

    // --- Image Formula & Amazon Link Generation (Using Keepa Image) ---
    const asin = row[asinListingIndex];
    let imageFormula = ""; // Default to blank if no ASIN or Keepa image
    let amazonLink = "";   // Default link to blank

    if (asin) {
      const keepaImageUrl = keepaMap[asin];
      // Use the Keepa image URL if available, otherwise a generic placeholder
      if (keepaImageUrl) {
        imageFormula = `=IMAGE("${keepaImageUrl}")`;
      } else {
        // Fallback placeholder image if Keepa data doesn't have it
        imageFormula = `=IMAGE("https://via.placeholder.com/${CONFIG.image.width}x${CONFIG.image.width/1.2}.png?text=No+Image")`;
      }

      // Create the direct product link using the ASIN
      amazonLink = `https://www.amazon.com/dp/${asin}`;
    }

    const pastMonthSales = sales.pastMonthSales || 0;
    const currentMonthSales = sales.currentMonthSales || 0;
    const monthlyDifference = currentMonthSales - pastMonthSales;

    allDashboardData.push([
      imageFormula, // Index 0
      row[nameListingIndex] || 'N/A', // Index 1
      asin || 'N/A', // Index 2
      sku, // Index 3
      amazonLink, // Index 4
      row[statusListingIndex] || 'N/A', // Index 5: Listing Status
      sales.yesterdaySales || 0, // Index 6
      inventory.available || 0, // Index 7
      inventory.daysOfSupply || 0, // Index 8: Days of Supply
      inventory.reservedCustomerOrder || 0, // Index 9
      inventory.reservedFcTransfer || 0, // Index 10
      inventory.reservedFcProcessing || 0, // Index 11
      inventory.unfulfillable || 0, // Index 12
      pastMonthSales, // Index 13
      currentMonthSales, // Index 14
      monthlyDifference // Index 15
    ]);
  });

  // --- 5. Apply Filtering and Sorting based on reportType ---
  let finalDashboardData = [...allDashboardData]; // Start with all data
  let reportTitle = "All Listings"; // Default title

  if (reportType === 'inactive') {
    reportTitle = "Inactive Listings (FBA-SKU Filtered)";
    finalDashboardData = finalDashboardData.filter(row => {
      const status = row[5]; // Index of 'Listing status'
      const sku = String(row[3]);    // Index of 'SKU', ensure it's a string
      return status === CONFIG.listingStatus.inactive && sku.toUpperCase().startsWith('FBA-');
    });
  } else if (reportType === 'active') {
    reportTitle = "Active Listings (Sorted by Days of Supply)";
    finalDashboardData = finalDashboardData.filter(row => {
      const status = row[5]; // Index of 'Listing status'
      return status === CONFIG.listingStatus.active;
    });
    // Sort active listings by 'Days of Supply' from lowest to highest
    finalDashboardData.sort((a, b) => {
      // Ensure numeric comparison; treat empty/invalid as Infinity for sorting to end
      const daysA = parseFloat(a[8]) || Infinity;
      const daysB = parseFloat(b[8]) || Infinity;
      return daysA - daysB;
    });
  }

  // --- 6. Write Data to Target Sheet ---
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) {
    targetSheet.clear();
  } else {
    targetSheet = ss.insertSheet(targetSheetName);
  }

  const dashboardHeaders = [
    'Image', 'Name', 'Asin', 'SKU', 'Amazon Link', 'Listing status', "Yesterday's Sales", 'Available',
    'Days of Supply', 'Customer Orders', 'FC Transfers', 'FC Processing', 'Unfulfillable',
    'Past 31-60 Day Sales', 'Current 30 Day Sales', 'Monthly Difference'
  ];

  let startRow = 1;

  // Add Context row if provided
  if (contextString) {
    targetSheet.getRange(startRow, 1).setValue(`Dashboard Context: ${contextString}`).setFontWeight('bold');
    targetSheet.getRange(startRow, 1, 1, dashboardHeaders.length).merge(); // Merge across header width
    startRow++; // Shift header and data rows down
  }

  // Add Report Type Title
  targetSheet.getRange(startRow, 1).setValue(`Report Type: ${reportTitle}`).setFontWeight('bold');
  targetSheet.getRange(startRow, 1, 1, dashboardHeaders.length).merge();
  startRow++; // Shift header and data rows down

  targetSheet.getRange(startRow, 1, 1, dashboardHeaders.length).setValues([dashboardHeaders]).setFontWeight('bold');

  if (finalDashboardData.length > 0) {
    const dataRange = targetSheet.getRange(startRow + 1, 1, finalDashboardData.length, finalDashboardData[0].length);
    dataRange.setValues(finalDashboardData); // Write all data as values first

    // Then, apply formulas to the image column
    const imageFormulas = finalDashboardData.map(row => [row[0]]);
    targetSheet.getRange(startRow + 1, 1, imageFormulas.length, 1).setFormulas(imageFormulas);
  }

  targetSheet.setColumnWidth(1, CONFIG.image.width);
  // Auto-resize from column 2 onwards, as column 1 (Image) has a fixed width.
  targetSheet.autoResizeColumns(2, dashboardHeaders.length - 1);
  ui.alert('Success!', `"${targetSheetName}" has been updated with ${finalDashboardData.length} products for "${reportTitle}".`, ui.ButtonSet.OK);
}

// --- Helper Functions ---

/** Processes FBA Inventory data. Now includes header validation. */
function processInventory(data, ui) {
  const map = {};
  const headers = data.shift();
  const required = [CONFIG.headers.invSku, CONFIG.headers.available, CONFIG.headers.daysOfSupply, CONFIG.headers.unfulfillable, CONFIG.headers.reservedCustomer, CONFIG.headers.reservedTransfer, CONFIG.headers.reservedProcessing];
  const missing = required.filter(h => headers.indexOf(h) === -1);
  if (missing.length > 0) {
    ui.alert('Error', `Missing columns in "${CONFIG.sheets.inventory}": ${missing.join(', ')}`, ui.ButtonSet.OK);
    return null;
  }

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
        available: row[availableIndex],
        daysOfSupply: row[daysOfSupplyIndex],
        unfulfillable: row[unfulfillableIndex],
        reservedCustomerOrder: row[reservedCustomerIndex],
        reservedFcTransfer: row[reservedTransferIndex],
        reservedFcProcessing: row[reservedProcessingIndex]
      };
    }
  });
  return map;
}

/** Processes Sales data. Now includes header validation and robust date logic. */
function processSales(data, ui) {
  const map = {};
  const headers = data.shift();
  const required = [CONFIG.headers.salesSku, CONFIG.headers.purchaseDate, CONFIG.headers.quantity];
  const missing = required.filter(h => headers.indexOf(h) === -1);
  if (missing.length > 0) {
    ui.alert('Error', `Missing columns in "${CONFIG.sheets.sales}": ${missing.join(', ')}`, ui.ButtonSet.OK);
    return null;
  }

  const skuIndex = headers.indexOf(CONFIG.headers.salesSku);
  const purchaseDateIndex = headers.indexOf(CONFIG.headers.purchaseDate);
  const quantityIndex = headers.indexOf(CONFIG.headers.quantity);

  const todayStart = new Date();
  todayStart.setHours(0, 0, 0, 0);

  const yesterdayStart = new Date(todayStart);
  yesterdayStart.setDate(todayStart.getDate() - 1);

  const thirtyDaysAgo = new Date(todayStart);
  thirtyDaysAgo.setDate(todayStart.getDate() - 30);

  const sixtyDaysAgo = new Date(todayStart);
  sixtyDaysAgo.setDate(todayStart.getDate() - 60);

  data.forEach(row => {
    const sku = row[skuIndex];
    // Robust date parsing: Handles various date formats
    const purchaseDate = new Date(row[purchaseDateIndex]);
    const quantity = parseInt(row[quantityIndex], 10);

    if (sku && !isNaN(quantity) && purchaseDate.getTime()) {
      if (!map[sku]) {
        map[sku] = { yesterdaySales: 0, currentMonthSales: 0, pastMonthSales: 0 };
      }

      // ROBUST DATE CHECK: Check if the sale happened between the start of yesterday and start of today
      if (purchaseDate >= yesterdayStart && purchaseDate < todayStart) {
        map[sku].yesterdaySales += quantity;
      }

      if (purchaseDate >= thirtyDaysAgo && purchaseDate < todayStart) {
        map[sku].currentMonthSales += quantity;
      }

      if (purchaseDate >= sixtyDaysAgo && purchaseDate < thirtyDaysAgo) {
        map[sku].pastMonthSales += quantity;
      }
    }
  });
  return map;
}

/** Processes Keepa data to map ASINs to Image URLs. */
function processKeepa(data, ui) {
  const map = {};
  const headers = data.shift();
  const required = [CONFIG.headers.keepaAsin, CONFIG.headers.keepaImage];
  const missing = required.filter(h => headers.indexOf(h) === -1);
  if (missing.length > 0) {
    ui.alert('Error', `Missing columns in "${CONFIG.sheets.keepa}": ${missing.join(', ')}`, ui.ButtonSet.OK);
    return null;
  }

  const asinIndex = headers.indexOf(CONFIG.headers.keepaAsin);
  const imageIndex = headers.indexOf(CONFIG.headers.keepaImage);

  data.forEach(row => {
    const asin = row[asinIndex];
    const imageUrl = row[imageIndex];
    if (asin && imageUrl) {
      map[asin] = imageUrl;
    }
  });
  return map;
}
```

### 7.2. `Sidebar.html`

```html
/**
 * @OnlyCurrentDoc
 * This script creates consolidated Amazon dashboards from three source reports,
 * now supporting distinct sheets for All, Active, and Inactive listings.
 * VERSION 10: Supports separate output tabs for different report types.
 */

// --- CONFIGURATION --- //
// Central configuration for all sheet names and column headers.
// If Amazon changes a report column name, you only need to update it here.
const CONFIG = {
  sheets: {
    listing: 'All_Listing_Report',
    inventory: 'FBA_Inventory_Snapshot',
    sales: 'FBA_Shipments_60_Days', // This is your All Orders Report
    keepa: 'Keepa',                  // New: Keepa sheet
    dashboard: 'Dashboard',          // Main dashboard for 'All' listings
    activeListings: 'Active_Listing',  // New: Sheet for Active listings
    inactiveListings: 'Inactive_Listing' // New: Sheet for Inactive listings
  },
  headers: {
    listingSku: 'seller-sku', // SKU header used in All_Listing_Report
    salesSku: 'sku',         // SKU header used in FBA_Shipments_60_Days (as per your feedback)
    invSku: 'sku',           // Specific SKU header for FBA_Inventory_Snapshot (already 'sku')
    asin: 'asin1',
    itemName: 'item-name',
    status: 'status',
    available: 'available',
    daysOfSupply: 'days-of-supply',
    unfulfillable: 'unfulfillable-quantity',
    reservedCustomer: 'Reserved Customer Order',
    reservedTransfer: 'Reserved FC Transfer',
    reservedProcessing: 'Reserved FC Processing',
    purchaseDate: 'purchase-date',
    quantity: 'quantity',
    // Keepa specific headers
    keepaAsin: 'ASIN',
    keepaImage: 'Image'
  },
  image: {
    width: 100 // Width of the image column in pixels
  },
  // Define exact strings for listing statuses from your reports
  listingStatus: {
    active: 'Active',   // Confirm this exact string from your 'status' column
    inactive: 'Inactive' // Confirm this exact string from your 'status' column
  }
};
// --- END CONFIGURATION ---

/**
 * Creates a custom menu and opens the sidebar when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Dashboard Tools')
    .addItem('Open Dashboard Controls', 'showSidebar') // Emphasize sidebar as primary control
    .addToUi();
  showSidebar(); // Automatically open the sidebar for the user
}

/**
 * Shows the custom HTML sidebar.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Dashboard Controls')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Wrapper function called by the sidebar to dispatch to the correct report generator.
 * This acts as the entry point for sidebar interactions.
 * @param {string} contextString - The optional context string provided by the user.
 * @param {string} reportType - The type of report to generate ('all', 'inactive', 'active').
 */
function createConsolidatedDashboardFromSidebar(contextString, reportType) {
  const ui = SpreadsheetApp.getUi();
  try {
    if (reportType === 'all') {
      generateAllListingsDashboard(contextString);
    } else if (reportType === 'inactive') {
      generateInactiveListingsDashboard(contextString);
    } else if (reportType === 'active') {
      generateActiveListingsDashboard(contextString);
    } else {
      throw new Error('Invalid report type specified. Please select "all", "inactive", or "active".');
    }
  } catch (e) {
    ui.alert('Generation Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Generates the 'All Listings' dashboard on the 'Dashboard' tab.
 * @param {string} contextString - An optional string to add as context.
 */
function generateAllListingsDashboard(contextString = '') {
  _generateDashboardReport(contextString, 'all', CONFIG.sheets.dashboard);
}

/**
 * Generates the 'Inactive Listings' dashboard on the 'Inactive_Listing' tab.
 * @param {string} contextString - An optional string to add as context.
 */
function generateInactiveListingsDashboard(contextString = '') {
  _generateDashboardReport(contextString, 'inactive', CONFIG.sheets.inactiveListings);
}

/**
 * Generates the 'Active Listings' dashboard on the 'Active_Listing' tab.
 * @param {string} contextString - An optional string to add as context.
 */
function generateActiveListingsDashboard(contextString = '') {
  _generateDashboardReport(contextString, 'active', CONFIG.sheets.activeListings);
}


/**
 * Core function to create a consolidated dashboard based on specified report type and target sheet.
 * This version correctly generates image formulas using Keepa data and adds a direct link to the Amazon page,
 * and now supports filtering and sorting based on report type, writing to a specified sheet.
 * @param {string} [contextString=''] - An optional string to add as context to the dashboard.
 * @param {string} [reportType='all'] - The type of report to generate ('all', 'inactive', 'active').
 * @param {string} targetSheetName - The name of the sheet where the report will be written.
 */
function _generateDashboardReport(contextString = '', reportType = 'all', targetSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // --- 1. Get Source Sheets ---
  const listingSheet = ss.getSheetByName(CONFIG.sheets.listing);
  const inventorySheet = ss.getSheetByName(CONFIG.sheets.inventory);
  const salesSheet = ss.getSheetByName(CONFIG.sheets.sales);
  const keepaSheet = ss.getSheetByName(CONFIG.sheets.keepa);

  if (!listingSheet || !inventorySheet || !salesSheet || !keepaSheet) {
    ui.alert('Error', `Please make sure all source sheets exist: ${CONFIG.sheets.listing}, ${CONFIG.sheets.inventory}, ${CONFIG.sheets.sales}, ${CONFIG.sheets.keepa}`, ui.ButtonSet.OK);
    return;
  }

  // --- 2. Get and Validate Data ---
  const listingData = listingSheet.getDataRange().getValues();
  const inventoryData = inventorySheet.getDataRange().getValues();
  const salesData = salesSheet.getDataRange().getValues();
  const keepaData = keepaSheet.getDataRange().getValues();

  // Helper function to find missing headers
  const findMissingHeaders = (actualHeaders, requiredHeaders) => {
    const headerSet = new Set(actualHeaders);
    return requiredHeaders.filter(h => !headerSet.has(h));
  };

  // Validate Listing Sheet Headers
  const listingHeaders = listingData[0];
  let missingListingHeaders = findMissingHeaders(listingHeaders, [CONFIG.headers.listingSku, CONFIG.headers.asin, CONFIG.headers.itemName, CONFIG.headers.status]);
  if (missingListingHeaders.length > 0) {
    ui.alert('Error', `Missing columns in "${CONFIG.sheets.listing}": ${missingListingHeaders.join(', ')}`, ui.ButtonSet.OK);
    return;
  }

  // --- 3. Process Data into Maps ---
  const inventoryMap = processInventory(inventoryData, ui);
  if (!inventoryMap) return;
  const salesMap = processSales(salesData, ui);
  if (!salesMap) return;
  const keepaMap = processKeepa(keepaData, ui);
  if (!keepaMap) return;

  // Get column indexes from Listing Report
  const skuListingIndex = listingHeaders.indexOf(CONFIG.headers.listingSku);
  const asinListingIndex = listingHeaders.indexOf(CONFIG.headers.asin);
  const nameListingIndex = listingHeaders.indexOf(CONFIG.headers.itemName);
  const statusListingIndex = listingHeaders.indexOf(CONFIG.headers.status);

  // --- 4. Build ALL Dashboard Data (before filtering/sorting) ---
  const allDashboardData = [];
  listingData.slice(1).forEach(row => { // Use slice(1) to skip header row
    const sku = row[skuListingIndex];
    if (!sku) return; // Skip rows with no SKU

    const inventory = inventoryMap[sku] || {};
    const sales = salesMap[sku] || {};

    // --- Image Formula & Amazon Link Generation (Using Keepa Image) ---
    const asin = row[asinListingIndex];
    let imageFormula = ""; // Default to blank if no ASIN or Keepa image
    let amazonLink = "";   // Default link to blank

    if (asin) {
      const keepaImageUrl = keepaMap[asin];
      // Use the Keepa image URL if available, otherwise a generic placeholder
      if (keepaImageUrl) {
        imageFormula = `=IMAGE("${keepaImageUrl}")`;
      } else {
        // Fallback placeholder image if Keepa data doesn't have it
        imageFormula = `=IMAGE("https://via.placeholder.com/${CONFIG.image.width}x${CONFIG.image.width/1.2}.png?text=No+Image")`;
      }

      // Create the direct product link using the ASIN
      amazonLink = `https://www.amazon.com/dp/${asin}`;
    }

    const pastMonthSales = sales.pastMonthSales || 0;
    const currentMonthSales = sales.currentMonthSales || 0;
    const monthlyDifference = currentMonthSales - pastMonthSales;

    allDashboardData.push([
      imageFormula, // Index 0
      row[nameListingIndex] || 'N/A', // Index 1
      asin || 'N/A', // Index 2
      sku, // Index 3
      amazonLink, // Index 4
      row[statusListingIndex] || 'N/A', // Index 5: Listing Status
      sales.yesterdaySales || 0, // Index 6
      inventory.available || 0, // Index 7
      inventory.daysOfSupply || 0, // Index 8: Days of Supply
      inventory.reservedCustomerOrder || 0, // Index 9
      inventory.reservedFcTransfer || 0, // Index 10
      inventory.reservedFcProcessing || 0, // Index 11
      inventory.unfulfillable || 0, // Index 12
      pastMonthSales, // Index 13
      currentMonthSales, // Index 14
      monthlyDifference // Index 15
    ]);
  });

  // --- 5. Apply Filtering and Sorting based on reportType ---
  let finalDashboardData = [...allDashboardData]; // Start with all data
  let reportTitle = "All Listings"; // Default title

  if (reportType === 'inactive') {
    reportTitle = "Inactive Listings (FBA-SKU Filtered)";
    finalDashboardData = finalDashboardData.filter(row => {
      const status = row[5]; // Index of 'Listing status'
      const sku = String(row[3]);    // Index of 'SKU', ensure it's a string
      return status === CONFIG.listingStatus.inactive && sku.toUpperCase().startsWith('FBA-');
    });
  } else if (reportType === 'active') {
    reportTitle = "Active Listings (Sorted by Days of Supply)";
    finalDashboardData = finalDashboardData.filter(row => {
      const status = row[5]; // Index of 'Listing status'
      return status === CONFIG.listingStatus.active;
    });
    // Sort active listings by 'Days of Supply' from lowest to highest
    finalDashboardData.sort((a, b) => {
      // Ensure numeric comparison; treat empty/invalid as Infinity for sorting to end
      const daysA = parseFloat(a[8]) || Infinity;
      const daysB = parseFloat(b[8]) || Infinity;
      return daysA - daysB;
    });
  }

  // --- 6. Write Data to Target Sheet ---
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) {
    targetSheet.clear();
  } else {
    targetSheet = ss.insertSheet(targetSheetName);
  }

  const dashboardHeaders = [
    'Image', 'Name', 'Asin', 'SKU', 'Amazon Link', 'Listing status', "Yesterday's Sales", 'Available',
    'Days of Supply', 'Customer Orders', 'FC Transfers', 'FC Processing', 'Unfulfillable',
    'Past 31-60 Day Sales', 'Current 30 Day Sales', 'Monthly Difference'
  ];

  let startRow = 1;

  // Add Context row if provided
  if (contextString) {
    targetSheet.getRange(startRow, 1).setValue(`Dashboard Context: ${contextString}`).setFontWeight('bold');
    targetSheet.getRange(startRow, 1, 1, dashboardHeaders.length).merge(); // Merge across header width
    startRow++; // Shift header and data rows down
  }

  // Add Report Type Title
  targetSheet.getRange(startRow, 1).setValue(`Report Type: ${reportTitle}`).setFontWeight('bold');
  targetSheet.getRange(startRow, 1, 1, dashboardHeaders.length).merge();
  startRow++; // Shift header and data rows down

  targetSheet.getRange(startRow, 1, 1, dashboardHeaders.length).setValues([dashboardHeaders]).setFontWeight('bold');

  if (finalDashboardData.length > 0) {
    const dataRange = targetSheet.getRange(startRow + 1, 1, finalDashboardData.length, finalDashboardData[0].length);
    dataRange.setValues(finalDashboardData); // Write all data as values first

    // Then, apply formulas to the image column
    const imageFormulas = finalDashboardData.map(row => [row[0]]);
    targetSheet.getRange(startRow + 1, 1, imageFormulas.length, 1).setFormulas(imageFormulas);
  }

  targetSheet.setColumnWidth(1, CONFIG.image.width);
  // Auto-resize from column 2 onwards, as column 1 (Image) has a fixed width.
  targetSheet.autoResizeColumns(2, dashboardHeaders.length - 1);
  ui.alert('Success!', `"${targetSheetName}" has been updated with ${finalDashboardData.length} products for "${reportTitle}".`, ui.ButtonSet.OK);
}

// --- Helper Functions ---

/** Processes FBA Inventory data. Now includes header validation. */
function processInventory(data, ui) {
  const map = {};
  const headers = data.shift();
  const required = [CONFIG.headers.invSku, CONFIG.headers.available, CONFIG.headers.daysOfSupply, CONFIG.headers.unfulfillable, CONFIG.headers.reservedCustomer, CONFIG.headers.reservedTransfer, CONFIG.headers.reservedProcessing];
  const missing = required.filter(h => headers.indexOf(h) === -1);
  if (missing.length > 0) {
    ui.alert('Error', `Missing columns in "${CONFIG.sheets.inventory}": ${missing.join(', ')}`, ui.ButtonSet.OK);
    return null;
  }

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
        available: row[availableIndex],
        daysOfSupply: row[daysOfSupplyIndex],
        unfulfillable: row[unfulfillableIndex],
        reservedCustomerOrder: row[reservedCustomerIndex],
        reservedFcTransfer: row[reservedTransferIndex],
        reservedFcProcessing: row[reservedProcessingIndex]
      };
    }
  });
  return map;
}

/** Processes Sales data. Now includes header validation and robust date logic. */
function processSales(data, ui) {
  const map = {};
  const headers = data.shift();
  const required = [CONFIG.headers.salesSku, CONFIG.headers.purchaseDate, CONFIG.headers.quantity];
  const missing = required.filter(h => headers.indexOf(h) === -1);
  if (missing.length > 0) {
    ui.alert('Error', `Missing columns in "${CONFIG.sheets.sales}": ${missing.join(', ')}`, ui.ButtonSet.OK);
    return null;
  }

  const skuIndex = headers.indexOf(CONFIG.headers.salesSku);
  const purchaseDateIndex = headers.indexOf(CONFIG.headers.purchaseDate);
  const quantityIndex = headers.indexOf(CONFIG.headers.quantity);

  const todayStart = new Date();
  todayStart.setHours(0, 0, 0, 0);

  const yesterdayStart = new Date(todayStart);
  yesterdayStart.setDate(todayStart.getDate() - 1);

  const thirtyDaysAgo = new Date(todayStart);
  thirtyDaysAgo.setDate(todayStart.getDate() - 30);

  const sixtyDaysAgo = new Date(todayStart);
  sixtyDaysAgo.setDate(todayStart.getDate() - 60);

  data.forEach(row => {
    const sku = row[skuIndex];
    // Robust date parsing: Handles various date formats
    const purchaseDate = new Date(row[purchaseDateIndex]);
    const quantity = parseInt(row[quantityIndex], 10);

    if (sku && !isNaN(quantity) && purchaseDate.getTime()) {
      if (!map[sku]) {
        map[sku] = { yesterdaySales: 0, currentMonthSales: 0, pastMonthSales: 0 };
      }

      // ROBUST DATE CHECK: Check if the sale happened between the start of yesterday and start of today
      if (purchaseDate >= yesterdayStart && purchaseDate < todayStart) {
        map[sku].yesterdaySales += quantity;
      }

      if (purchaseDate >= thirtyDaysAgo && purchaseDate < todayStart) {
        map[sku].currentMonthSales += quantity;
      }

      if (purchaseDate >= sixtyDaysAgo && purchaseDate < thirtyDaysAgo) {
        map[sku].pastMonthSales += quantity;
      }
    }
  });
  return map;
}

/** Processes Keepa data to map ASINs to Image URLs. */
function processKeepa(data, ui) {
  const map = {};
  const headers = data.shift();
  const required = [CONFIG.headers.keepaAsin, CONFIG.headers.keepaImage];
  const missing = required.filter(h => headers.indexOf(h) === -1);
  if (missing.length > 0) {
    ui.alert('Error', `Missing columns in "${CONFIG.sheets.keepa}": ${missing.join(', ')}`, ui.ButtonSet.OK);
    return null;
  }

  const asinIndex = headers.indexOf(CONFIG.headers.keepaAsin);
  const imageIndex = headers.indexOf(CONFIG.headers.keepaImage);

  data.forEach(row => {
    const asin = row[asinIndex];
    const imageUrl = row[imageIndex];
    if (asin && imageUrl) {
      map[asin] = imageUrl;
    }
  });
  return map;
}
```
