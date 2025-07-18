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
 * This script creates consolidated Amazon dashboards and a professional analysis tab.
 * VERSION 14.1: Refined for enhanced error handling and robustness.
 * - Refined the global error handler to safely operate in any execution context (UI or background).
 * - Made the showSidebar function more resilient by ensuring a UI instance is always available.
 * - Maintained existing performance, robustness, and maintainability enhancements.
 */

// --- CONFIGURATION --- //
const CONFIG = {
  sheets: {
    listing: 'All_Listing_Report',
    inventory: 'FBA_Inventory_Snapshot',
    sales: 'FBA_Shipments_60_Days',
    keepa: 'Keepa',
    dashboard: 'Dashboard',
    activeListings: 'Active_Listing',
    inactiveListings: 'Inactive_Listing',
    analysis: 'Analysis',
    chartData: '_AnalysisChartData' // Hidden sheet for chart source data
  },
  headers: {
    listingSku: 'seller-sku', salesSku: 'sku', invSku: 'sku', asin: 'asin1',
    itemName: 'item-name', status: 'status', available: 'available',
    daysOfSupply: 'days-of-supply', unfulfillable: 'unfulfillable-quantity',
    reservedCustomer: 'Reserved Customer Order', reservedTransfer: 'Reserved FC Transfer',
    reservedProcessing: 'Reserved FC Processing', purchaseDate: 'purchase-date',
    quantity: 'quantity', itemPrice: 'item-price', orderId: 'amazon-order-id',
    keepaAsin: 'ASIN', keepaImage: 'Image'
  },
  image: { width: 100 },
  listingStatus: { active: 'Active', inactive: 'Inactive' },
  chartLimits: { topNProducts: 10 }
};
// --- END CONFIGURATION ---

// --- UI & INITIALIZATION --- //

/**
 * Adds a custom menu to the spreadsheet UI when the file is opened.
 * Also displays the HTML sidebar.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi(); // Get UI once here
  try {
    ui.createMenu('Dashboard Tools')
      .addItem('Open Dashboard Controls', 'showSidebar')
      .addToUi();
    showSidebar(ui); // Pass UI instance to avoid an extra API call
  } catch (e) {
    _handleError(e, ui, 'onOpen');
  }
}

// --- REFINEMENT ---
// The `showSidebar` function has been updated to be more robust. It now ensures a UI 
// instance exists, whether it's passed in as an argument or fetched directly. This prevents
// errors if the function is ever called from a context that doesn't provide the `ui` object.
/**
 * Displays the HTML sidebar for user interaction.
 * @param {GoogleAppsScript.Spreadsheet.Ui} [ui] The spreadsheet UI instance (optional).
 */
function showSidebar(ui) {
  const effectiveUi = ui || SpreadsheetApp.getUi();
  try {
    const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Dashboard Controls')
      .setWidth(300);
    effectiveUi.showSidebar(html);
  } catch (e) {
    _handleError(e, effectiveUi, 'showSidebar');
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
        return generateAnalysisTab(ss, ui);
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
 * Generates the 'All_Listing_Report' dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @param {GoogleAppsScript.Spreadsheet.Ui} ui The spreadsheet UI instance.
 * @param {string} c User-provided context string.
 * @returns {string} Success message.
 */
function generateAllListingsDashboard(ss, ui, c) {
  return _generateDashboardReport(ss, ui, c, 'all', CONFIG.sheets.dashboard);
}

/**
 * Generates the 'Inactive_Listing' dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @param {GoogleAppsScript.Spreadsheet.Ui} ui The spreadsheet UI instance.
 * @param {string} c User-provided context string.
 * @returns {string} Success message.
 */
function generateInactiveListingsDashboard(ss, ui, c) {
  return _generateDashboardReport(ss, ui, c, 'inactive', CONFIG.sheets.inactiveListings);
}

/**
 * Generates the 'Active_Listing' dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @param {GoogleAppsScript.Spreadsheet.Ui} ui The spreadsheet UI instance.
 * @param {string} c User-provided context string.
 * @returns {string} Success message.
 */
function generateActiveListingsDashboard(ss, ui, c) {
  return _generateDashboardReport(ss, ui, c, 'active', CONFIG.sheets.activeListings);
}

/**
 * Generates the professional 'Analysis' tab with KPIs and charts.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet instance.
 * @param {GoogleAppsScript.Spreadsheet.Ui} ui The spreadsheet UI instance.
 * @returns {string} Success message.
 */
function generateAnalysisTab(ss, ui) {
  let analysisSheet = ss.getSheetByName(CONFIG.sheets.analysis);
  if (analysisSheet) {
    const charts = analysisSheet.getCharts();
    charts.forEach(chart => analysisSheet.removeChart(chart));
    analysisSheet.clear();
  } else {
    analysisSheet = ss.insertSheet(CONFIG.sheets.analysis);
  }

  let chartDataSheet = ss.getSheetByName(CONFIG.sheets.chartData);
  if (chartDataSheet) {
    chartDataSheet.clear();
  } else {
    chartDataSheet = ss.insertSheet(CONFIG.sheets.chartData);
    chartDataSheet.hideSheet();
  }

  const salesDataObject = _getValidatedData(ss, CONFIG.sheets.sales, [CONFIG.headers.purchaseDate, CONFIG.headers.quantity, CONFIG.headers.itemPrice, CONFIG.headers.orderId, CONFIG.headers.salesSku]);
  const listingDataObject = _getValidatedData(ss, CONFIG.sheets.listing, [CONFIG.headers.listingSku, CONFIG.headers.itemName]);

  const weeklySales = _processWeeklySalesData(salesDataObject);
  const marketShareData = _processMarketShareData(salesDataObject, listingDataObject);
  const topNSalesData = _processTopNSalesData(salesDataObject, listingDataObject);

  _createProfessionalLayout(analysisSheet, weeklySales);
  _createSalesTrendChart(analysisSheet, chartDataSheet, weeklySales);
  _createMarketShareChart(analysisSheet, chartDataSheet, marketShareData);
  _createTopNSalesChart(analysisSheet, chartDataSheet, topNSalesData);

  analysisSheet.setColumnWidths(1, 12, 110).setFrozenRows(1);
  ss.setActiveSheet(analysisSheet);

  return `The professional "${CONFIG.sheets.analysis}" tab has been generated.`;
}

// --- ANALYSIS TAB HELPERS --- //

/**
 * Creates the formatted layout with KPI cards and titles.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The target sheet.
 * @param {object} sales Processed weekly sales data including KPIs.
 */
function _createProfessionalLayout(sheet, sales) {
  sheet.getRange('A1:L1').merge().setValue('Amazon Sales Performance Dashboard').setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('A1:L1').setBackground('#2c3e50').setFontColor('white');

  const now = new Date();
  sheet.getRange('A2').setValue(`Last Updated: ${now.toLocaleString()}`).setFontSize(9).setFontColor('#7f8c8d');
  sheet.getRange('A2:L2').merge().setHorizontalAlignment('right');

  sheet.getRange('B4:E4').merge().setValue('Key Performance Indicators (Last 7 Days)').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('B4:E11').setBorder(true, true, true, true, true, true).setBackground('#ecf0f1');

  const kpiLabels = ['Total Sales Revenue:', 'Total Units Sold:', 'Total Orders:', 'Average Order Value:', 'Average Units Per Order:'];
  const kpiValues = [sales.weekTotalSales, sales.weekTotalUnits, sales.weekTotalOrders, sales.avgOrderValue, sales.avgUnitsPerOrder];
  const kpiStartRow = 6;
  kpiLabels.forEach((label, i) => {
    sheet.getRange(`C${kpiStartRow + i}`).setValue(label).setFontWeight('bold').setHorizontalAlignment('right');
    sheet.getRange(`D${kpiStartRow + i}`).setValue(kpiValues[i]).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('right');
  });

  sheet.getRange(`D${kpiStartRow}`).setNumberFormat('$#,##0.00');
  sheet.getRange(`D${kpiStartRow + 1}`).setNumberFormat('#,##0');
  sheet.getRange(`D${kpiStartRow + 2}`).setNumberFormat('#,##0');
  sheet.getRange(`D${kpiStartRow + 3}`).setNumberFormat('$#,##0.00');
  sheet.getRange(`D${kpiStartRow + 4}`).setNumberFormat('#,##0.00');

  sheet.getRange('F4:L4').merge().setValue('Daily Sales Trend (Last 7 Days)').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('F4:L16').setBorder(true, true, true, true, true, true).setBackground('#ecf0f1');

  sheet.getRange('B18:E18').merge().setValue('Product Market Share (by Sales Revenue)').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('B18:E36').setBorder(true, true, true, true, true, true).setBackground('#ecf0f1');

  sheet.getRange('F18:L18').merge().setValue(`Top ${CONFIG.chartLimits.topNProducts} Selling Products (Last 60 Days)`).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('F18:L36').setBorder(true, true, true, true, true, true).setBackground('#ecf0f1');

  sheet.getRange('A1:L36').setVerticalAlignment('middle');
  sheet.getRange('B4:L36').setFontSize(10);
}

/**
 * Creates and inserts the daily sales trend line chart.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The target display sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet The sheet to store chart data.
 * @param {object} sales Processed weekly sales data.
 */
function _createSalesTrendChart(sheet, dataSheet, sales) {
  const chartData = [['Date', 'Units Ordered', 'Ordered Product Sales']];
  const sortedDates = Object.keys(sales.daily).sort((a, b) => new Date(a).getTime() - new Date(b).getTime());

  sortedDates.forEach(date => {
    const data = sales.daily[date];
    chartData.push([new Date(date), data.units, data.sales]);
  });

  if (chartData.length < 2) {
    sheet.getRange('F5').setValue('Not enough sales trend data to generate chart.').setFontStyle('italic').setHorizontalAlignment('center');
    sheet.getRange('F5:L5').merge();
    return;
  }

  const chartDataRange = dataSheet.getRange(1, 1, chartData.length, chartData[0].length);
  chartDataRange.setValues(chartData);

  const lineChart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(chartDataRange)
    .setPosition(5, 6, 0, 0)
    .setOption('title', null)
    .setOption('legend', { position: 'top', textStyle: { fontSize: 10 } })
    .setOption('hAxis', { title: 'Date', format: 'M/d', textStyle: { fontSize: 9 }, titleTextStyle: { fontSize: 10 } })
    .setOption('vAxes', {
      0: { title: 'Units Ordered', textStyle: { fontSize: 9 }, titleTextStyle: { fontSize: 10 }, format: '#,##0' },
      1: { title: 'Product Sales', textStyle: { fontSize: 9 }, titleTextStyle: { fontSize: 10 }, format: '$#,##0.00' }
    })
    .setOption('series', {
      0: { targetAxisIndex: 0, color: '#3498db', type: 'bars' },
      1: { targetAxisIndex: 1, color: '#e74c3c', type: 'line', lineWidth: 2, pointSize: 4 }
    })
    .setOption('chartArea', { left: '15%', top: '15%', width: '70%', height: '70%' })
    .setOption('focusTarget', 'category')
    .build();

  sheet.insertChart(lineChart);
}

/**
 * Creates and inserts the product market share treemap chart.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The target display sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet The sheet to store chart data.
 * @param {Array<Array<any>>} marketShareData Processed market share data.
 */
function _createMarketShareChart(sheet, dataSheet, marketShareData) {
  if (marketShareData.length < 3) {
    sheet.getRange('B19').setValue('Not enough sales data to generate market share chart.').setFontStyle('italic').setHorizontalAlignment('center');
    sheet.getRange('B19:E19').merge();
    return;
  }

  const chartDataRange = dataSheet.getRange(1, 6, marketShareData.length, marketShareData[0].length);
  chartDataRange.setValues(marketShareData);

  const treemapChart = sheet.newChart()
    .setChartType(Charts.ChartType.TREEMAP)
    .addRange(chartDataRange)
    .setPosition(19, 2, 0, 0)
    .setOption('title', null)
    .setOption('minColor', '#e6f4ea')
    .setOption('midColor', '#87c5a4')
    .setOption('maxColor', '#225c4e')
    .setOption('headerHeight', 15)
    .setOption('fontColor', 'black')
    .setOption('showScale', true)
    .setOption('generateTooltip', true)
    .setOption('chartArea', { left: '5%', top: '10%', width: '90%', height: '85%' })
    .build();

  sheet.insertChart(treemapChart);
}

/**
 * Creates and inserts the top N selling products bar chart.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The target display sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet The sheet to store chart data.
 * @param {Array<Array<any>>} topNSalesData Processed top N sales data.
 */
function _createTopNSalesChart(sheet, dataSheet, topNSalesData) {
  if (topNSalesData.length < 2) {
    sheet.getRange('F19').setValue(`Not enough sales data to generate top ${CONFIG.chartLimits.topNProducts} products chart.`).setFontStyle('italic').setHorizontalAlignment('center');
    sheet.getRange('F19:L19').merge();
    return;
  }

  const chartDataRange = dataSheet.getRange(1, 10, topNSalesData.length, topNSalesData[0].length);
  chartDataRange.setValues(topNSalesData);

  const barChart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(chartDataRange)
    .setPosition(19, 6, 0, 0)
    .setOption('title', null)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: 'Sales Revenue', format: '$#,##0.00', textStyle: { fontSize: 9 }, titleTextStyle: { fontSize: 10 } })
    .setOption('vAxis', { title: 'Product Name', textStyle: { fontSize: 9 }, titleTextStyle: { fontSize: 10 } })
    .setOption('series', { 0: { color: '#2ecc71' } })
    .setOption('chartArea', { left: '25%', top: '10%', width: '70%', height: '80%' })
    .build();

  sheet.insertChart(barChart);
}

/**
 * Processes sales data for weekly trends and KPIs.
 * @param {{headers: string[], data: any[][]}} salesDataObject The validated sales data object.
 * @returns {object} Processed sales data including KPIs.
 */
function _processWeeklySalesData(salesDataObject) {
  const { headers, data } = salesDataObject;
  const dateIdx = headers.indexOf(CONFIG.headers.purchaseDate);
  const qtyIdx = headers.indexOf(CONFIG.headers.quantity);
  const priceIdx = headers.indexOf(CONFIG.headers.itemPrice);
  const orderIdIdx = headers.indexOf(CONFIG.headers.orderId);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const sevenDaysAgo = new Date(today);
  sevenDaysAgo.setDate(today.getDate() - 7);

  const weeklyData = {};
  const weeklyOrderIds = new Set();

  for (let i = 0; i < 7; i++) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    weeklyData[d.toLocaleDateString()] = { units: 0, sales: 0 };
  }

  data.forEach(row => {
    const purchaseDate = new Date(row[dateIdx]);
    if (!isNaN(purchaseDate.getTime()) && purchaseDate >= sevenDaysAgo && purchaseDate <= today) {
      const key = purchaseDate.toLocaleDateString();
      if (key in weeklyData) {
        const quantity = parseFloat(row[qtyIdx]) || 0;
        const price = parseFloat(row[priceIdx]) || 0;
        weeklyData[key].units += quantity;
        weeklyData[key].sales += quantity * price;
        weeklyOrderIds.add(row[orderIdIdx]);
      }
    }
  });

  const weekTotalOrders = weeklyOrderIds.size;
  const weekTotalUnits = Object.values(weeklyData).reduce((sum, day) => sum + day.units, 0);
  const weekTotalSales = Object.values(weeklyData).reduce((sum, day) => sum + day.sales, 0);
  const avgOrderValue = weekTotalOrders > 0 ? (weekTotalSales / weekTotalOrders) : 0;
  const avgUnitsPerOrder = weekTotalOrders > 0 ? (weekTotalUnits / weekTotalOrders) : 0;

  return { daily: weeklyData, weekTotalOrders, weekTotalUnits, weekTotalSales, avgOrderValue, avgUnitsPerOrder };
}

/**
 * Helper to process market share data (for Treemap chart).
 * @param {{headers: string[], data: any[][]}} salesDataObject The validated sales data object.
 * @param {{headers: string[], data: any[][]}} listingDataObject The validated listing data object.
 * @returns {Array<Array<any>>} Formatted data for treemap chart.
 */
function _processMarketShareData(salesDataObject, listingDataObject) {
  const { headers: salesHeaders, data: salesData } = salesDataObject;
  const salesSkuIdx = salesHeaders.indexOf(CONFIG.headers.salesSku);
  const salesQtyIdx = salesHeaders.indexOf(CONFIG.headers.quantity);
  const salesPriceIdx = salesHeaders.indexOf(CONFIG.headers.itemPrice);

  const salesBySku = {};
  salesData.forEach(row => {
    const sku = String(row[salesSkuIdx] || '').trim();
    if (sku) {
      const quantity = parseFloat(row[salesQtyIdx]) || 0;
      const price = parseFloat(row[salesPriceIdx]) || 0;
      salesBySku[sku] = (salesBySku[sku] || 0) + (quantity * price);
    }
  });

  const skuToNameMap = {};
  if (listingDataObject && listingDataObject.data.length > 0) {
    const { headers: listingHeaders, data: listingData } = listingDataObject;
    const listingSkuIdx = listingHeaders.indexOf(CONFIG.headers.listingSku);
    const listingNameIdx = listingHeaders.indexOf(CONFIG.headers.itemName);
    listingData.forEach(row => {
      const sku = String(row[listingSkuIdx] || '').trim();
      const name = String(row[listingNameIdx] || '').trim();
      if (sku && name) skuToNameMap[sku] = name;
    });
  } else {
    Logger.log(`Warning: Listing report sheet "${CONFIG.sheets.listing}" is empty or has no data, using SKUs for product names.`);
  }

  const treemapData = [['Product', 'Parent', 'Sales Revenue']];
  treemapData.push(['All Products', null, 0]);

  for (const [sku, totalRevenue] of Object.entries(salesBySku)) {
    if (totalRevenue > 0) {
      treemapData.push([skuToNameMap[sku] || sku, 'All Products', totalRevenue]);
    }
  }
  return treemapData;
}

/**
 * Helper to process sales data for top N selling products.
 * @param {{headers: string[], data: any[][]}} salesDataObject The validated sales data object.
 * @param {{headers: string[], data: any[][]}} listingDataObject The validated listing data object.
 * @returns {Array<Array<any>>} Formatted data for bar chart.
 */
function _processTopNSalesData(salesDataObject, listingDataObject) {
  const { headers: salesHeaders, data: salesData } = salesDataObject;
  const salesSkuIdx = salesHeaders.indexOf(CONFIG.headers.salesSku);
  const salesQtyIdx = salesHeaders.indexOf(CONFIG.headers.quantity);
  const salesPriceIdx = salesHeaders.indexOf(CONFIG.headers.itemPrice);
  const salesDateIdx = salesHeaders.indexOf(CONFIG.headers.purchaseDate);

  const salesBySku = {};
  const sixtyDaysAgo = new Date();
  sixtyDaysAgo.setDate(sixtyDaysAgo.getDate() - 60);
  sixtyDaysAgo.setHours(0, 0, 0, 0);

  salesData.forEach(row => {
    const purchaseDate = new Date(row[salesDateIdx]);
    if (!isNaN(purchaseDate.getTime()) && purchaseDate >= sixtyDaysAgo) {
      const sku = String(row[salesSkuIdx] || '').trim();
      if (sku) {
        const quantity = parseFloat(row[salesQtyIdx]) || 0;
        const price = parseFloat(row[salesPriceIdx]) || 0;
        salesBySku[sku] = (salesBySku[sku] || 0) + (quantity * price);
      }
    }
  });

  const skuToNameMap = {};
  if (listingDataObject && listingDataObject.data.length > 0) {
    const { headers: listingHeaders, data: listingData } = listingDataObject;
    const listingSkuIdx = listingHeaders.indexOf(CONFIG.headers.listingSku);
    const listingNameIdx = listingHeaders.indexOf(CONFIG.headers.itemName);
    listingData.forEach(row => {
      const sku = String(row[listingSkuIdx] || '').trim();
      const name = String(row[listingNameIdx] || '').trim();
      if (sku && name) skuToNameMap[sku] = name;
    });
  }

  const sortedProducts = Object.entries(salesBySku)
    .map(([sku, revenue]) => ({ name: skuToNameMap[sku] || sku, revenue: revenue }))
    .sort((a, b) => b.revenue - a.revenue)
    .slice(0, CONFIG.chartLimits.topNProducts);

  const chartData = [['Product', 'Sales Revenue']];
  sortedProducts.forEach(p => chartData.push([p.name, p.revenue]));

  return chartData;
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
  const listingDataObject = _getValidatedData(ss, CONFIG.sheets.listing, [CONFIG.headers.listingSku, CONFIG.headers.asin, CONFIG.headers.itemName, CONFIG.headers.status]);
  const inventoryDataObject = _getValidatedData(ss, CONFIG.sheets.inventory, [CONFIG.headers.invSku, CONFIG.headers.available, CONFIG.headers.daysOfSupply, CONFIG.headers.unfulfillable, CONFIG.headers.reservedCustomer, CONFIG.headers.reservedTransfer, CONFIG.headers.reservedProcessing]);
  const salesDataObject = _getValidatedData(ss, CONFIG.sheets.sales, [CONFIG.headers.salesSku, CONFIG.headers.purchaseDate, CONFIG.headers.quantity]);
  const keepaDataObject = _getValidatedData(ss, CONFIG.sheets.keepa, [CONFIG.headers.keepaAsin, CONFIG.headers.keepaImage]);

  const inventoryMap = processInventory(inventoryDataObject);
  const salesMap = processSales(salesDataObject);
  const keepaMap = processKeepa(keepaDataObject);

  const { headers: listingHeaders, data: listingData } = listingDataObject;
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

    const pastMonthSales = sales.pastMonthSales || 0;
    const currentMonthSales = sales.currentMonthSales || 0;

    allDashboardData.push([
      imageFormula, row[nameListingIndex] || 'N/A', asin || 'N/A', sku, amazonLink,
      row[statusListingIndex] || 'N/A', sales.yesterdaySales || 0, inventory.available || 0,
      inventory.daysOfSupply || 0, inventory.reservedCustomerOrder || 0,
      inventory.reservedFcTransfer || 0, inventory.reservedFcProcessing || 0,
      inventory.unfulfillable || 0, pastMonthSales, currentMonthSales, currentMonthSales - pastMonthSales
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

  const dashboardHeaders = ['Image', 'Name', 'Asin', 'SKU', 'Amazon Link', 'Listing status', "Yesterday's Sales", 'Available', 'Days of Supply', 'Customer Orders', 'FC Transfers', 'FC Processing', 'Unfulfillable', 'Past 31-60 Day Sales', 'Current 30 Day Sales', 'Monthly Difference'];
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
  const { headers, data } = inventoryDataObject;
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
 * Processes sales data into a SKU-based map for daily, current month, and past month sales.
 * @param {{headers: string[], data: any[][]}} salesDataObject The validated sales data object.
 * @returns {object} A map with SKU as key and sales details as value.
 */
function processSales(salesDataObject) {
  const map = {};
  const { headers, data } = salesDataObject;
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
    const sku = String(row[skuIndex]).trim();
    const purchaseDate = new Date(row[purchaseDateIndex]);
    const quantity = parseInt(row[quantityIndex], 10);
    if (sku && !isNaN(quantity) && !isNaN(purchaseDate.getTime())) {
      if (!map[sku]) map[sku] = { yesterdaySales: 0, currentMonthSales: 0, pastMonthSales: 0 };
      if (purchaseDate >= yesterdayStart && purchaseDate < todayStart) map[sku].yesterdaySales += quantity;
      if (purchaseDate >= thirtyDaysAgo && purchaseDate < todayStart) map[sku].currentMonthSales += quantity;
      if (purchaseDate >= sixtyDaysAgo && purchaseDate < thirtyDaysAgo) map[sku].pastMonthSales += quantity;
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
  const { headers, data } = keepaDataObject;
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

// --- REFINEMENT ---
// The error handler is updated to prevent the "Cannot read properties of undefined (reading 'alert')"
// error. It now checks if a UI context is available before attempting to show a UI alert.
// It will always log the full error for debugging, regardless of the context.
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
  if (data.length < 1) throw new Error(`Sheet "${sheetName}" is empty.`);
  
  const headers = data[0].map(h => String(h).trim());
  const headerSet = new Set(headers);
  const missingHeaders = requiredHeaders.filter(h => !headerSet.has(h));

  if (missingHeaders.length > 0) {
    throw new Error(`Missing required columns in "${sheetName}": ${missingHeaders.join(', ')}`);
  }

  return { headers: headers, data: data.slice(1) };
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
