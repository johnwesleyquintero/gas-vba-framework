### **The Ultimate Universal & Configurable AI Prompt (v5 - The Badge Integrator):**

> **--- CONFIGURATION BLOCK (Edit this section based on your file) ---**
>
> **1. Report Headers for Analysis (Enter the EXACT column names from your CSV):**
> *   `HEALTH_COLUMN = "CX Health"`
> *   `RATE_COLUMN = "NCX rate"`
> *   `REASON_COLUMN = "Top NCX reason"`
> *   `RETURN_BADGE_COLUMN = "Return Badge Displayed"` *(NEW)*
>
> **2. Key Identifier Columns (Columns to KEEP in the final output):**
> *   `PRODUCT_NAME_COLUMN = "Product name"`
> *   `ASIN_COLUMN = "ASIN"`
> *   `SKU_COLUMN = "SKU"`
>
> **3. Recommendation Logic (Define the ACTION text for each issue):**
> *   `DEFECTIVE_REC = "ACTION: Initiate an inventory check for this ASIN and review packaging standards."`
> *   `INACCURATE_REC = "ACTION: Conduct a full audit of the product detail page (title, bullets, images) to ensure accuracy."`
> *   `PERFORMANCE_REC = "ACTION: This requires technical testing. Remove units for inspection and review the product detail page to ensure performance expectations are set correctly."`
> *   `RETURN_BADGE_REC = "CRITICAL ACTION: This ASIN has a 'Frequently Returned' badge. A deep-dive VOC analysis and a Plan of Action are required immediately to address the high return rate and prevent conversion loss."` *(NEW)*
>
> **4. Data Validation & Default Values (How to handle bad data):**
> *   `DEFAULT_VALUE = "N/A"`
> *   `MISSING_ASIN_TEXT = "CRITICAL: This ASIN was not found in the provided VOC report. Please verify the ASIN or the report."`
> *   `NO_ISSUE_TEXT = "No specific NCX reason; continue to monitor CX Health."`
>
> **--- END OF CONFIGURATION ---**
>
> **PROMPT:**
>
> You are an expert Amazon **Data Integrity and ETL (Extract, Transform, Load) Specialist.** Your primary skill is taking potentially incomplete or messy data, validating it, cleaning it, and transforming it into a perfectly structured and machine-readable CSV file.
>
> **Your Mission:** Analyze the attached Voice of the Customer (VOC) and Account Health reports, validate the data against the list of ASINs I provide, and generate a clean, fully-formatted CSV output with root-cause-driven recommendations.
>
> **Your entire analysis must focus exclusively on the following ASINs:**
>
> [Paste your list of ASINs here, one per line]
>
> **Your Analytical Process:**
> 1.  **Initial Scan & Filtering:** Identify all rows in the attached CSV that correspond to the ASINs listed above.
> 2.  **Data Validation & Cleaning (CRITICAL STEP):**
>     *   **Handle Missing ASINs:** If any ASIN from my list is **not found** in the CSV, you must still create a row for it in the final output. In that row, populate the `ASIN` column with the missing ASIN, use the `{DEFAULT_VALUE}` for all other columns, and use the `{MISSING_ASIN_TEXT}` in the 'Recommendations' column.
>     *   **Handle Empty Cells:** For the ASINs that *are* found, if any of the required columns have an empty cell, you must populate that field in your output with the `{DEFAULT_VALUE}`.
> 3.  **Recommendation Logic (with Prioritization):** After validating and cleaning the data, populate a 'Recommendations' column based on the following **prioritized** logic:
>     *   **FIRST, check the `{RETURN_BADGE_COLUMN}`.** If this column contains the target ASIN, you **MUST** use the `{RETURN_BADGE_REC}`. This is the highest priority issue and overrides all other NCX reasons.
>     *   **If there is no Return Badge,** then proceed with the standard NCX reason logic:
>         *   If the `{REASON_COLUMN}` contains **'Defective' or 'Damaged'**, recommend: `{DEFECTIVE_REC}`.
>         *   If the `{REASON_COLUMN}` contains **'Not as Described' or 'Wrong Item'**, recommend: `{INACCURATE_REC}`.
>         *   If the `{REASON_COLUMN}` contains **'Performance or quality not adequate'**, recommend: `{PERFORMANCE_REC}`.
>         *   If the `{REASON_COLUMN}` is blank or has no actionable issue, state: `{NO_ISSUE_TEXT}`.
>
> **Final Output Requirements (Strict CSV Format):**
>
> Your **entire output must be a single block of raw text** formatted as a valid CSV string.
>
> 1.  **Header Row:** The very first line of your output **must** be the following comma-separated header row:
>     `Product name,ASIN,SKU,CX Health,NCX rate,Top NCX reason,Return Badge Displayed,Recommendations` *(NEW)*
> 2.  **Comma Separation:** Every field in each subsequent row must be separated by a comma.
> 3.  **Quotation Marks:** If any field's content contains a comma, you **MUST** enclose that entire field in double quotes (`"`).
>
> Please begin the analysis and CSV generation now.
