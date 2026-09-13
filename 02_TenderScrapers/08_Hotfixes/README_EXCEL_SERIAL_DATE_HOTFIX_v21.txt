TenderScrapers Excel Serial Date Handling Hotfix v21
===================================================

Purpose
-------
Fixes a central consolidation date-parsing issue where valid tender closing dates read from Excel source workbooks could be lost during consolidation.

Issue Addressed
---------------
Some source workbooks, especially eTenders outputs, may expose dates to the consolidation layer as Excel serial date values or numeric strings. Example:

46191.4166666667 = 2026-06-18 10:00:00

Pandas' default pd.to_datetime parser can misread numeric Excel serial values as 1970-based nanosecond timestamps or NaT. This caused valid eTenders dates to become blank/invalid in consolidated_master.xlsx and caused rows to appear in Needs Manual Date Check.

Fix
---
- Adds robust Excel serial date parsing in core/date_utils.py.
- Updates central consolidation date handling to use the robust parser consistently.
- Ensures eTenders closing_date values are correctly mapped into consolidated closing_datetime.
- Ensures duplicate date coalescing and daily review logic use the robust date parser.
- Keeps closing_date aligned to the date part of closing_datetime.
- Recalculates is_current and days_to_close after repair.

Expected Result
---------------
- eTenders rows with valid source closing dates should no longer appear in Needs Manual Date Check solely due to date conversion failure.
- The TCTA/eTenders duplicate tender 051/2025/PMID/STUDY/RFB should retain the 2026-06-18 10:00 closing date from eTenders and/or TCTA.
- Needs Manual Date Check should only contain rows where no valid closing date exists after parsing and duplicate coalescing.

Apply with:

Expand-Archive `
  -LiteralPath "C:\Users\Danie le Roux\Downloads\TenderScrapers_Excel_Serial_Date_Hotfix_v21.zip" `
  -DestinationPath "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers" `
  -Force

Then run:

cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

No install_source_scrapers.py step is required unless you also need to reinstall source-level scripts. This patch changes the central consolidation layer and schema/date utility files only.
