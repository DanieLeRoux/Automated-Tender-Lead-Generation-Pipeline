TenderScrapers Daily Review Eligibility + Reporting Repair Hotfix v30
=====================================================================

Purpose
-------
This patch repairs the daily-review generation layer after final QA showed that
consolidated_master.xlsx correctly marked excluded tenders, but the daily-review
workbook still displayed some excluded rows.

What changes
------------
1. Rebuilds daily-review operational tabs from review-eligible records only.
   Applied to:
   - Very High Priority
   - High Priority
   - Medium Priority
   - Closing Soon
   - New Today
   - Updated Today
   - Needs Manual Date Check

2. Enforces exclusion logic in the daily-review writer:
   - review_eligible must be TRUE
   - excluded_by_filter must not be TRUE
   - main priority tabs must have a valid closing date
   - Needs Manual Date Check only receives eligible/non-excluded records with no reliable closing date

3. Restores reporting tabs:
   - By Source
   - Source Coverage
   - Scraper Run Summary

4. Adds a QA Checks tab to the daily review workbook.

5. Keeps robust Excel serial-date handling from v21.

6. Adds source defaults/aliases for final expanded sources:
   - City of Ekurhuleni
   - Johannesburg Water
   - City of Mbombela
   - City of Polokwane
   - NamWater

7. Includes the latest sources_config.json with COGTA and Sibanye-Stillwater disabled.

Apply
-----
Expand this zip into the TenderScrapers root folder with -Force, then run:

cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

No install_source_scrapers.py run is required unless you want to reinstall source scripts.
