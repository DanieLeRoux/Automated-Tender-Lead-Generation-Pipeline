TenderScrapers Combined Hotfix v1

Purpose
-------
This hotfix addresses the first quality-assurance issues found in the consolidated daily review:

1. City of Cape Town rows now default to:
   - client = City of Cape Town
   - province = Western Cape

2. The consolidation layer now recalculates is_current centrally from closing_datetime / closing_date.
   It no longer trusts inconsistent source-level is_current flags.

3. Daily review tabs now use only current, non-cancelled, non-archived tenders with valid closing dates.
   Rows without a valid closing date are separated into a new "Needs Manual Date Check" tab.

4. South African source date handling is standardised with dayfirst=True for key affected scrapers.

5. World Bank rows now use:
   - client = World Bank
   - country_name remains the country/region value

6. Source URL fallbacks are inserted centrally where a source scraper output has blank source_url.

7. DBSA rows now default to client = Development Bank of Southern Africa.
   DBSA date parsing has been broadened to detect numeric dates in addition to month-name dates.

8. Daily review now includes:
   - By Source: primary source summary
   - Source Coverage: expanded summary using source_list for duplicate/merged records

Files included
--------------
core/consolidation.py
core/date_utils.py
source_scripts/CityOfCapeTown/coct_tender_pull_and_clean.py
source_scripts/ACSA/acsa_tender_pull_and_clean.py
source_scripts/Transnet/transnet_esupplier_tender_pull_and_clean.py
source_scripts/DBSA/dbsa_procurement_pull_and_clean_v2.py
source_scripts/WorldBank/worldbank_procurement_pull_and_clean_v5.py

Application
-----------
Extract this ZIP directly into:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers

Allow Windows to replace/overwrite the existing files.

Then run from PowerShell:

cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" install_source_scrapers.py --overwrite

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

Afterwards, review:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers\Outputs\Consolidated\consolidated_daily_review.xlsx
