TenderScrapers Document / Action Link Hotfix v2
=============================================

This hotfix supersedes the earlier CoCT-only hotfix and builds on the combined hotfix v1.

Purpose
-------
This patch treats eTenders as a discovery feed and improves review usability by adding
standard link/action handling across the consolidated outputs.

Changes included
----------------
1. Adds two standard consolidated columns:
   - action_url
   - document_access_status

2. eTenders handling:
   - eTenders remains a discovery feed.
   - Where no direct document URL is available, action_url points to the eTenders portal.
   - document_access_status is set to "Manual eTenders lookup required".

3. Direct-document handling:
   - Where document_urls exists, action_url is set to the first direct document URL.
   - document_access_status is set to "Direct document available".

4. Source/detail handling:
   - Where no direct document exists but a source/detail page exists, action_url is set to that page.
   - document_access_status is set to "Source/detail page only".

5. uMngeni-uThukela Water:
   - Prefers row-level view/title URLs over the generic portal URL.
   - This gives reviewers a better click-through path from consolidated outputs.

6. ACSA:
   - Deduplicates duplicate rows by document URL / description.
   - Uses earliest closing date when the same ACSA document appears more than once.
   - Improves opportunity_key generation where no tender number is visible.

7. DBSA:
   - Broadens DBSA date extraction to handle hyphenated dates and labelled closing-date text.
   - Keeps DBSA client/source defaults from the previous combined hotfix.

Files updated
-------------
core/standard_schema.py
core/consolidation.py
core/date_utils.py
source_scripts/ACSA/acsa_tender_pull_and_clean.py
source_scripts/DBSA/dbsa_procurement_pull_and_clean_v2.py
source_scripts/uMngeniUthukelaWater/umngeni_uthukela_tender_pull_and_clean.py

Application steps
-----------------
Extract this ZIP into:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers

Allow overwrite when prompted.

Then run:
cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" install_source_scrapers.py --overwrite
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

Expected result
---------------
The consolidated outputs should include action_url and document_access_status.

For eTenders rows without direct document links:
document_access_status = Manual eTenders lookup required

action_url = https://www.etenders.gov.za/Home/opportunities?id=1

For source portal rows with direct documents:
document_access_status = Direct document available

action_url = first direct document URL
