TenderScrapers NECSA Source Hotfix v10
=====================================

Adds NECSA as an enabled source in the TenderScrapers framework.

Source URL:
https://www.necsa.co.za/tenders/

Working folder created/used:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\NECSA

Expected outputs:
NECSA\Outputs\necsa_raw.csv
NECSA\Outputs\necsa_raw_full.xlsx
NECSA\Outputs\necsa_raw_page.html
NECSA\Outputs\necsa_master.xlsx
NECSA\Outputs\necsa_delta.xlsx
NECSA\Outputs\necsa_updates.xlsx
NECSA\Outputs\necsa_top_delta.xlsx

Apply:
Expand-Archive `
  -LiteralPath "C:\Users\Danie le Roux\Downloads\TenderScrapers_NECSA_Source_Hotfix_v10.zip" `
  -DestinationPath "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers" `
  -Force

Then run:
cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" install_source_scrapers.py --overwrite
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

Notes:
- The NECSA page is static HTML and exposes tender details on the Current Tenders page.
- The scraper captures BID NUMBER, BID DESCRIPTION, CLOSING DATE/TIME, enquiries, briefing/site-meeting detail and related document ZIP links.
- Central filtering/scoring still determines whether NECSA items appear in the bid-manager daily review.
