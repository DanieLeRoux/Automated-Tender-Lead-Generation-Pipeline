TenderScrapers Johannesburg Water Source Hotfix v23
===================================================

This patch adds Johannesburg Water as an enabled source.

It scrapes:
- Advertised Tenders: https://scm.johannesburgwater.co.za/supply-chain/tenders/all-open-tenders/
- Advertised RFPs: https://scm.johannesburgwater.co.za/supply-chain/request-for-proposal/all-advertised-rfps/

It intentionally does not scrape the RFQ page in v1 because RFQs are expected to be noisy and mostly excluded by the central bid-manager filter.

Apply by extracting this ZIP into:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers

Then run:
cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" install_source_scrapers.py --overwrite
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py
