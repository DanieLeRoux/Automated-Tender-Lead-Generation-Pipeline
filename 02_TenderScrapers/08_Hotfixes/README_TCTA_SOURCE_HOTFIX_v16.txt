TCTA Source Hotfix v16
======================

Adds the Trans-Caledon Tunnel Authority (TCTA) tenders page as an enabled source in the TenderScrapers framework.

Apply:

Expand-Archive `
  -LiteralPath "C:\Users\Danie le Roux\Downloads\TenderScrapers_TCTA_Source_Hotfix_v16.zip" `
  -DestinationPath "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers" `
  -Force

Then run:

cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" install_source_scrapers.py --overwrite

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

Check:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TCTA\Outputs\tcta_master.xlsx
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers\Outputs\Consolidated\consolidated_daily_review.xlsx
