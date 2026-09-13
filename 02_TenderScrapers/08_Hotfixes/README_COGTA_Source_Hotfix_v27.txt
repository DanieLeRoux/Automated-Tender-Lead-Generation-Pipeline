TenderScrapers COGTA Source Hotfix v27
======================================

This patch adds a COGTA source scraper to the TenderScrapers framework.

New source:
COGTA

Working folder:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\COGTA

Expected outputs:
COGTA\Outputs\cogta_raw.csv
COGTA\Outputs\cogta_raw_full.xlsx
COGTA\Outputs\cogta_raw_page.html
COGTA\Outputs\cogta_master.xlsx
COGTA\Outputs\cogta_delta.xlsx
COGTA\Outputs\cogta_updates.xlsx
COGTA\Outputs\cogta_top_delta.xlsx

Apply:
Expand-Archive `
  -LiteralPath "C:\Users\Danie le Roux\Downloads\TenderScrapers_COGTA_Source_Hotfix_v27.zip" `
  -DestinationPath "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers" `
  -Force

Then run:
cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" install_source_scrapers.py --overwrite

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py
