TenderScrapers GTAC Source Hotfix v9
====================================

This patch adds GTAC as a new enabled source in the TenderScrapers framework.

Files included:
- config/sources_config.json
- config/scoring_config.json
- source_scripts/GTAC/gtac_tender_pull_and_clean.py
- metadata/gtac_scraper_metadata.txt

Apply this patch to:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers

PowerShell apply command:
Expand-Archive `
  -LiteralPath "C:\Users\Danie le Roux\Downloads\TenderScrapers_GTAC_Source_Hotfix_v9.zip" `
  -DestinationPath "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers" `
  -Force

Then run:
cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" install_source_scrapers.py --overwrite

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

Expected new folder:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\GTAC

Expected outputs:
GTAC\Outputs\gtac_master.xlsx
GTAC\Outputs\gtac_delta.xlsx
GTAC\Outputs\gtac_updates.xlsx
GTAC\Outputs\gtac_top_delta.xlsx
