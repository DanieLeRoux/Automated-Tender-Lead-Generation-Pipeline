TenderScrapers COGTA Deactivation and COEGA Link Cleanup Hotfix v28
==================================================================

Files changed:
- config/sources_config.json
- source_scripts/Coega/coega_tender_pull_and_clean_v2.py
- metadata/cogta_deactivation_coega_link_cleanup_hotfix_v28_metadata.txt

This hotfix disables COGTA and improves COEGA document-link filtering.

Apply from PowerShell:

Expand-Archive `
  -LiteralPath "C:\Users\Danie le Roux\Downloads\TenderScrapers_COGTA_Deactivation_COEGA_Link_Cleanup_Hotfix_v28.zip" `
  -DestinationPath "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers" `
  -Force

Then run:

cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" install_source_scrapers.py --overwrite

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py
