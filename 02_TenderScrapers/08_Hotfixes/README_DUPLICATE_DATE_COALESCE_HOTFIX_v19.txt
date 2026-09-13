TenderScrapers Duplicate Date Coalescing Hotfix v19
==================================================

This patch fixes the issue where a duplicated tender captured from both eTenders and a source portal, such as TCTA, could be placed in Needs Manual Date Check because the eTenders row had a blank closing date and overwrote/controlled the merged record.

Apply with:

Expand-Archive `
  -LiteralPath "C:\Users\Danie le Roux\Downloads\TenderScrapers_Duplicate_Date_Coalesce_Hotfix_v19.zip" `
  -DestinationPath "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers" `
  -Force

Then run:

cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

No install_source_scrapers.py run is required because this patch only changes the central consolidation layer.
