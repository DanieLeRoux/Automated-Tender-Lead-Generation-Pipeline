TenderScrapers NNR Source Hotfix v12
====================================

Adds National Nuclear Regulator (NNR) as an enabled source.

This patch adds:
- config/sources_config.json updated to include NNR
- source_scripts/NNR/nnr_tender_pull_and_clean.py
- metadata/nnr_scraper_metadata.txt

Apply
-----
Expand into:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers

Then run:
cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" install_source_scrapers.py --overwrite
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

Expected working folder:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\NNR

Expected outputs:
NNR\Outputs\nnr_master.xlsx
NNR\Outputs\nnr_delta.xlsx
NNR\Outputs\nnr_updates.xlsx
NNR\Outputs\nnr_top_delta.xlsx

Notes
-----
The current NNR Open Tenders visible during investigation include employee wellness and laptop supply tenders. These should scrape into the NNR master but should not appear in the consolidated bid-manager daily review because the central filtering rules exclude employee wellness, equipment/laptops and supply-only tenders.
