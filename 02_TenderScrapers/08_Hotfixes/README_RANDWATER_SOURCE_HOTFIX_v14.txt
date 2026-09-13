TenderScrapers Rand Water Source Hotfix v14
==========================================

This patch adds Rand Water as an enabled source scraper.

Files included:
- config/sources_config.json
- source_scripts/RandWater/randwater_tender_pull_and_clean.py
- metadata/randwater_scraper_metadata.txt

Apply into:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers

Then run:
cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" install_source_scrapers.py --overwrite
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

After running, check:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\RandWater\Outputs\randwater_master.xlsx

Then check:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers\Outputs\Consolidated\consolidated_daily_review.xlsx
