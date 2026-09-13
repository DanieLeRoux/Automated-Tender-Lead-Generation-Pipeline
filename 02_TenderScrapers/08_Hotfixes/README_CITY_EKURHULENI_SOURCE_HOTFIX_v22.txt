TenderScrapers City of Ekurhuleni Source Hotfix v22
===================================================

Adds City of Ekurhuleni as an enabled source scraper.

Apply:

Expand-Archive `
  -LiteralPath "C:\Users\Danie le Roux\Downloads\TenderScrapers_CityEkurhuleni_Source_Hotfix_v22.zip" `
  -DestinationPath "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers" `
  -Force

Then run:

cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" install_source_scrapers.py --overwrite

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

Check:

C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\CityEkurhuleni\Outputs\city_ekurhuleni_master.xlsx

and:

C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers\Outputs\Consolidated\consolidated_daily_review.xlsx
