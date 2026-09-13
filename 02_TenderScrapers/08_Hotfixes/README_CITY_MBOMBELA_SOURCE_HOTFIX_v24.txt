TenderScrapers City of Mbombela Source Hotfix v24
=================================================

This patch adds the City of Mbombela tender source.

Source:
https://www.mbombela.gov.za/downloads/tenders/

New working folder created by install/run:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\CityMbombela

Expected outputs:
CityMbombela\Outputs\city_mbombela_raw.csv
CityMbombela\Outputs\city_mbombela_raw_full.xlsx
CityMbombela\Outputs\city_mbombela_raw_page.html
CityMbombela\Outputs\city_mbombela_master.xlsx
CityMbombela\Outputs\city_mbombela_delta.xlsx
CityMbombela\Outputs\city_mbombela_updates.xlsx
CityMbombela\Outputs\city_mbombela_top_delta.xlsx

Important:
The City of Mbombela website does not always expose closing dates in the HTML. The source scraper therefore preserves records with missing dates so central consolidation can resolve dates from eTenders duplicates where available.

Apply:
Expand this zip into the TenderScrapers root folder with -Force.
Run install_source_scrapers.py --overwrite.
Run run_all_scrapers.py.
Run validate_outputs.py.
