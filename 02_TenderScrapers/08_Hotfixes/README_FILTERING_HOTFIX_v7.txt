TenderScrapers Filtering Hotfix v7
==================================

This patch refines the daily-review exclusion filter following the v6 output review.

It excludes additional low-relevance tender types from the bid-manager daily review, including:
- Microsoft 365 / Azure / SQL / ICT / software / subscription / platform tenders
- legal research platform and database subscription tenders
- transformers, circuit breakers, substation equipment and other equipment/hardware tenders
- electrical/fire COC and certificate of compliance tenders
- annual report, copyediting, proofreading, layout, printing, publication and corporate communication tenders
- consumables, canteen supplies, supply-and-delivery-only tenders

The patch does NOT delete these tenders from consolidated_master.xlsx. It only excludes them from:
- consolidated_daily_review.xlsx priority/review tabs
- consolidated_top_delta.xlsx

How to apply
------------
1. Extract this ZIP into:
   C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers

2. Allow files to be overwritten.

3. Run:
   cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

   & "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py

   & "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

You do not need to run install_source_scrapers.py because this patch only updates the central consolidation/filtering layer.
