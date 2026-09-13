TenderScrapers Filtering Hotfix v6
====================================

Purpose
-------
Adds the bid-manager daily-review exclusion filter requested after reviewing the highlighted eTenders top-delta sample.

What changes
------------
1. Keeps all tenders in consolidated_master.xlsx for audit/history.
2. Adds these fields to the consolidated schema:
   - review_eligible
   - excluded_by_filter
   - exclusion_reason
   - excluded_keyword_matches
3. Excludes non-core tenders from the consolidated_daily_review priority tabs and consolidated_top_delta.xlsx.
4. Updates the priority-client scoring list to include the new priority clients provided.

Excluded themes
---------------
Repairs, Maintenance, Financial Advisory, Employee Wellness, Labour, Cleaning, Software, Training, Equipment, Appliances, AI / Artificial Intelligence, Document Storage, Blasting, Vehicle, CIDB below 7, Licences / Licenses, Digital Media, Invasive Plants / Vegetation, Landscaping / Gardening, Auctioning, Newsletters, Grave Preparations, Animals, and Towing Services.

Important
---------
This patch does not remove records from the master register. It only removes them from the bid-manager daily review and consolidated top-delta shortlist.

How to apply
------------
Extract this zip into:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers

Allow overwrite.

Then run:
cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

No install_source_scrapers.py run is required for this patch because it updates only the central consolidation/scoring configuration layer.
