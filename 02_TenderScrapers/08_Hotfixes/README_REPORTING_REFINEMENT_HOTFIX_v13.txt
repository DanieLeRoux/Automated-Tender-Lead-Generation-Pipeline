TenderScrapers Reporting Refinement Hotfix v13
=============================================

Purpose
-------
Refines the reporting tabs in consolidated_daily_review.xlsx so that source reporting is clearer.

Changes
-------
1. Refines the By Source tab so it now shows all configured scrapers, including sources with zero daily-review rows.
2. Adds source-level status/count fields such as source_master_rows, source_delta_rows, source_updates_rows, source_top_delta_rows, consolidated_master_rows, review_eligible_current_rows, excluded_rows, daily_review_rows, closing_soon_rows, new_today_rows, updated_today_rows and manual_date_check_rows.
3. Refines Source Coverage so it clearly represents sources appearing in the bid-manager review list, including primary-source rows and duplicate/source-list coverage.
4. Adds a Scraper Run Summary tab as an explicit all-scraper operational summary.
5. Adds NNR source defaults and SA-source location handling in the consolidation layer.

Notes
-----
This patch changes the central consolidation/reporting layer only. It does not change any source scraper parser.

Apply
-----
Expand this zip into:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers

Then run:
cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

No install_source_scrapers.py run is required for this patch.
