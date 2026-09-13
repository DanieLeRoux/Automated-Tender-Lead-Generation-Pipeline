TenderScrapers Final QA Hotfix v3
=================================

Purpose
-------
This hotfix addresses the remaining issues identified in consolidated_daily_review.xlsx after Document/Action Link Hotfix v2.

Changes
-------
1. Aligns closing_date to closing_datetime in the consolidation layer.
   - If closing_datetime exists, closing_date is forced to the date component of closing_datetime.
   - If only closing_date exists, closing_datetime is backfilled from closing_date.
   - This removes confusing conflicts such as closing_date = 09 Jun 2026 but closing_datetime = 06 Sep 2026.

2. Improves merged duplicate action_url handling.
   - eTenders remains a discovery feed.
   - If a merged record includes eTenders plus a source portal, the consolidated action_url now prefers the non-eTenders source/detail/document URL.
   - If a direct document URL exists, action_url uses the document URL.
   - If no better source exists, eTenders remains the fallback with document_access_status = Manual eTenders lookup required.

Files included
--------------
core/consolidation.py

Apply instructions
------------------
Extract this zip directly into:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers

Allow overwrite when prompted.

Then run:
cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

No install_source_scrapers.py step is required for this hotfix, because this patch only changes the central consolidation layer.
