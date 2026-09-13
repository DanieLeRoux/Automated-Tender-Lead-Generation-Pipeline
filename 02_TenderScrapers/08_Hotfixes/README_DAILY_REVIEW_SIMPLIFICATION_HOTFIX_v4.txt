TenderScrapers Daily Review Simplification Hotfix v4
====================================================

Purpose
-------
This patch simplifies consolidated_daily_review.xlsx so the priority/review tabs
act as a bid-manager action list instead of a full technical data export.

It does not remove technical fields from consolidated_master.xlsx,
consolidated_delta.xlsx, consolidated_updates.xlsx or the source outputs.

Changes included
----------------
1. Simplifies the following daily review tabs:
   - Very High Priority
   - High Priority
   - Medium Priority
   - Closing Soon
   - New Today
   - Updated Today
   - Needs Manual Date Check

2. Keeps the following compact review columns only:
   - priority_rank
   - score_band
   - weight_score
   - source
   - source_list
   - tender_reference
   - client
   - description
   - location
   - closing_date_time
   - days_to_close
   - document_access_status
   - action_url
   - matched_keywords_all
   - scoring_notes
   - change_type
   - notice_type
   - date_published
   - document_count
   - contact_email

3. Consolidates tender_number and reference_number into:
   - tender_reference

4. Shows only one practical link column:
   - action_url

   The full source_url, source_url_list, document_urls and document_names fields
   remain available in the consolidated master files for audit/debugging.

5. Removes duplicate/low-value fields from the daily review tabs:
   - source_url
   - source_url_list
   - tender_number
   - reference_number
   - title
   - sector
   - country_name
   - region
   - closing_date
   - closing_datetime
   - submission_method
   - raw_status
   - raw_category
   - source_original_weight_score

6. Creates one display location column:
   - Province, South Africa where province is available
   - country_name for World Bank/international rows
   - region where country_name is unavailable
   - South Africa for known South African sources with no province

7. Creates one display deadline column:
   - closing_date_time

8. Keeps summary tabs unchanged:
   - By Source
   - Source Coverage

Files updated
-------------
core/consolidation.py

Application steps
-----------------
Extract this ZIP into:
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers

Allow overwrite when prompted.

Then run:
cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py

Note
----
You do not need to run install_source_scrapers.py for this patch because it only
updates the central consolidation layer.
