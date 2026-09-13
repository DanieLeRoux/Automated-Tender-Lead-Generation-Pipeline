TenderScrapers eTenders Document Link Hotfix v5
================================================

Purpose
-------
Adds direct eTenders document-link extraction to the eTenders source scraper.

Confirmed behaviour
-------------------
The eTenders PaginatedTenderOpportunities API response already includes a supportDocument array for each tender. Each supportDocument record includes supportDocumentID, fileName, extension and tendersID. The public document download URL can be constructed using:

https://www.etenders.gov.za/home/Download/?blobName=<supportDocumentID><extension>&downloadedFileName=<URL-encoded fileName>

Example:
https://www.etenders.gov.za/home/Download/?blobName=ea85ea7f-d086-4ea6-ba16-3f2d1f9497b8.pdf&downloadedFileName=Terms%20of%20Reference%20for%20the%20review%20of%20the%20EC%20EIP%204th%20Edition%20-%20signed.pdf

Files changed
-------------
source_scripts/eTenders/etenders_pull_and_clean_weighted_v2.py

New/updated eTenders output columns
-----------------------------------
document_names
document_urls
document_count

Expected consolidation impact
-----------------------------
The central consolidation layer already uses document_urls to populate action_url and document_access_status. After this patch, eTenders rows with support documents should show:

document_access_status = Direct document available
action_url = first available eTenders document download link

Application steps
-----------------
1. Extract this hotfix into the TenderScrapers root folder and allow overwrite.
2. Run install_source_scrapers.py --overwrite so the updated eTenders source script is copied into the active eTenders working folder.
3. Run run_all_scrapers.py.
4. Run validate_outputs.py.
5. Open consolidated_daily_review.xlsx and check eTenders rows in the priority tabs.

Notes
-----
The source-level eTenders scoring remains unchanged. The new document fields are not included in the eTenders row_hash to avoid artificially marking existing tenders as UPDATED solely because document links were added by this patch. The master file will still backfill document fields because incoming current rows replace older master rows during the master update step.
