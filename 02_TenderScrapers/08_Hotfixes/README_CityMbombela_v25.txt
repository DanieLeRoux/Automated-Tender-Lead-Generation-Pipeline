TenderScrapers City Mbombela Robust Pagination Hotfix v25

This patch replaces the City of Mbombela source scraper only.

Fixes:
- Stops pagination gracefully when a later listing page returns HTTP 500.
- Keeps rows parsed from earlier pages instead of discarding the whole run.
- Keeps listing rows even when individual tender detail pages return HTTP 500.
- Preserves rows with missing closing dates in the source master so central consolidation can use eTenders duplicates or manual date handling.

Apply by extracting into TenderScrapers, then run:
install_source_scrapers.py --overwrite
run_all_scrapers.py
validate_outputs.py
