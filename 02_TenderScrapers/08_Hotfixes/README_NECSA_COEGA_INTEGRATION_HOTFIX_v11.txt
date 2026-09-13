TenderScrapers NECSA / COEGA Integration Hotfix v11
===================================================

Purpose
-------
This patch addresses two issues observed after adding NECSA:

1. NECSA source scrape succeeded, but relevant NECSA opportunities were not appearing in the consolidated daily review because the central review eligibility filter was too strict for facility upgrade/design language.
2. COEGA Selenium rendering timed out and returned a failure code, which caused the daily run to report COEGA as failed even though existing COEGA outputs were available from a previous successful run.

Changes
-------
- Adds central source defaults for GTAC and NECSA.
- Adds NECSA-friendly core-scope terms, including facility design, facility upgrade(s), design and upgrade(s), production facility, HVAC system, and cGMP.
- Removes the overly broad standalone 'layout' exclusion from publication/media filtering. It now only excludes graphic/page/print layout contexts.
- Updates the Coega scraper so Selenium timeout/failure is handled gracefully. If the dynamic render fails, the scraper returns no newly parsed rows and exits successfully, preserving previous output files for consolidation.

Expected Result
---------------
- Relevant NECSA professional/facility upgrade opportunities should now appear in the consolidated daily review where they pass the central review eligibility filter.
- NECSA supply/equipment/training-only opportunities should remain excluded from the bid-manager daily review, while remaining available in consolidated_master.xlsx.
- COEGA timeout should no longer break the full daily run.

Application
-----------
Extract this zip into the TenderScrapers root folder and allow overwrite.
Then run:

cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" install_source_scrapers.py --overwrite

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py

& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py
