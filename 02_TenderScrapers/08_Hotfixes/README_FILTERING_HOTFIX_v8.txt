TenderScrapers Filtering Hotfix v8
==================================

This patch tightens the bid-manager daily-review filter after review of the post-v7 consolidated output.

It adds two improvements:

1. Additional exclusion themes
   - Security / guarding services
   - Travel management, shuttle and staff transport services
   - SAP / ERP / ICT consulting and system upgrade work
   - Furniture, office and facilities items
   - HR, language, medical and admin support services
   - Legal, audit and financial administration services
   - Waste, refuse, sludge and recycling operational services

2. Core scope gate
   Priority clients no longer make a tender review-eligible by themselves.
   To appear in the bid-manager daily review, a tender must have at least one genuine professional/built-environment scope indicator, such as:
   - professional service provider / PSP
   - consultant / consultancy
   - engineering services
   - project management / programme management / PMO
   - feasibility study / master plan / implementation plan
   - technical advisory / transaction advisory
   - detailed design / construction monitoring / contract administration

The patch does NOT delete excluded tenders from consolidated_master.xlsx. It only excludes them from:
- consolidated_daily_review.xlsx review tabs
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
