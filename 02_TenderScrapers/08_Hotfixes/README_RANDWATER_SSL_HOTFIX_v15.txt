TenderScrapers Rand Water SSL Hotfix v15
======================================

Purpose
-------
Fixes Rand Water scraper failure caused by SSL certificate verification errors on some Windows/Python installations.

Change
------
- Rand Water scraper first tries normal HTTPS certificate verification.
- If Rand Water returns an SSL certificate-chain error, the scraper retries without certificate verification for this public tender page only.
- If HTTPS still fails, it also attempts HTTP fallback.
- If fetching still fails, it creates empty source outputs instead of breaking the full daily run.

Apply
-----
Extract this patch into the TenderScrapers root folder and run install_source_scrapers.py --overwrite.
