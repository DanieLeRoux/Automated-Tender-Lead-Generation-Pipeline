# TenderScrapers Full Implementation v1

This package contains the complete practical tender-scraper implementation for the current Henry Fagan tender-monitoring workflow.

It includes:

- the working source scraper scripts developed so far;
- a central scoring engine;
- a standard tender schema;
- a consolidation engine;
- a daily runner;
- installer/validation helpers; and
- metadata files for the source scrapers.

## Default root

The implementation assumes the existing Data Analytics root:

```text
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics
```

The framework folder should be installed here:

```text
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers
```

## Included active sources

- eTenders
- World Bank
- City of Cape Town
- DBSA
- Vaal Central Water Board / VCWB
- Transnet eSupplier
- ACSA
- COEGA
- Department of Human Settlements
- uMngeni-uThukela Water

Sibanye-Stillwater is included as a disabled source because its site blocked automated access through Incapsula/Imperva during testing.

## Folder structure

```text
TenderScrapers
│
├── core
│   ├── standard_schema.py
│   ├── scoring_engine.py
│   ├── consolidation.py
│   ├── excel_utils.py
│   ├── date_utils.py
│   └── text_utils.py
│
├── config
│   ├── sources_config.json
│   └── scoring_config.json
│
├── source_scripts
│   ├── eTenders
│   ├── WorldBank
│   ├── CityOfCapeTown
│   ├── DBSA
│   ├── VaalCentralWater
│   ├── Transnet
│   ├── ACSA
│   ├── Coega
│   ├── DHS
│   └── uMngeniUthukelaWater
│
├── metadata
├── run_all_scrapers.py
├── run_consolidation.py
├── install_source_scrapers.py
├── validate_outputs.py
├── daily_run.ps1
└── requirements.txt
```

## One-time setup

Unzip the package to:

```text
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers
```

Install Python dependencies:

```powershell
cd "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" -m pip install -r requirements.txt
```

Install/update the packaged source scrapers into their expected folders:

```powershell
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" install_source_scrapers.py --overwrite
```

## Daily run

Run all configured source scrapers, then consolidate:

```powershell
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_all_scrapers.py
```

Or use the PowerShell helper:

```powershell
.\daily_run.ps1
```

## Consolidation only

If the source scrapers have already run, rebuild the consolidated outputs only:

```powershell
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" run_consolidation.py
```

## Validate outputs

Check which expected source output files exist:

```powershell
& "C:/Users/Danie le Roux/AppData/Local/Programs/Python/Python314/python.exe" validate_outputs.py
```

## Consolidated output location

```text
C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers\Outputs\Consolidated
```

Generated files:

```text
consolidated_master.xlsx
consolidated_delta.xlsx
consolidated_updates.xlsx
consolidated_top_delta.xlsx
consolidated_daily_review.xlsx
source_top_delta_standardised.xlsx
```

`consolidated_daily_review.xlsx` includes:

```text
Very High Priority
High Priority
Medium Priority
Closing Soon
New Today
Updated Today
By Source
```

## Central scoring model

The central scoring engine recalculates `weight_score` across all sources and ignores source-specific scores for consolidated ranking.

| Criteria | Score |
|---|---:|
| Priority client/source match | +10 |
| Strategic sector keyword | +5 each |
| Technical keyword | +4 each |
| Service keyword | +4 each |
| Procurement/delivery keyword | +3 each |
| Priority notice type | +2 |
| Closing within 14 days | +2 |
| Has downloadable documents | +1 |
| Has contact email | +1 |

Service keywords are weighted the same as technical keywords because project/programme management is a core service line.

## Score bands

| Score | Band |
|---:|---|
| 20+ | Very High |
| 12–19 | High |
| 6–11 | Medium |
| 1–5 | Low |
| 0 | No Match |

## Important implementation note

The package is complete for operational use, but it deliberately retains the proven source-level scrapers as source scripts. The central framework then standardises, rescoring and consolidates the outputs. This avoids breaking working parsers while still achieving one comparable scoring and review layer across sources.
