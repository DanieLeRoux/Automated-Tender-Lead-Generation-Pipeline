# Automated Tender Lead Generation Pipeline

## Overview

This repository implements a registry-driven tender-monitoring pipeline for collecting procurement notices from multiple public-sector and development-finance sources, standardising them, identifying new or changed opportunities, and producing a ranked review pack.

The business problem is information fragmentation. Tender opportunities are published on separate portals with different page structures, field names, date formats, document-link conventions, and update behaviour. Manual monitoring makes it difficult to compare opportunities consistently and increases the risk of overlooking relevant work. The pipeline reduces that manual effort by applying the same schema, validation, scoring, deduplication, and review filters after source-specific collection.

The scoring layer is a transparent rule-based classification mechanism. It is not a predictive machine-learning model and does not claim to estimate win probability.

## Architecture

The operational flow is:

```text
Source registry
	|
	v
Active source scraper processes
	|
	v
Per-source raw data and workbooks
	|
	v
Standardisation, date handling, scoring, deduplication and change detection
	|
	v
Central consolidation
	|
	+--> Master, delta, updates and top-delta workbooks
	+--> Daily review workbook
	+--> Validation reports and run logs
```

`02_TenderScrapers/run_all_scrapers.py` reads `00_Config/sources_registry.xlsx`, skips inactive sources, runs each configured script as a child Python process, records the return code and elapsed time, and then invokes the consolidation stage unless disabled. Each child process writes its own process log. The runner writes a timestamped CSV summary and, when supported by the installed Excel dependencies, an XLSX summary.

`02_TenderScrapers/03_Consolidation/run_consolidation.py` calls the shared consolidation engine. The engine loads source outputs, applies the central scoring configuration, merges duplicate records, recalculates current status, applies review eligibility rules, and writes the final workbooks.

## Repository structure

```text
.
|-- README.md
`-- 02_TenderScrapers
    |-- 00_Config
    |   |-- settings.json
    |   |-- scoring_config.json
    |   `-- sources_registry.xlsx
    |-- 01_Core/core
    |   |-- consolidation.py
    |   |-- date_utils.py
    |   |-- excel_utils.py
    |   |-- scoring_engine.py
    |   |-- standard_schema.py
    |   `-- text_utils.py
    |-- 02_Sources
    |   `-- <source>/00_Scripts, 01_RawData, 02_Outputs, 03_Logs, 04_State, 05_Metadata
    |-- 03_Consolidation/run_consolidation.py
    |-- 04_FinalOutputs
    |-- 05_SystemLogs
    |   |-- DailyRun
    |   |-- RunLogs
    |   `-- Validation
    |-- 06_SystemState/Awarded
    |-- 07_SystemMetadata
    |-- 08_Hotfixes
    |-- 09_Tests
    |-- 99_Archive
    |-- daily_run.ps1
    |-- requirements.txt
    |-- run_all_scrapers.py
    `-- validate_outputs.py
```

The source registry currently marks these adapters active: ACSA, City of Ekurhuleni, City of Mbombela, City of Cape Town, City of Polokwane, Coega, COGTA, DBSA, DHS, eTenders, Johannesburg Water, NamWater, NECSA, NNR, Rand Water, TCTA, Transnet, uMngeni-uThukela Water, Vaal Central Water, and World Bank. GTAC and Sibanye Stillwater are present but marked inactive in the registry; the runner therefore does not execute them by default. The registry, rather than this list, is the source of truth for execution.

## Data collection

Each source adapter lives under `02_Sources/<source>/00_Scripts` and is responsible for the source-specific acquisition and parsing problem. The implementations use HTTP requests and, depending on the source, HTML parsing, tabular downloads, JSON-like responses, or source workbooks. They extract tender identifiers, titles, descriptions, client and location fields, publication and closing dates, notice types, status, documents, contacts, submission information, and source links where those fields are available.

Source adapters write source-level files into the configured source folders. The common framework does not assume that every source provides every field. Missing values remain explicit and are handled by the standardisation and validation stages. World Bank execution can optionally request non-English translation; the optional translation packages are listed in `requirements.txt`.

## Cleaning and standardisation

The shared schema in `01_Core/core/standard_schema.py` defines the comparable tender record. It includes source provenance, identifiers, notice content, dates, status, documents and contact details, change-tracking fields, scoring fields, and review-filter fields.

Cleaning and standardisation include:

- whitespace cleanup and case/character normalisation for matching;
- consistent handling of empty values;
- preservation of original and translated descriptions when supplied by a source;
- normalisation of source-specific fields into the standard column names;
- parsing of ISO, day-first, native datetime, and Excel serial dates, including fractional Excel dates;
- calculation of `is_current` while respecting archived and cancelled records; and
- Excel output formatting with tables, frozen headers, stable column widths, and URL strings kept as text.

Records with missing closing dates are not silently treated as current. They are retained for auditability and surfaced in the daily review's manual date-check sheet when appropriate.

## Deduplication and change detection

The consolidation engine constructs an opportunity key from available source identifiers, reference numbers, URLs, and tender content. It also creates a row hash for change tracking. Exact duplicate candidates are grouped using normalised identifiers and stable document/source information. Duplicate groups retain provenance through fields such as `source_list`, `source_url_list`, `duplicate_group_id`, and `duplicate_count`.

When duplicate rows provide different closing dates, the engine coalesces the date fields and selects the best valid date according to the implemented current/future-date rules. It also fills missing scalar fields from other rows in the same duplicate group. The resulting master, delta, updates, and top-delta datasets are deduplicated before final output.

Delta and update datasets are source-level change feeds that are consolidated and rescored centrally. The central `weight_score` is used for consolidated ranking rather than trusting source-specific scores.

## Classification and scoring

`01_Core/core/scoring_engine.py` searches a combined text representation of the title, descriptions, client, sector, category, notice type, conditions, and briefing details. It records matched keyword groups and a transparent scoring note for each row.

The weights are configured in `00_Config/scoring_config.json`:

| Signal | Weight |
| --- | ---: |
| Priority client or source match | +10 |
| Strategic-sector keyword group | +5 per matched group |
| Technical keyword group | +4 per matched group |
| Service keyword group | +4 per matched group |
| Procurement or delivery keyword group | +3 per matched group |
| Priority notice type | +2 |
| Closing within the configured 14-day window | +2 |
| Has documents | +1 |
| Has a contact email | +1 |

Scores are assigned to these configured bands: Very High (20+), High (12-19), Medium (6-11), Low (1-5), and No Match (0). The keyword dictionaries cover strategic sectors, technical disciplines, professional services, procurement/delivery models, priority clients, and notice types. Configuration changes therefore change classification behaviour without changing the scoring engine.

## Filtering and review eligibility

The consolidation layer applies a core-scope gate and a central exclusion filter to the review outputs. The configured exclusion themes include operational or supply-only categories such as repairs, routine maintenance, cleaning, training, software or subscriptions, equipment, consumables, licences, publications, landscaping, and other non-core themes defined in `scoring_config.json`.

Excluded records are not deleted from `consolidated_master.xlsx`. They remain available for traceability, with `excluded_by_filter`, `exclusion_reason`, and matched exclusion fields populated. They are omitted from the operational priority tabs and the consolidated top-delta shortlist. Current, review-eligible records with missing closing dates are separated into `Needs Manual Date Check` rather than being ranked as closing-soon opportunities.

## Outputs and validation

The consolidation stage writes the following files to `02_TenderScrapers/04_FinalOutputs`:

- `consolidated_master.xlsx`: the deduplicated consolidated register;
- `consolidated_delta.xlsx`: current records classified as new or changed by the source feeds;
- `consolidated_updates.xlsx`: consolidated update records;
- `consolidated_top_delta.xlsx`: the current, eligible top-delta shortlist, limited by the engine's configured top-50 selection;
- `consolidated_daily_review.xlsx`: multi-sheet operational review output; and
- `source_top_delta_standardised.xlsx`: standardised source top-delta records.

The daily review workbook contains priority bands, closing-soon, new-today, updated-today, manual-date-check, by-source, source coverage, and QA-oriented sheets as produced by the current engine. The exact row counts are data-dependent.

`validate_outputs.py` reads the source registry and checks active source folders, standard subfolders, configured scripts, required output paths, and optional logs, state, and metadata paths. It writes timestamped validation reports to `02_TenderScrapers/05_SystemLogs/Validation`. Validation checks file and folder presence; it is not a content-quality or business-approval test.

## Installation

Requirements are Python 3.10 or newer with the packages in `02_TenderScrapers/requirements.txt`. The package uses pandas, Requests, Beautiful Soup, lxml/html5lib, openpyxl, XlsxWriter, date utilities, and optional Argos/langdetect support for World Bank translation.

From the repository root, create and activate an isolated environment, then install the dependencies:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r .\02_TenderScrapers\requirements.txt
```

The source adapters access external sites and therefore require network access and any source-specific access conditions imposed by those sites. No credentials are documented here. If a source requires credentials or private configuration in a deployment environment, provide them through that environment's approved secret-management process rather than committing them to this repository.

## Configuration

Before running, review the following files without committing confidential values:

- `00_Config/sources_registry.xlsx`: source names, stable keys, active flags, folder names, script paths, output paths, and optional log/state/metadata paths;
- `00_Config/scoring_config.json`: scoring weights, keyword groups, score bands, closing-soon window, priority entities, and review exclusions; and
- `00_Config/settings.json`: framework folder names and deployment settings.

Paths in deployment-specific settings should be adapted to the local checkout. The README intentionally uses repository-relative paths and does not reproduce local absolute paths or workbook contents.

## Execution

Run all active sources, consolidate, and produce the run summary:

```powershell
python .\02_TenderScrapers\run_all_scrapers.py
```

Useful runner options include:

```powershell
python .\02_TenderScrapers\run_all_scrapers.py --skip-consolidation
python .\02_TenderScrapers\run_all_scrapers.py --stop-on-failure
python .\02_TenderScrapers\run_all_scrapers.py --translate-worldbank
python .\02_TenderScrapers\run_all_scrapers.py --project-root <path-to-02_TenderScrapers>
```

To rebuild only the consolidated files after source outputs already exist:

```powershell
python .\02_TenderScrapers\03_Consolidation\run_consolidation.py
```

The consolidation command also supports `--project-root`, `--sources-registry`, `--scoring-config`, and `--output-root`. To run the PowerShell wrapper, use:

```powershell
.\02_TenderScrapers\daily_run.ps1
```

The wrapper resolves `py`, `python`, or an explicitly supplied Python command, records a transcript under `05_SystemLogs/DailyRun`, and forwards options for World Bank translation, consolidation, and failure handling. Validate the resulting source structure and configured output files with:

```powershell
python .\02_TenderScrapers\validate_outputs.py
```

## Reproducibility and limitations

Results depend on the live source sites, source availability, source HTML or download formats, network conditions, the registry's active flags, and the version of the scoring configuration. A failed child scraper is recorded by the runner and causes a non-zero overall result after the remaining configured work completes, unless `--stop-on-failure` is selected. A missing consolidation script is reported as skipped by the runner, although this repository currently contains `03_Consolidation/run_consolidation.py`.

The repository contains no automated tests under `09_Tests`; validation is currently implemented as the executable output-structure checker described above. Source-specific parsing and the rule-based classification should therefore be reviewed against representative source outputs whenever a portal changes.