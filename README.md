# SAP ↔ Azhur Supplier Reconciliation

A repeatable monthly finance control that reconciles supplier balances and then drills into open-document composition. It preserves source detail, distinguishes supplier identity from financial/document/accounting matches, and never forces uncertain matches.

## What Sprint 1 discovered

The source inspection is documented fully in [`docs/assumptions.md`](docs/assumptions.md). In short, `4011.xlsx` is an Azhur EUR open-balance report, `4012.xlsx` is the corresponding foreign-currency report (with original and EUR reporting amounts), and `доставчици.xlsx` is a richer SAP vendor line-item population. Azhur's stated work date is 13 August 2026. SAP includes trade payables plus unbilled, deposit, advance, and other accounts, so only configurable trade-payable categories enter the main comparison.

## Architecture

- `src/reconciliation/xlsxio.py` — dependency-free OOXML input/output, useful in locked-down finance environments.
- `src/reconciliation/engine.py` — normalization, scope, matching, controls, reconciliation, exceptions and metrics.
- `config/settings.json` — filenames, currencies, G/L classifications, scope, tolerances and output paths.
- `config/supplier_mapping.csv` — persistent manual mapping template.
- `tests/` — synthetic unit tests; no confidential supplier data.
- `docs/assumptions.md` — empirical source inspection and unresolved accounting assumptions.
- `output/` — generated report and run log (ignored by Git).

## Setup and run

Python 3.10+ is sufficient at runtime; no third-party package is required.

```bash
PYTHONPATH=src python -m reconciliation --config config/settings.json
python -m unittest discover -s tests
```

The configured files can remain at repository root for the development run. For future months, place confidential exports under `input/` and update the paths in `config/settings.json`. Do not commit future input files or generated reports; `.gitignore` excludes them.

## Methodology

The funnel is: **population → supplier identity → supplier balance → currency → document → exception → root cause**.

### Sign and amount basis

Positive canonical balance means an amount owed to the supplier. Source signed EUR/company-code amounts are reversed only after source inspection confirmed negative credits/liabilities. Original-currency nominal values are retained separately and are never directly compared with EUR. `difference = SAP balance − Azhur balance`.

### SAP accounting scope

G/L accounts are classified in configuration as `TRADE_PAYABLE`, `UNBILLED_PAYABLE`, `ADVANCE`, `DEPOSIT`, `OTHER`, or `OUT_OF_SCOPE`. Only configured reconciliation categories enter the primary balance comparison; all SAP lines remain visible and controlled.

### Supplier matching

1. Confirmed manual mapping (highest priority).
2. Exact normalized supplier name.
3. Strong fuzzy name at the configured auto threshold.
4. Lower-scoring fuzzy candidate for manual review only.
5. Unmatched.

Normalization uppercases, Unicode-normalizes, collapses spaces, removes non-material punctuation and conservatively strips a trailing Bulgarian legal form. Original names remain unchanged in normalized output. Because current files have no shared VAT/EIK, the trusted-ID tier is dormant rather than invented.

### Manual mappings

Add a row to `config/supplier_mapping.csv`, set `manual_override` to `true`, and provide both source IDs. Manual rows take precedence on the next run. Keep confidential operational mappings in `config/supplier_mapping.local.csv` or secure storage and point `mapping_file` to that ignored path.

### Documents

Within a mapped supplier, the engine attempts:

- `EXACT_REFERENCE_AMOUNT`: normalized invoice/reference plus amount;
- `STRONG_AMOUNT_CURRENCY_DATE`: amount, currency, and date within the configured window;
- unmatched: exposed without forced pairing.

Invoice normalization removes common formatting separators and safe redundant numeric leading zeroes. Split/residual duplicates are retained and flagged. The architecture supports multiple rows per invoice; no destructive deduplication occurs.

## Tolerances and materiality

Defaults in `config/settings.json` are:

- rounding: EUR 0.01;
- document amount: EUR 0.05;
- supplier balance: EUR 0.05;
- date-assisted match: ±5 calendar days;
- material exception: EUR 1,000;
- fuzzy auto/review scores: 92/80.

Differences below materiality remain visible. Materiality only changes prioritization, not whether a record appears.

## Exception codes

`00_MATCH`, `01_SAP_ONLY`, `02_AZHUR_ONLY`, `03_BALANCE_DIFFERENCE`, `04_DOCUMENT_DIFFERENCE`, `05_CURRENCY_DIFFERENCE`, `06_TIMING_DIFFERENCE`, `07_SUPPLIER_MAPPING_DIFFERENCE`, `08_GL_CLASSIFICATION_DIFFERENCE`, `09_MIGRATION_OR_OPENING_BALANCE`, `10_ROUNDING_OR_FX`, `11_PARTIAL_PAYMENT_OR_RESIDUAL`, and `12_MANUAL_REVIEW` form the stable taxonomy. The engine assigns only conclusions supported by available evidence; unresolved cause remains manual review rather than speculation.

## Output workbook

`output/SAP_Azhur_Reconciliation.xlsx` contains:

0. **Control** — executive totals, populations, coverage and exception metrics.
1. **Supplier Recon** — one row per mapped/unmapped economic supplier with independent Financial, Document and Accounting statuses.
2. **Document Recon** — SAP and Azhur detail side by side.
3. **Exceptions** — unresolved suppliers sorted by absolute financial difference.
4. **Supplier Mapping** — cross-system identity evidence.
5. **SAP Normalized** — full normalized SAP population, including non-trade categories.
6. **Azhur Normalized** — combined 401/1 and 401/2 detail with source lineage.
7. **Controls** — data-quality, metadata and transformation-total checks.

The report uses filters, frozen headers, number/date formats, restrained corporate headers, widths and exception highlighting. `output/run_log.json` records timestamp, inputs/sheets/rows, totals, mapping counts, exception counts and output path.

### Coverage interpretation

Value-weighted reconciliation coverage is:

> absolute SAP balance of financially reconciled mapped suppliers ÷ absolute total in-scope SAP balance.

This avoids overstating coverage through many immaterial suppliers. Supplier-count coverage and document-reconciled value are shown separately.

## Monthly operating procedure

1. Save new source exports outside Git (prefer `input/`).
2. Confirm report dates and SAP selection/scope with Finance.
3. Update input paths and, if necessary, G/L classifications/currencies.
4. Preserve and update confirmed supplier mappings.
5. Run the engine and unit tests.
6. Review failed data-quality controls before accounting exceptions.
7. Investigate the largest Exceptions rows, record Finance comments in the working report, and promote confirmed mappings for the next month.
8. Archive input, config snapshot, run log, and output in the approved confidential finance repository—not public source control.

## Privacy

Accounting workbooks and reports contain confidential information. The initial development files remain tracked because they were supplied in repository history, but all future inputs/outputs are ignored. Never paste real supplier identities into documentation, tickets, or test fixtures; examples and tests use synthetic names only.
