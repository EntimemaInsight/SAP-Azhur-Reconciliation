# Source inspection and accounting assumptions

## Inspection completed before implementation

All workbooks were inspected at OOXML cell level (including dimensions, types, style-backed dates, merged cells, blank rows, and embedded totals) before reconciliation logic was implemented.

| File | Sheet / range | Header | Detail | Interpretation |
|---|---|---:|---:|---|
| `4011.xlsx` | `Sheet1`, A1:F518 | row 10 | 508 | Azhur open supplier balances for account **401/1 — suppliers in EUR**. Columns are counterparty ID/name, invoice, invoice date, due date and signed EUR balance. |
| `4012.xlsx` | `Sheet1`, A1:I186 | row 10 | 176 | Azhur open supplier balances for account **401/2 — suppliers in foreign currency**. It includes original currency/code and amounts plus EUR reporting balance. |
| `доставчици.xlsx` | `Data`, A1:R932 | row 1 | 930 | SAP vendor line-item/accounting population. Row 2 is an embedded grand total and is excluded. It contains vendor, G/L, two amount measures, currency, H/S indicator, document date, reference, document, text, assignment and offsetting G/L. |

There are no merged ranges or blank detail rows. Azhur has nine report/title rows before the header. SAP has one embedded total row. SAP dates are Excel serial dates; Azhur dates are `DD.MM.YYYY` text. Amounts are signed: credits/liabilities are normally negative and SAP `H` records are negative, while `S` records are positive. Currencies observed are EUR, USD, BGN and HRK. Azhur is explicitly a “with balance” open-item view at 13 August 2026. SAP is a line-item population spanning fiscal periods and includes opening-balance text; the SAP reporting cut-off is not explicitly labelled.

## Canonical sign convention

**Positive canonical amount means a liability owed to a supplier.** Both extracts' company/reporting-currency signed amounts are multiplied by negative one because source credits/liabilities are negative. This is empirically supported by Azhur balances and SAP H/S signs. Raw, canonical, and in-scope totals are retained as controls.

## Scope

Only SAP 4011/4012/4013/4014 accounts are initially reconciled to the combined Azhur 401/1 and 401/2 population. SAP 4015 is `UNBILLED_PAYABLE`; 4921 is `DEPOSIT`; 4221 is `ADVANCE`; 4951 and 4983 are `OTHER`. They remain visible in SAP Normalized but are not mixed with trade payable reconciliation. Account scope is configuration, not code.

## Unresolved assumptions

1. **Cut-off:** Azhur states 13 August 2026; SAP does not state a cut-off. Timing differences cannot be conclusively identified without owner confirmation.
2. **Open item versus movement:** Azhur is an open-balance report. SAP appears to be a line-item/opening migration population but the export selection criteria are unavailable.
3. **Migration/opening balances:** multiple SAP lines say “НАЧАЛНО САЛДО”; the engine exposes them but does not automatically classify every unmatched line as migration.
4. **4013/4014 comparability:** these group/related-party trade payable accounts are included based on their account descriptions, but Finance should confirm whether Azhur 401 populations contain them.
5. **Foreign currency:** Azhur `Салдо` is treated as EUR reporting balance and `Валута` as original-currency balance. SAP company-code currency is EUR. HRK legacy balances and zero original values require review.
6. **Partial payments:** the design preserves split rows and does not force them. Sprint 1 flags unmatched residual compositions for review; more advanced subset matching can follow confirmed accounting rules.
7. **Identifiers:** no common VAT/EIK field is present, so matching relies on manual overrides, normalized exact names, and conservative fuzzy names.
