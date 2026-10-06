# DT-305 independent review and B3 gate

Review date: 2026-08-03.

## Independent path

`tools/regulations/audit_bqp_2025.py` reads the UTF-8 TSV and official PDF evidence
directly with Python standard-library code. It does not call the C# package reader,
catalog parsers or calculators used by the application.

The audit checks:

- exact module/type counts and all 248 unique record identities;
- mandatory fields, verified status, source document and valid page range;
- the frozen legal key sets for 34 norms, 33 machines and 37 cost rules;
- norm variant/rate dimensions, duplicate resources and non-negative consumption;
- machine code sequence, required fields, percentage/shift/price bounds;
- K1/K4 arithmetic, K2/K5 thresholds, rate dimensions and current BXD basis;
- 34 unique locality rows, valid zones and fallback membership;
- all six official PDF SHA-256 values and text-layer page counts;
- all 248 unique, verified change-reason rows;
- cited PDF text contains every Norm code and every M010 machine code;
- deterministic rebuild is byte-identical for the seven bundle files.

## Finding and correction

The first independent run found `NORM-040.0500` cited only at page 40 although the
heading and legal code start at page 39. The source generator was corrected to pages
39-40. Consumption, variants, calculation behavior and source document did not change.

The corrected package identity is:

- Package: `C23A3AE526841473AEAB6E36CCC76B0663A56C5FE63D3351F7AFBFA4639447F2`.
- Norm: `2CFF3C031ACE079157333AAE97F60A55B3CDED4A926127AFCC0013FDC3EFD0E9`.
- CostRule: `73FC98145982998E5463E1C7C9F1378FD8BB3800F38B8809A525A5D133346860`.
- MachineRate: `6352542DC1F488FEA3D6B5D05D16E37D80464FEC91C09A0486FA4E7839BE8CBE`.

Final independent result: `PASS`, 6,534 checks, 248 records, 248 change reasons,
six PDFs, zero errors and zero warnings. Machine OCR variants such as `MO 10.023`
were normalized only for source-text matching; catalog codes remain exact.

## Regression

- Release x64: 38/38 core tests passed.
- BQP 2021 and BQP 2025 package validation passed.
- Machine rate: 33 machines and 10 workbook totals passed.
- Norm/cost: six core gates, 12 norm cells, 55 K5 cells and 10 K2 cells passed.
- B1 role variants and sheet-change coordinator passed.
- Pin/migration cancel, rollback, apply, persistence and Excel state passed on isolated
  working copies.
- `DT-305-end.xlsm`: profile schema 2, package Available, roles valid, 92 inherited
  formula errors and `THKP-TC!E27=1198731000`.

## Gate decision

Gate B3 is `PASS`. No serious data error, missing key, duplicate, unexplained outlier or
unverified record remains. The exact-1000-kg K6 gap and logical machine choices remain
explicit validation decisions for the B4 workflow, not hidden assumptions.
