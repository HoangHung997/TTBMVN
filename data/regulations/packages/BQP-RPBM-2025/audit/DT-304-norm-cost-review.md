# DT-304 norm and cost review

Review date: 2026-08-03.

## Scope

- Norm source: VBHN97-2025-BQP Appendix I, PDF pages 18-42.
- BQP cost source: VBHN97-2025-BQP Appendix II, PDF pages 43-52.
- Current K2 cross-check: TT36/2026/TT-BXD Appendix III, table 3.7, PDF page 8.
- Current K5 cross-check: TT38/2026/TT-BXD Appendix VIII, table 2.24, PDF page 47.
- Workbook oracle: `Dutoanmau/00. Du toan TP 3.xlsm` opened in Excel.

## Norm contract

All 34 `NormCatalog` records use the same strict invariant schema:

- Input: norm key, variant code, work quantity and an explicit set of conditions.
- Data: `variantCodes`, `rates`, optional `adjustments`, optional `constraints`.
- Output: material, labor and machine consumption grouped by resource code and unit.
- Rounding: none; the engine preserves `decimal` consumption and scales by work quantity.
- Locator: every definition retains its VBHN97 document/page/section locator.

Covered conditions:

- labor factor 1.10 for slope over 25 degrees;
- additional QNCN 8/10 labor for UXO signals;
- additional pumping shift for excavation in water;
- labor and machine factors 1.10, 1.25 and 1.50 for three current-speed bands;
- reject current speed over 2 m/s;
- require the legal project condition for depth 0.5 m or 1 m in `NORM-020.0500`;
- mark drilling and deep-water cases that require a separately approved estimate.

Logical resources `M010.DIVING`, `M010.002-OR-M010.003` and
`M010.018-OR-M010.028` intentionally preserve choices stated by the norm. The unit-rate
workflow must resolve the exact machine from depth and approved construction method; it must
not silently select one machine.

## Cost contract

- Direct cost: `T = VL + NC + M`.
- Common cost: `C = 40% * NC`.
- K1/K4: eight terrain rows; K1 minimum 6,300,000 VND for area up to 4 ha; K4 remains
  independently calculated.
- K2: upper-inclusive step table on direct cost, matching current TT36 table 3.7.
- K3: three Z brackets, clamped to 2,000,000-60,000,000 VND.
- K5: linear interpolation, first rate for values up to 10 billion VND, separate estimate
  above 10,000 billion VND, matching current TT38 table 2.24.
- K6: 5% below 1,000 kg, 3% above 1,000 kg, minimum 6,300,000 VND up to 4 ha.
  The source has no rule at exactly 1,000 kg, so the engine rejects that value pending an
  approved decision instead of guessing.
- VAT: rate is an external input from current tax law; template 04 excludes K3 and K4 from
  the taxable base.
- Money rounding: nearest 1 VND, `MidpointRounding.AwayFromZero`.
- Every result exposes the BQP source locator; K2/K5 tables also expose
  `currentExternalBasis` for TT36/TT38.

## Verification

- Core: 38/38 tests passed, including six Norm/Cost gates and vi-VN/de-DE locale tests.
- Package: 248 records, package checksum
  `C23A3AE526841473AEAB6E36CCC76B0663A56C5FE63D3351F7AFBFA4639447F2`.
- Norm module checksum:
  `2CFF3C031ACE079157333AAE97F60A55B3CDED4A926127AFCC0013FDC3EFD0E9`.
- CostRule module checksum:
  `73FC98145982998E5463E1C7C9F1378FD8BB3800F38B8809A525A5D133346860`.
- Workbook: 12 norm cells in `DG Can`, 55 K5 table cells and 10 K2 table cells in
  `ChiPhi` matched.
- Four BXD 2026 PDF checksums passed in `test-norm-cost-catalog.ps1`.

## Workbook discrepancies

- `ChiPhi!A3` still labels K5 as TT12/2021, which TT38/2026 replaced from 2026-07-01.
- `ChiPhi!P7`/`P13` extrapolate between the 10 and 20 billion VND rows when Z is below
  10 billion VND. The current legal rule uses the first table rate for that range.
- Existing `#REF!`/`#N/A` cells recorded in `docs/du-toan/BASELINE.md` remain inherited
  workbook defects. DT-304 did not rewrite workbook formulas.

Decision: legal text is authoritative. The workbook is a regression oracle only where its
table/formula agrees with the effective legal source.

DT-305 independent review corrected the source locator for `NORM-040.0500` from page 40
to pages 39-40. The consumption data did not change; the package and Norm checksums above
include the corrected locator.
