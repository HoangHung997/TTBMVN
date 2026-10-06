# Quy trinh kiem thu

## Workbook chuan

- Working workbook: `Dutoanmau/00. Du toan TP 3.xlsm`.
- Nguoi dung da cho phep sua truc tiep file nay de test.
- Workbook hien dang mo trong Excel; file khoa: `Dutoanmau/~$00. Du toan TP 3.xlsm`.
- `DT-101` phai tao baseline bat bien va checkpoint truoc khi bat dau thay doi chuc nang.

Mo hinh su dung sau `DT-101`:

- `Baseline`: anh chup bat bien cua workbook chuan tai thoi diem chot.
- `Working`: workbook tich luy cac task da nghiem thu.
- `Checkpoint start/end`: ban sao truoc va sau moi task co tac dong vao Excel.

Khong ghi de baseline. Neu working workbook loi, phuc hoi tu checkpoint gan nhat da dat, khong phuc hoi tu mot file khong ro nguon.

## Build va core tests

Chay tu thu muc goc:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
```

Ket qua toi thieu:

- Build `Release|x64` thanh cong.
- `ExcelAddIn1.Core`, `ExcelAddIn1`, `ExcelAddIn1.Tests`, `ExcelAddIn1.LicenseTool` build thanh cong.
- Tat ca test console dat.
- Khong chap nhan warning moi neu task khong ghi ro ly do.

## Test gate cho moi task

1. **Unit test**: bat buoc cho logic khong phu thuoc Excel.
2. **Integration test**: bat buoc cho persistence, package, resolver va mapping.
3. **Live Excel test**: bat buoc neu task doc/ghi workbook, form, ribbon hoac Excel event.
4. **Regression ho dao**: bat buoc neu sua infrastructure dung chung, Excel context, settings, logging, license hoac workbook event.
5. **Save/reopen**: bat buoc neu task thay doi metadata, cong thuc, format hoac setting workbook.
6. **Evidence**: ghi command, ket qua, workbook/checkpoint va cell so sanh vao `WORKLOG.md`.

## Quy trinh live Excel

1. Xac nhan dung workbook theo full path.
2. Luu workbook va tao checkpoint dau task.
3. Ghi lai sheet dang active, calculation mode va cac input baseline.
4. Thuc hien dung kich ban task, khong sua ngoai pham vi neu khong can.
5. Recalculate workbook.
6. Kiem tra output value, formula, number format, merged cells, named ranges va link sheet.
7. Kiem tra khong co `#REF!`, `#VALUE!`, `#NAME?` moi trong pham vi bi tac dong.
8. Luu, dong, mo lai va lap lai cac kiem tra persistence.
9. Kiem tra `runtime.log` khong co exception bi nuot.
10. Tao checkpoint cuoi task va ghi evidence.

## Baseline nghiep vu da chot tai DT-101

Chi tiet cell, formula, checksum va loi ke thua nam tai `BASELINE.md`.

| Luong | Dau vao can ghi | Dau ra can ghi | Trang thai |
|---|---|---|---|
| VL-NC-M | Gia nhan cong bac 5/7/8, gia ca may va vat lieu mau | F7=450000; F11=495000; F15=517500; cac tong may F23:F72 theo `BASELINE.md` | CHOT |
| DG Can | Khoi luong va he so cua 9 block cong tac | Tong VL/NC/M tai row 11,29,46,63,75,93,111,127,145 | CHOT |
| DG Nuoc | Khoi luong va he so cua 8 block cong tac | Tong VL/NC/M tai row 11,34,57,74,88,103,121,138 | CHOT |
| Gia DT TC | Khoi luong THKL va don gia DG | I27=12540585.6195888; J27=268414820.1; K27=558195620.98722 | CHOT |
| THKP-TC | Tong VL/NC/M va cac ty le chi phi | E26=1198731169.1621242; E27=1198731000 | CHOT |
| Ho dao | Bo min/max mac dinh 3m/5m va cong thuc chop cut | V 3m=0.1286781592..27.1066238629; V 5m=0.3155159274..52.0720420121; moi tin hieu bat buoc co output | CHOT |

`DT-101` khong dat neu bat ky output nao chua chot hoac expected output khong co checkpoint
va checksum de tai lap.

## Dung sai

- Ma, text, role, package ID va cong thuc: phai khop chinh xac.
- Tien va don gia: theo quy tac lam tron cua package/workbook; neu chua chot thi phai ghi ro so chu so.
- Phep tinh double noi bo: dung sai phai duoc quy dinh trong test, khong dung so sanh bang tuyet doi.
- Locale: ky tu thap phan va list separator phai lay tu Excel, khong lay mac dinh tu Windows.

## Gate B1 da nghiem thu

- Build: `Release|x64`, khong warning; 13/13 core tests dat.
- Role regression: 7/7 role; rename `DG Can` van resolve theo `UnitRateLand`.
- ProjectProfile: schema 1, payload 1747 ky tu, checksum
  `49723AFBE82F9E6C582CFC50F48DD8E89D3612222E7291E68050967D25BE03B4`.
- Sheet coordinator: add/rename/activate/delete dat; range va DataGridView chua luu khong doi.
- Variant `DT-105-no-tags.xlsm`: goi y, apply, save/reopen dat.
- Variant `DT-105-missing-sheet.xlsm`: validation va mapping draft deu tu choi.
- Variant `DT-105-all-renamed.xlsm`: 16 sheet doi ten, manual mapping, save/reopen va resolve
  bay role dat.
- Checkpoint `DT-105-end.xlsm`: SHA-256
  `2090595D4BB3F8A4FA59EEF6310F1347437C9C314D9540F7AD1625C8093EEF27`;
  `THKP-TC!E27=1198731000`.

Lenh regression B1:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-worksheet-roles.ps1 -RenameProbe
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-project-profile.ps1 -Reopen
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-sheet-change-coordinator.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-b1-role-mapping.ps1 -Overwrite
```

## Package store DT-202

- Valid install va import lai idempotent.
- Module checksum sai, file thieu va manifest schema cu bi tu choi.
- Cung PackageId/DataVersion nhung checksum khac bi conflict; ban da cai van load dung.
- Inject exception tai `BeforeCommit`: khong co package duoc cai va `.staging` rong.
- Filesystem reparse point bi tu choi; commit khong ghi de thu muc package cu.

## Resolver DT-203

- Bien ngay: 04/11/2021 khong co package; 05/11/2021 chon goi 2021;
  27/10/2025 chon goi 2021; 28/10/2025 chon goi 2025.
- Ho so phe duyet 25/10/2025, danh gia nam 2026: giu goi 2021 theo
  `BQP-TT101-2025-ARTICLE-6` va co source `TT101-2025-BQP`.
- Ho so phe duyet sau moc va ho so chua phe duyet danh gia sau moc: chon goi 2025.
- Gap, overlap, ngay co time component, approval sau evaluation va package Withdrawn
  deu tra decision loi ro rang, khong auto-select.

## Workbook pin DT-204

- ProjectProfile schema 1 thuc te migrate sang schema 2 ma khong tu tao pin.
- Pin test: package ID `TEST-DT103-PACKAGE-NOT-FOR-CALCULATION`, version `1.0.0`,
  checksum `3E4E1F72611295B4961174E58E1A82C7E9C2CFA1E831CDF334E18E86073EA8BB`.
- Default store co them version `2.0.0` nhung verify van nap version `1.0.0`.
- `DT-204-machine-copy.xlsm` voi kho rong tra Missing; voi kho day du tra Available.
- Save/close/reopen: profile payload 1872 ky tu, checksum
  `9FAA6A5EE5116A7A275D2AD74AB6235F9B1E483F6623A9B08C18A1E71702DDA1`;
  role valid va `THKP-TC!E27=1198731000`.

## Gate B2 - Migration DT-205

- Build `Release|x64` sach; 23/23 console tests dat.
- Preview v1 -> v2: 7 thay doi, trong do 6 module tinh toan; PlanId
  `MIG-8B6CF6D083A5F14D0895EF02`.
- Cancel: khong tao backup va profile payload khong doi.
- Inject loi sau `ProfilePinned`: exception duoc ghi log, rollback ve v1 va backup ton tai.
- Apply: backup giu v1; working workbook save/reopen pin v2 checksum
  `6D044E98A906BC5C0FB7A92CCED66075F9924A039106A5210D4EFCB1C961BDA7`.
- Truoc/sau migration deu co 92 o loi ke thua; role valid; `THKP-TC!E27=1198731000`.
- Excel state sau reopen: Calculation Automatic, ScreenUpdating/EnableEvents/DisplayAlerts True.
- B1 regression: no-tags PASS, missing-sheet PASS-REJECTED, all-renamed PASS.

Lenh regression B2:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-package-pin.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-package-migration.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-b1-role-mapping.ps1 -Overwrite
```

## Package BQP 2021 - DT-301

- Source TSV co 198 record: 59 quy trinh, 34 dinh muc, 34 chi phi, 33 ca may,
  35 dia ly va 3 compliance.
- Moi record phai co key unique, data, locator trang va
  `VerifiedAgainstOfficialSource`.
- PDF SHA-256 phai khop `sources.tsv`; bundle phai build lai deterministic va import
  duoc vao package store.
- Review A doi chieu anh PDF; review B doi chieu workbook chuan trong pham vi workbook co data.
- `DG Nuoc` workbook chuan chua co nhom `040.x`; khong duoc dung workbook de loai nhom nay.
- Checkpoint DT-301 start/end phai co role/profile/package Available, 92 o loi cong thuc ke thua
  va `THKP-TC!E27=1198731000`.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-bqp-2021-package.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 -CheckpointPath "Dutoanmau\Checkpoints\DT-301-end.xlsm"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-b1-role-mapping.ps1 -Overwrite
```

## Package BQP 2025 - DT-302

- Source TSV co 248 record: 60 quy trinh, 42 dinh muc, 37 chi phi, 33 ca may,
  69 dia ly va 7 compliance.
- Sau PDF chinh thuc phai khop SHA-256 va page count trong `sources.tsv` va
  `data/regulations/text/2025/text-manifest.tsv`.
- Record diff 2021 -> 2025 phai dung `248 = 50 added + 0 removed + 198 changed`;
  `change-reasons.tsv` phai co dung mot dong verified cho moi change.
- Gia tri gate: 020.0500; tam tinh 8 vung/doi tuong; K2; hai muc 6.300.000;
  Bieu mau 05; M010.001-033; 34 dia ban va vung bien.
- Review anh doc lap dung TT101 trang PDF 35-38, VBHN96 trang 15-19 va VBHN97
  trang 5, 14, 42, 44-46, 52.
- Checkpoint start/end phai co profile schema 2, role/package Available, 92 loi cong thuc
  ke thua va `THKP-TC!E27=1198731000`.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-bqp-2025-package.ps1 -RegenerateSource
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 -CheckpointPath "Dutoanmau\Checkpoints\DT-302-end.xlsm"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-b1-role-mapping.ps1 -Overwrite
```

## Catalog may va gia ca may - DT-303

- Core phai load dung 33 key cho ca M010 va M011, parse strict bang invariant culture.
- Test rieng nam thanh phan, an mon 1,05, thu hoi 30 trieu, loai nhien lieu, gia cho/gia gio.
- Doi chieu workbook `VL-NC-M`: F23, F29, F35, F41, F46, F51, F56, F62, F67, F72.
- M011.024 va tong M011.009/010 co sai khac bang in da ghi trong audit; khong ep engine
  theo tong in sai.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-bqp-2025-package.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-machine-rate-catalog.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 -CheckpointPath "Dutoanmau\Checkpoints\DT-303-end.xlsm"
```

Ket qua chot: 32/32 core dat; 33 may va 10 tong workbook dat; checkpoint semantic dat.

## Catalog dinh muc va chi phi - DT-304

- Core parse strict invariant, hao phi khong lam tron va tien lam tron 1 VND AwayFromZero.
- Kiem tra 34 key dinh muc, bien the VL/NC/M, dieu kien doc/dong chay/tin hieu/bom nuoc.
- Kiem tra K1-K6, minimum, boundary, noi suy K5, VAT input va cac truong hop bi tu choi.
- Doi chieu workbook: 12 o dinh muc `DG Can`, 55 o K5 va 10 o K2 `ChiPhi`.
- Khoa checksum bon PDF BXD 2026 dung de doi chieu bang dong hien hanh.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-bqp-2025-package.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-norm-cost-catalog.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-package-migration.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 -CheckpointPath "Dutoanmau\Checkpoints\DT-304-end.xlsm"
```

Ket qua chot: 38/38 core dat; package 248 record dat; checkpoint semantic dat.

## Independent review va Gate B3 - DT-305

- Auditor Python doc truc tiep source TSV/text-layer/PDF hash, khong goi parser C#.
- Gate bat buoc: 248 unique record, 248 verified change reason, 6 PDF, key set day du,
  locator hop le, khong duplicate/outlier va deterministic bundle 7/7 file.
- Finding `NORM-040.0500` phai co locator trang 39-40; checksum package sau sua la
  `C23A3AE526841473AEAB6E36CCC76B0663A56C5FE63D3351F7AFBFA4639447F2`.
- Test Excel phu phai dung `excel-test-process.ps1` va working copy; khong mo variant trong
  Excel giao dien cua nguoi dung.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-b3-independent-audit.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-bqp-2021-package.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-bqp-2025-package.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-machine-rate-catalog.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-norm-cost-catalog.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-package-pin.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-package-migration.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-b1-role-mapping.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-sheet-change-coordinator.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 -CheckpointPath "Dutoanmau\Checkpoints\DT-305-end.xlsm"
```

Ket qua chot: 6534 independent check, zero error/warning; 38/38 core; full B1/B2/B3
regression va checkpoint semantic dat. Gate B3: `DAT`.

## Project setup va sheet mapping - DT-401

- `ProjectSetupPlanner` la pure Core: validate ngay/profile/mapping va resolve package.
- Form chi giu draft; constructor/preview va nut Huy khong ghi workbook.
- Commit profile + role mapping trong `ExcelWriteContext`; loi sau khi ghi profile phai
  khoi phuc dung payload va role snapshot cu.
- Production output phai chua du hai bundle BQP-RPBM-2021/2025 va bootstrap vao local store.
- Live test chi sua `Dutoanmau/Variants/DT-401-working-copy.xlsm`.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-project-setup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-sheet-change-coordinator.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-project-profile.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 -CheckpointPath "Dutoanmau\Checkpoints\DT-401-end.xlsm"
```

Ket qua chot: 39/39 core; form smoke; cancel no-write; fault rollback; untagged va
doi ten ca 7 sheet; save/reopen; 92 loi ke thua va `THKP-TC!E27=1198731000` deu dat.

## Tra cuu dinh muc - DT-402

- Search exact ma neu ton tai phai chi tra dung ma; ma dung pattern nhung khong ton tai
  tra zero, khong fuzzy sang ma gan giong.
- Fuzzy chuan hoa Unicode/khong dau va typo nho; filter moi truong, do sau, bien the,
  VL/NC/M; ket qua luon kem package identity va source locator.
- Package 2021/2025 co cung ma phai la hai ket qua rieng. Package 2021 chua so hoa hao phi
  chi tiet phai canh bao, khong lay hao phi 2025 gan nguoc.
- UI `NormLookupControl` trong shell `Dutoan` chi doc package da pin va khong ghi workbook.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-norm-lookup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-project-setup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 -CheckpointPath "Dutoanmau\Checkpoints\DT-402-end.xlsm"
```

Ket qua chot: 41/41 core; exact/fuzzy/filter/locale; lookup 2021/2025; UI shell,
save/reopen; 92 loi ke thua va `THKP-TC!E27=1198731000` deu dat.

## Rule engine ChiPhi - DT-403

- CostRule phai duoc nap tu package pin trong ProjectProfile, khong doc he so tu sheet.
- Package TT123/2021 duoc parse theo schema phap ly cua chinh no: K5 toi 2.000 ty,
  khong co muc toi thieu 6,3 trieu va khong gan can cu BXD 2026.
- Package VBHN 2025 bat buoc metadata interpolation/currentExternalBasis va ap minimum K1/K6.
- UI phai hien calculated/applied/source; override chi hop le khi co amount khong am va ly do.
- Live test chi doc/tinh trong app, khong thay doi cong thuc hoac so loi workbook.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-cost-rule-ui.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-norm-lookup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-project-setup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 -CheckpointPath "Dutoanmau\Checkpoints\DT-403-end.xlsm"
```

Ket qua chot: 43/43 core; CostRule 2021/2025; K1-K6 va override; UI shell,
save/reopen; regression DT-401/402; 92 loi ke thua va `THKP-TC!E27=1198731000` deu dat.

## VL-NC-M va PriceProfile - DT-404

- `PriceProfile` phai round-trip deterministic, validate ID/version/date/source/gia/alias va
  luu atomic; cung ID/version khac checksum phai bi tu choi.
- Coverage phai bao missing/kind/unit mismatch, khong zero-fill. Snapshot workbook phai con
  nguyen sau save/reopen va ProjectProfile phai pin dung PriceProfile ID.
- Import legacy phai co 34 entry. Batch apply chi ghi 21 gia vat lieu F81:F101; F7 va F23
  dai dien cong thuc nhan cong/may phai khong doi.
- Doi ten sheet `ResourcePrices` phai van resolve theo role. Gia override phai cap nhat o tra
  lien ket tren `DG Can`, sau do khoi phuc baseline trong working copy.
- UI Bang gia phai hien metadata, checksum, grid, coverage/import/export/save va bao toan
  snapshot khi khong sua. Khong test ghi tren workbook dang mo cua nguoi dung.
- Runtime COM khong duoc `FinalReleaseComObject` RCW dung chung; regression DT-401..403
  bat buoc sau sua ownership.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-price-profile.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-cost-rule-ui.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-norm-lookup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-project-setup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 -CheckpointPath "Dutoanmau\Checkpoints\DT-404-end.xlsm"
```

Ket qua chot: 47/47 core; 34 entry, 39 missing requirement duoc bao, 21 vat lieu batch-write,
formula preserve, UI/save-reopen/role rename/override dat; regression DT-401..403 va semantic
checkpoint giu 92 loi ke thua, `THKP-TC!E27=1198731000`.

## DG Can va DG Nuoc - DT-405

- Don gia phai duoc tinh tu NormCatalog va PriceProfile da pin, khong doc cong thuc legacy
  lam input. Ket qua mot don vi cong tac dung decimal va policy `none-invariant-decimal`.
- Test tren can: `NORM-020.0200/density-1`. Test duoi nuoc: `NORM-030.0100/water-0.5-12`.
  Tong VL/NC/M phai khop cac dong oracle legacy da chot.
- Bat/tat he so luu toc phai chi thay NC/M dung factor. Dieu kien khong duoc ap ngam.
- `M010.DIVING` va ma logic `OR` phai bi chan neu chua binding; binding hop le can ma gia may
  cu the va ly do. Missing price phai chan tinh, khong zero-fill.
- UI phai hien package/profile, source, hao phi, don gia, thanh tien va tong VL/NC/M; doi ten
  hai role sheet, save/reopen va locale vi-VN/de-DE khong lam thay ket qua.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-unit-rate.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-price-profile.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-cost-rule-ui.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-norm-lookup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-project-setup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 -CheckpointPath "Dutoanmau\Checkpoints\DT-405-end.xlsm"
```

Ket qua chot: 50/50 core; can/nuoc baseline, he so, hai binding may, missing price,
UI/role rename/save-reopen dat; regression DT-401..404 va semantic checkpoint giu 92 loi
ke thua, `THKP-TC!E27=1198731000`.

## Gia DT TC - DT-406

- Import phai nhan 13 dong theo worksheet role, khong theo ten sheet; mapping variant sparse
  va binding `M010.DIVING` phai deterministic.
- Dong co khoi luong phai co don gia; dong zero quantity unresolved chi la canh bao.
- Preview giu ca ket qua goc va ket qua co he so. Writer ghi batch F:N, rollback khi verify
  tong sai, cong thuc D/E thay doi phai tu recalculate va subtotal phai gom du row 18/25.
- Oracle phap ly row 22 dung `water-3-12`: `H22=34270.8`, khong ep khop hao phi may 0,014
  cua workbook cu. Tong sau ghi: I27 `12540585.6195888`, J27 `268414820.1`,
  K27 `558860046.12222`; THKP-TC E27 `1199577000`.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-bqp-2025-package.ps1 -RegenerateSource
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-b3-independent-audit.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-estimate-appendix.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-unit-rate.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-price-profile.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-cost-rule-ui.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-norm-lookup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-project-setup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 `
  -CheckpointPath "Dutoanmau\Checkpoints\DT-406-end.xlsm" -ExpectedE27 1199577000
```

Ket qua chot: package 2.0.1/audit 6534 check dat; 52/52 core; UI, batch write,
input recalculate, save/reopen va regression DT-401..405 dat; 92 loi ke thua khong tang.

## THKP-TC va Gate B4 - DT-407

- Import VL/NC/M tu `EstimateAppendix`, cau hinh legacy tu `CostSummary` va package 2.0.1.
- K2 phai dung T; K1/K3/K4/K5/K6 phai dung Z. K5 canonical la 2.598%, khong dung 2.873% cu.
- Writer ghi batch D10:E27/A28, cap nhat tien bang chu, thay input phu luc phai tu recalculate.
- Save/reopen giu E27, formula K5 va khong tang 92 loi Excel ke thua.
- UI/shell phai mo `CostSummaryControl` va hien dung tong/phien ban package.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-cost-summary.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-estimate-appendix.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-unit-rate.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-price-profile.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-cost-rule-ui.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-norm-lookup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-project-setup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 `
  -CheckpointPath "Dutoanmau\Checkpoints\DT-407-end.xlsm" -ExpectedE27 1201557000
```

Ket qua chot: 55/55 core; UI, import, preview, batch write, input recalculate,
save/reopen va regression DT-401..406 dat; 92 loi ke thua khong tang; Gate B4 dat.

## Batch Excel - DT-501

- Moi batch phai dat value parity voi cach ghi cu, rollback duoc khi nem loi giua giao dich
  va phuc hoi ScreenUpdating/Calculation/Events/DisplayAlerts ve dung gia tri truoc khi chay.
- Benchmark cung mot Excel PID va cung 4.000 o; ghi batch phai nhanh hon cell-by-cell va
  khong doi cong thuc/format ngoai vung so huu.
- Duong ho dao, PriceProfile, `Gia DT TC`, `THKP-TC` va ProjectSetup phai qua regression.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\inventory-excel-write-paths.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-excel-batch-write.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-price-profile.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-estimate-appendix.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-cost-summary.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-project-setup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 `
  -CheckpointPath "Dutoanmau\Checkpoints\DT-501-end.xlsm" -ExpectedE27 1201557000
```

Ket qua chot: Release x64 sach, 55/55 core; batch 4.000 o 41 ms so voi 5.899 ms
cell-by-cell, nhanh hon 142,47 lan. Value parity, exception rollback, bon Excel state,
batch path ho dao va cac regression deu dat; 92 loi ke thua khong tang.

## Truy vet can cu - DT-502

- Phu luc phai tao 13 entry dong + 1 block; THKP phai tao 6 entry K + block tong + tien chu.
- Chon `Gia DT TC!K22` phai tra `NORM-030.0300/water-3-12`, package 2.0.1,
  VBHN97 trang 32-33 va cac nguon gia; chon `THKP-TC!E22` phai tra K5 va locator dung.
- Audit phai con sau doi ten sheet/save/reopen; plan phu luc phai tai lai duoc khi F:H da
  la value. UI `Truy vet` phai doc active cell va khong hien metadata noi bo.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-estimate-appendix.ps1 `
  -WorkbookPath .\Dutoanmau\Checkpoints\DT-406-start.xlsm -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-cost-summary.ps1 `
  -WorkbookPath .\Dutoanmau\Variants\DT-406-working-copy.xlsm -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-result-audit.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-price-profile.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-unit-rate.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-cost-rule-ui.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-norm-lookup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-project-setup.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 `
  -CheckpointPath .\Dutoanmau\Checkpoints\DT-502-end.xlsm -ExpectedE27 1201557000
```

Ket qua chot: Release x64 sach, 58/58 Core; 22 audit entry, active-cell/UI,
rename/save/reopen, plan restore va full DT-401..405 regression dat; 92 loi ke thua khong tang.

## Validation va reconciliation - DT-503

- Clean checkpoint phai co 0 loi chan; broken names ke thua duoc gom mot warning.
- Fault matrix bat buoc: missing price -> dung dong, missing role -> dung role,
  changed rate -> dung total, changed formula -> dung cell, broken app name -> dung name.
- Scanner la read-only; moi fault duoc phuc hoi va report cuoi/save-reopen phai bang clean.
- UI phai mo tu shell va dieu huong issue co address den dung active cell.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-validation.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 `
  -CheckpointPath .\Dutoanmau\Checkpoints\DT-503-end.xlsm -ExpectedE27 1201557000
```

Ket qua chot: Release x64 sach, 60/60 Core; 5 fault, navigation, UI, restore va
save/reopen dat. Clean report co 0 loi chan, 1 canh bao tong hop 4.034 name ke thua.

## Chuyen package va rollback - DT-504

- Preview 2021 -> 2025 phai khong ghi workbook, hien 21 thay doi package, 6 module tinh
  thay doi, 13 dong phu luc, 22 audit entry va 7 tong gia tri truoc/sau.
- Gia tri nguon lay tu workbook legacy; package 2021 khong co hao phi chi tiet nen khong
  duoc gia lap kha nang tai tinh 2021. Target 2025 bat buoc tinh du tat ca dong.
- Cancel khong tao backup/khong doi fingerprint. Fault tai `BackupCreated`, `ProfilePinned`,
  `EstimateWritten`, `CostSummaryWritten`, `ValidationCompleted`, `WorkbookSaved` deu tao
  backup; moi fault sau mutation phai rollback payload + range dung fingerprint.
- Apply thanh cong phai tao audit 2025, validation sach, save/reopen va backup khoi phuc
  duoc source 2021. UI co hai tab ket qua/package va chi enable apply khi preview hop le.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-package-migration-results.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-package-migration-ui.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 `
  -CheckpointPath .\Dutoanmau\Checkpoints\DT-504-end.xlsm -ExpectedE27 1201557000
```

Ket qua chot: Release x64 sach, 60/60 Core; migration BQP-RPBM-2021@1.0.0 ->
BQP-RPBM-2025@2.0.1, cancel, 6 fault phase, restore checkpoint, UI va save/reopen dat.
Checkpoint co 22 audit entry, 92 loi ke thua, `THKP-TC!E27=1201557000`.

## Runtime diagnostics va Gate B5 - DT-505

- Log bat buoc co correlation ID 32 hex, UTC, task/phase, workbook token, worksheet role,
  package id/version/checksum, exception type va HResult.
- Fault oracle la `COMException` `0x800A03EC`. Context an toan duoc giu; customer name,
  license key, data address, workbook path/name, sheet name va Machine ID khong duoc xuat hien.
- Support package gom dung 5 file: `app-info.txt`, `workbook-info.txt`,
  `daodat-settings.json`, `runtime.log`, `manifest.sha256`. Setting phai co
  `IncludeWorkbookFields=false`; manifest phai hash du 4 payload.
- Gate B5 chay tuan tu, audit luon dung working copy rieng; checkpoint start/end khong duoc
  bi test rename/save lam thay doi hash.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-runtime-diagnostics.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-b5-gate.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 `
  -CheckpointPath .\Dutoanmau\Checkpoints\DT-505-end.xlsm -ExpectedE27 1201557000
```

Ket qua chot: 8/8 nhom Gate B5 PASS trong 611,1 giay; Release x64, 60/60 Core,
DT-501..504, COM diagnostics/support privacy va checkpoint dat. Batch 4.000 o 33 ms so voi
6.210 ms cell-by-cell, nhanh hon 185,74 lan. Checkpoint 22 audit entry, 92 loi ke thua,
`THKP-TC!E27=1201557000`.

## Goi cap nhat offline - DT-601

- Core bat buoc test package hop le, public-key round-trip, extract/read bundle, tamper,
  wrong signature, downgrade, minimum app version, already-installed va version conflict.
- Evidence production phai duoc verify bang chinh `OfflineUpdateTrustCatalog.Production`;
  public artifact phai khop modulus compile trong Core va workspace khong co private update key.
- Checkpoint workbook khong thay doi vi DT-601 chi them ha tang phat hanh/update.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-offline-update.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 `
  -CheckpointPath .\Dutoanmau\Checkpoints\DT-601-end.xlsm -ExpectedE27 1201557000
```

Ket qua chot: Release x64 sach, 65/65 Core; evidence `BQP-RPBM-2025@2.0.1`
ky boi `TTBMVN-OFFLINE-2026-01`, archive SHA-256
`AC9859DD09DEC61C362A07BA9E703CF67EE1F4D1DEAEFB68C4BC0C9155918FE2`, disposition
`Ready`. Checkpoint package/role dat, 92 loi ke thua va `THKP-TC!E27=1201557000`.

## Update center - DT-602

- Core: install fresh, duplicate, activation rollback, restart, provider Disabled/mat mang
  va offline fallback.
- Live UI: preview diff, install vao store test, preferred song qua restart, shell mo dung
  control va ProjectProfile workbook khong doi.
- Regression: PackageMigration van thay toan bo package, Project Setup dung preferred list.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-update-center.ps1 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-package-migration-ui.ps1 `
  -WorkbookPath .\Dutoanmau\Checkpoints\DT-602-start.xlsm -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-project-setup.ps1 `
  -WorkbookPath .\Dutoanmau\Checkpoints\DT-401-start.xlsm -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 `
  -CheckpointPath .\Dutoanmau\Checkpoints\DT-602-end.xlsm -ExpectedE27 1201557000
```

Ket qua chot: 68/68 Core, install/restart/rollback/offline-only, UI 21 diff va shell dat;
migration 21 package changes/7 result rows dat; Project Setup tren oracle lich su dat.
Checkpoint package/role dat, 92 loi ke thua va `THKP-TC!E27=1201557000`.

## Ma tran tuong thich - DT-603

- Build ca `Any CPU` va `x64`; khong coi build AnyCPU la live Excel x86.
- Inventory tu OS/Office registry va isolated Excel PID.
- Test decimal/group/list separator theo Excel, format number va `FormulaR1C1` multi-arg.
- Moi host khong co may that ghi `NOT_RUN`; pham vi pilot chi gom dong PASS that.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-compatibility-matrix.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 `
  -CheckpointPath .\Dutoanmau\Checkpoints\DT-603-end.xlsm -ExpectedE27 1201557000
```

Ket qua chot tai `COMPATIBILITY.md`; evidence JSON trong `tmp`. Win11/O365 x64 va hai
separator scenario dat; AnyCPU/x64 build 68/68; host con lai `NOT_RUN`.

## Support va chan doan - DT-604

- Shell Du toan phai co nut `Ho tro`; dialog phai nhan workbook ro rang, khong phu thuoc
  ActiveWorkbook thay doi khi form dang mo.
- Goi ho tro gom 8 payload va mot manifest SHA-256. Manifest regulation la ban sealed da
  cai; validation report chi giu role/code/message/remediation da sanitize.
- Cam lo Machine ID/key, workbook/path/sheet/address, project ID, PriceProfile ID va noi dung
  override. Chi token SHA-256 rut gon va aggregate nghiep vu duoc phep xuat.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-runtime-diagnostics.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-estimate-support.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-workbook-checkpoint.ps1 `
  -CheckpointPath .\Dutoanmau\Checkpoints\DT-604-end.xlsm -ExpectedE27 1201557000
```

Ket qua chot: Release x64 sach, 68/68 Core; support package 9 file/8 hash, privacy PASS;
shell/modal/activation va anh UI dat. Checkpoint package/role dat, 92 loi ke thua,
`THKP-TC!E27=1201557000`.

## Acceptance phat hanh - DT-605

- Oracle phai tach theo checkpoint: DT-101 old workbook, DT-405 don gia, DT-406 phu luc,
  DT-407 tong hop va DT-605 cho current validation/migration/update/support.
- Moi script live Excel dung PID cach ly; audit chi ghi working copy. Checkpoint bat bien.
- Lifecycle package dung store trong `tmp`, khong go VSTO dang duoc nguoi dung su dung.
- Evidence JSON phai co dung checkpoint SHA va 14 ket qua PASS. Resume chi hop le khi SHA
  checkpoint khong doi va chi tai su dung prefix PASS lien tuc.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-release-acceptance.ps1 -Overwrite
```

Ket qua chot: `14/14 PASS`, 1044,4 giay; migration 6 fault phase/restore dat; semantic
checkpoint co package/role dat, 92 loi ke thua va `THKP-TC!E27=1201557000`.
Chi tiet: `RELEASE-ACCEPTANCE.md`.

## Gate phat hanh - DT-606

- Build/publish phai lay Code Signing certificate tu Windows store; source/release cam PFX.
- `mage.exe -Verify` deployment va application manifest; `setup.exe` signer phai khop
  `release-info.json`; channel ngoai PILOT bat buoc trust status Valid.
- ZIP verifier doi chieu dung tap file va SHA-256; negative test sua release notes phai bi tu choi.
- Install/uninstall preflight chay `-VerifyOnly` khi Excel nguoi dung con mo; khong mutate
  current registration. Actual deployment duoc ghi `NOT_RUN` thay vi suy dien PASS.

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\publish-release.ps1 `
  -Version 1.0.0 -Overwrite
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-release-gate.ps1 `
  -Version 1.0.0 -Overwrite
```

Ket qua chot: Release x64/68 Core, 37 file, signatures/checksum/tamper/preflight/checkpoint
deu PASS. ZIP SHA-256
`7F3FBE1B65793B130A45A2462F0D16C03C84452095F7FC0AAB92BBFAA977D0F9`.

## Tieu chi dung khan cap

Dung task va ghi `BLOCKED` khi:

- Workbook khong the phuc hoi tu checkpoint.
- Ket qua baseline thay doi ma khong co thay doi nghiep vu da duoc phe duyet.
- Package phap ly khong xac minh duoc nguon.
- Excel state khong duoc khoi phuc sau exception.
- Phat hien nguy co ghi nham workbook ngoai `Dutoanmau`.

## Estimate Workspace - DT-701

Lenh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-estimate-workspace.ps1 -Overwrite
```

Oracle live dung ban sao rieng
`Dutoanmau/Variants/DT-701-estimate-workspace.xlsm` tu checkpoint `DT-606-end.xlsm`.

- 6 dong: 5 cong tac tinh va 1 dong van ban.
- 4 don gia: hai dong land/density-1/KHLNS dung chung; density-2, HLNS va water/binding tach rieng.
- 5 sheet sinh theo nhom dang dung; 5 dong duoc link formula, dong van ban khong bi ghi.
- Writer chay hai lan van 4 don gia/5 sheet, khong tao trung.
- Sheet Phu luc doi ten, save/reopen van load dung workspace va 4 don gia.
- Smoke UI va anh: `tmp/DT-701-estimate-workspace.png`, `tmp/DT-701-rate-catalog.png`.
- Ket qua chot: Release x64 sach, 71/71 Core, live Excel PASS.
