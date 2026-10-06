# Bien ban acceptance phat hanh DT-605

Ngay chay: 2026-08-03. Trang thai: `PASS`.

## Pham vi

- Build Release x64 va 68 Core test.
- Clean install package `BQP-RPBM-2025@2.0.0` trong store cach ly.
- Upgrade bang goi ky `.ttbupdate` len `2.0.1`, giu hai version va activation qua restart.
- Go store cach ly va reinstall `2.0.1` tu goi ky.
- Mo workbook DT-101 chua co role/profile; form first-run de xuat du 7 role va khong ghi ngam.
- Regression batch write, don gia, phu luc, tong hop, audit, validation va migration.
- Update Center, runtime/support package, support UI va semantic checkpoint.

Vong doi o day la package/application-data lifecycle. Cai/go cai VSTO that, publish manifest
va bo cai phat hanh thuoc DT-606; acceptance khong go add-in dang duoc nguoi dung su dung.

## Oracle

| Nhom | Workbook |
|---|---|
| Old workbook | `DT-101-end.xlsm` |
| Don gia | `DT-405-start.xlsm` |
| Phu luc | `DT-406-start.xlsm` |
| Tong hop | `DT-407-start.xlsm` |
| Validation/migration/update/support | `DT-605-start.xlsm` |

Tach oracle theo checkpoint la bat buoc: ket qua phap ly da thay doi co kiem soat qua cac
task, nen khong duoc ap mot gia tri E27 lich su cho moi regression.

## Ket qua

- `14/14 PASS`, tong thoi gian `1044.4 giay`.
- Migration 6 fault phase va restore checkpoint: PASS.
- Checkpoint SHA-256:
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`.
- Package hien hanh: `BQP-RPBM-2025@2.0.1`, checksum
  `E2537C08E7156414585BEFC446E74DD51E79AB8D67A3774B1F5250E7EECCDA90`.
- Workbook hien hanh: 92 formula error ke thua, `THKP-TC!E27=1201557000`.
- Evidence may doc: `tmp/DT-605-release-acceptance.json`.
- Evidence lifecycle: `tmp/DT-605-release-lifecycle.json`.

## Chay lai

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-release-acceptance.ps1 -Overwrite
```

Neu mot buoc fail do adapter/harness va checkpoint hash khong doi, co the sua harness roi
resume cac buoc PASS da ghi:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\test-release-acceptance.ps1 -Overwrite -Resume
```

Resume bi tu choi neu SHA-256 checkpoint thay doi.
