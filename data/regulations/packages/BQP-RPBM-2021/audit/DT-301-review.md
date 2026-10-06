# DT-301 - Bien ban doi chieu goi BQP 2021

Ngay kiem tra: 2026-08-03.

## Pham vi

- TT121/2021/TT-BQP: 59 dieu quy trinh ky thuat.
- TT123/2021/TT-BQP, Phu luc I: 34 ma dinh muc.
- TT123/2021/TT-BQP, Phu luc II: 34 quy tac/bang chi phi.
- TT122/2021/TT-BQP, Phu luc II: 33 dong du lieu co ban ca may.
- TT123/2021/TT-BQP: 35 record dia ban, mat do, rung va cap dat.
- Ba quy tac hieu luc/compliance.
- Tong: 198 record, tat ca `VerifiedAgainstOfficialSource` va co source locator.

## Nguon

| Van ban | Trang | SHA-256 PDF |
|---|---:|---|
| TT121/2021/TT-BQP | 99 | `9C0C9870EFD49FA49F7CC345B9D6D538D62AB5917EC9B68AE0B0C1E167363BA3` |
| TT122/2021/TT-BQP | 28 | `C5F78850F99411518267445B83F97124843666EEECDE06D5946C3DC24D0407EC` |
| TT123/2021/TT-BQP | 39 | `DE436E3BE81CE8B62A089D5B36C8A7A931BDB9EAC468D65B110455CB1F0DCF2E` |

PDF la ban scan. OCR chi dung de tim trang; gia tri phat hanh duoc doi chieu tren anh trang PDF.

## Luot A - PDF chinh thuc

- Quy trinh: doi chieu Dieu 1, 24, 29, 57 va 59 cua TT121; ma, title va page locator khop.
- Ca may: doi chieu truc quan cac trang 13-16 TT122; chon M010.001, M010.004,
  M010.016, M010.025 va M010.033. So ca, ty le, nguyen gia va VAT khop.
- Dinh muc: doi chieu 000.0100, 020.0800, 030.0700 va 040.0600 tai TT123.
- Chi phi: C=40%*NC; bang K1/K4; K3 min 2 trieu/max 60 trieu; bang K5; K6 5%/3% khop.
- Dia ban/mat do: bang tren can, duoi nuoc va duoi bien tai trang 4-7 khop.

Ket qua: `PASS`.

## Luot B - Workbook chuan

Workbook: `Dutoanmau/00. Du toan TP 3.xlsm`.

- `Tracuu!C22:G22`: `245;25;6;1,5;0,2`, khop khu vuc 4 tren can.
- `ChiPhi!D8`: `3,285`, khop moc dau bang K5 cong trinh dan dung.
- `DG Can` co ma `020.0800`; `DG Nuoc` co ma `030.0700`.
- Workbook `DG Nuoc` chua co nhom `040.x` duoi bien. Vi vay nhom nay chi doi chieu
  theo PDF chinh thuc va khong bi loai khoi package.

Ket qua: `PASS` cho pham vi workbook co du lieu; ghi nhan khoang thieu `040.x` de B4 xu ly.

## Test tu dong

- Build `Release|x64`: sach, 27/27 console tests dat.
- Bundle build lai cho cung package checksum va import package store dat.
- Key unique, count, page range, verification, unit catalog, quan he ty le ca may dat.
- `scripts/test-bqp-2021-package.ps1`: 198 record, PDF hash va hai luot review dat.
- B1 regression: `no-tags PASS`, `missing-sheet PASS-REJECTED`, `all-renamed PASS`.

Package checksum:
`BBE95EFDF3B2DA0780CAFEEF828C9418644452655C733C1BF2417F234C3E8008`.

## Checkpoint workbook

- Start: `DT-301-start.xlsm`, 670220 byte, SHA-256
  `81E38B7179319A4BDAC1CA22B94FFF93D46EA65C9187481A02808897C6F7C89D`.
- End: `DT-301-end.xlsm`, 670221 byte, SHA-256
  `AA3013E8AD420F97960A3CCF775F69066D41C14899D3B7089D7413D3C6D327E1`.
- Ca hai: profile schema 2, pin v2 Available, role PASS, 92 o loi cong thuc ke thua,
  `THKP-TC!E27=1198731000`.
- Hash byte khac do hai lan `SaveCopyAs`; kiem tra semantic khong co thay doi nghiep vu.
