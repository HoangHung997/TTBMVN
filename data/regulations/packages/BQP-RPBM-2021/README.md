# BQP-RPBM-2021

Goi du lieu snapshot ap dung tu 05/11/2021 den 27/10/2025, duoc dung tu ba PDF
chinh thuc TT121, TT122 va TT123 nam 2021.

## Cau truc

- `source/package.properties`: identity va khoang hieu luc.
- `source/sources.tsv`: van ban nguon, URL chinh thuc va SHA-256 PDF.
- `source/modules/*.tsv`: du lieu nguon co the review va build lai.
- `bundle/manifest.ttbmanifest`: manifest package da niem phong.
- `bundle/modules/*.data`: sau module runtime.
- `audit/DT-301-review.md`: bien ban doi chieu va test.

## Build va validate

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
.\ExcelAddIn1.RegulationTool\bin\Release\ExcelAddIn1.RegulationTool.exe build `
  .\data\regulations\packages\BQP-RPBM-2021\source `
  .\data\regulations\packages\BQP-RPBM-2021\bundle
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\scripts\test-bqp-2021-package.ps1
```

Builder khong ghi de bundle co cung ID/version nhung checksum khac. Muon phat hanh thay doi
du lieu phai tang `dataVersion`, khong sua ngam bundle da phat hanh.

## Checksum

- Package: `BBE95EFDF3B2DA0780CAFEEF828C9418644452655C733C1BF2417F234C3E8008`.
- `RegulationDataModule.ContentChecksum` bao ve canonical body ben trong module.
- `RegulationPackageModuleManifest.ContentChecksum` la SHA-256 byte cua file `.data` va la
  checksum ma package store kiem tra khi import.
