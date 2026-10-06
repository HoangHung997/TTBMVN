# Nhat ky thuc hien

File nay ghi theo thu tu thoi gian va khong xoa lich su cu. Moi task phai co it nhat mot ban ghi khi bat dau va mot ban ghi khi ket thuc/bi chan.

## 2026-08-03 - DT-000 DONE

- Pham vi: tao bo tai lieu dieu phoi phat trien module du toan RPBM.
- File tao: `README.md`, `ROADMAP.md`, `PROGRESS.md`, `TASKS.md`, `TESTING.md`, `ARCHITECTURE.md`, `SOURCES.md`, `HANDOFF.md`, `WORKLOG.md`.
- Workbook test xac dinh: `Dutoanmau/00. Du toan TP 3.xlsm`.
- Build command: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1`.
- Build result: `Release|x64` thanh cong.
- Tests: 5/5 dat.
- Ghi chu: workbook dang mo nen chua doc duoc SHA-256 tren file dia; se xu ly tai DT-101.
- Task tiep theo: `DT-101`.

## 2026-08-03 - DT-101 IN_PROGRESS

- Muc tieu: chot workbook tren man hinh thanh baseline kiem thu tai lap duoc.
- Nguoi thuc hien: Codex.
- Workbook: `Dutoanmau/00. Du toan TP 3.xlsm`.
- Trang thai dau task: workbook dang mo trong Excel; se dung Save/SaveCopyAs de bao toan thay doi tren man hinh.
- Next action: thu thong tin active workbook, tao baseline/checkpoint va lap bang expected output.

## 2026-08-03 - DT-101 DONE

- Muc tieu: chot workbook tren man hinh thanh baseline kiem thu tai lap duoc.
- File/tai lieu da thay doi: `docs/du-toan/BASELINE.md`, `TESTING.md`, `README.md`,
  `PROGRESS.md`, `HANDOFF.md`, `WORKLOG.md`.
- Baseline: `Dutoanmau/Baseline/00. Du toan TP 3.baseline.xlsm`.
- Checkpoint: `Dutoanmau/Checkpoints/DT-101-start.xlsm` va `DT-101-end.xlsm`.
- SHA-256 cua ca ba file:
  `E4BFDBD5FFA3AB47835A59D794F8E907633820F91E2FDF18EC9696559C89EADA`.
- Inventory: 16 sheet; 15 sheet co noi dung da render thanh PNG; `DosauHLAT` trong.
- Package: 73 Open XML entry; khong co `xl/vbaProject.bin` du mang duoi `.xlsm`.
- Defined names ke thua: 4826 name, trong do 4034 co `#REF!` va 120 co `#N/A`.
- Live Excel reopen: mo baseline, dieu huong `THKP-TC!E27`; actual `1198731000`,
  formula `=ROUND(E26,-3)`; dong bang `Don't Save`; checksum sau test khong doi.
- Expected output: da chot cell/formula cho THKL, VL-NC-M, DG Can, DG Nuoc,
  Gia DT TC, THKP-TC va invariant ho dao trong `BASELINE.md`; `TESTING.md` khong con
  gia tri chua chot.
- Loi ke thua: da ghi rieng trong `BASELINE.md`; gate sau chi cam loi moi.
- Build command: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1`.
- Build result: `Release|x64` thanh cong; 5/5 core tests dat; khong co warning.
- Quyet dinh: cached value cua Microsoft Excel live la oracle; render engine chi dung
  xem bo cuc vi co sai lech recalculate cong thuc dong.
- Task tiep theo: `DT-102`.

## 2026-08-03 - DT-102 IN_PROGRESS

- Muc tieu: nhan dien sheet bang role ben vung, khong phu thuoc ten tab.
- Nguoi thuc hien: Codex.
- Dau vao: role list trong `TASKS.md` va workbook chuan DT-101.
- Next action: them model/validator trong Core, service doc-ghi Worksheet.CustomProperties,
  unit test va test rename tren workbook checkpoint.

## 2026-08-03 - DT-102 DONE

- Muc tieu: nhan dien sheet theo role, khong phu thuoc ten tab.
- Code: them `ExcelAddIn1.Core/WorksheetRoles.cs`,
  `ExcelAddIn1/Funtion/WorksheetRoleService.cs`, test trong `ExcelAddIn1.Tests/Program.cs`
  va script `scripts/test-worksheet-roles.ps1`.
- Property: `TTBMVN.WorksheetRole` trong `Worksheet.CustomProperties`.
- Role da gan: ResourcePrices=VL-NC-M; UnitRateLand=DG Can; UnitRateWater=DG Nuoc;
  EstimateAppendix=Gia DT TC; CostSummary=THKP-TC; NormLookupView=Tracuu;
  CostRuleView=ChiPhi.
- Checkpoint start: `DT-102-start.xlsm`, SHA-256
  `E4BFDBD5FFA3AB47835A59D794F8E907633820F91E2FDF18EC9696559C89EADA`.
- Checkpoint end: `DT-102-end.xlsm`, SHA-256
  `5ADED6B7B01E244EAB98AFED4566784137B2ACB617CE3D0648C2B604028EB833`.
- Live test: validate 7/7 role; doi `DG Can` thanh `DG Can ROLE TEST`, resolver van tim
  theo UnitRateLand; tra ten cu; save, close/open working workbook va validate lai dat.
- Checkpoint reopen read-only: 7/7 role resolve dung; `THKP-TC!E27=1198731000`.
- Package inspection: marker role xuat hien trong dung 7 worksheet XML.
- Build/test: Release x64 thanh cong, khong warning; 8/8 core tests dat.
- Regression ho dao: 4 test core cu van dat; Excel events `True`, calculation
  `xlCalculationAutomatic` sau integration test.
- Quyet dinh: resolver khong fallback theo ten sheet; thieu/trung/unknown role la loi.
- Task tiep theo: `DT-103`.

## 2026-08-03 - DT-103 IN_PROGRESS

- Muc tieu: workbook tu mo ta duoc du an, package phap ly va price profile dang dung.
- Nguoi thuc hien: Codex.
- Next action: model ProjectProfile v1 trong Core, serializer/migration, service doc-ghi
  workbook custom properties va round-trip save/reopen tren workbook chuan.

## 2026-08-03 - DT-103 DONE

- Muc tieu: luu project profile co version ben trong workbook, doc lap may nguoi dung.
- Code: `ExcelAddIn1.Core/ProjectProfile.cs`,
  `ExcelAddIn1/Funtion/WorkbookProjectProfileService.cs`, test Core va
  `scripts/test-project-profile.ps1`.
- Model schema 1: ProjectId, PreparedDate, ApprovalDate nullable, PriceDate,
  RegulationPackageId, PriceProfileId va OverrideSummary.
- Serializer: key/value thu tu on dinh, Base64 UTF-8 cho text, date `yyyy-MM-dd`,
  SHA-256; migration schema 0 sang 1 khong ghi ngam.
- Store: mot manifest va cac chunk 240 ky tu theo generation; manifest commit sau cung;
  save cung payload tra `false` va khong nhan ban property.
- Phat hien va sua loi live: Office hien tai khong cast duoc CustomDocumentProperties sang
  `Office.DocumentProperties` (`E_NOINTERFACE`); service chuyen sang late-bound COM.
- Profile test trong working workbook: ProjectId `TEST-DT-103-00-DU-TOAN-TP3`, package
  `TEST-DT103-PACKAGE-NOT-FOR-CALCULATION`, price profile
  `TEST-DT103-PRICE-BASELINE`; day la marker test, khong phai du lieu phap ly.
- Payload: 1747 ky tu; checksum
  `49723AFBE82F9E6C582CFC50F48DD8E89D3612222E7291E68050967D25BE03B4`;
  `docProps/custom.xml` co dung 1 manifest va 8 part.
- Checkpoint start: `DT-103-start.xlsm`, SHA-256
  `5ADED6B7B01E244EAB98AFED4566784137B2ACB617CE3D0648C2B604028EB833`.
- Checkpoint end: `DT-103-end.xlsm`, SHA-256
  `8A53BB1486ADD44AA727A694D42BB0FAC5F3293724114C0EC99DEA5DA9FF3364`.
- Save/close/open: profile doc lai dung; save lan hai idempotent; checkpoint mo read-only
  doc duoc profile; role DT-102 valid; `THKP-TC!E27=1198731000`.
- Build/test: Release x64 thanh cong, khong warning; 11/11 core tests dat.
- Task tiep theo: `DT-104`.

## 2026-08-03 - DT-104 IN_PROGRESS

- Muc tieu: form dang mo cap nhat danh sach sheet sau add/delete/rename/activate ma khong
  reset du lieu nguoi dung dang nhap.
- Nguoi thuc hien: Codex.
- Next action: rao soat `ThisAddIn`, `FrmDaodat` va form du toan de tach event refresh danh
  sach khoi luong load setting/toan form.

## 2026-08-03 - DT-104 DONE

- Muc tieu: cap nhat danh sach sheet trong form dang mo ma khong reset input.
- Code: `ExcelAddIn1.Core/WorkbookSheetList.cs`,
  `ExcelAddIn1/Funtion/WorkbookSheetChangeCoordinator.cs`, tich hop `ThisAddIn` va
  `FrmDaodat`, test Core va `scripts/test-sheet-change-coordinator.ps1`.
- Coordinator: Excel events + WinForms timer 350 ms; periodic signature bat rename;
  subscribe theo workbook luc mo form thay vi `ActiveWorkbook` hien tai.
- Stable key: CodeName neu co; workbook chuan khong co VBA nen CodeName rong, live test
  phat hien va chuyen fallback sang COM IUnknown identity trong phien.
- FrmDaodat: ComboBox chua descriptor key/name; rename giu selection theo key; delete chon
  sheet hop le dau tien; callback chi goi `ApplyWorkbookSheets`.
- Live test tren `DT-104-start.xlsm`: 16 sheet ban dau; add, rename, activate va delete deu
  PASS voi 6 callback; range text chua luu va DataGridView value chua luu khong doi.
- Workbook test dong khong luu sau khi xoa sheet thu; working workbook khong nhan thay doi
  cau truc tu bai test.
- Checkpoint start: SHA-256
  `8A53BB1486ADD44AA727A694D42BB0FAC5F3293724114C0EC99DEA5DA9FF3364`.
- Checkpoint end: `DT-104-end.xlsm`, SHA-256
  `868C70CAF5AB142814409D5846510FD58B557F87E5FC36B5A4D898EE1DD9EE20`.
- Checkpoint end reopen: 16 sheet, role valid, ProjectProfile schema 1,
  `THKP-TC!E27=1198731000`.
- Build/test: Release x64 thanh cong, khong warning; 12/12 core tests dat; role/profile
  regression dat; Excel Events/DisplayAlerts=True, Calculation=Automatic.
- Task tiep theo: `DT-105`.

## 2026-08-03 - DT-105 IN_PROGRESS

- Muc tieu: workbook chua tag hoac mat sheet co the duoc anh xa lai qua wizard va gate B1.
- Nguoi thuc hien: Codex.
- Next action: model mapping draft trong Core, form wizard bảy role, validate xung dot,
  persistence va bo test workbook bien the.

## 2026-08-03 - DT-105 DONE

- Muc tieu: anh xa lai bay worksheet role cho workbook chua tag/mat sheet va dong gate B1.
- Code: `ExcelAddIn1.Core/WorksheetRoleMapping.cs`,
  `ExcelAddIn1/Winform/FrmWorksheetRoleMapping.cs`, `WorksheetRoleService.ApplyMapping`,
  gate trong `Ribbon1.btnDutoan_Click`, core test va `scripts/test-b1-role-mapping.ps1`.
- Validation: du bay role, khong trung role, sheet da chon con ton tai, mot sheet khong
  nhan nhieu role; ten tab chuan chi la goi y lan dau.
- Persistence: apply xoa/ghi marker trong mot transaction logic, validate sau ghi va rollback
  marker cu neu co exception.
- Variant no-tags: PASS, 16 sheet; goi y/apply/save/reopen thanh cong.
- Variant missing-sheet: PASS-REJECTED, 15 sheet; role validation va draft mapping tu choi.
- Variant all-renamed: PASS, 16 sheet; manual mapping/save/reopen va resolve 7/7 role dat.
- Regression: Release x64 khong warning, 13/13 core tests; role rename, ProjectProfile
  save/reopen va sheet coordinator add/rename/activate/delete deu dat.
- Checkpoint end: `Dutoanmau/Checkpoints/DT-105-end.xlsm`, SHA-256
  `2090595D4BB3F8A4FA59EEF6310F1347437C9C314D9540F7AD1625C8093EEF27`.
- Reopen checkpoint: ProjectProfile schema 1/checksum dung, 7/7 role valid,
  `THKP-TC!E27=1198731000`; runtime log khong co exception moi.
- Gate B1: `DAT`.
- Task tiep theo: `DT-201`.

## 2026-08-03 - DT-201 IN_PROGRESS

- Muc tieu: dinh nghia `RegulationPackage` bat bien, co version/checksum va khong phu
  thuoc Excel Interop hay WinForms.
- Nguoi thuc hien: Codex.
- Dau vao: ADR-003 den ADR-005, task DT-201 va nguon phap ly da lap danh muc.
- Next action: kiem ke core hien tai, chot invariant/schema, trien khai serializer va bo test
  validation/round-trip/package sai.

## 2026-08-03 - DT-201 DONE

- Muc tieu: mo hinh hoa package phap ly bat bien, co version/checksum va doc lap Excel.
- Code: `ExcelAddIn1.Core/RegulationPackage.cs`, `RegulationPackageSerializer.cs` va
  ba test moi trong `ExcelAddIn1.Tests/Program.cs`.
- Schema 1: ID, data version, effective from/to, Draft/Published/Superseded/Withdrawn,
  transition note, source documents, sau module manifest va package checksum.
- Validation: ID/version/date/status, URL nguon, SHA-256, source/module duplicate,
  module thieu va field/count/schema payload sai.
- Immutability: factory sao chep, trim, sap xep source/module va niêm phong checksum;
  collection chi doc, input list thay doi sau Create khong tac dong package.
- Serializer: canonical UTF-8/Base64, thu tu deterministic; payload bi sua va field la bi tu choi.
- Build/test: Release x64 khong warning; 16/16 test dat.
- B1 regression sau thay doi cuoi: no-tags PASS, missing-sheet PASS-REJECTED,
  all-renamed PASS; role/profile/coordinator da dat trong cung task.
- Checkpoint start/end: `DT-201-start.xlsm`, `DT-201-end.xlsm`, cung SHA-256
  `2090595D4BB3F8A4FA59EEF6310F1347437C9C314D9540F7AD1625C8093EEF27`;
  task khong ghi workbook.
- Task tiep theo: `DT-202`.

## 2026-08-03 - DT-202 IN_PROGRESS

- Muc tieu: cai nhieu package/version song song bang staging, atomic install va rollback.
- Nguoi thuc hien: Codex.
- Next action: chot layout storage/archive, threat model path traversal va cac ket qua import
  truoc khi trien khai repository thuần .NET.

## 2026-08-03 - DT-202 DONE

- Muc tieu: kho package version song song, staging import, atomic install va rollback.
- Code: `ExcelAddIn1.Core/RegulationPackageStore.cs` va ba test store trong
  `ExcelAddIn1.Tests/Program.cs`.
- Layout: manifest co dinh + sau file module theo enum; khong doc path tu payload.
- Validation: schema/checksum package, checksum module streaming, file thieu/thua,
  manifest >10 MB, Draft va reparse point deu bi tu choi.
- Atomicity: stage duoi repository, file lock lien tien trinh, receipt truoc commit,
  `Directory.Move` la diem commit duy nhat; fault injection BeforeCommit de lai kho rong.
- Version policy: import lai cung checksum tra AlreadyInstalled; cung ID/version checksum
  khac nem `RegulationPackageVersionConflictException`, ban cu van load dung.
- Build/test: Release x64 khong warning; 19/19 test dat; temp test directory duoc don sach.
- B1 regression: no-tags PASS, missing-sheet PASS-REJECTED, all-renamed PASS.
- Checkpoint start/end: `DT-202-start.xlsm`, `DT-202-end.xlsm`, cung SHA-256
  `2090595D4BB3F8A4FA59EEF6310F1347437C9C314D9540F7AD1625C8093EEF27`;
  task khong ghi workbook.
- Task tiep theo: `DT-203`.

## 2026-08-03 - DT-203 IN_PROGRESS

- Muc tieu: resolver package theo ngay hieu luc va quy dinh chuyen tiep, tra ket qua co ly do.
- Nguoi thuc hien: Codex.
- Next action: mo hinh decision input/result/rule va test cac moc 05/11/2021, 28/10/2025.

## 2026-08-03 - DT-203 DONE

- Muc tieu: resolve package theo ngay hieu luc va quy dinh chuyen tiep co can cu.
- Xac minh VBPL: TT121/122/123 hieu luc 05/11/2021; TT101 hieu luc 28/10/2025;
  Dieu 6 TT101 bao luu phuong an/du toan da phe duyet truoc ngay hieu luc.
- Code: `ExcelAddIn1.Core/RegulationPackageResolver.cs` va hai test resolver.
- Input: PreparedDate, ApprovalDate nullable, EvaluationDate; validation date-only va thu tu ngay.
- Output: package, checksum identity, decision code, reference date, rule ID, reason,
  source documents, candidates/errors.
- Cases: before/at 05/11/2021, before/at 28/10/2025, approval truoc/sau,
  pending, gap, overlap, missing historical package va Withdrawn.
- Build/test: Release x64 khong warning; 21/21 test dat.
- B1 regression: no-tags PASS, missing-sheet PASS-REJECTED, all-renamed PASS.
- Checkpoint start/end: `DT-203-start.xlsm`, `DT-203-end.xlsm`, cung SHA-256
  `2090595D4BB3F8A4FA59EEF6310F1347437C9C314D9540F7AD1625C8093EEF27`;
  task khong ghi workbook.
- Task tiep theo: `DT-204`.

## 2026-08-03 - DT-204 IN_PROGRESS

- Muc tieu: workbook pin package/version/checksum va mo dung ban cu tren may khac.
- Nguoi thuc hien: Codex.
- Next action: nang ProjectProfile schema/migration, them pin verifier va live save/reopen.

## 2026-08-03 - DT-204 DONE

- Muc tieu: workbook pin package/version/checksum va mo dung ban tren may/store khac.
- Code: ProjectProfile schema 2, `RegulationPackagePin.cs`,
  `WorkbookRegulationPackageService.cs`, AppPaths package store va test/script pin.
- Migration: schema 0/1 -> schema 2 trong bo nho, version/checksum trong; strict field count.
- Unit: Pin clone profile; Available/Unpinned/Missing/Corrupt/Invalid; store co newer version
  van nap dung version/checksum da pin.
- Live: `scripts/test-package-pin.ps1 -Overwrite`; pin v1, cai them v2, save/close/reopen,
  default store Available v1, machine-copy + empty store Missing, sau do Available khi co store.
- Profile sau pin: schema 2, payload 1872, payload checksum
  `9FAA6A5EE5116A7A275D2AD74AB6235F9B1E483F6623A9B08C18A1E71702DDA1`;
  package checksum `3E4E1F72611295B4961174E58E1A82C7E9C2CFA1E831CDF334E18E86073EA8BB`.
- Build/test: Release x64 khong warning; 22/22 test dat; B1 ba variant dat.
- Checkpoint start: `DT-204-start.xlsm`, SHA-256 `2090595D...3EEF27`.
- Checkpoint end: `DT-204-end.xlsm`, SHA-256
  `B8D0EB86CDC75F5A7AF5FA06FBB5F3CBDF3F869EC5A78957A3A5072FFAF7CF46`;
  reopen profile/role/package pin valid, `THKP-TC!E27=1198731000`.
- Task tiep theo: `DT-205`.

## 2026-08-03 - DT-205 DONE

- Muc tieu: preview migration co diff/checkpoint/confirm/rollback va dong gate B2.
- Code: `RegulationPackageMigration.cs`, `WorkbookPackageMigrationService.cs`, test
  `RegulationPackageDiffAndPlan` va `scripts/test-package-migration.ps1`.
- Preview: PlanId `MIG-8B6CF6D083A5F14D0895EF02`, 7 thay doi, 6 module tinh toan.
- Live: cancel khong doi payload/khong tao file; injected fault sau pin rollback v1; apply
  tao backup v1 va save/reopen working workbook pin v2.
- Loi phat hien: save khi Calculation=Manual lam mode Manual bi luu vao workbook. Da doi
  migration de save backup/final chi khi state da khoi phuc; test state sau reopen dat.
- Build/test: Release x64 sach; 23/23 test dat; B1 no-tags/missing/all-renamed dat.
- Expected/actual: 92 o loi Excel ke thua khong tang; role valid;
  `THKP-TC!E27=1198731000`.
- Target package checksum:
  `6D044E98A906BC5C0FB7A92CCED66075F9924A039106A5210D4EFCB1C961BDA7`.
- Checkpoint end: `DT-205-end.xlsm`, 670222 byte, SHA-256
  `43E11CE892840E6CA2BF784A79ED19072D4014DEA0B807E6EF23569EA9CF4F39`;
  reopen profile schema 2/pin v2/role/E27 dat.
- Gate B2: `DAT`.
- Task tiep theo: `DT-301`.

## 2026-08-03 - DT-301 IN_PROGRESS

- Muc tieu: tao goi du lieu BQP 2021 tu TT121/122/123, moi record co source locator.
- Nguoi thuc hien: Codex.
- Next action: chot schema catalog/raw evidence, nhap nguon chinh thuc va them validation.

## 2026-08-03 - DT-301 DONE

- Muc tieu: tao goi BQP 2021 tu TT121/122/123, moi record co source locator.
- Raw evidence: ba PDF chinh thuc, 99/28/39 trang; SHA-256 da khoa trong `sources.tsv`.
- OCR: tool `tools/ocr/ocr_regulations.py`, output theo trang va manifest; OCR chi dung tim trang.
- Code: `RegulationDataModule`, source TSV reader, atomic bundle builder/reader va
  `ExcelAddIn1.RegulationTool` build/validate.
- Data: 198 record gom TechnicalProcess 59, Norm 34, CostRule 34, MachineRate 33,
  Geography 35 va Compliance 3; tat ca verified truc tiep nguon chinh thuc.
- Loi phat hien: package manifest can SHA-256 byte file module, nhung validator moi da so voi
  checksum canonical ben trong module. Da tach hai contract va them test package store.
- Package checksum:
  `BBE95EFDF3B2DA0780CAFEEF828C9418644452655C733C1BF2417F234C3E8008`.
- Review A PDF: quy trinh/dinh muc/ca may/chi phi/dia ban sample dat.
- Review B workbook: `Tracuu!C22:G22`, `ChiPhi!D8`, ma DG Can/DG Nuoc dat. Ghi nhan
  `DG Nuoc` chua co nhom 040.x nen van ban la nguon quyet dinh cho nhom nay.
- Build/test: Release x64 sach, 27/27 tests dat; `test-bqp-2021-package.ps1` dat;
  B1 no-tags/missing/all-renamed dat.
- Checkpoint start SHA
  `81E38B7179319A4BDAC1CA22B94FFF93D46EA65C9187481A02808897C6F7C89D`;
  end SHA `AA3013E8AD420F97960A3CCF775F69066D41C14899D3B7089D7413D3C6D327E1`.
- Ca hai checkpoint: profile schema 2, package v2 Available, role PASS, 92 o loi ke thua,
  `THKP-TC!E27=1198731000`.
- Audit: `data/regulations/packages/BQP-RPBM-2021/audit/DT-301-review.md`.
- Task tiep theo: `DT-302`.

## 2026-08-03 - DT-302 IN_PROGRESS

- Muc tieu: tao snapshot resolved BQP 2025 hieu luc 28/10/2025 va diff giai thich duoc
  tung thay doi so voi package 2021.
- Nguon bat buoc: TT101/2025 va VBHN 94-98/BQP tu cong thong tin chinh thuc.
- Next action: tai raw evidence/checksum, OCR neu can va lap ma tran thay doi theo sau module.

## 2026-08-03 - DT-302 DONE

- Raw evidence: TT101 va 94-98/VBHN-BQP, 69/30/99/35/52/76 trang; PDF va text-layer
  SHA-256 da khoa trong source/manifest.
- Data: package `BQP-RPBM-2025@2.0.0`, 248 record gom TechnicalProcess 60, Norm 42,
  CostRule 37, MachineRate 33, Geography 69 va Compliance 7.
- Thay doi nghiep vu: 020.0500, tam tinh theo ha, K2, hai muc toi thieu, Bieu mau 05,
  gia/hao phi may, 34 dia ban, vung bien va quy tac chuyen tiep TT101 Dieu 6.
- Code: source generator deterministic, text-layer extractor, diff cap record/field va lenh
  `RegulationTool diff`; validator cho phep van ban hop nhat ky sau moc hieu luc noi dung.
- Audit: `248 = 50 added + 0 removed + 198 changed`; moi record diff co ly do, locator va
  `VerifiedAgainstOfficialSource` trong `change-reasons.tsv`.
- Review doc lap: anh TT101 trang 35-38; VBHN96 trang 15-19; VBHN97 trang 5, 14,
  42, 44-46, 52 deu khop source.
- Build/test: Release x64 sach, 29/29 tests dat; package script regenerate/hash/page/count/diff
  dat; B1 role/profile/coordinator/ba variant dat.
- Package checksum:
  `F4618DF4D256378C59A5719912CFF4D9EED410B52FCF157AEAD1D4EB6EC49A4D`.
- Checkpoint start SHA
  `0AD1CCD5BD610268BF303E71982CF83A93AF2435994CD2A146DCBC8FD83AAF89`;
  end SHA `1AB5B6487827DFDD55803BCF602C880598414B1C42F17B9AAFA891710EB79272`.
- Ca hai checkpoint: profile schema 2, package Available, role PASS, 92 loi ke thua,
  `THKP-TC!E27=1198731000`.
- Audit chinh: `data/regulations/packages/BQP-RPBM-2025/audit/DT-302-review.md`.
- Task tiep theo: `DT-303`.

## 2026-08-03 - DT-303 IN_PROGRESS

- Muc tieu: trien khai catalog va engine gia ca may theo TT122/2021 va Phu luc IV TT101.
- Checkpoint start: `DT-303-start.xlsm`, 670201 byte, SHA-256
  `2780D4F0D3419FEF7EA0539FA09C6DEE6DBCFC399F5A02883339B395F8C826EC`;
  profile/role/package Available, 92 loi ke thua va E27 dung baseline.
- Next action: trich cong thuc (1)-(6), dinh nghia input/rounding, sau do code Core engine va
  unit test tung thanh phan truoc khi doi chieu DG Can/DG Nuoc.

## 2026-08-03 - DT-303 DONE

- Muc tieu: catalog va engine gia ca may TT122/VBHN96 cho hai doi tuong luong.
- Code: them `MachineRateCatalog.cs`, `MachineRateCalculator.cs`; parser strict invariant,
  decimal engine, component rounding, gia cho, gia gio, an mon, thu hoi va loai nhien lieu.
- Data: 33 key logic resolve M010/M011; sua crew tau thanh si quan + thuy thu; bo sung
  override M011.020 va M011.024; locator bao phu Bang 01/03 trang 15-29.
- Phat hien nguon: Bang 04 in sai M011.024 sua chua/tong va lech 1 VND tai tong M011.009/010;
  engine giu ket qua cong thuc, audit tai `DT-303-machine-rate-review.md`.
- Package 2025 moi: checksum
  `D58CAC64422E14151FD68F1C92944FAB7A9859908165C59F92089E7F13998B7C`;
  MachineRate checksum
  `6352542DC1F488FEA3D6B5D05D16E37D80464FEC91C09A0486FA4E7839BE8CBE`.
- Test: Release x64 sach 32/32; package 2021/2025 dat; 3 machine core gate dat;
  `VL-NC-M` khop 10 tong may; pin/migration/profile/coordinator/B1 variants deu dat.
- Baseline doc sua cell tong may hut/xoi bun cat tu F57 thanh F56 theo workbook thuc.
- Checkpoint end: `DT-303-end.xlsm`, 670222 byte, SHA-256
  `EF55831509852668DA6F154D7751D239DC53F230B906A56CB799FF0D03C8CC4C`;
  profile schema 2, package Available, role PASS, 92 loi ke thua, E27=1198731000.
- Task tiep theo: `DT-304`.

## 2026-08-03 - DT-304 IN_PROGRESS

- Muc tieu: catalog dinh muc va engine quy tac chi phi TT123/VBHN97.
- Checkpoint start: `DT-304-start.xlsm`, 670224 byte, SHA-256
  `18C240708AA4F3031758D686F0DE5D74AC4E24ECD3CA2C67867E70523B54320E`;
  semantic checkpoint dat.
- Next action: inventory schema Norm/CostRule, trich cong thuc/dieu kien/noi suy va chot
  contract input-output-rounding truoc khi code.

## 2026-08-03 - DT-304 DONE

- Muc tieu: catalog dinh muc va engine quy tac chi phi TT123/VBHN97, co can cu BXD dong.
- Code: them `NormCatalog.cs`, `NormCalculator.cs`, `CostRuleCatalog.cs`,
  `CostRuleCalculator.cs`; parser strict invariant va ket qua kem locator/rounding.
- Data: 34 dinh muc day du VL/NC/M, bien the va dieu kien; K1-K6, chi phi chung,
  truc tiep, tong hop va bang K2/K5 hien hanh.
- Nguon bo sung: TT36/2026/TT-BXD va TT38/2026/TT-BXD cung phu luc chinh thuc;
  bon PDF da khoa SHA-256 trong `data/regulations/raw/2026/README.md`.
- Package 2025: checksum
  `C23A3AE526841473AEAB6E36CCC76B0663A56C5FE63D3351F7AFBFA4639447F2`;
  Norm `2CFF3C03...EFD0E9`, CostRule `73FC9814...346860` (locator 040.0500
  duoc audit doc lap DT-305 sua tu trang 40 thanh 39-40, khong doi hao phi).
- Test: Release x64 sach, 38/38 core; package 2021/2025, machine, pin/migration,
  project profile, sheet coordinator va B1 variants deu dat.
- Workbook: 12 o dinh muc, 55 o K5, 10 o K2 khop; 92 loi ke thua khong tang;
  `THKP-TC!E27=1198731000`.
- Phat hien: `ChiPhi!A3` con nhan TT12/2021; P7/P13 ngoai suy duoi moc 10 ty.
  Engine theo van ban hien hanh, khong sao chep cong thuc workbook sai.
- Sua test migration: persistence/backup duoc mo trong Excel an rieng, khong dong/mo lai
  workbook lon trong phien Excel giao dien; cancel/rollback/apply/reopen deu PASS.
- Checkpoint end: `DT-304-end.xlsm`, 670239 byte, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  profile/role/package Available, 92 loi ke thua, E27 dung baseline.
- Audit: `data/regulations/packages/BQP-RPBM-2025/audit/DT-304-norm-cost-review.md`.
- Task tiep theo: `DT-305`.

## 2026-08-03 - DT-305 IN_PROGRESS

- Muc tieu: review doc lap toan bo du lieu B3 va dong gate B3.
- Checkpoint start: `DT-305-start.xlsm`, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  semantic checkpoint dat.
- Next action: viet/chay audit TSV doc lap missing/duplicate/outlier, sample PDF va
  deterministic checksum truoc khi lap bien ban B3.

## 2026-08-03 - DT-305 DONE

- Muc tieu: review doc lap toan bo du lieu B3 va dong gate B3.
- Auditor: `tools/regulations/audit_bqp_2025.py`, chi dung Python standard library de doc
  source/PDF evidence, doc lap voi C# runtime.
- Ket qua: 6534 check, 248 record, 248 change reason, 6 PDF, zero error/warning;
  deterministic rebuild khop 7/7 file.
- Finding: locator `NORM-040.0500` ghi trang 40 trong khi ma/heading bat dau trang 39.
  Da sua generator thanh 39-40; hao phi va behavior khong doi.
- Package final B3 checksum:
  `C23A3AE526841473AEAB6E36CCC76B0663A56C5FE63D3351F7AFBFA4639447F2`;
  Norm `2CFF3C031ACE079157333AAE97F60A55B3CDED4A926127AFCC0013FDC3EFD0E9`.
- Ha tang test: them `excel-test-process.ps1`, Excel `/x` theo PID/native object,
  resolver workbook theo full path va working copy cho pin/migration/B1/coordinator.
- Loi da loai: mo workbook lon trong Excel giao dien gay resource warning,
  `RPC_E_CALL_REJECTED` va ROT tro nham instance.
- Regression: Release x64 38/38; BQP 2021/2025, machine, norm/cost, pin/migration,
  B1 variants, coordinator, profile/role va workbook checkpoint deu dat.
- Checkpoint end: `DT-305-end.xlsm`, 670239 byte, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  profile schema 2, package Available, role PASS, 92 loi ke thua, E27=1198731000.
- Bien ban: `data/regulations/packages/BQP-RPBM-2025/audit/DT-305-gate-review.md`.
- Gate B3: `DAT`.
- Task tiep theo: `DT-401`.

## 2026-08-03 - DT-401 IN_PROGRESS

- Muc tieu: project setup va sheet mapping thanh mot luong co validate/cancel/restore.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-401-start.xlsm`, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  semantic checkpoint dat.
- Next action: inventory entry point/UI du toan, chot contract draft/commit va unit test
  truoc khi noi resolver package, pin va wizard mapping.

## 2026-08-03 - DT-401 DONE

- Core: them `ProjectSetupDraft`, planner/result/plan va transition catalog TT101 Dieu 6;
  test package 2021/2025, ngay hien hanh, transition, mapping thieu va ngay sai.
- Add-in: them `FrmProjectSetup`, bootstrap bundled package, transactional
  `WorkbookProjectSetupService`, restore/clear ProjectProfile va worksheet role snapshot.
- Ribbon: hop nhat project profile, moc gia, can cu va mapping vao mot dialog truoc khi mo
  form thong tin du toan; form dang mo van duoc toggle ma khong reload setup.
- Package deploy: 14 file bundle 2021/2025 duoc copy vao
  `bin/x64/Release/RegulationPackages` va import idempotent vao local store.
- Test Core: Release x64 sach, 39/39 dat.
- Live Excel: `test-project-setup.ps1 -Overwrite` dat form smoke, cancel no-write, fault
  rollback exact, file bo tag, doi ten ca 7 sheet theo CodeName va save/reopen.
- Regression: coordinator bao toan input, project profile/role valid; 92 loi Excel ke thua
  khong tang va `THKP-TC!E27=1198731000`.
- Ha tang test: COM factory chi duoc chap nhan khi tao PID Excel moi; fallback `/x`/native
  object. Khong dung ROT va khong cham workbook giao dien cua nguoi dung.
- Checkpoint end: `DT-401-end.xlsm`, 670239 byte, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  semantic checkpoint dat.
- Task tiep theo: `DT-402`.

## 2026-08-03 - DT-402 IN_PROGRESS

- Muc tieu: thay nguon sheet `Tracuu` bang NormCatalog trong app, co exact/fuzzy/filter,
  hao phi va can cu phap ly.
- Nguoi thuc hien: Codex.
- Next action: inventory UI/ham tra cuu hien co va chot query/result contract Core.

## 2026-08-03 - DT-402 DONE

- Inventory: khong co UI/ham tra cuu dang hoat dong; sheet `Tracuu` chi con vai tro mapping.
- Core: them `NormSearchIndex/Query/Result`, exact key, fuzzy khong dau/typo, filter moi
  truong/do sau/variant/resource va multi-package identity.
- Store/runtime: them `LoadBundleRequired` validate full bundle va
  `WorkbookNormCatalogService` nap module Norm theo profile pin.
- UI: them `NormLookupControl` va shell `Dutoan`; ket qua hien ma/ten/don vi, hao phi theo
  variant, dieu chinh/rang buoc, package va source locator.
- Du lieu cu: package 2021 tim duoc 34 ma/title/source nhung chua co hao phi chi tiet; UI
  canh bao ro, khong su dung hao phi package 2025 thay the.
- Test: Release x64 sach, 41/41 core; exact/fuzzy/filter/vi-VN/de-DE, missing code va cung
  ma trong hai package deu dat.
- Live Excel: `test-norm-lookup.ps1 -Overwrite` dat 34 row, package 2021/2025, shell
  navigation, screenshot, save/reopen; 92 loi ke thua va E27 dung baseline.
- Regression DT-401: form smoke, cancel no-write, fault rollback, untagged/renamed va
  save/reopen van dat.
- Checkpoint start/end: `DT-402-start.xlsm`/`DT-402-end.xlsm`, 670239 byte, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  semantic checkpoint dat.
- Task tiep theo: `DT-403`.

## 2026-08-03 - DT-403 IN_PROGRESS

- Muc tieu: thay bang hard-code va sheet `ChiPhi` bang rule engine theo package da pin.
- Nguoi thuc hien: Codex.
- Next action: inventory DataDutoan/FormThongtinchung/ChiPhi va chot request-result contract.

## 2026-08-03 - DT-403 DONE

- Core: them `CostRuleEngine` voi request K1-K6, ket qua calculated/applied, summary VAT,
  rounding 1 VND AwayFromZero va override bat buoc ly do.
- Versioning: `CostRuleCatalog` nap du ca TT123/2021 va VBHN 2025. Adapter 2021 giu bang K5
  toi 2.000 ty, khong ap minimum 6,3 trieu va khong gan currentExternalBasis 2026.
- Runtime: them `WorkbookCostRuleService`, chi nap CostRule bang package ID/version/checksum
  trong ProjectProfile; xoa `DataDutoan.cs` va field hard-code khong con su dung.
- UI: them `CostRuleControl` vao shell `Dutoan`; nhap bien so, chon K1-K6, hien calculated,
  applied, source, override/reason va tong T/C/Z/truc thue/VAT/sau thue.
- Test: Release x64 sach, 43/43 core. `test-cost-rule-ui.ps1 -Overwrite` dat package
  2021/2025, 6 component, override K5, UI/screenshot, shell, save/reopen.
- Regression: `test-norm-lookup.ps1` va `test-project-setup.ps1` deu dat; 92 loi Excel
  ke thua khong tang va `THKP-TC!E27=1198731000`.
- Checkpoint start/end: 670239 byte, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  profile schema 2, package Available, role PASS va semantic gate dat.
- Quyet dinh: ADR-025; sheet `ChiPhi` khong con la data source, override tach khoi quy tac goc.
- Task tiep theo: `DT-404`.

## 2026-08-03 - DT-404 IN_PROGRESS

- Muc tieu: quan ly VL-NC-M va PriceProfile theo dia diem/thoi diem, tach khoi package phap ly.
- Nguoi thuc hien: Codex.
- Next action: inventory sheet `VL-NC-M`, ma nguon luc, cong thuc/link va persistence hien co
  de chot schema PriceProfile va test contract truoc khi code.

## 2026-08-03 - DT-404 DONE

- Inventory: `VL-NC-M` A1:N117, 174 cong thuc phu thuoc, 32 ten VLOOKUP; chot 34 entry
  legacy, 21 vat lieu nhap truc tiep va 10 block ca may. Norm 2025 tao 39 requirement chua
  co gia trong workbook mau; missing duoc bao, khong gan 0.
- Core: them `PriceProfile`, entry/source/override/audit, coverage validator, deterministic
  serializer/checksum, atomic local store va adapter sang `MachineRatePriceProfile`.
- Runtime: them generic `WorkbookCustomPayloadStore`, snapshot/pin giao dich,
  `WorkbookPriceRequirementService`, adapter import legacy va batch-write F81:F101 theo role.
- UI: them `PriceProfileControl` va muc `Bang gia` trong shell; ho tro nhap tay, override,
  coverage, doc VL-NC-M, import/export, local store va snapshot workbook.
- On dinh COM: thay bon `FinalReleaseComObject` bang `ReleaseComObject`; loi RCW bi tach khoi
  doi tuong goc da duoc tai hien khi service resolve lai sheet va da loai bo.
- Test harness: sua ep kieu mang COM PowerShell, collection generic reflection va host WinForms
  that de screenshot/assert metadata; day la loi harness, duoc phan biet voi loi RCW runtime.
- Test: Release x64 sach, 47/47 core. `test-price-profile.ps1 -Overwrite` dat 34 entry,
  39 missing, 21 batch update, linked unit rate, formula preservation, UI, override save/reopen
  va baseline restore.
- Regression: DT-401, DT-402, DT-403 chay song song trong ba Excel PID rieng deu dat;
  92 loi ke thua khong tang va `THKP-TC!E27=1198731000`.
- Checkpoint start/end: 670239 byte, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  profile schema 2, package Available, role PASS va semantic gate dat.
- Quyet dinh: ADR-026. PriceProfile tach package phap ly; workbook giu full snapshot;
  legacy writer chi ghi vat lieu va coverage khong zero-fill.
- Task tiep theo: `DT-405`.

## 2026-08-03 - DT-405 IN_PROGRESS

- Muc tieu: tinh don gia chi tiet `DG Can`/`DG Nuoc` tu NormCatalog va PriceProfile da pin,
  co tuy chon he so ro rang va khong doc sheet legacy lam data source.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-405-start.xlsm`, 670239 byte, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  semantic checkpoint dat.
- Next action: inventory hai sheet, cong thuc/ma/he so, doi chieu NormCatalog va chot
  request-result/test vectors truoc khi code engine.

## 2026-08-03 - DT-405 DONE

- Inventory: `DG Can` A1:U170 va `DG Nuoc` A1:U438; doi chieu ma/cong thuc/he so legacy
  tai `tmp/DT-405-unit-rate-inventory.json`. `DG Nuoc!K1=1,1` la he so luu toc cu.
- Core: them `UnitRateCalculator` tinh VL/NC/M bang `decimal`, khong lam tron, xu ly vat lieu
  phan tram sau tong vat lieu truc tiep va chan gia thieu/sai loai/sai don vi.
- Dieu kien: he so doc, tin hieu vat no, dao nuoc va ba khoang luu toc chi ap dung khi duoc
  chon ro. Van toc tren 2 m/s va dieu kien trai nhau tiep tuc do `NormCalculator` chan.
- Binding: ma logic nhu `M010.DIVING` khong duoc tu chon may; nguoi dung phai anh xa ma gia
  may cu the va nhap ly do/phuong an. Ket qua giu source locator va danh tinh PriceProfile.
- Runtime/UI: them `WorkbookUnitRateService`, `UnitRateControl` va muc `Don gia` trong shell;
  ho tro can/nuoc, bien the, dieu kien, binding, preview hao phi va tong VL/NC/M.
- Test: Release x64 sach, 50/50 core. `test-unit-rate.ps1 -Overwrite` dat land/water
  baseline, he so, hai binding may, missing price, UI, doi ten role sheet va save/reopen.
- Regression: DT-401..404 deu dat trong Excel PID rieng; 92 loi ke thua khong tang va
  `THKP-TC!E27=1198731000`.
- Checkpoint end va start task sau: `DT-405-end.xlsm`/`DT-406-start.xlsm`, 670239 byte,
  SHA-256 `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  profile schema 2, package Available, role PASS va semantic gate dat.
- Quyet dinh: ADR-027. Don gia chi nhan input ro rang; khong doc cong thuc legacy lam data
  source, khong zero-fill va khong tu resolve ma may logic.
- Task tiep theo: `DT-406`.

## 2026-08-03 - DT-406 IN_PROGRESS

- Muc tieu: ap don gia tu engine vao phu luc `Gia DT TC`, giu ro khoi luong, nguon, precision
  va lien ket thay doi dau vao.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-406-start.xlsm`, 670239 byte, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  semantic checkpoint dat.
- Next action: inventory cau truc/cell/formula/defined name cua `Gia DT TC` va doi chieu
  don gia can/nuoc de chot contract Core truoc khi noi writer Excel.

## 2026-08-03 - DT-406 DONE

- Inventory: `Gia DT TC` A1:V32, 13 dong cong tac (7 can, 6 nuoc), dong he so 26 va
  tong 27; phat hien subtotal legacy bo sot row 18/25 va `Gia DT TC!A3=#REF!` ke thua.
- Core: them `EstimateAppendixCalculator` voi request/result/validation decimal, tong theo
  nhom va khong lam tron trung gian. `UnitRateCalculator` bo qua hao phi bang 0.
- Runtime: them `WorkbookEstimateAppendixService` import mapping theo worksheet role,
  variant sparse `NORM-020.0500`, binding may lan va canh bao dong zero quantity.
- Sua phap ly: package `BQP-RPBM-2025@2.0.1`, checksum
  `E2537C08E7156414585BEFC446E74DD51E79AB8D67A3774B1F5250E7EECCDA90`;
  `NORM-020.0500/.1000` dung co do 0,4x0,6 m. Resolver chon patch cao nhat trong cung
  PackageId/effective date, nhung van bao ambiguous giua cac PackageId khac nhau.
- Sai lech oracle: row 22 workbook dung ma suffix `.1`/hao phi may 0,014 nhung mo ta va
  khoi luong la nuoc 3-12 m. Theo VBHN97 trang 32-33, adapter chon `water-3-12`/0,016;
  tong may tang dung `664425.135`, `Gia DT TC!K27=558860046.12222`.
- Writer: `WorkbookEstimateAppendixWriter` ghi batch F:N, rollback khi loi, giu cong thuc
  dong theo D/E, sua subtotal du dong, dua phan chenh he so vao row 26 va verify tong sau ghi.
- UI: them `EstimateAppendixControl` va muc `Phu luc DT`; chon he so nuoc, preview 13 dong,
  tooltip canh bao, tong VL/NC/M va ghi/lưu workbook. Screenshot:
  `tmp/DT-406-estimate-appendix.png`.
- Test: Release x64 sach, 52/52 core. `test-estimate-appendix.ps1 -Overwrite` dat mapping,
  preview, binding 029/030, batch write, thay input tu tinh lai, UI va save/reopen.
- Regression: DT-401..405 deu dat voi package 2.0.1; 92 loi ke thua khong tang.
- Checkpoint end/start task sau: `DT-406-end.xlsm`/`DT-407-start.xlsm`, 673771 byte,
  SHA-256 `288324623F921259C5633A4180FC9F8062AD138DA0A37718982C48F09C247B22`;
  package/profile/role Available, `THKP-TC!E27=1199577000`.
- Quyet dinh: ADR-028. Phu luc ghi don gia goc va he so tach dong; sai legacy co can cu
  phap ly duoc sua va ghi ro delta thay vi ep khop workbook cu.
- Task tiep theo: `DT-407`.

## 2026-08-03 - DT-407 IN_PROGRESS

- Muc tieu: tinh va ghi `THKP-TC` tu phu luc/CostRule da pin, sau do dong Gate B4 bang
  test end-to-end tren checkpoint DT-407-start.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-407-start.xlsm`, 673771 byte, SHA-256
  `288324623F921259C5633A4180FC9F8062AD138DA0A37718982C48F09C247B22`.
- Next action: inventory cell/formula/he so/minimum/thue/defined name cua `THKP-TC` va chot
  contract tong hop truoc khi code.

## 2026-08-03 - DT-407 DONE

- Inventory: `THKP-TC` A1:Y47, 14 merged area, cong thuc lien ket `Gia DT TC`/`ChiPhi` va
  1 print-area; khong co `#REF!` moi trong sheet. Bang mau la Bieu 04 nguon von khac.
- Core: them `CostSummaryCalculator` cho Bieu 01-04, tinh C/TL/Z, K1-K6 voi K2 tren T,
  cac K con lai tren Z, VAT loai K3/K4, lam tron VND/1.000 VND va override co ly do.
- Tien bang chu: them `VietnameseMoneyWords`; dong A28 duoc cap nhat theo ket qua lam tron.
- Adapter/UI: them `WorkbookCostSummaryService`, `WorkbookCostSummaryWriter`,
  `CostSummaryControl` va muc `Tong hop KP`; import cau hinh legacy theo role, preview va
  batch write D10:E27/A28 co backup/rollback/verify.
- Sua phap ly: workbook cu tinh K5 tren T voi 2.873%. VBHN97 trang 45/51 va bang hien hanh
  trong package quy dinh K5 tren Z; du an nong nghiep duoi 10 ty dung 2.598%.
- Ket qua: T=839815452; C=107365928; TL=52094976; Z=999276356; K=114386486;
  VAT=87893896; H=1201556738; lam tron `THKP-TC!E27=1201557000`.
- Test: Release x64 sach, 55/55 core. `test-cost-summary.ps1 -Overwrite` dat import,
  preview, K5, batch write, thay input tu tinh lai, UI, tien bang chu va save/reopen.
- Regression: DT-401..406 chay tuan tu deu dat; 92 loi Excel ke thua khong tang.
- Checkpoint end/start task sau: `DT-407-end.xlsm`/`DT-501-start.xlsm`, 673693 byte,
  SHA-256 `1D878CA03CFA0B623FBA75C59C9B7E1D9E7F9163EDEEC19BBFA295FB551D2325`;
  package 2.0.1/profile/role Available, E27 va semantic gate dat.
- Quyet dinh: ADR-029. THKP dung engine template, writer chi ghi khoi tinh toan theo role;
  sai cong thuc legacy co can cu phap ly duoc sua va ghi ro delta.
- Gate B4: `DONE`.
- Task tiep theo: `DT-501`.

## 2026-08-03 - DT-501 IN_PROGRESS

- Muc tieu: kiem ke va hop nhat cac duong ghi Excel thanh batch co transaction, phuc hoi
  ScreenUpdating/Calculation/Events/DisplayAlerts va co bang chung performance/exception.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-501-start.xlsm`, 673693 byte, SHA-256
  `1D878CA03CFA0B623FBA75C59C9B7E1D9E7F9163EDEEC19BBFA295FB551D2325`.
- Next action: inventory writer/range assignment hien co, phan loai batch va cell-by-cell.

## 2026-08-03 - DT-501 DONE

- Inventory: `tmp/DT-501-excel-write-inventory.json`; phan loai duong `Value2`, `Formula`,
  `Insert/Delete`, `ExcelWriteContext`, transaction va cac vong ghi theo cell.
- Runtime: them `ExcelBatchWriteTransaction`, snapshot Formula theo worksheet/address,
  rollback nguoc thu tu va giai phong COM snapshot. Writer PriceProfile, `Gia DT TC` va
  `THKP-TC` dung transaction chung thay cho ba cach rollback lap lai.
- `ExcelWriteContext`: constructor tu phuc hoi neu loi giua luc tat ScreenUpdating,
  Calculation, Events, DisplayAlerts; `Dispose` idempotent va phuc hoi qua mot ham chung.
- Ho dao: duong selection ghi mot ma tran 7 cot/block; header, auxiliary link va link-back
  duoc gom theo hang; cac range nong duoc release ro rang.
- Ribbon legacy: random nhan he so va cong gia tri vung chon chuyen sang matrix/transaction.
- Test: Release x64 sach, 55/55 core. `test-excel-batch-write.ps1 -Overwrite` tren 4.000 o:
  batch 41 ms, cell-by-cell 5.899 ms, nhanh hon 142,47 lan; value parity, exception rollback,
  bon Excel state va batch path ho dao deu PASS.
- Regression: PriceProfile, `Gia DT TC`, `THKP-TC` va ProjectSetup deu dat; 92 loi ke thua
  khong tang va ket qua nghiep vu giu nguyen.
- Checkpoint end/start task sau: `DT-501-end.xlsm`/`DT-502-start.xlsm`, 673693 byte,
  SHA-256 `1D878CA03CFA0B623FBA75C59C9B7E1D9E7F9163EDEEC19BBFA295FB551D2325`;
  semantic checkpoint dat, workbook khong can thay doi.
- Quyet dinh: ADR-030. Ghi Excel theo batch transaction va phuc hoi trang thai dung chung.
- Task tiep theo: `DT-502`.

## 2026-08-03 - DT-502 IN_PROGRESS

- Muc tieu: moi block ket qua truy duoc package, PriceProfile, ma dinh muc va source locator.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-502-start.xlsm`, 673693 byte, SHA-256
  `1D878CA03CFA0B623FBA75C59C9B7E1D9E7F9163EDEEC19BBFA295FB551D2325`.
- Next action: inventory identity/source hien co va chot schema audit block deterministic.

## 2026-08-03 - DT-502 DONE

- Core: them `ResultAuditTrail` schema 1, immutable entry/source, validation, deterministic
  Base64 serializer, SHA-256, replace-scope va lookup vung nho nhat; 3 test moi, tong 58/58.
- Runtime: `WorkbookResultAuditService` luu chunk/checksum trong custom properties; writer
  `Gia DT TC` va `THKP-TC` ghi audit trong cung giao dich sau verify, truoc commit.
- Audit: 13 dong + 1 block phu luc, 6 K + block tong + dong tien chu; giu package,
  PriceProfile, norm/variant, locator van ban/trang/muc, nguon gia/khoi luong/output.
- Kho payload: sua doc chunk tu O(n^2) COM thanh mot lan quet dictionary; test audit giam
  tu vuot 184 giay xuong khoang 30 giay cho luong ghi/phuc hoi phu luc.
- Kha nang mo lai: metadata noi bo giu line/group/unit-rate-role/binding; khi F:H da la gia tri,
  phu luc tai dung 13-line plan tu audit va van tinh lai duoc.
- UI: them muc `Truy vet`, doc active cell, hien package/profile/norm va bang nguon; metadata
  noi bo khong lam roi bang nguon nguoi dung.
- Test: `test-estimate-appendix` dat audit + plan restore; `test-cost-summary` dat; 
  `test-result-audit` dat 22 entry, VBHN97 trang 32-33, 6 nguon gia, K5, UI,
  active-cell va rename/save/reopen. Regression DT-401..405 dat; 92 loi ke thua khong tang.
- Checkpoint end/start task sau: `DT-502-end.xlsm`/`DT-503-start.xlsm`, 679410 byte,
  SHA-256 `60AFD5F96B102BE99B38E7169B2423D7EE387FA409CCC1BB7DC014BA452F5704`;
  package/profile/role Available va E27=1201557000.
- Quyet dinh: ADR-031. Audit la payload workbook theo CodeName/role, khong gan vao ten sheet.
- Task tiep theo: `DT-503`.

## 2026-08-03 - DT-503 IN_PROGRESS

- Muc tieu: phat hien thieu gia, sai role/mapping, sai tong, stale formula va broken name;
  thong bao dung worksheet role/address va khong sua workbook khi scan.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-503-start.xlsm`, 679410 byte, SHA-256
  `60AFD5F96B102BE99B38E7169B2423D7EE387FA409CCC1BB7DC014BA452F5704`.
- Next action: inventory cac loi co the bo co chu dich va chot contract validation report Core.

## 2026-08-03 - DT-503 DONE

- Core: them `WorkbookValidationReport/Issue`, severity, code, location/remediation,
  inherited warning va deterministic sorting; 2 test moi, tong 60/60.
- Runtime: `WorkbookValidationService` scan role mapping, profile/audit identity, preview
  thieu gia, formula map batch va tong phu luc/THKP; scanner khong ghi workbook.
- Formula contract: hai writer phu luc/THKP expose expected formula map dung chung, nen
  validation va writer khong lech quy tac. Audit block luu dieu kien he so de tai tinh dung.
- Defined name: 4.034 name `#REF!` ke thua duoc gom thanh mot canh bao khong chan;
  name app `TTBMVN_*` hong la loi chan rieng, tranh lam UI treo boi hang nghin dong.
- UI: them muc `Kiem tra`, grid severity/code/role-address/message/remediation va nut
  dieu huong den dung o Excel.
- Fault test: bo `MAT-CONCRETE-STAKE` bao MissingPrice `F13:N13`; go role bao
  ResourcePrices; sua F13 bao IncorrectTotal `I27`; sua K22 bao StaleFormula `K22` va
  dieu huong dung; name `TTBMVN_TEST_BROKEN` duoc phat hien. Tat ca duoc rollback.
- Test: Release x64 sach, 60/60 Core; `test-workbook-validation.ps1 -Overwrite` dat UI,
  5 fault, navigation, restore va save/reopen. Clean checkpoint: 0 loi chan, 1 canh bao.
- Checkpoint end/start task sau: `DT-503-end.xlsm`/`DT-504-start.xlsm`, 679410 byte,
  SHA-256 `60AFD5F96B102BE99B38E7169B2423D7EE387FA409CCC1BB7DC014BA452F5704`;
  workbook khong doi, 92 loi ke thua va E27=1201557000.
- Quyet dinh: ADR-032. Validation read-only, formula contract dung chung va loi ke thua
  khong duoc coi la blocking regression.
- Task tiep theo: `DT-504`.

## 2026-08-03 - DT-504 IN_PROGRESS

- Muc tieu: migration 2021 -> 2025 co preview, cancel, rollback toan khoi va compare report.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-504-start.xlsm`, 679410 byte, SHA-256
  `60AFD5F96B102BE99B38E7169B2423D7EE387FA409CCC1BB7DC014BA452F5704`.
- Next action: rao soat `WorkbookPackageMigrationService` va mo rong transaction cho audit/gia/ket qua.

## 2026-08-03 - DT-504 DONE

- Code: them migration impact/report, snapshot ProjectProfile/PriceProfile/ResultAudit va
  range phu luc/THKP; preview tinh package dich read-only, stale fingerprint guard, writer
  phu luc + THKP + validation trong mot giao dich co backup/rollback.
- Legacy rule: package 2021 khong co hao phi chi tiet nen gia tri nguon doc tu workbook;
  chi package dich phai tinh du, khong suy dien hao phi 2021.
- UI: them `PackageMigrationControl` va muc `Chuyen goi`, hai tab so sanh ket qua/package,
  backup bat buoc va apply chi enable khi target hop le.
- Test: Release x64 sach, 60/60 Core; live 2021 -> 2025 co 21 package change, 6 calculation
  change, 13 dong, 22 audit entry va 7 value row. Cancel + 6 fault phase rollback dung
  fingerprint; restore backup, UI screenshot, shell, validation va save/reopen dat.
- Checkpoint end/start task sau: `DT-504-end.xlsm`/`DT-505-start.xlsm`, 679296 byte,
  SHA-256 `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  package 2.0.1 Available, roles dat, 92 loi ke thua, E27=1201557000.
- Quyet dinh: ADR-033.
- Task tiep theo: `DT-505`.

## 2026-08-03 - DT-505 IN_PROGRESS

- Muc tieu: runtime log co correlation context, support package loc du lieu va full B5 regression.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-505-start.xlsm`, 679296 byte, SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`.
- Next action: rao soat `RuntimeLogger`, catch boundary va `SupportPackage`.

## 2026-08-03 - DT-505 DONE

- Code: runtime log schema co correlation/task/phase/workbook token/role/package, UTC,
  exception HResult/stack, privacy filter va rotation 5 MiB/3 archive. Cac boundary chinh
  DT-401/404/406/407/502/503/504 ghi context cu the.
- Support: 5 file app/workbook aggregate, setting khong workbook field, sanitized runtime
  log va SHA-256 manifest; Machine ID/workbook/sheet/path/license/customer value khong lo.
- Test: controlled COM `0x800A03EC`, context redaction, support manifest va 5 file dat.
  Full `test-b5-gate.ps1`: 8/8 nhom PASS trong 611,1 giay; 60/60 Core, DT-501..504,
  runtime diagnostics va semantic checkpoint dat. Batch 4.000 o nhanh hon 185,74 lan.
- Test isolation: phat hien script audit rename/save input; Gate B5 da doi sang working copy
  rieng va checkpoint start da khoi phuc/dung hash truoc khi chot.
- Checkpoint end/start task sau: `DT-505-end.xlsm`/`DT-601-start.xlsm`, 679296 byte,
  SHA-256 `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  package 2.0.1 Available, roles dat, 92 loi ke thua va E27=1201557000.
- Quyet dinh: ADR-034. Gate B5: `DAT`.
- Task tiep theo: `DT-601`.

## 2026-08-03 - DT-601 IN_PROGRESS

- Muc tieu: goi cap nhat offline co manifest/checksum/chu ky, chan tamper/downgrade/duplicate.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-601-start.xlsm`, 679296 byte, SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`.
- Next action: rao soat package tool/store va chot threat model/trust root update.

## 2026-08-03 - DT-601 DONE

- Code: `.ttbupdate` co manifest UTF-8/LF canonical, detached RSA-SHA256 signature,
  length/SHA-256 tung payload, minimum app version va policy version. Verifier gioi han
  archive/entry, chan path traversal/file thua-thieu-trung, doc lai regulation bundle va
  doi chieu identity truoc khi tra ket qua.
- Trust: public key `TTBMVN-OFFLINE-2026-01` compile trong Core va co public artifact;
  private key non-exportable chi nam trong Windows CAPI container cua publisher tool,
  khong co private update-key file trong workspace.
- Tool: them `keygen`, `build-update`, `verify-update`. Evidence 2025@2.0.1 co SHA-256
  `AC9859DD09DEC61C362A07BA9E703CF67EE1F4D1DEAEFB68C4BC0C9155918FE2`, disposition `Ready`.
- Test: Release x64 sach, 65/65 Core; valid, tamper, wrong signature, unknown trust boundary,
  downgrade, app qua cu, already-installed va cung version khac checksum dat.
- Checkpoint end/start task sau: `DT-601-end.xlsm`/`DT-602-start.xlsm`, 679296 byte,
  SHA-256 `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  package Available, roles dat, 92 loi ke thua va E27=1201557000.
- Quyet dinh: ADR-035; threat model/van hanh tai `OFFLINE-UPDATE.md`.
- Rui ro ghi cho DT-606: `ExcelAddIn1_TemporaryKey.pfx` la key ClickOnce cu, khong phai
  update key; can thay bang chung thuong mai va quy trinh code-signing chinh thuc.
- Task tiep theo: `DT-602`.

## 2026-08-03 - DT-602 IN_PROGRESS

- Muc tieu: Update center offline co danh sach package, preview diff, install va rollback;
  interface online de san nhung khong phu thuoc mang.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-602-start.xlsm`, 679296 byte, SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`.
- Next action: rao soat shell Du toan, bootstrap/store va migration UI de tai su dung cac
  giao dich da co thay vi tao duong cai dat song song.

## 2026-08-03 - DT-602 DONE

- Core: `OfflineUpdateCenterService` inspect/diff/install/activate va provider interface;
  online provider mac dinh Disabled, exception mat mang tra Unavailable. Install xac minh
  lai archive, extract co checksum, import atomic va khong tu dong migrate workbook.
- Rollback: `RegulationPackageActivationStore` luu con tro mac dinh atomic/checksum theo
  PackageId. Version cu/moi cung ton tai; rollback chi doi ban dung cho du an moi, workbook
  da pin van tai dung checksum cu. Migration van thay tat ca package da cai.
- UI: nut `Cap nhat` trong shell, danh sach installed/preferred/current workbook, chon
  `.ttbupdate`, 21 dong diff, install, rollback default va dieu huong sang Chuyen goi.
  Thu muc mo file gan nhat duoc nho; online button hien ro chua bat.
- Test: Release x64 sach, 68/68 Core; valid install, duplicate, restart, activation rollback,
  network unavailable/offline fallback. `test-update-center.ps1 -Overwrite` dat UI/install/
  restart va xac minh workbook pin khong doi; `test-package-migration-ui.ps1` dat 21 diff;
  Project Setup regression dat tren oracle `DT-401-start`.
- Screenshot: `tmp/DT-602-update-center.png`, 830x614, khong overlap/cat text nghiem trong.
- Checkpoint end/start task sau: `DT-602-end.xlsm`/`DT-603-start.xlsm`, 679296 byte,
  SHA-256 `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  package/role dat, 92 loi ke thua va E27=1201557000.
- Quyet dinh: ADR-036.
- Task tiep theo: `DT-603`.

## 2026-08-03 - DT-603 IN_PROGRESS

- Muc tieu: lap va chay ma tran Windows 10/11, Office 2016/2019/2021/365, x86/x64,
  decimal/list separator theo Excel; cau hinh khong co may that phai ghi `NOT_RUN`.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-603-start.xlsm`, 679296 byte, SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`.
- Next action: inventory OS/Office/Excel bitness/locale hien tai va xay compatibility harness.

## 2026-08-03 - DT-603 DONE

- Host that: Windows 11 Pro 10.0.26200 x64, Microsoft 365/O365HomePremRetail
  16.0.20228.20124 x64, Office client `vi-vn`, process culture `en-US`, .NET release 533509.
- Build: `Release|Any CPU` va `Release|x64` deu sach, moi build 68/68 Core.
- Excel live: default decimal/group/list `,`/`.`/`;`; custom `,`/`.` va `.`/`,` deu
  format dung, `FormulaR1C1` SUM/ROUND cho 3.75. Separator duoc khoi phuc sau test.
- Matrix: `COMPATIBILITY.md`; Office 2016/2019/2021, Excel x86 va Windows 10 ghi
  `NOT_RUN`, khong suy dien PASS. Pilot support chi chot Win11 + O365 x64.
- Evidence: `tmp/DT-603-compatibility-current.json`.
- Checkpoint end/start task sau: `DT-603-end.xlsm`/`DT-604-start.xlsm`, 679296 byte,
  SHA-256 `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  package/role dat, 92 loi ke thua, E27=1201557000.
- Task tiep theo: `DT-604`.

## 2026-08-03 - DT-604 IN_PROGRESS

- Muc tieu: nut Ho tro trong shell Du toan va goi chan doan co activation, project profile,
  package manifest, validation report va log da loc du lieu nhay cam.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-604-start.xlsm`, 679296 byte, SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`.
- Next action: rao soat `FrmSupport`, `SupportPackage`, activation/update state va validation.

## 2026-08-03 - DT-604 DONE

- Code: nut `Ho tro` duoc them vao shell Du toan va mo `FrmSupport` gan voi workbook hien
  tai. Form hien thong tin app, Machine ID rut gon, trang thai/han license, kich hoat offline;
  online duoc hien ro la tam tat.
- Support package: 9 file gom 8 payload va `manifest.sha256`; bo sung activation-info,
  project-profile, regulation package manifest va validation report. Workbook/project/
  price-profile/machine duoc token hoa; key, path, sheet name/address va override content
  khong duoc xuat.
- Test: Release x64 sach, 68/68 Core; `test-runtime-diagnostics.ps1` dat 9 file, checksum va
  privacy; `test-estimate-support.ps1` dat constructor gan workbook, activation, shell va
  modal dialog. Anh `tmp/DT-604-estimate-support.png` khong overlap/cat noi dung.
- Checkpoint end/start task sau: `DT-604-end.xlsm`/`DT-605-start.xlsm`, 679296 byte,
  SHA-256 `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  package 2.0.1 Available, roles dat, 92 loi ke thua va E27=1201557000.
- Quyet dinh: ADR-037.
- Task tiep theo: `DT-605`.

## 2026-08-03 - DT-605 IN_PROGRESS

- Muc tieu: mot acceptance harness phat hanh tu dong, co evidence may-doc-duoc cho build,
  clean install, upgrade, old workbook, migration va uninstall/reinstall.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-605-start.xlsm`, 679296 byte, SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`.
- Next action: rao soat cac script gate hien co va thiet ke isolated acceptance store de
  khong go cai dat add-in dang dung cua nguoi dung.

## 2026-08-03 - DT-605 DONE

- Harness: them `test-release-lifecycle.ps1`, `test-old-workbook-open.ps1` va
  `test-release-acceptance.ps1`; evidence JSON co checkpoint hash, tung test va duration.
  Resume chi dung prefix PASS va bi tu choi neu checkpoint hash thay doi.
- Fixture: luu package lich su hop le `BQP-RPBM-2025@2.0.0` checksum `D58CAC...` trong
  `Dutoanmau/Fixtures`; upgrade bang `.ttbupdate` ky sang 2.0.1, restart/gỡ/reinstall dat.
- Old workbook: DT-101 chua role/profile mo duoc; first-run de xuat du 7 role, E27 cu dung
  va hash working copy khong doi khi dong khong luu.
- Full acceptance: 14/14 PASS trong 1044,4 giay. Build 68 Core, batch, DT-405/406/407 oracle,
  audit, validation, migration 6 fault, update center, diagnostics/support va semantic gate dat.
- Evidence: `tmp/DT-605-release-acceptance.json`, `tmp/DT-605-release-lifecycle.json` va
  `docs/du-toan/RELEASE-ACCEPTANCE.md`.
- Checkpoint end/start task sau: `DT-605-end.xlsm`/`DT-606-start.xlsm`, 679296 byte,
  SHA-256 `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  package/role dat, 92 loi ke thua va E27=1201557000.
- Quyet dinh: ADR-038.
- Task tiep theo: `DT-606`.

## 2026-08-03 - DT-606 IN_PROGRESS

- Muc tieu: tao bo artifact pilot co version, publish/manifest, checksum, installer docs,
  uninstall/update docs va release notes; chot ro tinh trang code-signing.
- Nguoi thuc hien: Codex.
- Checkpoint start: `DT-606-start.xlsm`, 679296 byte, SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`.
- Next action: inventory VSTO publish prerequisites, certificate va current installation;
  khong go add-in dang dung trong Excel PID 3208.

## 2026-08-03 - DT-606 DONE

- Signing: key tam het han `BD5A...` duoc bo khoi project va chuyen ra ngoai workspace.
  Build/publish lay certificate-store thumbprint; co the dung CA qua
  `TTBMVN_SIGNING_CERT_THUMBPRINT`. Pilot certificate RSA 3072/code-signing, private key
  non-exportable; publisher `CN=TTBMVN Pilot Publisher`, thumbprint `A659D6...E328C`.
- Publish: them `signing-certificate.ps1`, `publish-release.ps1`, installer/uninstaller va
  release verifier. `setup.exe` co Authenticode signer; VSTO deployment/application manifest
  RSA-SHA256 deu duoc `mage.exe` xac minh hop le. Release khong chua PFX/private key.
- Artifact: `artifacts/TTBMVN-Excel-Tools-1.0.0-PILOT-x64.zip`, SHA-256
  `7F3FBE1B65793B130A45A2462F0D16C03C84452095F7FC0AAB92BBFAA977D0F9`;
  gom 37 file, package update 2.0.1, public update key, workbook mau va tai lieu cai-go.
- Gate: Release x64 + 68/68 Core, checksum inventory, manifest signatures, install/uninstall
  VerifyOnly, tamper rejection, DT-605 acceptance 14/14 va semantic checkpoint deu PASS.
- Gioi han trung thuc: setup signer status `UnknownError/UntrustedRoot` tren may chua import
  pilot cert la ky vong; day khong phai CA commercial. Actual VSTO mutation khong chay vi
  Excel PID 3208 cua nguoi dung dang mo; script va preflight da dat, pilot deployment can may sach.
- Checkpoint cuoi: `DT-606-end.xlsm`, 679296 byte, SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  package/role dat, 92 loi ke thua va E27=1201557000.
- Quyet dinh: ADR-039. Gate B6 pilot: `DAT`. Toan bo DT-101 den DT-606: `DONE`.

## Mau ghi task

```text
## YYYY-MM-DD - DT-xxx STATUS

- Muc tieu:
- File/code da thay doi:
- Workbook/checkpoint:
- Test command:
- Test result:
- Expected/actual quan trong:
- Loi con lai/blocker:
- Quyet dinh moi:
- Task tiep theo:
```
