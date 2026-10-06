# Ban giao hien tai

Cap nhat: 2026-08-03.

## Doc dau tien

1. `docs/du-toan/README.md`.
2. `docs/du-toan/PROGRESS.md`.
3. Muc `DT-606` trong `docs/du-toan/TASKS.md`.
4. `docs/du-toan/TESTING.md`.

## Trang thai code

- Solution: `ExcelAddIn1.sln`.
- VSTO add-in: `ExcelAddIn1` tren .NET Framework 4.8.1.
- Core logic: `ExcelAddIn1.Core`.
- Console tests: `ExcelAddIn1.Tests`.
- License generator: `ExcelAddIn1.LicenseTool`.
- Build `Release|x64` ngay 2026-08-03: thanh cong.
- 68 test hien co: tat ca dat.
- `DT-101` da DONE; baseline va expected output nam trong `docs/du-toan/BASELINE.md`.
- `DT-102` da DONE; bay worksheet role da luu vao working workbook.
- `DT-103` da DONE; ProjectProfile schema 1 da luu vao working workbook.
- `DT-104` da DONE; sheet list coordinator da tich hop vao FrmDaodat.
- `DT-105` da DONE; wizard mapping va gate B1 da dat.
- `DT-201` da DONE; RegulationPackage schema 1, validator va serializer deterministic.
- `DT-202` da DONE; package store staging/atomic install/version conflict/rollback.
- `DT-203` da DONE; resolver ngay hieu luc va TT101 Dieu 6 transition.
- `DT-204` da DONE; ProjectProfile schema 2 va workbook pin package identity day du.
- `DT-205` da DONE; diff/migration co cancel, backup, rollback va apply/reopen; Gate B2 dat.
- `DT-301` da DONE; package BQP 2021 co 198 record, build/validate/import store va hai luot review dat.
- `DT-302` da DONE; package BQP 2025 co 248 record, 6 PDF, record diff va review doc lap dat.
- `DT-303` da DONE; engine gia ca may decimal, hai doi tuong M010/M011 va 10 tong
  `VL-NC-M` da doi chieu dat.
- `DT-304` da DONE; engine dinh muc/chi phi decimal, K1-K6, 12 o dinh muc va bang K2/K5
  workbook da doi chieu dat.
- `DT-305` da DONE; independent audit 6534 check, zero error/warning va Gate B3 dat.
- `DT-401` da DONE; project setup hop nhat profile/package/moc gia/mapping, cancel no-write
  va rollback giao dich; bundled package 2021/2025 duoc cai vao local store.
- `DT-402` da DONE; NormSearch exact/fuzzy/khong dau/filter, package identity va UI tra cuu
  trong shell Du toan; package 2021 khong co hao phi chi tiet duoc canh bao, khong suy dien.
- `DT-403` da DONE; CostRuleEngine K1-K6 theo package pin, override bat buoc ly do,
  UI Chi phi va adapter TT123/2021 khong ap muc toi thieu cua VBHN 2025.
- `DT-404` da DONE; PriceProfile deterministic/store/snapshot, import 34 gia legacy,
  coverage khong zero-fill, batch-write 21 gia vat lieu va UI Bang gia da qua save/reopen.
- `DT-405` da DONE; engine don gia tu NormCatalog + PriceProfile, dieu kien/he so chon ro
  rang, binding ma may logic bat buoc ly do va UI `Don gia` da qua live Excel/save-reopen.
- `DT-406` da DONE; engine phu luc, import 13 dong legacy, preview goc/co he so, writer batch
  co rollback va UI `Phu luc DT` da qua thay input/save-reopen. Ban va package 2.0.1 sua co do
  lon va hao phi may row 22 theo VBHN97/2025; tong phap ly moi da duoc chot.
- `DT-407` da DONE; engine tong hop 4 bieu mau, dung co so T/Z cho K1-K6, TL/VAT theo
  bieu 04, lam tron VND/1.000 VND va tien bang chu. UI `Tong hop KP`, batch writer,
  input recalculate/save-reopen va regression DT-401..406 dat; Gate B4 dat.
- `DT-501` da DONE; `ExcelBatchWriteTransaction` dung chung, `ExcelWriteContext` phuc hoi
  ca khi constructor/ghi loi, cac writer phu luc/tong hop/bang gia va duong nong ho dao dung
  batch. Benchmark chot 4.000 o nhanh hon 142,47 lan; rollback/state/value parity deu dat.
- `DT-502` da DONE; 22 audit entry gan package/profile/norm/source vao 13 dong phu luc,
  khoi tong va K1-K6. UI `Truy vet` doc theo active cell; audit song qua rename/save/reopen
  va phuc hoi duoc plan phu luc sau khi F:H da chuyen thanh gia tri.
- `DT-503` da DONE; scanner batch phat hien role, profile/audit stale, thieu gia, sai tong,
  stale formula va broken name; UI `Kiem tra` dieu huong den cell. Nam fault test va
  save/reopen dat; 4.034 name ke thua duoc gom thanh mot canh bao khong chan.
- `DT-504` da DONE; preview 2021 -> 2025 tinh target truoc khi ghi, giao dich snapshot
  profile/gia/audit/range, 6 fault phase rollback dung fingerprint, backup restore, UI va
  save/reopen dat. Package 2021 thieu hao phi chi tiet duoc xu ly nhu legacy source, khong
  gia lap kha nang tai tinh.
- `DT-505` da DONE; runtime log co correlation/task/phase/workbook token/role/package,
  HResult va redaction. Support package 5 file khong chua Machine ID, workbook/sheet name,
  path, license key hay workbook-bound setting. Full Gate B5 8/8 nhom dat.
- `DT-601` da DONE; `.ttbupdate` co manifest/hash/RSA signature, trust root chi public key,
  private key non-exportable trong publisher CAPI container. Security matrix va evidence
  2025@2.0.1 ky boi `TTBMVN-OFFLINE-2026-01` deu dat.
- `DT-602` da DONE; Update center inspect/diff/install/activation rollback/restart,
  provider online Disabled va UI shell dat. Cai package khong tu dong doi workbook pin.
- `DT-603` da DONE; Win11/O365 x64 va hai Excel separator scenario dat, AnyCPU/x64 build
  dat. Office cu/x86/Win10 la `NOT_RUN`, pham vi pilot chua tuyen bo ho tro cac host nay.
- `DT-604` da DONE; shell Du toan co nut Ho tro, dialog gan workbook va support package
  9 file co activation/project/package/validation da sanitize cung manifest checksum.
- `DT-605` da DONE; acceptance 14/14 PASS trong 1044,4 giay, co lifecycle store, old
  workbook, B4/B5, migration, update/support va evidence JSON.
- `DT-606` da DONE; Release 1.0.0 PILOT x64 co manifest/setup signing, checksum/tamper gate,
  installer docs va ZIP artifact. Toan bo DT-101 den DT-606 da hoan thanh.
- COM ownership da sua: cac service dung `ReleaseComObject` mot lan, khong dung
  `FinalReleaseComObject` lam vo hieu RCW worksheet/collection dang duoc caller giu.
- Live Excel test lay COM instance moi va xac minh PID rieng, fallback `/x`, resolve workbook
  theo full path va dung working copy;
  khong con dong/mo workbook lon trong Excel giao dien.
- Cong cu package: `ExcelAddIn1.RegulationTool`.
- Thu muc khong phai Git repository; khong co commit de khoi phuc.

Lenh build da xac minh:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
```

## Workbook test

- Full path: `D:\VSstudio\EXCEL_TTBMVN\Dutoanmau\00. Du toan TP 3.xlsm`.
- Nguoi dung xac nhan day la file rieng cho test va cho phep sua.
- Baseline: `Dutoanmau/Baseline/00. Du toan TP 3.baseline.xlsm`.
- Checkpoint DT-101 start/end co cung SHA-256
  `E4BFDBD5FFA3AB47835A59D794F8E907633820F91E2FDF18EC9696559C89EADA`.
- Baseline da mo lai bang Excel va xac minh `THKP-TC!E27=1198731000`.

## Trang thai ban giao

Khong con task `IN_PROGRESS`. Toan bo lo trinh da `DONE` cho kenh PILOT.

Hanh dong van hanh tiep theo, khong phai task con do:

1. Chay actual install/uninstall tren mot may pilot sach da dong Excel.
2. Doi chieu publisher thumbprint va luu bien ban pilot deployment.
3. Mua CA code-signing + timestamp neu phat hanh thuong mai cong khai.
4. Mo rong compatibility matrix sang Office cu/x86/Windows 10 truoc khi tuyen bo ho tro.

## Rui ro can nho

- Workbook co nhieu defined name va cong thuc cu; khong sua hang loat ngoai pham vi task.
- Code du toan hien tai co ten sheet hard-code va bang he so hard-code.
- Form dang mo khong duoc reload lam mat input khi workbook event xay ra.
- Moi logic tinh toan moi phai tach khoi Excel Interop de unit test.
- Hieu luc van ban va quy dinh chuyen tiep phai duoc model hoa; khong su dung mot goi `latest` duy nhat.
- Checkpoint B1 moi nhat: `Dutoanmau/Checkpoints/DT-105-end.xlsm`, SHA-256
  `2090595D4BB3F8A4FA59EEF6310F1347437C9C314D9540F7AD1625C8093EEF27`.
- Checkpoint DT-201 start/end cung SHA-256 B1 vi task khong ghi workbook.
- Checkpoint DT-202 start/end cung SHA-256 B1 vi task khong ghi workbook.
- Checkpoint DT-203 start/end cung SHA-256 B1 vi task khong ghi workbook.
- Checkpoint DT-204 start SHA `2090595D...3EEF27`; end SHA
  `B8D0EB86CDC75F5A7AF5FA06FBB5F3CBDF3F869EC5A78957A3A5072FFAF7CF46`.
- Checkpoint DT-205 end SHA
  `43E11CE892840E6CA2BF784A79ED19072D4014DEA0B807E6EF23569EA9CF4F39`;
  package v2, profile/role/E27 valid sau reopen.
- Package DT-301 checksum
  `BBE95EFDF3B2DA0780CAFEEF828C9418644452655C733C1BF2417F234C3E8008`;
  audit tai `data/regulations/packages/BQP-RPBM-2021/audit/DT-301-review.md`.
- Checkpoint DT-301 end SHA
  `AA3013E8AD420F97960A3CCF775F69066D41C14899D3B7089D7413D3C6D327E1`;
  92 o loi ke thua, role/profile/package Available va E27 dung baseline.
- Package 2025 hien tai sau DT-303 checksum
  `D58CAC64422E14151FD68F1C92944FAB7A9859908165C59F92089E7F13998B7C`;
  audit tai `data/regulations/packages/BQP-RPBM-2025/audit/DT-303-machine-rate-review.md`.
- Checkpoint DT-302 end SHA
  `1AB5B6487827DFDD55803BCF602C880598414B1C42F17B9AAFA891710EB79272`.
- Checkpoint DT-303 start SHA
  `2780D4F0D3419FEF7EA0539FA09C6DEE6DBCFC399F5A02883339B395F8C826EC`;
  92 loi ke thua, role/profile/package Available va E27 dung baseline.
- Checkpoint DT-303 end SHA
  `EF55831509852668DA6F154D7751D239DC53F230B906A56CB799FF0D03C8CC4C`;
  92 loi ke thua, role/profile/package Available va E27 dung baseline.
- Checkpoint DT-304 start SHA
  `18C240708AA4F3031758D686F0DE5D74AC4E24ECD3CA2C67867E70523B54320E`;
  semantic checkpoint dat.
- Checkpoint DT-304 end va DT-305 start SHA
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  package Available, role/profile dat, 92 loi ke thua va E27 dung baseline.
- Gate B3 package checksum
  `C23A3AE526841473AEAB6E36CCC76B0663A56C5FE63D3351F7AFBFA4639447F2`;
  audit tai `data/regulations/packages/BQP-RPBM-2025/audit/DT-305-gate-review.md`.
- Checkpoint DT-305 end SHA
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  semantic gate dat.
- Checkpoint DT-401 start cung SHA voi DT-305 end; profile/role/package Available,
  92 loi ke thua va E27 dung baseline.
- Checkpoint DT-401 end SHA
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  39 core, form smoke, cancel/rollback, untagged/renamed/save-reopen va semantic gate dat.
- Checkpoint DT-402 start/end SHA
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  41 core, UI/live lookup package 2021/2025 va semantic gate dat.
- Checkpoint DT-403 start/end SHA
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  43 core, CostRule 2021/2025, override/UI/live Excel va regression DT-401/402 dat.
- Checkpoint DT-404 start/end va DT-405 start SHA
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  47 core, PriceProfile UI/live Excel, batch material, save/reopen, regression DT-401..403
  va semantic gate dat.
- Checkpoint DT-405 end va DT-406 start SHA
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  50 core, UnitRate can/nuoc, he so, binding may, missing price, UI/live Excel,
  save/reopen, regression DT-401..404 va semantic gate dat.
- Package hien hanh `BQP-RPBM-2025@2.0.1`, checksum
  `E2537C08E7156414585BEFC446E74DD51E79AB8D67A3774B1F5250E7EECCDA90`.
  Checksum 2.0.0 trong bien ban DT-305 la bang chung Gate B3 lich su, khong phai bundle moi.
- Checkpoint DT-406 end va DT-407 start SHA
  `288324623F921259C5633A4180FC9F8062AD138DA0A37718982C48F09C247B22`;
  package/profile/role dat, 92 loi ke thua, `Gia DT TC!K27=558860046.12222` va
  `THKP-TC!E27=1199577000`.
- Checkpoint DT-407 end va DT-501 start SHA
  `1D878CA03CFA0B623FBA75C59C9B7E1D9E7F9163EDEEC19BBFA295FB551D2325`;
  55 core, package/profile/role dat, 92 loi ke thua, `THKP-TC!E27=1201557000` va Gate B4 dat.
- Checkpoint DT-501 end va DT-502 start: 673693 byte, cung SHA-256
  `1D878CA03CFA0B623FBA75C59C9B7E1D9E7F9163EDEEC19BBFA295FB551D2325`;
  batch transaction/state/performance dat, workbook khong doi nghiep vu.
- Checkpoint DT-502 end va DT-503 start: 679410 byte, cung SHA-256
  `60AFD5F96B102BE99B38E7169B2423D7EE387FA409CCC1BB7DC014BA452F5704`;
  22 audit entry va semantic checkpoint dat, 92 loi ke thua, E27=1201557000.
- Checkpoint DT-503 end va DT-504 start khong doi workbook: 679410 byte, cung SHA-256
  `60AFD5F96B102BE99B38E7169B2423D7EE387FA409CCC1BB7DC014BA452F5704`;
  validation gate dat.
- Checkpoint DT-504 end va DT-505 start: 679296 byte, cung SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  60 core, migration/rollback/UI/restore dat, 22 audit entry, 92 loi ke thua va
  `THKP-TC!E27=1201557000`.
- Checkpoint DT-505 end va DT-601 start khong doi workbook: 679296 byte, cung SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  Gate B5 8/8 nhom dat, runtime/support privacy dat, 22 audit entry, 92 loi ke thua va
  `THKP-TC!E27=1201557000`.
- Checkpoint DT-601 end va DT-602 start khong doi workbook: 679296 byte, cung SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  65 Core/security tests, signed production evidence va semantic checkpoint dat.
- Checkpoint DT-602 end va DT-603 start khong doi workbook: 679296 byte, cung SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  68 Core, Update center UI/install/restart, migration regression va semantic gate dat.
- Checkpoint DT-603 end va DT-604 start khong doi workbook: 679296 byte, cung SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  compatibility current-host/separator/build matrix va semantic gate dat.
- Checkpoint DT-604 end va DT-605 start khong doi workbook: 679296 byte, cung SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  support package/privacy/UI va semantic gate dat.
- Checkpoint DT-605 end va DT-606 start khong doi workbook: 679296 byte, cung SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  acceptance 14/14 PASS va semantic gate dat.
- Checkpoint DT-606 end: 679296 byte, SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  Release Gate B6 PILOT dat. ZIP SHA-256
  `7F3FBE1B65793B130A45A2462F0D16C03C84452095F7FC0AAB92BBFAA977D0F9`.

## DT-701 - Estimate Workspace

- Da thay luong Phu luc cu bang workspace quet tu range Excel, mapping cot va gan dinh muc
  chi tiet tren app. Dong khong gan dinh muc duoc coi la van ban va khong tham gia tinh.
- Don gia duoc gom theo ngu canh day du; cung norm nhung khac variant, can/nuoc,
  HLNS/KHLNS, profile, dieu kien hoac binding se tao don gia rieng.
- Workbook luu hai PriceProfile HLNS/KHLNS, source ID cua sheet, hidden row names va workspace
  payload. Doi ten sheet va save/reopen da test.
- Writer tao dung cac sheet VL-NC-M/DG dang dung, ghi batch va link formula ve Phu luc.
  Chay lai lan hai da test khong tao trung.
- UI `Phu luc DT` co grid cong tac va panel chi tiet; UI `Don gia` hien cac don gia duy nhat
  va hao phi cua dong dang chon.
- Lenh nghiem thu: `scripts/test-estimate-workspace.ps1 -Overwrite`.
- Bang chung: 71/71 Core, live PASS, anh trong `tmp/DT-701-*.png`.
- Thu muc khong co `.git`; phai bao toan cac file nguoi dung dang sua khi tiep tuc.

## Bo cai 1.0.1

- Release: `artifacts/TTBMVN-Excel-Tools-1.0.1-PILOT-x64`.
- ZIP: `artifacts/TTBMVN-Excel-Tools-1.0.1-PILOT-x64.zip`.
- SHA-256: `B4E498F209B8F07D50D4F2F145C20C9224ADB8807B97E63AB86422FA3EA608C0`.
- Nguoi dung giai nen va nhap dup `Cai-dat-TTBMVN.cmd`; khong mo `.vsto` truc tiep khi
  dang co ban cai tu vi tri khac.
- Release gate PASS: 38 file, manifest/install/uninstall/tamper/checkpoint deu dat.

## Hotfix 1.0.2

- Sua SelectionChanged chay trong luc DataGridView clear/add lam `RowDraft` null.
- Quet duoc bang hai tang header; bo qua dong header phu khong co du lieu chinh.
- Cot nghiem thu tuy chon xu ly text/dau gach/ma loi Excel nhu o trong.
- Regression doc truc tiep `Gia DT TC!A9:L27` va reload grid khi control dang hien thi PASS.
- ZIP: `artifacts/TTBMVN-Excel-Tools-1.0.2-PILOT-x64.zip`.
- SHA-256: `4026F6785039756FE506F62A2447A077B2CF4DF797D34F68CEB4356BB5AA3902`.
- Release gate PASS; setup signature va hai manifest deu Valid.

## Modeless 1.0.3

- Ribbon giu mot `Dutoan` theo COM identity cua workbook; bam lai chi activate/bring-to-front.
- Form chinh modeless; dialog anh xa/xac nhan van modal.
- `WorkbookSheetChangeCoordinator.WorkbookClosing` dong form khi workbook tuong ung dong.
- Live test: form visible, Excel active cell B2, close-with-workbook va support regression PASS.
- Script: `scripts/test-dutoan-modeless.ps1`.
- ZIP: `artifacts/TTBMVN-Excel-Tools-1.0.3-PILOT-x64.zip`.
- SHA-256: `29903B8F2F237CEA0BC664548EA92EEC49BA6E526CC0419C94192860278F1FF1`.
- Release gate PASS; setup signature va hai manifest deu Valid.

## Installer stale-registration 1.0.4

- Nguyen nhan -401: registry tro vao `bin/x64/Release/ExcelAddIn1.vsto|vstolocal`, nhung
  customization khong co trong VSTO ClickOnce cache de `/Uninstall`.
- Installer nhan dien duong `bin`, sao luu registry vao `InstallBackups`, xoa dung khoa
  `ExcelAddIn1` va cai ban publish. Neu install fail thi khoi phuc khoa cu.
- Uninstaller cung xu ly debug registration, manifest mat va exit -401.
- May test con ClickOnce subscription `TTBMVN Excel Tools 1.0.1.0` trong HKCU Uninstall;
  installer 1.0.4 doc `UrlUpdateInfo` cua entry nay va go subscription truoc khi cai moi.
