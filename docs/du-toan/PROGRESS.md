# Tien do chinh thuc

Cap nhat lan cuoi: 2026-10-06.

## Tong quan hoan thanh

- Dieu phoi + baseline + B1-B7: `35/35 task DONE` - `100%`.
- Rieng task B1-B7: `33/33 task DONE` - `100%`.
- Giai doan hien tai: Sua installer debug/stale registration, chuan bi ban `1.0.4 PILOT x64`.

| Giai doan | DONE/Tong | Hoan thanh | Gate |
|---|---:|---:|---|
| Khoi dong | 2/2 | 100% | Dat |
| B1 | 4/4 | 100% | Dat |
| B2 | 5/5 | 100% | Dat |
| B3 | 5/5 | 100% | Dat |
| B4 | 7/7 | 100% | Dat |
| B5 | 5/5 | 100% | Dat |
| B6 | 6/6 | 100% | Dat |
| B7 | 1/1 | 100% | Dat |

## Task hien tai

| Truong | Gia tri |
|---|---|
| Current task | `NONE` |
| Ten | Hoan thanh DT-101 den DT-701 |
| Trang thai | `DONE` |
| Nguoi thuc hien | Codex |
| Bat dau | 2026-08-03 |
| Blocker | Khong co |
| Next action | Cai `1.0.4 PILOT x64`, xac minh nang cap tu dang ky Visual Studio stale |

## Bang tien do

Trang thai hop le: `NOT_STARTED`, `IN_PROGRESS`, `BLOCKED`, `DONE`.

| Task | Trang thai | Test gate | Ghi chu |
|---|---|---|---|
| DT-000 | DONE | Build Release x64 + 5 core tests dat | Tao bo tai lieu dieu phoi |
| DT-101 | DONE | Baseline/checkpoint/hash + Excel reopen + expected outputs | SHA-256 E4BFDBD5...89EADA |
| DT-102 | DONE | 8 core tests + live rename/save/reopen | 7 role trong Worksheet.CustomProperties |
| DT-103 | DONE | 11 core tests + chunk/checksum + save/reopen | ProjectProfile schema 1 |
| DT-104 | DONE | 12 core tests + live add/rename/activate/delete | State form duoc bao toan |
| DT-105 | DONE | 13 core tests + 3 workbook variants + save/reopen | Gate B1 dat |
| DT-201 | DONE | 16 core tests + B1 regression | Package model/serializer/checksum |
| DT-202 | DONE | 19 core tests + B1 regression | Atomic package store |
| DT-203 | DONE | 21 core tests + legal boundary + B1 regression | Effective-date resolver |
| DT-204 | DONE | 22 core tests + live save/reopen/machine-store | Workbook package pin |
| DT-205 | DONE | 23 core tests + live cancel/rollback/apply/reopen | Gate B2 dat |
| DT-301 | DONE | 27 core tests + PDF/workbook review + package store | 198 record, package 2021 |
| DT-302 | DONE | 29 core tests + 6 PDF + record diff + B1 regression | 248 record, package 2025 |
| DT-303 | DONE | 32 core + PDF + 10 workbook totals | May va don gia ca may |
| DT-304 | DONE | 38 core + PDF/workbook/regression | Dinh muc va chi phi |
| DT-305 | DONE | 6534 independent checks + full regression | Gate B3 dat |
| DT-401 | DONE | 39 core + form/live Excel transaction | Project setup |
| DT-402 | DONE | 41 core + UI/live Excel | Tra cuu dinh muc |
| DT-403 | DONE | 43 core + UI/live Excel + regression DT-401/402 | Chi phi theo package da pin |
| DT-404 | DONE | 47 core + UI/live Excel + regression DT-401..403 | PriceProfile, coverage va batch material write |
| DT-405 | DONE | 50 core + UI/live Excel + regression DT-401..404 | DG Can/DG Nuoc, he so va binding may ro rang |
| DT-406 | DONE | 52 core + UI/live Excel + regression DT-401..405 | Gia DT TC, batch writer va baseline phap ly |
| DT-407 | DONE | 55 core + UI/live Excel + regression DT-401..406 | THKP-TC va Gate B4 dat |
| DT-501 | DONE | 55 core + benchmark/rollback/state/regression | Batch Excel writes |
| DT-502 | DONE | 58 core + live audit/UI/save-reopen/regression | Truy vet can cu |
| DT-503 | DONE | 60 core + 5 fault/live UI/save-reopen | Kiem tra sai lech |
| DT-504 | DONE | 60 core + live 2021->2025/6 fault/UI/restore | Migration toan khoi |
| DT-505 | DONE | 8 nhom regression + COM/support privacy | Gate B5 dat |
| DT-601 | DONE | 65 core + signed evidence/security/checkpoint | Goi cap nhat offline |
| DT-602 | DONE | 68 core + UI/install/restart/migration/checkpoint | Update center |
| DT-603 | DONE | Win11/O365 x64 + separator/build matrix | Office/Windows/locale |
| DT-604 | DONE | Support package/privacy/UI/checkpoint | Goi chan doan 9 file |
| DT-605 | DONE | 14/14 acceptance PASS | Evidence JSON + bien ban |
| DT-606 | DONE | Build/publish/sign/hash/tamper/checkpoint PASS | Gate B6 pilot dat |
| DT-701 | DONE | 71 core + live Excel/UI/save-reopen/idempotent writer | Estimate Workspace va don gia theo ngu canh |

## Baseline ky thuat hien tai

- Build `Release|x64`: dat ngay 2026-08-03.
- Test hien co: 68/68 dat; gom ho dao/license, B1, package, gia ca may, dinh muc, chi phi,
  PriceProfile, don gia, phu luc, tong hop kinh phi va bao mat goi update offline.
- Gate B2: `DT-205-end.xlsm`, SHA-256
  `43E11CE892840E6CA2BF784A79ED19072D4014DEA0B807E6EF23569EA9CF4F39`;
  package v2 pin dung sau reopen, 92 o loi ke thua khong tang va `THKP-TC!E27=1198731000`.
- Repository duoc quan ly tai `https://github.com/HoangHung997/TTBMVN`, nhanh `main`.
- Baseline/checkpoint DT-101: 667469 byte; SHA-256
  `E4BFDBD5FFA3AB47835A59D794F8E907633820F91E2FDF18EC9696559C89EADA`.
- Baseline mo lai bang Excel, `THKP-TC!E27=1198731000`; dong khong luu va checksum khong doi.
- Expected output va loi workbook ke thua: `docs/du-toan/BASELINE.md`.
- Gate B1: `DT-105-end.xlsm`, SHA-256
  `2090595D4BB3F8A4FA59EEF6310F1347437C9C314D9540F7AD1625C8093EEF27`;
  role/profile valid va `THKP-TC!E27=1198731000`.
- DT-301 package checksum:
  `BBE95EFDF3B2DA0780CAFEEF828C9418644452655C733C1BF2417F234C3E8008`;
  198 record, raw PDF hash/PDF review/workbook review/package store dat.
- DT-301 end: `DT-301-end.xlsm`, SHA-256
  `AA3013E8AD420F97960A3CCF775F69066D41C14899D3B7089D7413D3C6D327E1`;
  profile/role/package pin valid, 92 o loi ke thua va `THKP-TC!E27=1198731000`.
- Package 2025 hien tai sau DT-305:
  `C23A3AE526841473AEAB6E36CCC76B0663A56C5FE63D3351F7AFBFA4639447F2`;
  248 record, MachineRate bao phu ca Bang 01/03 va 248/248 thay doi co ly do/locator.
- Ban va du lieu hien hanh `BQP-RPBM-2025@2.0.1`:
  `E2537C08E7156414585BEFC446E74DD51E79AB8D67A3774B1F5250E7EECCDA90`;
  sua co do lon cua `NORM-020.0500/.1000`, audit doc lap 6534 check dat. Checksum 2.0.0
  o tren van duoc giu lam bang chung Gate B3 lich su.
- DT-302 end: `DT-302-end.xlsm`, SHA-256
  `1AB5B6487827DFDD55803BCF602C880598414B1C42F17B9AAFA891710EB79272`;
  profile/role/package pin valid, 92 o loi ke thua va `THKP-TC!E27=1198731000`.
- DT-303 start: `DT-303-start.xlsm`, SHA-256
  `2780D4F0D3419FEF7EA0539FA09C6DEE6DBCFC399F5A02883339B395F8C826EC`;
  semantic checkpoint dat.
- DT-303 end: `DT-303-end.xlsm`, SHA-256
  `EF55831509852668DA6F154D7751D239DC53F230B906A56CB799FF0D03C8CC4C`;
  profile/role/package Available, 92 loi ke thua va `THKP-TC!E27=1198731000`.
- DT-304 start: `DT-304-start.xlsm`, SHA-256
  `18C240708AA4F3031758D686F0DE5D74AC4E24ECD3CA2C67867E70523B54320E`;
  semantic checkpoint dat.
- DT-304 end va DT-305 start: SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  profile/role/package Available, 92 loi ke thua va `THKP-TC!E27=1198731000`.
- Gate B3: independent audit 6534 check, 248 record/reason, 6 PDF, zero error/warning;
  package checksum `C23A3AE526841473AEAB6E36CCC76B0663A56C5FE63D3351F7AFBFA4639447F2`.
- DT-305 end: `DT-305-end.xlsm`, 670239 byte, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  semantic checkpoint dat.
- DT-401 start: `DT-401-start.xlsm`, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  semantic checkpoint dat.
- DT-401 end: `DT-401-end.xlsm`, 670239 byte, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  form smoke, cancel no-write, fault rollback, untagged/renamed mapping va save/reopen dat;
  92 loi ke thua, `THKP-TC!E27=1198731000`.
- DT-402 start/end: 670239 byte, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  41 core, exact/fuzzy/filter/locale, package 2021/2025, UI shell va save/reopen dat;
  92 loi ke thua, `THKP-TC!E27=1198731000`.
- DT-403 start/end: 670239 byte, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  43 core, CostRule 2021/2025, K1-K6, override co ly do, UI shell va save/reopen dat;
  regression DT-401/402 dat, 92 loi ke thua va `THKP-TC!E27=1198731000`.
- DT-404 start/end va DT-405 start: 670239 byte, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  47 core, import 34 gia legacy, bao 39 gia thieu, batch-write 21 vat lieu, override,
  UI/save-reopen va regression DT-401..403 dat; 92 loi ke thua va
  `THKP-TC!E27=1198731000`.
- DT-405 end va DT-406 start: 670239 byte, SHA-256
  `8E3486BEAA652EE0FA9455A874A6426E6637C3478B84C18B9FF56E7A8C71CF51`;
  50 core, can/nuoc, bat/tat he so, binding may co ly do, missing price, locale,
  UI/save-reopen va regression DT-401..404 dat; 92 loi ke thua va
  `THKP-TC!E27=1198731000`.
- DT-406 end va DT-407 start: 673771 byte, SHA-256
  `288324623F921259C5633A4180FC9F8062AD138DA0A37718982C48F09C247B22`;
  52 core, import 13 dong, preview, batch write/rollback, input recalculate, UI/save-reopen
  va regression DT-401..405 dat. Package `2.0.1`, 92 loi ke thua;
  `Gia DT TC!K27=558860046.12222`, `THKP-TC!E27=1199577000` theo sua loi phap ly row 22.
- DT-407 end va DT-501 start: 673693 byte, SHA-256
  `1D878CA03CFA0B623FBA75C59C9B7E1D9E7F9163EDEEC19BBFA295FB551D2325`;
  55 core, import cau hinh legacy, 4 bieu mau, batch write/rollback, input recalculate,
  UI/save-reopen va regression DT-401..406 dat; 92 loi ke thua, K5 dung 2.598% tren Z va
  `THKP-TC!E27=1201557000`. Gate B4 dat.
- DT-501 end va DT-502 start: 673693 byte, cung SHA-256
  `1D878CA03CFA0B623FBA75C59C9B7E1D9E7F9163EDEEC19BBFA295FB551D2325`;
  transaction batch dung chung, exception rollback, phuc hoi 4 trang thai Excel va benchmark
  4.000 o dat. Workbook khong doi nghiep vu; 92 loi ke thua va E27 giu nguyen.
- DT-502 end va DT-503 start: 679410 byte, cung SHA-256
  `60AFD5F96B102BE99B38E7169B2423D7EE387FA409CCC1BB7DC014BA452F5704`;
  22 audit entry, UI/active-cell/rename/save-reopen/plan restore dat; 92 loi ke thua va
  `THKP-TC!E27=1201557000`.
- DT-503 end va DT-504 start khong doi workbook: 679410 byte, cung SHA-256
  `60AFD5F96B102BE99B38E7169B2423D7EE387FA409CCC1BB7DC014BA452F5704`;
  validation fault matrix dat, 0 loi chan va 1 canh bao tong hop name ke thua.
- DT-504 end va DT-505 start: 679296 byte, cung SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  migration BQP 2021 -> 2025 preview 21 thay doi/7 dong ket qua, cancel va 6 fault phase
  rollback dung fingerprint; restore backup, UI, save/reopen va validation dat. Checkpoint co
  22 audit entry, 92 loi ke thua va `THKP-TC!E27=1201557000`.
- DT-505 end va DT-601 start khong doi workbook: 679296 byte, cung SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  runtime log correlation/task/phase/workbook token/role/package dat; COM `0x800A03EC`,
  support package 5 file va redaction du lieu nhay cam dat. Gate B5 gom 8 nhom test dat
  trong 611,1 giay; batch 4.000 o 33 ms so voi 6.210 ms, nhanh hon 185,74 lan.
- DT-601 end va DT-602 start khong doi workbook: 679296 byte, cung SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  65/65 Core dat cho valid/tamper/wrong-signature/downgrade/duplicate/app-version.
  Evidence `BQP-RPBM-2025@2.0.1` ky boi `TTBMVN-OFFLINE-2026-01` co archive SHA-256
  `AC9859DD09DEC61C362A07BA9E703CF67EE1F4D1DEAEFB68C4BC0C9155918FE2` va `Ready`.
- DT-602 end va DT-603 start khong doi workbook: 679296 byte, cung SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  68/68 Core, offline install/activate/restart, online unavailable, UI 21 dong diff,
  migration regression va semantic checkpoint dat. Install khong doi workbook pin.
- DT-603 end va DT-604 start khong doi workbook: 679296 byte, cung SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  Win11 Pro build 26200 + O365 16.0.20228.20124 x64 PASS, hai separator scenario PASS,
  AnyCPU/x64 build deu 68/68. Host Office cu/x86/Win10 duoc ghi ro `NOT_RUN`.
- DT-604 end va DT-605 start khong doi workbook: 679296 byte, cung SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  support package 9 file/8 payload co checksum, activation/project/package/validation da
  sanitize, shell Du toan va dialog UI dat; semantic checkpoint giu 92 loi ke thua va
  `THKP-TC!E27=1201557000`.
- DT-605 end va DT-606 start khong doi workbook: 679296 byte, cung SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`;
  acceptance 14/14 PASS trong 1044,4 giay, gom clean/upgrade/reinstall store, old workbook,
  batch/B4/B5, update/support va semantic checkpoint.
- DT-606 end khong doi workbook: 679296 byte, SHA-256
  `56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C`.
  Release `TTBMVN Excel Tools 1.0.0 PILOT x64` co 37 file, manifest/setup signer
  `CN=TTBMVN Pilot Publisher`, tamper bi tu choi va Gate B6 dat. ZIP SHA-256:
  `7F3FBE1B65793B130A45A2462F0D16C03C84452095F7FC0AAB92BBFAA977D0F9`.
