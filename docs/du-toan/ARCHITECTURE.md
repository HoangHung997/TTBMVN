# Kien truc va quyet dinh da chot

## Muc tieu nghiep vu

Luong chinh:

```text
ProjectProfile + RegulationPackage + PriceProfile
                        |
                        v
VL-NC-M -> DG Can / DG Nuoc -> Gia DT TC -> THKP-TC
```

Ten sheet la giao dien hien thi, khong phai danh tinh nghiep vu.

## ADR-001 - Nhan dien sheet theo vai tro

- Quyet dinh: luu role trong `Worksheet.CustomProperties`.
- Property key: `TTBMVN.WorksheetRole`; value la mot trong bay ID on dinh:
  `ResourcePrices`, `UnitRateLand`, `UnitRateWater`, `EstimateAppendix`, `CostSummary`,
  `NormLookupView`, `CostRuleView`.
- Ly do: nguoi dung co the doi ten sheet; role van di theo sheet.
- Workbook chua tag phai mo wizard chon sheet.
- Xoa sheet da gan role phai danh dau mapping thieu va yeu cau chon lai.
- Resolver chi chap nhan dung mot sheet cho moi role. Role thieu, trung hoac khong nhan biet
  la validation error; khong fallback ngam ve ten tab.

## ADR-002 - Workbook giu ProjectProfile

- Quyet dinh: workbook luu project profile co schema version trong custom document properties.
- Noi dung: ngay, package, price profile, workbook schema va override summary.
- Khong luu du lieu phap ly day du vao custom properties; chi luu ID, version, checksum va snapshot can audit.
- Store dung `TTBMVN.ProjectProfile.Manifest` va cac property
  `TTBMVN.ProjectProfile.Part.<generation>.<index>` toi da 240 ky tu moi part de khong vuot
  gioi han string cua Office.
- Manifest gom store version, generation, part count va SHA-256. Ghi part moi truoc,
  commit manifest sau cung, roi moi xoa generation cu.
- Cung payload thi save idempotent, khong tao property trung. Checksum sai hoac thieu part
  la loi du lieu, khong tra profile mot phan.
- Migration schema chi tao model hien tai trong bo nho; khong tu dong ghi lai workbook.

## ADR-003 - Goi phap ly bat bien

- Khong co mot bang `latest` bi ghi de.
- Package moi cai song song package cu.
- Workbook pin package da dung.
- Ho so moi duoc de xuat package co hieu luc moi nhat phu hop ngay.
- Ho so cu chi chuyen package khi nguoi dung chu dong va xac nhan diff.

## ADR-004 - Tach bon loai phien ban

- `AppVersion`: ma nguon va giao dien.
- `RegulationPackageVersion`: van ban, dinh muc, he so, quy tac.
- `PriceProfileVersion`: VL, NC, nhien lieu, gia may theo noi/thoi diem.
- `WorkbookSchemaVersion`: metadata va cau truc lien ket workbook.

Mot loai cap nhat khong duoc ngam thay doi loai con lai.

## ADR-005 - Module du lieu phap ly

`RegulationPackage` gom:

- `TechnicalProcessCatalog`: TT121 va phan sua doi.
- `NormCatalog`: hao phi VL/NC/M cua TT123 va phan sua doi.
- `CostRuleCatalog`: co cau va ty le chi phi cua TT123 va phan sua doi.
- `MachineRateCatalog`: phuong phap TT122 va du lieu may da duoc TT101 thay the/sua doi.
- `GeographyCatalog`: dia ban, khu vuc/vung va ngay hieu luc.
- `ComplianceCatalog`: quy chuan, tieu chuan va yeu cau quan ly chat luong dung de validation.

Moi record phai co `SourceDocument`, `Appendix`, `TableOrSection`, `SourceRow`, `EffectiveFrom` va package checksum.

## ADR-006 - Tracuu va ChiPhi

- `Tracuu` chuyen thanh man hinh tim catalog cua app.
- `ChiPhi` chuyen thanh rule engine co phien ban.
- Sheet cung ten neu con ton tai chi la view/output de nguoi dung kiem tra, khong la source of truth.

## ADR-007 - PriceProfile tach khoi phap ly

- Gia VL, NC, nhien lieu va gia thi truong thay doi theo dia diem/thoi gian.
- Dinh muc hao phi va quy tac tinh nam trong package phap ly.
- Gia ca may la ket qua ket hop du lieu/phuong phap phap ly voi PriceProfile va thong so du an.

## ADR-008 - Override cua nguoi dung

- Khong sua package chinh thuc.
- Override la lop rieng gom gia tri cu, gia tri moi, ly do, nguon tai lieu, nguoi sua va thoi diem.
- Bao cao ket qua phai the hien co override.

## ADR-009 - Excel la view va kenh tuong tac

- Logic nghiep vu test duoc nam trong `ExcelAddIn1.Core` hoac module khong phu thuoc Interop.
- VSTO chi xu ly event, mapping, doc/ghi va UI.
- Ghi Excel theo mang/block trong `ExcelWriteContext`, han che call tung cell.
- Moi context phai khoi phuc ScreenUpdating, Calculation, EnableEvents va DisplayAlerts khi dispose.

## ADR-010 - Khong nang cap ngam

- App chi thong bao co package moi.
- Migration tao checkpoint/ban sao, hien diff, yeu cau xac nhan va tao report.
- Cancel hoac exception phai giu nguyen workbook cu.

## ADR-011 - Dong bo danh sach sheet khi form dang mo

- `WorkbookSheetChangeCoordinator` la mot coordinator dung chung duoc tao/huy theo
  `ThisAddIn` startup/shutdown.
- Excel event danh dau refresh; WinForms timer 350 ms coalesce event va polling de bat rename
  vi Excel khong co event rename worksheet tin cay.
- Form subscribe theo workbook luc no duoc mo, khong doc `ActiveWorkbook` moi lan refresh.
- Sheet key uu tien `CodeName`; workbook khong co VBA/CodeName dung COM IUnknown identity
  de on dinh trong phien Excel. Role nghiep vu van dung `Worksheet.CustomProperties` de ben
  vung qua save/reopen.
- Callback chi thay danh sach item va selection. Khong duoc goi load setting/toan form,
  khong sua range text, DataGridView hoac field dang edit.

## ADR-012 - Wizard anh xa role la gate bat buoc

- Ribbon du toan validate bay worksheet role truoc khi mo luong nghiep vu.
- Workbook thieu/trung/unknown role mo `FrmWorksheetRoleMapping`; huy wizard thi khong mo
  form du toan.
- Ten sheet chuan chi duoc dung de goi y lan dau trong wizard, khong duoc dung lam fallback
  cua resolver.
- Mapping chi duoc ghi khi du bay role, sheet con ton tai va mot sheet khong nhan hai role.
- Apply mapping co rollback marker cu neu ghi mot phan bi loi; sau apply phai validate lai.
- Wizard subscribe cung `WorkbookSheetChangeCoordinator`, nen add/delete/rename khi dang mo
  cap nhat danh sach ma khong ghi ngam mapping.

## ADR-013 - RegulationPackage la manifest bat bien va tu kiem chung

- `RegulationPackage` nam trong Core thuần .NET, khong tham chieu Excel Interop/WinForms.
- Schema 1 gom PackageId, DataVersion, khoang hieu luc, status, transition note, danh sach
  van ban nguon, sau module manifest va PackageChecksum.
- Sau module on dinh: TechnicalProcess, Norm, CostRule, MachineRate, Geography va Compliance.
- Package khac Draft phai co du sau module va it nhat mot van ban nguon; Draft duoc phep
  chua day du de staging nhung van phai tu hop le va co checksum.
- Factory `Create` sao chep, trim/sap xep du lieu, chuan hoa checksum chu hoa va niêm phong
  SHA-256 tren canonical manifest. Collection va property sau tao chi doc.
- Serializer line-based UTF-8/Base64 co thu tu deterministic; deserialize tu choi schema la,
  field thua/thieu/trung, count sai va checksum khong khop.
- DataVersion dung `major.minor.patch` co the kem pre-release; package ID/module/source ID
  chi dung bo ky tu ASCII on dinh de lam key tren file system va audit.

## ADR-014 - Package store commit bang doi ten thu muc

- Bundle dau vao co layout co dinh: `manifest.ttbmanifest` va
  `modules/<RegulationModuleKind>.data`; manifest khong duoc chi dinh path tuy y.
- Import copy vao `.staging/<guid>`, tu choi reparse point, manifest qua 10 MB, schema/checksum
  sai, file module thieu/thua hoac content checksum khong khop.
- Package Draft khong vao kho chinh. Published/Superseded/Withdrawn duoc cai de ho tro
  ho so moi va ho so cu da pin.
- Kho luu song song theo `packages/<PackageId>/<DataVersion>/<PackageChecksum>`.
- Cung ID/version/checksum la idempotent `AlreadyInstalled`; cung ID/version nhung checksum
  khac la conflict, tuyet doi khong ghi de ban cu.
- Commit bang `Directory.Move` tu staging trong cung root khi dang giu `.store.lock` doc quyen.
  Loi truoc commit xoa staging; khong co central index can commit buoc hai.
- Moi ban da cai co `install.receipt`; load/list luon validate lai manifest va module files,
  nen kho bi sua ngoai app khong duoc tin tuong ngam.

## ADR-015 - Resolver tach ngay lap, phe duyet va ngay danh gia

- Request gom PreparedDate, ApprovalDate nullable va EvaluationDate; ca ba la date-only.
- Neu da phe duyet, ngay phe duyet la moc thong thuong. Neu chua phe duyet, ngay tinh/kiem tra
  la moc chon package, nen ho so cu chua phe duyet khong duoc bao luu ngam.
- Transition rule la du lieu co RuleId, effective date, previous/new package ID, source
  document ID va citation; khong hard-code mot cau `if year == 2025` trong UI.
- Rule `BQP-TT101-2025-ARTICLE-6`: ho so phe duyet truoc 28/10/2025 va duoc danh gia
  tu ngay nay tro di giu package 2021 theo TT101/2025/TT-BQP Dieu 6.
- Resolver chi de xuat Published/Superseded. Draft va Withdrawn khong duoc chon.
- Gap tra `NoApplicablePackage`; overlap tra `AmbiguousPackages`; thieu package cu can cho
  transition tra `RequiredTransitionPackageUnavailable`. Khong case nao tu chon ngam.
- Ket qua thanh cong luon co package identity kem checksum, reference date, decision code,
  rule ID, ly do va source document IDs de audit.

## ADR-016 - Workbook pin package bang ba thanh phan identity

- ProjectProfile schema 2 them RegulationPackageVersion va RegulationPackageChecksum;
  ID/version/checksum la identity day du, khong chi PackageId.
- Schema 0/1 migrate trong bo nho sang schema 2 voi version/checksum trong; day la
  `UnpinnedLegacyProfile`, khong duoc gan checksum gia hoac tu dong chon latest.
- Version va checksum phai cung co hoac cung trong; version dung semver va checksum la
  SHA-256 64 hex.
- `RegulationPackagePinService.Pin` clone profile, khong sua object dau vao, va copy identity
  tu package da validate.
- Verify load dung thu muc checksum da pin. Package version moi hon cai song song khong tac
  dong ket qua. Kho thieu tra Missing; file thieu/sai checksum/khong doc duoc tra Corrupt.
- Kho mac dinh nam tai `%LocalAppData%/TTBMVNExcelTools/RegulationPackages`; workbook khong
  luu duong dan theo may.

## ADR-017 - Migration package la giao dich co backup va rollback pin

- Preview tao diff bat bien giua source/target va PlanId deterministic; apply tu choi neu
  workbook khong con pin dung source luc preview.
- Cancel khong tao backup va khong ghi profile. Apply chi chay sau `confirmed=true`.
- Backup `SaveCopyAs` duoc tao truoc khi doi pin. Loi sau khi pin se ghi lai source profile,
  save va verify; ket qua exception cong khai rollback attempted/succeeded va backup path.
- Package target phai la Published/Superseded va duoc nap lai tu store bang checksum luc apply.
- Khong `Save()` workbook khi `ExcelWriteContext` dang dat Calculation=Manual. Calculation mode
  co the bi Excel luu vao file; save va verify chi chay sau khi context khoi phuc state nguoi dung.
- Diff phan loai metadata, source document va sau module; module tinh toan duoc dem rieng de UI
  canh bao muc do tac dong.

## ADR-018 - Du lieu phap ly la source bundle co the tai tao

- Du lieu phap ly khong hard-code trong add-in. Moi package co `source/package.properties`,
  `source/sources.tsv`, sau TSV module va bundle runtime duoc build deterministic.
- `RegulationPackageSourceReader` doc UTF-8 nghiem ngat, header/so cot/property exact va tu choi
  record thieu locator hoac chua xac minh.
- `ExcelAddIn1.RegulationTool` la CLI duy nhat de build/validate bundle. Builder ghi vao staging,
  doc lai xac minh roi `Directory.Move`; khong ghi de cung ID/version voi checksum khac.
- Module co hai checksum khac muc dich: checksum canonical ben trong
  `RegulationDataModule.ContentChecksum` va checksum byte file trong package manifest.
  Package store luon kiem tra checksum byte file; khong duoc tron hai contract.
- Raw PDF co SHA-256 va duoc giu cung source. OCR chi la chi muc tim trang, khong la nguon phat hanh.
- Workbook chuan la oracle hoi quy thu hai, khong thay the van ban. Noi workbook thieu data
  (hien tai `DG Nuoc` thieu nhom `040.x`) phai ghi ro trong audit, khong loai data phap ly.

## ADR-019 - Snapshot hop nhat va diff cap record

- Van ban hop nhat co the duoc ky sau ngay hieu luc cua noi dung duoc hop nhat. Vi vay
  `Source.IssuedDate` khong bi ep nho hon `Source.EffectiveFrom`; ca hai van phai la date-only
  hop le va khoang `EffectiveFrom/EffectiveTo` van phai dung thu tu.
- Package 2025 la snapshot resolved, khong phai patch runtime len package 2021. Tat ca record
  tro vao TT101 hoac VBHN 94-98, nen app khong phai tu ghep dieu khoan khi tinh.
- `RegulationDataBundleDiffer` so sanh theo `ModuleKind + Key`, phan loai Added/Removed/Changed
  va liet ke tung field khac: type, unit, title, data, locator va verification.
- Diff phat hanh phai co ma tran audit mot dong cho moi record thay doi, gom ly do, document,
  page range va trang thai review. DT-302 co 248/248 dong audit.
- Dia ban 2025 dung 34 key tinh/thanh pho; zone trong tung row la duy nhat, fallback phai thuoc
  danh sach zone va khi mot dia danh duoc neu ro thi uu tien zone cao hon fallback.
- CLI `RegulationTool diff` la dau ra review deterministic; package checksum van la identity
  runtime, diff khong thay the viec validate checksum.

## ADR-020 - Gia ca may la engine decimal va catalog hai doi tuong

- 33 key `MACHINE-M010.xxx` la identity logic; ma phap ly duoc resolve thanh M010 (huong
  luong ngan sach) hoac M011 (khong huong luong ngan sach), khong nhan doi may trong app.
- Record co gia tri nen chung va override `nonState*` chi khi Bang 03 khac Bang 01. M011.020
  co nhan cong bac 5/10; M011.024 co nguyen gia 350.000 thay vi 1.350.000.
- Thanh phan to lai tau la hai hao phi `si-quan` va `thuy-thu`; chuoi 6 x 20 khong duoc
  parse thanh so luong va bac tho.
- Engine Core dung `decimal` va invariant parsing. Nam thanh phan duoc lam tron 1 VND bang
  `AwayFromZero`, tong bang tong cac thanh phan da lam tron; khong dung double/locale Excel.
- Price profile cap gia nhien lieu va nhan cong theo code. He so nhien lieu phu mac dinh theo
  TT122; dieu kien an mon chi nhan 1,05 vao khau hao va sua chua.
- Khi nhien lieu da tinh trong vat lieu cong tac, request phai bat co ro rang de Cnl=0; engine
  khong tu suy tu ten may hay sheet.
- Sai so trong Bang 04 (M011.024, tong M011.009/010) duoc audit; engine giu cong thuc va
  rounding da cong bo, khong hard-code tong in sai.

## ADR-021 - Dinh muc la ma tran bien the; chi phi tach snapshot BQP va can cu BXD hien hanh

- 34 key dinh muc phap ly duoc giu on dinh. Moi record chua `variantCodes`, ma tran
  `rates`, dieu chinh co dieu kien va rang buoc; khong tach mot bien the thanh key nghiep vu moi.
- Hao phi dung `decimal` invariant va khong lam tron. Lam tron tien chi thuc hien trong engine
  chi phi tai 1 VND bang `AwayFromZero`.
- Dieu kien doc, dong chay, tin hieu vat no va bom nuoc phai duoc truyen ro. Engine tu choi
  dieu kien khong biet, nhieu khoang dong chay cung luc va van toc tren 2 m/s.
- K1 va K6 ap muc toi thieu 6.300.000 VND rieng; K4 van tinh doc lap. K6 dung 1000 kg bi
  tu choi vi VBHN97 chi quy dinh `duoi` va `tren` 1000 kg.
- K2/K5 trong VBHN97 co dan chieu BXD dong. `currentExternalBasis` ghi bang hien hanh:
  TT36/2026 Phu luc III bang 3.7 va TT38/2026 Phu luc VIII bang 2.24.
- K5 noi suy tuyen tinh trong cac moc, dung ty le moc dau khi <=10 ty va yeu cau lap du toan
  rieng khi vuot 10.000 ty. Workbook khong duoc phep ghi de quy tac phap ly neu cong thuc cu sai.

## ADR-022 - Live Excel test phai resolve theo PID/path va tach working copy

- `Marshal.GetActiveObject("Excel.Application")` khong du tin cay khi co nhieu Excel
  instance; ROT co the tro vao process test vua thoat hoac workbook khac.
- Script test dung `excel-test-process.ps1`: uu tien COM factory nhung chi chap nhan khi
  PID moi khong trung bat ky Excel dang mo; fallback `EXCEL.EXE /x` va native object theo
  PID/window handle. Process luon duoc giai phong trong `finally`.
- Test ghi/xoa/migration chi chay tren working copy trong `Dutoanmau/Variants`; workbook
  dang mo cua nguoi dung chi la oracle doc khi can doi chieu cell.
- Khi can workbook dang mo, resolver tim theo `Workbook.FullName` tren tung PID, khong theo
  active window, ten file ngan hay thu tu trong ROT.
- Checkpoint duoc mo read-only trong process `/x`; khong mo them workbook lon trong Excel
  giao dien. Quy tac nay ngan loi `RPC_E_CALL_REJECTED` va can tai nguyen COM/GDI.

## ADR-023 - Project setup la draft pure va commit workbook co rollback

- `ProjectSetupDraft` chua ma du an, ngay lap/phe duyet/danh gia, moc/ho so gia va mapping;
  preview chi goi Core, khong ghi workbook.
- Package phap ly duoc resolver chon theo ngay va transition catalog; UI khong tu chon
  package `latest`. Profile duoc pin day du ID/version/checksum.
- Output add-in mang theo bundle BQP-RPBM-2021/2025. Bootstrap import idempotent vao local
  package store va chi dua package co prefix production `BQP-RPBM-` vao resolver.
- Mapping uu tien `Worksheet.CodeName`, fallback ten sheet va goi y ten canonical. Doi ten
  sheet khong lam mat mapping; sheet bi xoa phai bao validation thay vi chon ngam sheet khac.
- Commit luu profile va mapping trong mot `ExcelWriteContext`. Truoc commit phai chup payload
  profile va role assignments; bat ky exception nao cung khoi phuc ca hai, ke ca workbook
  ban dau chua co profile/tag.
- Ribbon chi mo form thong tin du toan sau khi project setup tra `OK`; Huy dong luong ma
  khong tao dirty state trong workbook.

## ADR-024 - Tra cuu dinh muc theo package identity, khong dung sheet lam data source

- `NormSearchIndex` nhan package ID/version/checksum va module Norm; moi result luon giu
  identity + source locator, nen cung ma trong hai package khong bi tron.
- Exact key co uu tien tuyet doi. Chuoi dung pattern `NORM-xxx.xxxx` nhung khong ton tai
  tra zero; fuzzy khong duoc chuyen sang ma gan giong.
- Fuzzy chuan hoa Unicode FormD, bo dau tieng Viet, token prefix va Levenshtein co gioi han;
  ket qua deterministic theo score/key/package.
- Moi truong duoc phan loai theo nhom ma 000/010-020/030/040; do sau trich tu title phap ly;
  filter variant/resource chi ap dung khi package co du lieu chi tiet.
- Package 2021 hien chi co metadata NormCatalog, chua co bang hao phi chi tiet. UI hien canh
  bao ro va khong suy dien/copy hao phi tu package 2025.
- `RegulationPackageStore.LoadBundleRequired` validate receipt, byte checksum va module
  payload truoc khi runtime tao search index.
- Sheet `Tracuu` van duoc tag de tuong thich workbook, nhung UI va tinh toan khong doc sheet
  nay lam nguon du lieu.

## ADR-025 - ChiPhi dung engine co phien ban va override tach biet

- `CostRuleEngine` nhan mot request ro rang cho VL, NC, M, dia hinh, dien tich, loai du an,
  loai cong trinh, khoi luong xu ly, VAT va tap K1-K6; khong doc cell hay bang he so Excel.
- Moi ket qua giu ca calculated amount va applied amount. Override khong sua gia tri tinh goc,
  phai co ma K hop le, amount khong am va ly do 1-1024 ky tu; tong hop dung applied amount.
- `WorkbookCostRuleService` chi nap module CostRule bang ID/version/checksum da pin.
  UI hien ty le, hai gia tri, source locator va can cu ngoai hien hanh neu package co khai bao.
- Adapter TT123/2021 chi duoc nhan dien bang source document `TT123-2021-BQP`: K2 dung step,
  K5 dung linear, currentExternalBasis rong, khong ap minimum K1/K6 va K5 dung tai 2.000 ty.
  Package moi van bat buoc metadata day du; khong lay quy tac VBHN 2025 gan nguoc cho 2021.
- Bang `DataDutoan` hard-code cu khong con trong runtime. Sheet `ChiPhi` van duoc tag de tuong
  thich workbook, nhung khong la data source cua engine.

## ADR-026 - PriceProfile doc lap voi phap ly va snapshot trong workbook

- `PriceProfile` la bo gia theo dia diem, ngay gia va doi tuong luong; no co ID/version/checksum
  rieng, khong nam trong `RegulationPackage`. Ma nguon luc on dinh va alias tach khoi ten hien thi
  legacy de workbook doi ten/chuoi ma van resolve co kiem soat.
- Moi entry giu gia goc, loai nguon, locator, don vi va ngay gia. Override la lop rieng co gia
  cu/moi, ly do, nguon, nguoi sua va UTC; checksum serializer deterministic khong phu thuoc locale.
- Local `PriceProfileStore` cai qua staging/atomic move, idempotent theo ID/version/checksum va
  tu choi cung version khac noi dung. Workbook giu full snapshot chunked + SHA va pin
  `PriceProfileId` trong ProjectProfile, nen mo tren may khac khong phu thuoc local store.
- Coverage doi chieu requirement cua NormCatalog theo code/kind/unit. Gia thieu phai duoc bao
  ro va chan luong tinh lien quan; tuyet doi khong tu gan 0.
- Adapter `VL-NC-M` chi la cau noi tuong thich. DT-404 batch-write 21 gia vat lieu F81:F101;
  cong thuc nhan cong va ca may duoc bao toan de DT-405 tinh truc tiep tu PriceProfile.
- Service resolve sheet theo worksheet role, khong theo ten hien thi. COM object dung chung chi
  `ReleaseComObject` mot tham chieu do ham so huu; cam `FinalReleaseComObject` vi co the lam
  vo hieu RCW ma caller dang giu va gay loi runtime sau khi da ghi mot phan workbook.

## ADR-027 - Don gia la ket qua cua NormCatalog va PriceProfile da pin

- `UnitRateCalculator` chi nhan request Core: dinh muc, bien the, dieu kien, PriceProfile va
  binding nguon luc. Engine khong doc `DG Can`, `DG Nuoc` hoac cong thuc legacy lam data source.
- Don gia luon tinh cho mot don vi cong tac. Hao phi va thanh tien dung `decimal`, khong lam
  tron trong engine; policy cong bo la `none-invariant-decimal`, locale chi dung khi hien thi.
- Vat lieu phan tram duoc tinh sau tong vat lieu truc tiep. Gia thieu, sai loai hoac sai don vi
  la validation blocking; khong zero-fill va khong bo qua dong.
- Ma nguon luc logic (`M010.DIVING` va cac ma `OR`) bat buoc binding sang PriceProfile cu the
  kem ly do/phuong an 1-1024 ky tu. Runtime khong tu chon mot may thay cho bien phap duoc duyet.
- Dieu kien/he so la danh sach code duoc nguoi dung chon ro. He so legacy trong worksheet chi
  dung lam oracle doi chieu; khong duoc doc ngam vao request.
- Ket qua luu identity package/profile, source locator, ma hao phi, ma gia va trang thai
  override de task sau co the truy vet va ghi Excel theo batch.

## ADR-028 - Phu luc du toan giu baseline va he so tach dong

- `EstimateAppendixCalculator` nhan cac dong khoi luong + don gia da tinh, dung `decimal`,
  khong doc Excel va khong lam tron trung gian. Ket qua giu package/profile identity va
  quantity/output source reference.
- Adapter legacy chi dung formula F:H de nhan dien ma dong/variant ban dau. Ket qua tinh van
  den tu NormCatalog + PriceProfile; dong co khoi luong khong duoc bo qua khi mapping/gia sai.
- Preview luu hai lop: baseline khong he so va current co dieu kien. Writer ghi don gia baseline
  vao F:H, cong thuc dong I:N theo D/E, va phan delta vao row he so; vi vay input Excel thay doi
  van recalculate ma khong nhan he so hai lan.
- Writer resolve `EstimateAppendix` theo CodeName/role, ghi mot batch F:N, backup Formula,
  rollback neu COM/verify loi va tu tinh lai tong I:N truoc khi tra thanh cong.
- Legacy row 22 la hybrid sai: mo ta/khoi luong 3-12 m nhung may dung bien the 0,5-3 m.
  App uu tien VBHN97/2025 trang 32-33, ghi warning va chot delta; khong lam sai package de
  ep khop workbook cu.
- Patch package cung PackageId/effective date duoc collapse theo semantic DataVersion cao nhat.
  Cac PackageId khac cung hieu luc van bi resolver bao ambiguous, khong chon ngam.

## ADR-029 - Tong hop kinh phi dung template va co so tinh phap ly

- `CostSummaryCalculator` la Core pure cho Bieu 01-04. Bieu 04 co TL va VAT; Bieu 01-03
  khong tu them hai khoan nay. Moi khoan tien duoc lam tron VND AwayFromZero, tong cuoi lam
  tron 1.000 VND AwayFromZero.
- Co so tinh bat buoc: K2 tren T; K1, K3, K4, K5, K6 tren Z. VAT Bieu 04 tinh tren
  `Q-(K3+K4)`. Override khong sua calculated amount va bat buoc co ly do.
- Adapter workbook chi doc tong VL/NC/M tu `EstimateAppendix` va cau hinh legacy tu
  `CostSummary`; moi sheet duoc resolve bang role/CodeName, khong bang ten hien thi.
- Writer ghi mot batch `D10:E27`, cap nhat A28 bang tien chu, backup Formula/Value2,
  rollback neu write/Calculate/verify loi va xac minh E27 truoc khi tra thanh cong.
- Cong thuc legacy K5 tren T va ty le cu khong duoc coi la oracle. Khi package co can cu
  K5 tren Z va bang ngoai hien hanh, ket qua phap ly moi duoc chot trong BASELINE.

## ADR-030 - Ghi Excel la batch transaction co phuc hoi trang thai

- Moi writer ghi `Formula`/`Value2` bang mang hai chieu trong it COM round-trip nhat;
  khong lap `Cells[row, column] = value` tren duong nong neu co the gom thanh range.
- `ExcelBatchWriteTransaction` snapshot Formula theo worksheet/address, rollback nguoc thu tu
  neu chua `Commit` va giai phong COM snapshot khi `Dispose`. Caller van so huu range dau vao.
- `ExcelWriteContext` quan ly dong bo ScreenUpdating, Calculation, EnableEvents va
  DisplayAlerts. Constructor phai phuc hoi neu setup loi mot phan; `Dispose` idempotent.
- Performance gate dung cung workbook/PID va cung so o, so sanh mot batch voi cell-by-cell;
  ngoai toc do phai dat value parity, exception rollback va state restoration.
- Header/link sparse co the gom theo mot hang va bao toan o khong thuoc mapping bang snapshot
  ma tran. Dinh dang chi ap dung cho cell mapping de khong doi behavior ngoai y muon.

## ADR-031 - Truy vet la payload workbook theo CodeName va role

- Audit la model Core co schema/checksum deterministic, khong dung comment cell va khong
  gan identity vao ten sheet hien thi. Vung ket qua luu CodeName, worksheet role va toa do.
- Entry phu luc giu package, PriceProfile, norm/variant, locator phap ly, nguon gia,
  khoi luong/output; entry K1-K6 giu locator va can cu ngoai. Lookup chon entry co vung nho
  nhat de dong chi tiet uu tien hon block tong.
- Writer chi cap nhat audit sau khi Excel verify ket qua va truoc transaction commit.
  Custom payload store dung generation/manifest/checksum; cleanup chunk cu la best effort.
- Metadata noi bo du suc tai dung plan phu luc sau khi writer thay formula F:H bang value;
  metadata khong hien trong bang nguon nguoi dung.
- Doc custom properties phai quet collection mot lan thanh dictionary; cam tim tung chunk
  bang mot vong COM rieng vi tao O(n^2) va treo tren workbook audit lon.

## ADR-032 - Validation la read-only va dung chung formula contract voi writer

- Validation report la Core model, moi issue co code/severity/role/CodeName/address,
  message/remediation va co inherited. Chi Error khong inherited moi chan gate.
- Scanner khong ghi workbook; doc formula/value theo range batch va tai tinh qua cung engine,
  package, PriceProfile, plan va dieu kien da audit.
- Writer phu luc/THKP cung cap expected formula map dung chung. Validator khong duoc chep lai
  mot bo cong thuc khac vi se tao false positive khi writer thay doi.
- Audit identity stale khong ngan scanner doc metadata o che do validation; nhờ vay report
  van chi ra chinh xac ma/dong thieu gia. Luong chay tinh toan binh thuong van chan identity sai.
- Defined name `#REF!` ke thua duoc gom mot warning; name app `TTBMVN_*` hong la Error rieng.
  UI khong liet ke hang nghin name cu thanh hang nghin dong.

## ADR-033 - Migration package la giao dich toan khoi co preview tinh toan

- Preview khong ghi Excel. No giu package diff, PriceProfile identity, audit scope va so sanh
  6 tong phu luc + tong kinh phi giua gia tri workbook hien tai va engine package dich.
- Package 2021 khong co hao phi chi tiet cho engine moi; migration van doc gia tri nguon tu
  workbook legacy va chi cho apply khi package dich tinh du 13 dong. Khong suy dien bang hao
  phi 2021 va khong bo qua dong loi.
- Apply snapshot ProjectProfile, PriceProfile, ResultAudit va cac range `EstimateAppendix` /
  `CostSummary`; tao `SaveCopyAs` truoc mutation, sau do pin package, ghi phu luc, ghi THKP,
  validation va save theo thu tu ro rang.
- Moi phase co fault hook kiem thu. Loi tai bat ky phase sau backup phai khoi phuc payload va
  Formula bang snapshot, calculate/save lai va doi chieu fingerprint truoc khi bao rollback dat.
- Apply tu choi preview stale neu workbook, bang gia, audit, package store hoac gia tri nguon
  thay doi. UI khong co nut bo qua validation va luon yeu cau duong dan backup rieng.

## ADR-034 - Runtime log co context co cau truc va support package toi thieu du lieu

- Moi record log co UTC timestamp, correlation ID, task, phase, workbook token, worksheet
  role, package identity, action va exception type/HResult/stack. Record khong co operation
  van co correlation ID va gia tri `runtime/unspecified` thay vi bo trong truong.
- Workbook va Machine ID chi hien bang SHA-256 token rut gon; context key lien quan path,
  name, key, value, data, license, machine, customer va project bi redact. Message/stack
  duoc loc path, sheet reference, license-like token va gioi han kich thuoc.
- Logger khong duoc nem loi ra luong Excel; file xoay o 5 MiB, giu toi da 3 archive.
- Support package khong chep nguyen log, Machine ID, workbook/sheet name hay workbook-bound
  setting. Goi gom app-info, package/role aggregate, setting ho dao khong co workbook field,
  sanitized log va SHA-256 manifest.
- Cac boundary PriceProfile/phu luc/THKP/audit/validation/migration/project setup ghi task,
  phase, workbook token, role va package. HResult `0x800A03EC` la fault oracle cua DT-505.

## ADR-035 - Goi update co trust root bat bien va verifier fail-closed

- `.ttbupdate` la ZIP gioi han kich thuoc, chi co manifest canonical, detached signature va
  7 payload package bat buoc. Moi entry thua/thieu/trung hoac path khong an toan deu bi tu choi.
- Chu ky RSA-SHA256/PKCS#1 bao ve manifest; manifest bao ve length/SHA-256 tung payload.
  Regulation bundle duoc doc lai va doi chieu identity sau khi hash dat.
- Trust catalog trong Core chi chua public key. Private key non-exportable chi nam trong
  Windows CAPI user key container va chi publisher tool duoc mo de ky.
- Version policy tach khoi cryptographic validity: app qua cu, downgrade va cung version
  khac checksum deu khong duoc cai; cung version/cung checksum la `AlreadyInstalled`.
- Extract chi duoc goi bang verification result, doi chieu lai SHA-256 archive va payload
  de chan thay file giua luc preview va install.

Chi tiet dinh dang, trust boundary va lenh phat hanh: `docs/du-toan/OFFLINE-UPDATE.md`.

## ADR-036 - Update install va workbook migration la hai giao dich doc lap

- Cai `.ttbupdate` chi ghi package store va activation state; khong doi ProjectProfile,
  cong thuc hay ket qua workbook dang mo.
- Activation state chon mot patch mac dinh tren moi PackageId cho du an moi. Tat ca version
  van duoc giu de workbook pin checksum cu mo lai va de migration chon ro dich.
- Rollback update la doi activation pointer ve ban da cai, khong xoa package. State co
  serializer canonical/checksum va atomic replace; restart phai tai lai dung preference.
- `LoadPreferredPackages` chi dung cho project setup. Catalog/migration va pinned services
  van thay hoac tai toan bo package can thiet.
- Online update la `IRegulationUpdateProvider`; ban hien tai dung provider Disabled va
  chuyen exception mang thanh status Unavailable, khong anh huong file update offline.

## ADR-037 - Support package la snapshot chan doan co schema on dinh

- Shell Du toan truyen workbook ro rang vao `FrmSupport`; khong resolve lai ActiveWorkbook
  khi tao goi vi nguoi dung co the doi workbook/sheet trong luc dialog mo.
- Goi chan doan gom dung 8 payload nghiep vu va mot `manifest.sha256`. Moi payload co ten
  on dinh; loi doc mot thanh phan van tao file cung ten voi exception type/HResult toi thieu.
- Activation chi xuat token Machine ID, trang thai va han; project chi xuat token identity,
  ngay va package pin. Khong xuat license key, raw ID, path, sheet/address hay override value.
- Regulation manifest la ban sealed tu package store; validation duoc tinh read-only tai thoi
  diem tao goi va sanitize lai bang workbook/project/profile/sheet/machine token truoc khi ghi.
- Runtime log va moi text chan doan deu qua mot sanitizer dung chung; manifest hash la lop
  phat hien file support bi sua, khong thay the chu ky cua `.ttbupdate`.

## ADR-038 - Acceptance tach oracle lich su va co resume fail-closed

- Regression DT-405/406/407 chay tren checkpoint cua chinh hop dong nghiep vu do; current
  checkpoint dung cho validation, migration, update, support va semantic gate. Khong ep mot
  oracle E27 qua cac moc co thay doi phap ly da duoc duyet.
- Old workbook test mo DT-101 read-only, xac minh first-run suggestion va hash khong doi.
- Clean/upgrade/uninstall/reinstall package chay trong store cach ly. DT-605 khong duoc go
  VSTO hien dang duoc Excel nguoi dung nap; vong doi bo cai VSTO thuoc release drill DT-606.
- Acceptance fail-fast va ghi evidence ke ca FAIL. Resume chi tai su dung prefix PASS khi
  test name/order va SHA-256 checkpoint khop; bat ky thay doi checkpoint buoc chay lai tu dau.

## ADR-039 - Manifest signing dung certificate store va release channel ro rang

- Project khong tham chieu/luu PFX. Build/publish nhan thumbprint tu Windows certificate
  store; `TTBMVN_SIGNING_CERT_THUMBPRINT` la duong vao cho certificate CA ben ngoai.
- Khi khong co cert ngoai, tool tao `TTBMVN Pilot Publisher` RSA 3072, EKU Code Signing,
  private key non-exportable. Artifact mang nhan PILOT va public cert; khong duoc coi la CA trust.
- VSTO deployment/application manifest duoc ky trong MSBuild. `setup.exe` duoc Authenticode
  sign sau publish; release verifier buoc signer/thumbprint khop va channel ngoai PILOT phai Valid.
- Release la allowlist file co SHA-256. Verifier chan file thua/thieu/hash sai va PFX; gate
  co negative tamper test. ZIP co sidecar SHA-256 va khong tu ghi hash cua chinh no vao trong.
- Installer dong Excel, go manifest cu neu ton tai o vi tri khac, import public cert chi khi
  nguoi dung xac nhan `-TrustPilotCertificate`, cai VSTO va giu LocalAppData qua nang cap.
- Khong mutate VSTO tren may build dang co Excel nguoi dung. `-VerifyOnly` la preflight;
  actual pilot install/uninstall phai chay tren may test sach va duoc ghi rieng.

## No luc ky thuat hien tai

- `Winform/Dutoan.cs` la shell cho thong tin chung, tra cuu dinh muc va ChiPhi.
- `FormThongtinchung.cs` phan lon la ma du thao/comment va con tham chieu ten sheet truc tiep.
- `ComboBoxManager.cs` hard-code danh sach loai du toan/dia hinh.
- `CostRuleEngine` va `WorkbookCostRuleService` da thay bang `DataDutoan.cs` hard-code.
- `ReadWriteDutoan.cs` da co nen doc/ghi `CustomDocumentProperties`.
- `ExcelWriteContext.cs`, `RuntimeLogger.cs` va `SupportPackage.cs` co the tai su dung sau khi rao soat.

## ADR-040 - Phu luc DT la workspace va don gia duoc gom theo ngu canh

- `EstimateWorkspace` luu tung dong Phu luc DT, cong thuc/khoi luong, mapping cot, moi truong,
  doi tuong luong, dinh muc chi tiet, dieu kien va binding thiet bi logic.
- Dong khong co dinh muc/ma chi tiet la dong van ban: app bao toan noi dung va khong ghi ket qua.
- Sheet va dong nguon khong phu thuoc ten hien thi. Sheet co custom property source ID; dong co
  workbook-level hidden name. Resolver fallback CodeName/ten chi de tuong thich file cu.
- `EstimateRateIdentity` khong chua so hieu cong tac hay khoi luong. Hai dong chi dung chung
  don gia khi package, norm/variant, can-nuoc, HLNS/KHLNS, PriceProfile, conditions va bindings khop.
- `PriceProfilePortfolio` chua toi da mot profile HLNS va mot profile KHLNS trong workbook.
  Snapshot PriceProfile cu van duoc doc/ghi de cac luong DT-404..407 tiep tuc hoat dong.
- Writer chi tao sheet app-owned cho nhom dang su dung, danh dau bang custom property, ghi batch,
  tao workbook names an va ghi formula ve cac cot Phu luc DT da mapping. Chay lai la idempotent.
- Ten sheet dau ra la presentation; quyen so huu va resolve dua vao metadata, khong dua vao ten.
- Project C# dat code page 65001 de cac file UTF-8 khong BOM hien thi dung tren VSTO/.NET Framework.
