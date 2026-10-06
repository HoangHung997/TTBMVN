# HANDOVER — Dự toán RPBM V2

> File này là điểm vào cho người tiếp quản. Đọc file này trước, sau đó đọc `PROGRESS.md`, `UI-CONTRACT.md`, `ARCHITECTURE.md` và `MIGRATION-PLAN.md`.

## 1. Trạng thái bàn giao

Nhánh làm việc chính: `main`.

Mở rộng 2026-10-06: V2-802 bổ sung mục 7 và bảng quản lý/biên soạn gói.
Đọc [PACKAGE-MANAGEMENT.md](PACKAGE-MANAGEMENT.md) để biết luồng tạo phiên bản,
kiểm chứng, chặn xóa gói đang dùng, bản nháp và giới hạn. Build sạch; 83 test Core
PASS; UI tạo từ mẫu/kiểm tra/lưu nháp PASS. Full runtime regression vẫn pending.

Các task đã triển khai ở mức code:

| Task | Nội dung | Trạng thái |
|---|---|---|
| V2-001 | Bỏ gate khởi động Dự toán | DONE - implementation / runtime pending |
| V2-002 | CustomTaskPane + UI nền tảng | DONE - implementation / runtime pending |
| V2-101 | WorkItemId + binding định mức bền vững | DONE - implementation / runtime pending |
| V2-201 | VL-NC-M | DONE - implementation / runtime pending |
| V2-301 | DG Cạn / Nước / Biển | DONE - implementation / runtime pending |
| V2-401 | Gia DT TC + THKP-TC | DONE - implementation / runtime pending |
| V2-501 | Validation / safe repair | DONE - implementation / runtime pending |
| V2-601 | Legacy / migration / VT-DN / rename | DONE - implementation / runtime pending |
| V2-701 | Thiết lập chung | DONE - implementation / runtime pending |
| V2-801 | Gói pháp lý & Dữ liệu | DONE - implementation / targeted Excel service tests PASS; full UI regression pending |
| V2-901 | Báo cáo & Xuất in | DONE - implementation / PDF service tests PASS; full UI regression pending |

**Quan trọng:** đã có build, kiểm thử service Excel và smoke test task pane trên Windows,
nhưng chưa có full regression end-to-end hoặc đối chiếu tài chính toàn bộ. Không được
đổi toàn bộ trạng thái thành runtime PASS từ các kiểm thử mục tiêu này.

**Cập nhật tiếp quản Windows 2026-10-06:** đã sửa ba lỗi biên dịch và build Release x64
thành công trên MSBuild 18.10.1; 81 test Core chạy PASS. Full Excel/VSTO regression và đối chiếu
ảnh UI vẫn pending. Xem [WINDOWS-VERIFICATION.md](WINDOWS-VERIFICATION.md) để biết
lệnh chạy, lỗi đã sửa và giới hạn kiểm chứng. V2-801 và V2-901 hiện đã có code và
targeted Excel service tests; xem V2-801.md/V2-901.md. NEXT là regression runtime,
không phải tiếp tục dựa trên trạng thái cũ trong các mục lịch sử bên dưới.

File mẫu HoaLuNamDinh có sẵn lỗi #REF! (285 cached errors, 91 công thức tham chiếu hỏng).
Không dùng tổng có sẵn của file để khẳng định V2 đúng toàn bộ. Test riêng từng fixture,
giữ nguyên nguồn và xử lý tham chiếu mất bằng dữ liệu có căn cứ, không suy đoán số tiền.

## 2. Mục tiêu kiến trúc đã chốt

### Excel là hồ sơ chính

- workbook chứa khối lượng dự án, giá thị trường, tham số giá, công thức, đơn giá và các sheet in;
- database/package chỉ là thư viện định mức, hao phí, danh mục VL/NC/M, thông số và căn cứ pháp lý;
- add-in là engine hỗ trợ gắn định mức, sinh công thức/link, đồng bộ và validation.

### Không có số chết cho kết quả tính

Cho phép nhập trực tiếp:

- khối lượng;
- giá vật liệu thị trường;
- giá nhiên liệu/tham số đầu vào;
- hao phí gốc từ package.

Không cho phép engine C# tính xong rồi ghi snapshot số chết cho các kết quả:

```text
Giá đầu vào
 -> giá NC / giá ca máy
 -> đơn giá công tác
 -> thành tiền
 -> THKP
```

Chuỗi trên phải là công thức/link Excel và workbook Name.

### Số khu vực không phải biến kiến trúc

Dự án có 1, 2 hay 3 khu vực đều dùng chung:

- WorkItemId;
- binding định mức;
- RateId;
- resource aggregation;
- writer;
- validation.

Dòng nhóm khu vực không được biến thành WorkItem.

## 3. UI contract tuyệt đối không được phá

Nguồn chuẩn:

- `docs/du-toan-v2/UI-CONTRACT.md`
- `/mnt/data/chuan_UI`

Tại thời điểm bàn giao, thư mục `/mnt/data/chuan_UI` hiện có ảnh chuẩn:

1. `01-Tong-quan.png`
2. `02-Cong-tac.png`
3. `03-Gan-dinh-muc.png`
4. `04-VL-NC-M.png`
5. `05-DG-Can.png`
6. `06-DG-Nuoc.png`
7. `07-THKP-TC-Kiem-tra.png`
8. `08-Thiet-lap-chung.png`

Chưa thấy file ảnh 09–10 trong folder chuẩn hiện tại.

Nguyên tắc UI:

- UI chính là **CustomTaskPane bên phải Excel**, khoảng 430 px;
- Excel luôn là vùng làm việc chính;
- không quay lại form modal lớn/menu dọc của legacy;
- không được tạo startup wizard bắt buộc;
- dùng renderer icon nội bộ `EstimateUiIcons`;
- màn hình mới phải cùng ngôn ngữ xanh lá / vàng / xanh dương / đỏ, card trắng, Segoe UI;
- nếu UI 09/10 chưa có ảnh chuẩn khi bắt đầu task, phải dựng/chốt mockup cùng phong cách trước hoặc bám contract đã ghi; không tự chuyển sang kiểu UI khác.

## 4. Luồng V2 hiện tại

```text
Ribbon Dự toán
    |
    +--> mở Trợ lý Dự toán ngay
            |
            +--> Tổng quan
            +--> Công tác
            |      +--> đăng ký vùng công tác
            |      +--> reconcile WorkItemId
            |      +--> migration legacy on-demand
            |
            +--> Gắn định mức
            +--> VL-NC-M
            +--> DG Cạn / DG Nước / DG Biển
            +--> THKP-TC & Kiểm tra
            |
Ribbon Thiết lập Chung
    |
    +--> mở cùng task pane
            +--> Thiết lập chung
```

Không yêu cầu:

- có THKP trước;
- map đủ 7 role;
- có package trước khi mở module.

Dependency chỉ được chặn **đúng chức năng cần dependency đó**.

## 5. Identity và persistence

### WorkItem

Nguồn thật:

- Custom XML state trong workbook.

Mỗi công tác có các thông tin chính:

- WorkItemId;
- SourceKey;
- NormCode;
- VariantCode;
- PackageId;
- DataVersion;
- PackageChecksum;
- Kind;
- Fingerprint;
- IsOrphaned.

Không dùng RowIndex làm identity.

Anchor phụ trên sheet:

- `__TTB_ID`
- `__TTB_NORM`
- `__TTB_KIND`
- `__TTB_HASH`

Các cột này phải nằm ngoài vùng in/visible output và bị ẩn.

### Binding định mức

Ô `Định mức` trên sheet chỉ là display/cache.

Binding thật là:

```text
Custom XML state
+ PackageId
+ DataVersion
+ PackageChecksum
+ NormCode
+ VariantCode
```

Không được tự lấy text legacy ở ô Định mức rồi coi là binding pháp lý.

### Output sheet identity

Ưu tiên:

1. Worksheet Role;
2. worksheet CodeName;
3. custom environment metadata;
4. canonical/legacy alias fallback.

Mapping chính:

- `ResourcePrices` -> VL-NC-M;
- `UnitRateLand` -> DG Cạn;
- `UnitRateWater` -> DG Nước;
- `EstimateAppendix` -> Gia DT TC;
- `CostSummary` -> THKP-TC;
- DG Biển dùng `TTBMVN.EstimateV2.UnitRateEnvironment=Sea`.

Rename tab không được tạo sheet duplicate mới.

## 6. Các output chính

Mặc định:

- `THKP-TC`
- `Gia DT TC`
- `DG Can`
- `DG Nuoc`
- `VL-NC-M`

`DG Bien` chỉ tạo khi thực sự có rate biển đang dùng.

Không tạo:

- DG theo từng khu vực;
- lookup sheet visible;
- cost-rule sheet visible;
- technical sheet rác chỉ để code hoạt động.

## 7. Tóm tắt các task đã làm

### V2-101 — WorkItem/binding

File chính:

- `ExcelAddIn1.Core/EstimateV2State.cs`
- `WorkbookEstimateV2StateService.cs`
- `WorkbookEstimateV2RegistrationService.cs`

Đã có:

- save/reopen;
- insert/delete/sort/copy recovery;
- fingerprint recovery;
- duplicate WorkItemId split;
- orphan detection;
- restore norm display từ binding nếu setting cho phép.

### V2-201 — VL-NC-M

File chính:

- `EstimateV2Resources.cs`
- `EstimateV2ResourcePriceSheetProjector.cs`
- `WorkbookEstimateV2ResourceService.cs`
- `WorkbookEstimateV2ResourceSheetWriter.cs`
- `EstimateResourcesPaneView.cs`

Nguyên tắc:

- chỉ sinh resource thực sự được dùng;
- vật liệu là input thị trường;
- nhân công và máy là formula output;
- metadata ngoài A:F;
- workbook Name bền vững;
- giữ input người dùng khi refresh.

### V2-301 — DG

File chính:

- `EstimateV2Rates.cs`
- `WorkbookEstimateV2RateService.cs`
- `WorkbookEstimateV2RateSheetWriter.cs`
- `EstimateUnitRatesPaneView.cs`

Identity Rate:

```text
PackageIdentity + NormCode + VariantCode -> RateId
```

Một rate dùng chung cho mọi WorkItem dùng cùng định mức/variant.

DG link trực tiếp VL-NC-M bằng workbook Name/formula.

### V2-401 — Gia DT TC / THKP

File chính:

- `WorkbookEstimateV2CostLinkService.cs`
- `WorkbookEstimateV2CostLinkWriter.cs`
- `EstimateCostSummaryPaneView.cs`

Bảng công tác chuẩn:

```text
A TT
B Mã công tác
C Định mức
D Mô tả
E Đơn vị
F Khối lượng
G:I Đơn giá VL/NC/M
J:L Thành tiền VL/NC/M
M... metadata/legacy hidden
```

G:L phải là formula/link, không snapshot.

THKP chỉ quản lý các liên kết V2 mà writer sở hữu; không tự suy đoán tỷ lệ pháp lý khác.

### V2-501 — Validation

File chính:

- `EstimateV2Validation.cs`
- `WorkbookEstimateV2ValidationService.cs`

Kiểm tra:

- thiếu định mức;
- rate/package không resolve;
- thiếu giá;
- workbook Name thiếu;
- formula bị ghi đè;
- `#REF!`, `#VALUE!`, `#N/A`, ...;
- orphan/duplicate/mismatch;
- layout metadata chồng G:L.

Safe repair chỉ gọi writer sở hữu. Không tự điền market price.

### V2-601 — compatibility/migration

File chính:

- `EstimateV2Compatibility.cs`
- `WorkbookEstimateV2CompatibilityService.cs`

Hỗ trợ:

- legacy aliases;
- VT/DN;
- 1/2/3 khu vực;
- rename sheet;
- package pinned missing/corrupt;
- legacy Gia DT TC migration khi layout nhận diện chắc chắn.

Nếu VT/DN hoặc nhiều copy không có lựa chọn duy nhất: **không tự đoán**.

### V2-701 — Thiết lập chung

File chính:

- `EstimateV2Settings.cs`
- `WorkbookEstimateV2SettingsService.cs`
- `EstimateSettingsPaneView.cs`
- `EstimateTaskPaneControl.cs`
- `EstimateTaskPaneManager.cs`

Settings lưu trong workbook Custom Document Properties.

6 output mapping:

- THKP;
- Gia DT TC;
- DG Cạn;
- DG Nước;
- VL-NC-M;
- DG Biển.

Behavior:

- AutoRestoreNormDisplay;
- ValidateOnOpen;
- AutoSyncRows;
- WarnOnMappingLoss;
- AutoSaveEnabled / minutes.

Invariant không cho tắt:

- formula/link;
- Custom XML;
- hidden technical columns.

Runtime:

- SheetChange chỉ debounce registered work sheet;
- row reconcile timer khoảng 650 ms;
- auto-save mặc định 5 phút;
- không Save As workbook mới;
- mapping loss là warning mềm;
- SaveConfiguration có rollback.

## 8. Các nguyên tắc tuyệt đối không được phá

1. Không đưa `FrmProjectSetup` hoặc role/package validation trở lại startup gate.
2. Không yêu cầu THKP tồn tại mới được mở Dự toán.
3. Không tự nâng pinned package sang `latest`.
4. Không dùng RowIndex làm WorkItem identity.
5. Không biến text định mức legacy thành legal binding.
6. Không ghi kết quả tính thành số chết.
7. Không tự điền giá vật liệu thị trường bị thiếu.
8. Không tự suy đoán điều kiện/adjustment của định mức.
9. Không tự chọn VT hoặc DN khi workbook có cả hai.
10. Không xóa output legacy âm thầm; chỉ được ẩn sau khi output V2 chuẩn đã được rebuild an toàn.
11. Không tạo sheet kỹ thuật visible chỉ để phục vụ code.
12. Không hard-code tỷ lệ pháp lý minh họa từ mockup.
13. Không thay style UI đã chốt bằng form lớn/modal.
14. Không ghi runtime PASS nếu chưa chạy thật.

## 9. V2-701 — commit gần nhất cần biết

Chuỗi commit chính:

- `e8299a079c56` — settings policy;
- `0761c829530c` — persist settings/output mapping;
- `3dd6057a1b3a` — settings pane UI;
- `961318845a9c` — áp settings vào runtime;
- `0e9f79a379ef` — rollback settings + mappings;
- `b7274db94931` — debounce chỉ registered source;
- `c21b429ba119` / `dded04ebf9d6` — Ribbon entry;
- `14d867142cd1` — settings mở từ Ribbon, không sửa layout Tổng quan;
- `b20160b35b7f` — UI contract V2-701;
- `43e8a6d6e294` — expected output map + detect mapping loss;
- `840c02e9702e` / `5ea6e007e7f2` — surface mapping loss warning.

Commit cập nhật tiến độ bàn giao:

- `5bcd2d4c7ba5` — PROGRESS cập nhật qua V2-701 và roadmap 801/901.

## 10. Việc chưa được xác nhận runtime

Phải chạy trên máy Windows có Excel/VSTO.

### Build/test cơ bản

- build solution đúng target hiện tại;
- chạy `ExcelAddIn1.Tests`;
- không chấp nhận compile warning/error mới từ V2;
- mở Excel bằng add-in thật.

### Regression bắt buộc

#### Startup

- workbook trống vẫn mở Dự toán;
- missing package không crash;
- corrupt profile/state không khóa pane;
- không có modal gate.

#### Identity

- insert/delete/sort/copy;
- save/close/open;
- duplicate ID;
- orphan;
- xóa ô Định mức display.

#### Formula chain

- đổi giá vật liệu -> DG -> Gia DT TC -> THKP tự cập nhật;
- đổi giá NC/tham số máy -> chuỗi tự cập nhật;
- không có `#REF!` do writer/migration.

#### Rename/migration

- rename VL-NC-M;
- rename DG Cạn/Nước/Biển;
- rename THKP;
- rename registered Gia DT TC;
- workbook VT/DN;
- workbook 1/2/3 khu vực.

#### Settings

- UI ảnh 08 ở DPI 100/125/150%;
- save/reopen settings;
- mapping rename/delete;
- reset default giữ mapping;
- auto-sync chỉ registered source;
- auto-save không Save As file mới;
- close workbook dispose timer/event.

## 11. Task NEXT — V2-801 Gói pháp lý & Dữ liệu

Mục này giữ đặc tả bàn giao. V2-801 đã triển khai và kiểm thử service trên Excel ngày
2026-10-06; xem [V2-801.md](V2-801.md). Task NEXT hiện tại là V2-901.

Đây là task tiếp theo.

### Trước khi code

1. đọc lại file này;
2. đọc `PROGRESS.md`;
3. đọc `UI-CONTRACT.md`;
4. kiểm tra `/mnt/data/chuan_UI`;
5. nếu ảnh 09 vẫn chưa có, dựng/chốt mockup 09 cùng phong cách ảnh 01–08 trước hoặc bám contract đã chốt; không tự sáng tạo kiểu UI khác.

### Thành phần cũ nên tái sử dụng

Tìm và tận dụng:

- `RegulationPackageStore`;
- package bootstrap/store;
- offline update;
- effective-date resolver;
- transition resolver;
- `WorkbookProjectProfileService`;
- `WorkbookPackageMigrationService`;
- package migration preview/rollback;
- diagnostics hiện có.

Không viết lại package engine nếu logic cũ đã đúng.

### Behavior bắt buộc

Màn hình phải thể hiện rõ package workbook đang pin:

```text
PackageId
DataVersion
Checksum
```

Trạng thái phải phân biệt:

- chưa cấu hình;
- đã cài/sẵn sàng;
- missing;
- corrupt;
- package khác có sẵn để người dùng chọn.

Nếu workbook pin package A nhưng máy thiếu A:

- vẫn mở module;
- hiển thị package A đang thiếu;
- cho cài lại/chọn hướng xử lý;
- không tự chuyển sang B/latest.

Nếu người dùng chủ động đổi package:

- phải preview impact;
- xác nhận rõ;
- dùng migration/rollback;
- không âm thầm đổi legal binding.

Package UI không được chạy full scan nặng ở startup.

## 12. Sau V2-801

### V2-901 — Báo cáo & Xuất in

Cần làm:

- UI theo phong cách ảnh 10;
- chọn output sheet;
- validate PrintArea;
- xuất không gồm technical columns;
- hỗ trợ DG Biển khi có;
- workbook hiện tại là nguồn thật;
- không tạo sheet report rác.

Sau V2-901 mới làm một vòng regression/acceptance toàn V2 trên các workbook mẫu.

## 13. Tài liệu phải đọc theo thứ tự

1. `docs/du-toan-v2/HANDOVER.md`
2. `docs/du-toan-v2/PROGRESS.md`
3. `docs/du-toan-v2/UI-CONTRACT.md`
4. `docs/du-toan-v2/ARCHITECTURE.md`
5. `docs/du-toan-v2/MIGRATION-PLAN.md`
6. `docs/du-toan-v2/README.md`

Không dựa vào README cũ của module legacy để thay đổi các quyết định V2 đã chốt.

## 14. Cách tiếp tục task sau

Mỗi lượt chỉ làm **một task**.

Quy trình:

1. đọc trạng thái hiện tại;
2. kiểm tra code/commit liên quan;
3. đối chiếu UI chuẩn;
4. triển khai một task;
5. tự rà soát static;
6. nếu có runtime thì chạy thật;
7. cập nhật `PROGRESS.md`;
8. cập nhật `HANDOVER.md` nếu kiến trúc/roadmap thay đổi;
9. commit rõ ràng;
10. không sang task tiếp theo trong cùng lượt nếu task hiện tại chưa được chốt.
