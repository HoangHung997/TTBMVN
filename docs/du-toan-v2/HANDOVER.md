# HANDOVER — Dự toán RPBM V2

> **Bàn giao cập nhật: 2026-10-07.** Đây là file đầu tiên người tiếp quản phải đọc. Sau file này đọc lần lượt `PROGRESS.md`, `WINDOWS-VERIFICATION.md`, `UI-CONTRACT.md`, `ARCHITECTURE.md`, `MIGRATION-PLAN.md` và các tài liệu task chuyên biệt.

## 1. Trạng thái hiện tại

Nhánh chính: `main`.

Roadmap feature của Dự toán V2 đã được triển khai ở mức code qua V2-901. **Hiện không có task feature mới nào phải bắt đầu ngay. Việc NEXT là regression/acceptance end-to-end và đối chiếu tài chính**, sau đó mới quyết định có cần task sửa lỗi hay mở rộng tiếp.

| Task | Nội dung | Trạng thái |
|---|---|---|
| V2-001 | Bỏ gate khởi động Dự toán | DONE - implementation / runtime pending |
| V2-002 | CustomTaskPane + UI nền tảng | DONE - implementation / runtime pending |
| V2-101 | WorkItemId + binding định mức bền vững | DONE - implementation / runtime pending |
| V2-201 | VL-NC-M | DONE - implementation / runtime pending |
| V2-301 | DG Cạn / DG Nước / DG Biển | DONE - implementation / runtime pending |
| V2-401 | Gia DT TC + THKP-TC | DONE - implementation / runtime pending |
| V2-501 | Validation / safe repair | DONE - implementation / runtime pending |
| V2-601 | Legacy / migration / VT-DN / rename | DONE - implementation / runtime pending |
| V2-701 | Thiết lập chung | DONE - implementation / runtime pending |
| V2-801 | Gói pháp lý & Dữ liệu | DONE - implementation / targeted Excel service tests PASS; full regression pending |
| V2-802 | Quản lý, tạo/sửa/xóa gói và định mức | DONE - implementation / Core + UI smoke PASS; full regression pending |
| V2-803 | Editor bảng theo công tác, hao phí/variant, Khu vực | DONE - implementation / 84 Core tests + navigation/check smoke PASS; full editing regression pending |
| V2-901 | Báo cáo & Xuất in | DONE - implementation / PDF service tests PASS; full UI regression pending |

Bằng chứng Windows hiện có:

- build Release x64 đã chạy sạch trên Windows/MSBuild 18.10.1;
- bộ Core đã có các lần chạy PASS; mốc mới nhất được ghi nhận sau V2-803 là **84 tests PASS**;
- task pane V2 đã mở trong Excel thật;
- V2-801 đã test preview/cancel/pin/checksum/backup/rollback/save-reopen;
- V2-802 đã smoke test tạo từ mẫu, kiểm tra gói, lưu nháp;
- V2-803 đã smoke test navigation/check và lưu nháp ma trận;
- V2-901 đã test xuất PDF một/nhiều sheet, rename, chặn lỗi Excel, giữ file PDF cũ khi export lỗi, giữ active sheet/sheet count nguồn.

**Không được suy diễn các kiểm thử mục tiêu trên thành “full runtime PASS”.** Toàn bộ chuỗi Dự toán V2 vẫn chưa được regression end-to-end và chưa đối chiếu tài chính hoàn chỉnh.

Chi tiết bằng chứng: [WINDOWS-VERIFICATION.md](WINDOWS-VERIFICATION.md).

## 2. Cảnh báo về workbook mẫu

Workbook HoaLuNamDinh gốc có lỗi nguồn tồn tại sẵn:

- 285 cached error cells;
- 91 công thức chứa `#REF!`.

Vì vậy:

- không dùng tổng tiền có sẵn của file gốc để chứng minh engine V2 đúng;
- không sửa `#REF!` bằng cách đoán số;
- không bypass validation chỉ để xuất báo cáo;
- mọi test phải làm trên bản sao/fixture;
- file mẫu gốc phải giữ nguyên hash.

Con số tổng hiển thị trong PDF mẫu chỉ chứng minh chức năng xuất giữ nguyên dữ liệu hiện có, **không** chứng minh engine V2 đã tái tạo đúng toàn bộ dự toán.

## 3. Kiến trúc đã chốt — không được đảo ngược

### Excel là hồ sơ chính

Workbook giữ:

- khối lượng dự án;
- giá vật liệu thị trường;
- tham số giá;
- công thức;
- phân tích đơn giá;
- Gia DT TC;
- THKP-TC;
- các sheet in.

Database/package chỉ là thư viện:

- định mức;
- hao phí;
- VL/NC/M;
- thông số máy;
- cost rule;
- khu vực;
- căn cứ/văn bản;
- version/checksum.

Add-in chỉ là engine hỗ trợ:

- đăng ký bảng công tác;
- gắn định mức;
- sinh công thức/link;
- đồng bộ;
- migration;
- validation;
- package management;
- report/export.

### Không có số chết cho kết quả tính

Input có thể là giá trị trực tiếp:

- khối lượng;
- giá vật liệu thị trường;
- giá nhiên liệu/tham số;
- hao phí gốc package.

Kết quả phải đi bằng công thức/link Excel:

```text
Giá đầu vào
  -> Giá nhân công / Giá ca máy
  -> Đơn giá VL/NC/M
  -> Thành tiền công tác
  -> THKP
```

Không cho phép tính xong trong C# rồi ghi snapshot số cuối vào các output này.

### Không lấy khu vực làm biến kiến trúc

Dự án 1, 2 hay 3 khu vực vẫn dùng chung:

- WorkItemId;
- binding;
- RateId;
- resource aggregation;
- writer;
- validation.

Dòng nhóm/khu vực không được biến thành WorkItem.

## 4. Identity và persistence

### WorkItem

Nguồn thật là Custom XML trong workbook.

State chính:

```text
WorkItemId
SourceKey
NormCode
VariantCode
PackageId
DataVersion
PackageChecksum
Kind
Fingerprint
IsOrphaned
```

Không dùng RowIndex làm identity.

Anchor phụ trên sheet:

- `__TTB_ID`
- `__TTB_NORM`
- `__TTB_KIND`
- `__TTB_HASH`

Các cột này phải ẩn và nằm ngoài vùng in/output.

### Binding định mức

Ô “Định mức” trên worksheet chỉ là display/cache.

Binding pháp lý thật là:

```text
Custom XML state
+ PackageId
+ DataVersion
+ PackageChecksum
+ NormCode
+ VariantCode
```

**Không được lấy text legacy trong ô Định mức và tự coi là binding pháp lý.**

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
- DG Biển dùng metadata môi trường `Sea`.

Đổi tên tab không được làm tạo sheet duplicate.

## 5. Output chính và vùng dữ liệu

Mặc định chỉ cần các sheet hồ sơ:

- `THKP-TC`;
- `Gia DT TC`;
- `DG Can`;
- `DG Nuoc`;
- `VL-NC-M`;
- `DG Bien` chỉ khi thực sự có công tác biển.

Không tạo sheet lookup/rule/technical visible chỉ để code hoạt động.

Bảng công tác chuẩn:

```text
A  TT
B  Mã công tác
C  Định mức
D  Mô tả công việc
E  Đơn vị
F  Khối lượng
G:I  Đơn giá VL / NC / M
J:L  Thành tiền VL / NC / M
M... metadata/legacy hidden
```

G:L phải là formula/link.

## 6. UI contract bắt buộc

Nguồn chuẩn:

- `docs/du-toan-v2/UI-CONTRACT.md`;
- `/mnt/data/chuan_UI`.

Bộ ảnh chuẩn đang có:

1. `01-Tong-quan.png`
2. `02-Cong-tac.png`
3. `03-Gan-dinh-muc.png`
4. `04-VL-NC-M.png`
5. `05-DG-Can.png`
6. `06-DG-Nuoc.png`
7. `07-THKP-TC-Kiem-tra.png`
8. `08-Thiet-lap-chung.png`

Ảnh chuẩn 09–10 chưa có file trong bộ chuẩn đã bàn giao. Hai màn hình đã được triển khai theo UI-CONTRACT/phong cách 01–08, nhưng **không được tuyên bố pixel-match 100%** khi chưa có ảnh chuẩn để đối chiếu.

Quy tắc:

- UI chính là CustomTaskPane bên phải Excel, khoảng 430 px;
- Excel vẫn là vùng làm việc chính;
- không quay lại form lớn/menu dọc legacy;
- không startup wizard bắt buộc;
- không modal hóa luồng chính;
- icon dùng renderer nội bộ `EstimateUiIcons`;
- xanh lá / vàng / xanh dương / đỏ + card trắng + Segoe UI;
- mỗi màn hình mới phải cùng ngôn ngữ UI hiện tại.

## 7. Luồng sử dụng hiện tại

```text
Ribbon Dự toán
  -> Tổng quan
     -> Công tác
        -> đăng ký vùng
        -> reconcile WorkItemId
        -> migration legacy on-demand
     -> Gắn định mức
     -> VL-NC-M
     -> DG Cạn / Nước / Biển
     -> THKP-TC & Kiểm tra

Ribbon Thiết lập Chung
  -> EstimateSettingsPaneView

Ribbon Gói pháp lý / mục 7
  -> EstimatePackagesPaneView
     -> tạo/sửa gói
     -> editor theo công tác
     -> draft / check / install / pin / migration preview

Ribbon Báo cáo / Xuất in
  -> EstimateReportsPaneView
     -> chọn sheet
     -> validate
     -> PDF / Print Preview
```

Không yêu cầu có THKP, đủ role hoặc package chỉ để mở module. Dependency chỉ được chặn đúng chức năng cần dependency đó.

## 8. Tóm tắt từng khối code quan trọng

### WorkItem / binding

- `ExcelAddIn1.Core/EstimateV2State.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2StateService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2RegistrationService.cs`

Chịu trách nhiệm ID bền vững, fingerprint recovery, duplicate split, orphan, save/reopen và restore display.

### VL-NC-M

- `EstimateV2Resources.cs`
- `EstimateV2ResourcePriceSheetProjector.cs`
- `WorkbookEstimateV2ResourceService.cs`
- `WorkbookEstimateV2ResourceSheetWriter.cs`
- `EstimateResourcesPaneView.cs`

Chỉ sinh resource đang dùng; vật liệu là input thị trường; NC/M là formula output; giữ giá người dùng khi refresh.

### DG

- `EstimateV2Rates.cs`
- `WorkbookEstimateV2RateService.cs`
- `WorkbookEstimateV2RateSheetWriter.cs`
- `EstimateUnitRatesPaneView.cs`

Rate identity:

```text
PackageIdentity + NormCode + VariantCode -> RateId
```

Một RateId chỉ có một block dùng chung.

### Gia DT TC / THKP

- `WorkbookEstimateV2CostLinkService.cs`
- `WorkbookEstimateV2CostLinkWriter.cs`
- `EstimateCostSummaryPaneView.cs`

Writer chỉ quản lý vùng V2 sở hữu; không tự đoán tỷ lệ pháp lý còn lại trong THKP.

### Validation

- `EstimateV2Validation.cs`
- `WorkbookEstimateV2ValidationService.cs`

Phát hiện missing norm/rate/price/name, formula overwrite, lỗi Excel, orphan/duplicate/mismatch và layout cũ. Safe repair phải gọi đúng writer sở hữu.

### Compatibility / migration

- `EstimateV2Compatibility.cs`
- `WorkbookEstimateV2CompatibilityService.cs`

Hỗ trợ alias legacy, VT/DN, rename, 1/2/3 khu vực, package pin missing/corrupt. Nếu có nhiều VT/DN/copy không xác định được duy nhất thì **không tự chọn**.

### Settings

- `EstimateV2Settings.cs`
- `WorkbookEstimateV2SettingsService.cs`
- `EstimateSettingsPaneView.cs`
- `EstimateTaskPaneManager.cs`

Settings theo workbook; mapping output bền vững; auto-sync debounce; auto-save chỉ workbook đã có path.

### Package / legal data

Tài liệu chính:

- [V2-801.md](V2-801.md)
- [PACKAGE-MANAGEMENT.md](PACKAGE-MANAGEMENT.md)

File/code chính gồm:

- `RegulationPackageStore`;
- package bootstrap/update/migration;
- `WorkbookProjectProfileService`;
- `WorkbookPackageMigrationService`;
- `EstimatePackagesPaneView`;
- `RegulationPackageEditorForm`;
- `RegulationPackageAuthoring.cs`.

Package đã cài là immutable theo PackageId/version/checksum. Sửa phải tạo version mới.

### Report/export

Tài liệu chính:

- [V2-901.md](V2-901.md)

Pane:

- `EstimateReportsPaneView`.

Xuất PDF phải giữ workbook nguồn, PageSetup và formula chain; export lỗi không được ghi đè PDF cũ.

## 9. V2-802 / V2-803 — trạng thái editor package

V2-802 đã bổ sung quản lý/tạo/sửa/xóa gói và định mức ở mục 7.

V2-803 đã chuyển editor sang cách làm theo từng công tác/quy tắc thay vì một bảng phẳng dài:

- định mức 2025: hàng là hao phí VL/NC/M, cột là variant;
- giữ nguyên identity variant thật;
- có tab thông tin/căn cứ/điều kiện;
- Khu vực dùng GeographyZone và ProvisionalEstimateRate;
- chặn variant trùng, hao phí âm;
- gói 2021 không có rates chi tiết thì không tự bịa hao phí;
- thay đổi chưa kiểm chứng sẽ về trạng thái Unverified;
- dữ liệu Unverified được lưu draft nhưng không được đóng gói/cài.

Giới hạn vẫn phải regression:

- chưa test đầy đủ edit cell/đổi mã qua UI;
- chưa test đầy đủ thêm/bớt variant;
- điều kiện/adjustment phức tạp vẫn còn representation kỹ thuật;
- chưa phải editor “bản sao 100% mọi bảng văn bản”;
- cài/xóa/pin package qua toàn bộ UI cần test trên kho test, không đụng kho người dùng thật.

## 10. Những điều tuyệt đối không được phá

1. Không đưa `FrmProjectSetup` hoặc role/package validation trở lại startup gate.
2. Không bắt có THKP mới cho mở Dự toán.
3. Không tự nâng pinned package sang latest.
4. Không dùng RowIndex làm WorkItem identity.
5. Không biến text định mức legacy thành legal binding.
6. Không ghi output tính toán thành số chết.
7. Không tự điền market price bị thiếu.
8. Không tự suy đoán adjustment/điều kiện định mức.
9. Không tự chọn VT hoặc DN khi có nhiều bộ.
10. Không xóa output legacy âm thầm; chỉ ẩn sau rebuild an toàn.
11. Không tạo tab kỹ thuật visible.
12. Không hard-code tỷ lệ pháp lý từ mockup.
13. Không thay UI đã chốt bằng form lớn/modal.
14. Không ghi runtime PASS nếu chưa chạy thật.
15. Không sửa/cài/xóa gói đang dùng mà bỏ qua backup/validation/preview.
16. Không thay dữ liệu pháp lý “cho tiện test” trên kho người dùng thật.

## 11. Việc NEXT — regression/acceptance toàn V2

**Đây là phần người tiếp quản phải làm tiếp. Không mở feature task mới trước khi xử lý danh sách này.**

### Bước 1 — baseline build

Trên Windows có Excel x64/VSTO:

```powershell
./scripts/build-release.ps1 -Platform x64 -Configuration Release
./ExcelAddIn1.Tests/bin/Release/ExcelAddIn1.Tests.exe
```

Nếu số test tăng so với 84 thì ghi đúng số thực tế. Không hard-code “84 PASS” nếu lần chạy mới khác.

### Bước 2 — chạy regression script

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File scripts/test-estimate-v2.ps1
```

Luôn dùng bản sao/fixture. Không ghi vào file gốc.

### Bước 3 — end-to-end formula chain

Trên fixture sạch, chạy từ đầu đến cuối:

```text
Đăng ký Công tác
 -> Gắn định mức
 -> VL-NC-M
 -> DG
 -> Gia DT TC
 -> THKP
 -> Validation
 -> PDF
```

Sau đó thay đổi lần lượt:

- giá vật liệu;
- giá nhân công/tham số lương;
- nhiên liệu/tham số máy;
- khối lượng.

Kết quả phải tự cập nhật bằng công thức đến THKP, không cần C# ghi lại số cuối.

### Bước 4 — identity regression

Bắt buộc:

- insert row;
- delete row;
- sort;
- sort visible-only nếu có;
- copy/paste dòng;
- duplicate WorkItemId;
- orphan;
- xóa ô Định mức display;
- save/close/open.

Không được mất/gắn nhầm binding.

### Bước 5 — workbook variation

Test:

- 1 khu vực;
- 2 khu vực;
- 3 khu vực;
- VT;
- DN;
- workbook có cả VT/DN;
- rename VL-NC-M;
- rename DG Cạn/Nước/Biển;
- rename THKP;
- rename registered Gia DT TC;
- missing/corrupt package pin.

### Bước 6 — DG Nước/Biển và đối tượng lương

Đây là vùng chưa được runtime coverage đầy đủ:

- DG Nước;
- DG Biển;
- hai đối tượng lương;
- logical resource candidate;
- NC điều khiển máy;
- save/reopen.

### Bước 7 — package authoring regression

Trên **kho test riêng**:

- tạo package từ mẫu;
- sửa metadata;
- sửa hao phí;
- đổi mã hợp lệ/không hợp lệ;
- lưu draft;
- mở lại draft;
- check;
- build;
- install;
- pin/migrate preview;
- rollback;
- remove/restore;
- add/remove variant theo rule hiện có;
- kiểm tra Unverified bị chặn install.

Không test phá hủy trên package store người dùng thật.

### Bước 8 — UI/DPI

Đối chiếu ảnh chuẩn 01–08 tại:

- 100%;
- 125%;
- 150%.

Kiểm tra:

- không tràn pane;
- không che Excel quá mức;
- font/icon/card đúng ngôn ngữ đã chốt;
- theme tối Office không làm chữ trắng khó đọc.

09–10 chưa có ảnh chuẩn thì chỉ đánh giá theo UI-CONTRACT, không ghi pixel-match 100%.

### Bước 9 — report/print

- PrintArea đúng;
- technical columns không lọt vùng in;
- PDF một/nhiều sheet;
- lỗi export giữ PDF cũ;
- Print Preview không tự in;
- workbook nguồn không tự Save;
- active sheet/sheet count được phục hồi.

### Bước 10 — đối chiếu tài chính

Chỉ dùng fixture/workbook không có lỗi nguồn để so:

- VL;
- NC;
- M;
- đơn giá;
- thành tiền;
- THKP.

Mọi sai lệch phải trace ngược về input/hao phí/formula/package. Không “chỉnh số cho khớp”.

## 12. Tiêu chí để nâng trạng thái runtime PASS

Chỉ cập nhật `PROGRESS.md` sang runtime PASS cho phần nào khi có bằng chứng thật:

- build log;
- test output;
- workbook fixture;
- expected/actual;
- ảnh/screenshot khi là UI;
- PDF khi là report;
- hash nguồn khi cần bảo đảm không sửa file gốc.

Nếu chỉ code review/static check thì vẫn để runtime pending.

## 13. Commit mốc cần biết

Các mốc gần hiện tại:

- `f265c85f0129` — sửa build blockers + Windows takeover verification;
- `153b3e4a8f27` — V2-801 package pane + transactional binding preview;
- `36d53babc303` — V2-901 PDF reporting + Windows Excel runtime verification;
- `44fb90630cd2` — V2-802 package authoring/management;
- `0c135a6ebca3` — ghi tài liệu implementation package management;
- `d202c86e110e` — V2-803 editor định mức theo bảng/công tác + khu vực.

Tài liệu tiến độ đã được refresh sau các mốc trên. Luôn kiểm tra `git log`/commit mới hơn trước khi sửa tiếp.

## 14. Tài liệu phải đọc theo thứ tự

1. `docs/du-toan-v2/HANDOVER.md`
2. `docs/du-toan-v2/PROGRESS.md`
3. `docs/du-toan-v2/WINDOWS-VERIFICATION.md`
4. `docs/du-toan-v2/UI-CONTRACT.md`
5. `docs/du-toan-v2/ARCHITECTURE.md`
6. `docs/du-toan-v2/MIGRATION-PLAN.md`
7. `docs/du-toan-v2/V2-801.md`
8. `docs/du-toan-v2/PACKAGE-MANAGEMENT.md`
9. `docs/du-toan-v2/V2-901.md`
10. `docs/du-toan-v2/README.md`

Không dùng README/module legacy để đảo ngược các quyết định V2 đã chốt.

## 15. Cách người tiếp quản tiếp tục

Mỗi lượt chỉ làm **một task hoặc một nhóm regression có ranh giới rõ**.

Quy trình:

1. đọc trạng thái hiện tại;
2. kiểm tra commit/code liên quan;
3. nếu có UI thì đọc `/mnt/data/chuan_UI` trước;
4. tạo fixture/bản sao an toàn;
5. chạy test hiện có trước khi sửa;
6. sửa một vấn đề;
7. chạy lại test;
8. ghi bằng chứng thật;
9. cập nhật `PROGRESS.md`;
10. cập nhật file chuyên biệt/HANDOVER nếu kiến trúc hoặc giới hạn thay đổi;
11. commit message rõ ràng;
12. không sang việc khác nếu regression hiện tại chưa chốt.

**Điểm bắt đầu khuyến nghị:** baseline build + Core tests + `scripts/test-estimate-v2.ps1`, sau đó end-to-end formula chain trên fixture sạch.
