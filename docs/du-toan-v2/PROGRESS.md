# TIẾN ĐỘ TRIỂN KHAI DỰ TOÁN RPBM V2

> Tài liệu này là nguồn trạng thái chính cho quá trình tiếp quản và phát triển module Dự toán V2.
> Mỗi task phải ghi rõ: mục tiêu, thay đổi đã làm, file liên quan, trạng thái kiểm tra và việc tiếp theo.
> Người tiếp quản bắt đầu từ `HANDOVER.md`, sau đó đọc file này cùng README, ARCHITECTURE, MIGRATION-PLAN và UI-CONTRACT để biết dự án đang ở đâu.

## Trạng thái tổng quan

| Mã | Task | Trạng thái |
|---|---|---|
| V2-001 | Bỏ gate khởi động Dự toán | DONE - implementation / runtime pending |
| V2-002 | Task pane UI theo bộ ảnh đã chốt | DONE - shell + 3 màn hình đầu / runtime pending |
| V2-101 | WorkItemId + binding định mức bền vững | DONE - implementation / runtime pending |
| V2-201 | Tổng hợp và sinh VL-NC-M | DONE - implementation / runtime pending |
| V2-301 | Sinh DG Cạn / DG Nước / DG Biển | DONE - implementation / runtime pending |
| V2-401 | Link Gia DT TC + THKP-TC | DONE - implementation / runtime pending |
| V2-501 | Validation / phục hồi / phát hiện lỗi | DONE - implementation / runtime pending |
| V2-601 | Tương thích file cũ và migration | DONE - implementation / runtime pending |
| V2-701 | Thiết lập chung theo workbook | DONE - implementation / runtime pending |
| V2-801 | Gói pháp lý & Dữ liệu | NEXT - cần chốt/mockup UI 09 nếu ảnh chuẩn chưa có |
| V2-901 | Báo cáo & Xuất in | TODO |

---

## UI CONTRACT — BẮT BUỘC

UI V2 phải bám theo bộ ảnh/ngôn ngữ giao diện người dùng đã chốt và được mô tả trong:

- `docs/du-toan-v2/UI-CONTRACT.md`

Nguyên tắc không được tự ý thay đổi:

- UI chính là **CustomTaskPane dock bên phải Excel**, khoảng 430 px.
- Không quay lại form Dự toán lớn/menu dọc của bản legacy.
- Excel luôn là vùng làm việc chính bên trái.
- Màu xanh/vàng/xanh dương/đỏ, card trắng và mật độ UI phải đồng nhất với ảnh chốt.
- Icon thiếu phải tự vẽ bằng renderer nội bộ; không dùng bộ icon lệch phong cách.
- Màn hình nào chưa có mockup chi tiết thì phải dựng tiếp theo đúng ngôn ngữ thiết kế hiện tại trước khi làm khác.

---

## V2-001 — Bỏ gate khởi động Dự toán

**Trạng thái:** DONE - implementation / runtime pending  
**Commit gốc:** `63c54d886e2d3b463044a98a6328d4c04c8fe0c9`  
**Luồng UI hiện tại:** CustomTaskPane, không còn mở form `Dutoan` làm UI chính.

### Mục tiêu

Sửa hành vi cũ:

```text
Bấm Dự toán
  -> FrmProjectSetup modal
  -> map đủ role
  -> resolve package
  -> hợp lệ mới mở Dự toán
```

thành:

```text
Bấm Dự toán
  -> mở Trợ lý Dự toán ngay
  -> từng chức năng tự kiểm tra dependency khi cần
```

### Đã làm

- Bỏ `FrmProjectSetup` khỏi startup gate.
- `Ribbon1.btnDutoan_Click` hiện gọi `EstimateTaskPaneManager.Show(currentWorkbook)`.
- Không bắt map `THKP-TC` hoặc đủ 7 worksheet role trước khi vào module.
- Missing package không được phép chặn task pane.
- Task pane được đóng theo workbook trong `ThisAddIn.WorkbookBeforeClose`.

### File chính

- `ExcelAddIn1/Ribbon1.cs`
- `ExcelAddIn1/ThisAddIn.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneManager.cs`

### Runtime test còn thiếu

Chưa chạy build VSTO/Excel thật trong môi trường hiện tại. Khi chạy local phải kiểm tra:

- workbook chưa có role -> pane vẫn mở;
- workbook chưa pin package -> pane vẫn mở;
- workbook đóng -> pane đóng;
- không xuất hiện `FrmProjectSetup` trước pane.

---

## V2-002 — Task pane UI theo bộ ảnh chốt

**Trạng thái:** DONE - shell + 3 màn hình đầu / runtime pending

### Đã làm

Đã chuyển UI chính sang VSTO CustomTaskPane và dựng theo ảnh chốt:

1. **Tổng quan** — `EstimateTaskPaneControl`
2. **Công tác** — `EstimateWorkItemsPaneView`
3. **Gắn định mức** — `EstimateNormBindingPaneView`

Đã có renderer icon riêng:

- `EstimateUiIcons.cs`

Các icon hiện tại được vẽ bằng GDI+ để giữ đúng phong cách và tránh phụ thuộc file icon bên ngoài.

### Navigation đã nối

```text
Tổng quan
  -> Công tác
       -> Gắn định mức
  -> Gắn định mức
```

Các màn hình VL-NC-M, DG Cạn, DG Nước, THKP/Kiểm tra, Thiết lập, Gói pháp lý, Báo cáo sẽ được bổ sung ở đúng task nghiệp vụ tương ứng nhưng phải giữ nguyên UI contract.

### File chính

- `ExcelAddIn1/Winform/EstimateUiIcons.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneControl.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneManager.cs`
- `ExcelAddIn1/Winform/EstimateWorkItemsPaneView.cs`
- `ExcelAddIn1/Winform/EstimateNormBindingPaneView.cs`
- `docs/du-toan-v2/UI-CONTRACT.md`

---

## V2-101 — WorkItemId + binding định mức bền vững

**Trạng thái:** DONE - implementation / runtime pending

### Mục tiêu đã đạt ở code

Không dùng RowIndex làm identity.

Mỗi công tác có state:

```text
WorkItemId
SourceKey
NormCode
VariantCode
PackageId
DataVersion
Kind
Fingerprint
IsOrphaned
```

Nguồn sự thật là **Custom XML Part** trong workbook.

Các cột kỹ thuật ẩn làm anchor:

```text
__TTB_ID
__TTB_NORM
__TTB_KIND
__TTB_HASH
```

Cột kỹ thuật được đặt ngoài vùng dữ liệu đang dùng, sau đó hide. Không thêm sheet metadata visible.

### Hành vi đã cài đặt

- Đăng ký vùng công tác từ selection.
- Tự nhận diện các cột:
  - Mã công tác
  - Định mức
  - Mô tả công việc
  - Đơn vị
  - Khối lượng
- Sinh GUID cho dòng công tác mới.
- Lưu binding định mức trong Custom XML.
- Lưu SourceKey để nhiều bảng/sheet không ghi đè trạng thái của nhau.
- Mở task pane -> tự reconcile state đã lưu.
- Xóa ô “Định mức” hiển thị -> binding thật vẫn còn; reconcile có thể phục hồi lại.
- Chèn/xóa dòng -> không phụ thuộc RowIndex.
- Copy/paste gây trùng WorkItemId -> cấp ID mới cho dòng copy, giữ binding.
- Sort làm ID ẩn lệch với dữ liệu hiển thị -> fingerprint được dùng để phục hồi khi có match duy nhất.
- Dòng biến mất -> đánh orphaned theo đúng source, không làm hỏng source khác.
- “Bỏ gắn” trong UI mới thực sự xóa binding.

### Core / service đã thêm

- `ExcelAddIn1.Core/EstimateV2State.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2StateService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2RegistrationService.cs`

### UI đã nối

- `EstimateWorkItemsPaneView`
- `EstimateNormBindingPaneView`

### Test code đã thêm

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2StateRoundTrip`
- `EstimateV2Fingerprint`

**Chưa chạy test thực tế trong môi trường hiện tại**, vì chưa có build/runtime VSTO + Excel. Không ghi PASS giả.

### Checklist runtime bắt buộc

- register vùng công tác -> save -> đóng -> mở lại -> binding còn;
- delete ô định mức -> mở pane/reconcile -> định mức phục hồi;
- insert row -> dòng cũ giữ binding, dòng mới có ID mới;
- delete row -> dòng khác không lệch binding;
- sort toàn vùng -> binding đi theo công tác;
- sort chỉ cột visible -> fingerprint recovery không gắn nhầm;
- copy/paste row -> duplicate ID được tách;
- nhiều bảng/sheet -> orphan state không ghi đè lẫn nhau;
- thiếu package -> Công tác vẫn dùng được, Gắn định mức chỉ cảnh báo dependency.

---

## V2-201 — Tổng hợp và sinh VL-NC-M

**Trạng thái:** DONE - implementation / runtime pending

### Mục tiêu đã đạt ở code

Từ các WorkItem đã gắn định mức, V2 hiện:

1. Load đúng package/version/checksum đã pin trong từng binding.
2. Chỉ gom **resource thực sự đang dùng**, không dump toàn bộ catalog.
3. Tách resource theo vật liệu / nhân công / máy thi công.
4. Với resource logic (`-OR-`, `.DIVING`), VL-NC-M sinh các **ứng viên giá vật lý** cần thiết; lựa chọn chính xác dùng ở công tác nào để dành cho bước đơn giá.
5. Bổ sung cả **nhân công điều khiển máy** từ `MachineRateCatalog`, kể cả khi loại nhân công đó không xuất hiện trực tiếp trong hao phí định mức.
6. Sinh/cập nhật `VL-NC-M` bằng writer riêng, không ghi kết quả tính toán chết.
7. Giá người dùng đã nhập được ưu tiên giữ lại khi refresh.
8. Nhân công và giá ca máy dùng **công thức/liên kết Excel**; các đầu vào được neo bằng workbook Name ổn định.
9. Metadata đặt ngoài vùng in và bị ẩn; không tạo sheet kỹ thuật visible.

### UI đã hoàn thiện theo ảnh chuẩn 04

Màn hình `EstimateResourcesPaneView` đã đối chiếu trực tiếp:

- `/mnt/data/chuan_UI/04-VL-NC-M.png`

Đã có:

- 4 metric: Số vật liệu / Số nhân công / Số máy / Thiếu giá;
- 3 card tổng quan theo nhóm;
- icon Vật liệu / Nhân công / Máy thi công tự vẽ bằng GDI+;
- danh sách thiếu giá có STT, mã hiệu, tên tài nguyên, đơn vị, nhóm và nút điều hướng;
- 4 chức năng chính;
- icon `fx` cho công thức nhân công và calculator cho giá ca máy;
- thông báo cuối pane về formula/link, không dùng số chết;
- giao diện vẫn là CustomTaskPane bên phải, không chiếm vùng Excel chính.

### Writer VL-NC-M

File chính:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2ResourceSheetWriter.cs`

Hành vi chính:

- giữ A:F là vùng biểu mẫu in;
- giữ lại giá vật liệu, giá nhiên liệu, đầu vào nhân công đã có khi refresh;
- tạo workbook Name bền vững cho giá resource và input;
- tính nhân công bằng công thức Excel;
- tính giá ca máy bằng công thức Excel: nhiên liệu/năng lượng, sửa chữa, khấu hao, chi phí khác, nhân công điều khiển máy;
- áp dụng hệ số nhiên liệu phụ;
- áp dụng hệ số môi trường ăn mòn cho ngữ cảnh nước/biển theo calculator hiện có;
- không xóa các cột unrelated ngoài vùng writer quản lý;
- tìm và tái sử dụng vị trí metadata cũ để tránh metadata trôi sang phải sau mỗi lần refresh;
- ẩn toàn bộ cột phụ ngoài A:F.

### Đối chiếu mẫu Excel thực tế

Đã đọc cấu trúc file mẫu:

- `/mnt/data/Du toan RPBM HoaLuNamDinh_Ver1.xlsx`
- sheet `VL-NC-M`

Writer đã được chỉnh theo đặc điểm in của mẫu:

- dòng 1–2 ẩn;
- block in chính bắt đầu từ mục `I. GIÁ NHÂN CÔNG`;
- header phần giá ca máy được lặp khi in;
- phần `III. GIÁ VẬT LIỆU` là block in riêng;
- vật liệu có ghi chú `ĐG thị trường`;
- hướng giấy dọc;
- in đen trắng;
- scale 98%;
- chỉ A:F thuộc vùng in;
- các cột phụ bị ẩn và nằm ngoài vùng in.

Mẫu hiện hành có PrintArea dạng hai vùng; writer V2 tạo lại cùng nguyên tắc theo số dòng sinh thực tế, không hard-code số dòng của file mẫu.

### Core / service / UI đã thêm hoặc cập nhật

- `ExcelAddIn1.Core/EstimateV2Resources.cs`
- `ExcelAddIn1.Core/EstimateV2ResourcePriceSheetProjector.cs`
- `ExcelAddIn1.Core/EstimateV2ResourceNames.cs`
- `ExcelAddIn1.Core/EstimateV2ExcelNames.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2ResourceService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2ResourceSheetWriter.cs`
- `ExcelAddIn1/Winform/EstimateResourcesPaneView.cs`
- `ExcelAddIn1/Winform/EstimateUiIcons.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneControl.cs`

### Commit quan trọng

- `3d4ae4f3cead` — unique resource aggregation plan
- `0c21071546f7` — logical resource price projector
- `21a2c64a8908` — writer dùng các ứng viên giá vật lý
- `dc63fff7bce6` — đưa writer VL-NC-M vào project
- `85dc8569f4f9` — sửa công thức hệ số nhiên liệu phụ
- `9d046c2d3fe5` — căn vùng in VL-NC-M theo file mẫu
- `bbd823b83463` — icon riêng theo UI chuẩn
- `f1449438a336` — chỉnh card VL-NC-M theo ảnh chuẩn
- `b44fbac63b2d` — đưa nhân công điều khiển máy vào preview/tổng hợp

### Test code đã có

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2ResourcePlan`
- `EstimateV2PriceSheetProjection`

Test projection kiểm tra:

- `M010.DIVING` -> các máy lặn vật lý;
- `M010.002-OR-M010.003` -> hai ứng viên;
- vật liệu `MAT-GASOLINE-OR-DIESEL` -> xăng + dầu;
- merge resource trùng;
- output price-sheet không còn resource logic.

### Kiểm tra đã thực hiện trong môi trường hiện tại

Đã tự rà soát tĩnh các file V2-201 sau thay đổi:

- cân bằng ngoặc `{}`, `()`, `[]`;
- không còn `TODO/FIXME/NotImplementedException` trong các file V2-201 chính;
- đã sửa lỗi kiểu dữ liệu `PriceProfile.TryFind`: API trả `PriceProfilePrice`, không phải `PriceProfileEntry`;
- đã sửa tham chiếu hidden parameter của hệ số nhiên liệu;
- đã sửa metadata không bị dịch sang phải qua nhiều lần refresh;
- đã sửa preview để tính cả nhân công điều khiển máy.

**Chưa chạy build VSTO/Excel thật và chưa chạy bộ test console trong môi trường hiện tại. Không ghi PASS giả.**

### Checklist runtime bắt buộc khi chạy trên máy có Excel/VSTO

- mở task pane -> vào VL-NC-M không lỗi;
- số VL / NC / Máy trên pane khớp resource writer;
- công tác có máy yêu cầu NC điều khiển -> NC đó xuất hiện trong pane và sheet;
- logical resource -> đủ ứng viên giá vật lý, không còn mã logic trong VL-NC-M;
- sinh mới VL-NC-M -> công thức NC/M hoạt động;
- nhập giá vật liệu -> refresh -> giá vẫn còn;
- nhập giá nhiên liệu / NC -> refresh -> giá ca máy cập nhật;
- refresh nhiều lần -> metadata không chạy sang phải;
- G:P hoặc các cột phụ legacy không xuất hiện trong vùng in;
- PrintArea có 2 block giống nguyên tắc file mẫu;
- dòng tiêu đề 1–2 ẩn; header máy lặp khi in;
- đóng/mở workbook -> workbook Name và giá nhập vẫn đọc lại được;
- không có `#REF!`, `#VALUE!`, `#NAME?` trong công thức sinh.

### Việc tiếp theo

V2-201 đã chốt. V2-301 đã được triển khai ở section bên dưới.

---

## Các task tiếp theo

### V2-301 — DG Cạn / DG Nước / DG Biển

**Trạng thái:** DONE - implementation / runtime pending

#### UI đã đối chiếu ảnh chuẩn

Đã đọc và bám trực tiếp:

- `/mnt/data/chuan_UI/05-DG-Can.png`
- `/mnt/data/chuan_UI/06-DG-Nuoc.png`

View mới:

- `ExcelAddIn1/Winform/EstimateUnitRatesPaneView.cs`

DG Cạn hiện có đúng cấu trúc đã chốt:

- 4 metric: Định mức cần sinh / Đã sinh / Thiếu giá / Liên kết công thức;
- bước 1 chọn định mức cần sinh;
- bước 2 tùy chọn Sinh mới / Cập nhật từ VL-NC-M / Chỉ sinh định mức đang dùng;
- bước 3 preview VL / NC / M / Tổng cộng;
- nút xanh `Sinh đơn giá`;
- footer giải thích mỗi định mức chỉ sinh một block dùng chung.

DG Nước hiện có:

- 4 metric: Định mức nước / Đã sinh / Thiếu giá / Cảnh báo;
- quy trình 4 bước;
- danh sách định mức công tác dưới nước;
- trạng thái từng định mức;
- footer formula/link.

DG Biển dùng cùng ngôn ngữ UI của DG Nước và **chỉ hiện điều hướng khi thực sự có định mức biển đang dùng**. Không tạo sheet DG Biển rác khi dự toán không có công tác biển.

#### Identity đơn giá

Đã thêm core model:

- `ExcelAddIn1.Core/EstimateV2Rates.cs`

Đơn giá unique theo:

```text
PackageIdentity + NormCode + VariantCode
        -> RateId ổn định
```

Hai hoặc nhiều WorkItem dùng cùng package + định mức + variant chỉ sinh **một block đơn giá** và tăng `UsageCount`; không nhân bản block theo dòng công tác.

Đã xử lý binding legacy có variant rỗng nhưng định mức chỉ có một variant: normalize về variant thật trước khi tạo RateId để không sinh trùng.

#### Phân loại Cạn / Nước / Biển

- `NORM-000.*`, `NORM-010.*`, `NORM-020.*` -> DG Cạn;
- `NORM-030.*` -> DG Nước;
- `NORM-040.*` -> DG Biển.

Sheet chỉ được sinh khi môi trường đó có rate đang dùng.

Tên sheet tương thích mẫu hiện tại:

- `DG Can` / `DG Cạn`;
- `DG Nuoc` / `DG Nước`;
- `DG Bien` / `DG Biển`.

#### Writer đơn giá

Đã thêm:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2RateService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2RateSheetWriter.cs`

Writer:

1. cập nhật `VL-NC-M` trước để bảo đảm workbook Name giá tồn tại;
2. sinh block đơn giá từ định mức/variant đang dùng;
3. hao phí định mức nằm ở cột D;
4. đơn giá resource ở cột E là **formula link** tới workbook Name của `VL-NC-M`;
5. thành tiền:
   - VL -> F;
   - NC -> G;
   - M -> H;
6. vật liệu phần trăm tính bằng công thức từ tổng vật liệu trực tiếp;
7. dòng `Cộng:` giữ riêng VL / NC / M;
8. thêm dòng `Tổng cộng đơn giá` và workbook Name `TOTAL`;
9. metadata nằm từ cột I trở đi và bị ẩn;
10. vùng in chỉ A:H.

Không copy snapshot giá chết vào block đơn giá.

#### Logical resource

Các resource logic như:

- `M010.002-OR-M010.003`;
- `M010.DIVING`;
- `MAT-GASOLINE-OR-DIESEL`;

được mở thành các price candidate vật lý dùng chung với VL-NC-M.

Khi chưa có UI chọn candidate riêng cho từng công tác, V2-301 dùng policy xác định được: **chọn candidate đầu tiên đang có giá > 0** bằng công thức Excel. Pane đồng thời đánh cảnh báo để người dùng biết rate có lựa chọn cần rà soát.

#### Điều kiện / hệ số định mức

State V2 hiện mới lưu NormCode + Variant, chưa lưu tập điều kiện như lưu tốc dòng chảy, đào có nước, độ dốc...

Vì vậy V2-301 hiện sinh **hao phí cơ sở của variant đã gắn**. Nếu `NormDefinition` có adjustment/constraint, rate được đánh cảnh báo `RequiresConditionReview`.

Không tự đoán điều kiện và không ghi hệ số giả vào workbook.

Đây là giới hạn đã biết của V2-301, cần được xử lý ở luồng thiết lập/validation sau; không được âm thầm tính sai.

#### Đối chiếu sheet mẫu thực tế

Đã đọc/rendere:

- `/mnt/data/Du toan RPBM HoaLuNamDinh_Ver1.xlsx`
- `DG Can`
- `DG Nuoc`

Ảnh render dùng để đối chiếu:

- `/mnt/data/dg_can_sample.png`
- `/mnt/data/dg_nuoc_sample.png`

Writer đã chỉnh theo cấu trúc in thực tế:

- Times New Roman;
- A:H;
- title block;
- tên công tác / số hiệu định mức / đơn vị;
- header **hai tầng**:
  - A:E merge dọc;
  - F:H merge ngang `Thành tiền (đồng)`;
  - hàng dưới F/G/H = Vật liệu / Nhân công / Máy;
- section I Vật liệu / II Nhân công / III Máy thi công;
- dòng `Cộng:`;
- Portrait;
- BlackAndWhite;
- FitToPagesWide = 1;
- helper/metadata từ I trở đi ẩn.

#### Navigation

`EstimateTaskPaneControl` đã nối:

```text
Tổng quan
  -> DG Cạn
  -> DG Nước
       -> DG Biển (chỉ khi có định mức biển)
```

Output tile DG Biển cũng chỉ xuất hiện nếu workbook thực sự có sheet biển.

#### Test code đã thêm

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2RatePlan`

Test kiểm tra:

- 2 WorkItem cùng Land norm + variant -> 1 rate, UsageCount = 2;
- tách đúng Land / InlandWater / Sea;
- RateId ổn định;
- `M010.DIVING` có candidate máy lặn cụ thể;
- rate nước có adjustment/constraint được đánh cần rà soát;
- rate biển giữ đúng logical machine để writer liên kết candidate.

#### Commit quan trọng

- `8b40d688ca6d` — expose resource candidate expansion;
- `7e311d8ed845` — core rate plan / RateId;
- `f9e90e8314e3` — preview service DG;
- `58b6120af5e2` — formula-linked DG writer;
- `66960b684b07` — UI DG Cạn/Nước/Biển;
- `8edc18670b88` — navigation vào task pane;
- `71e4764eba62` — core rate tests;
- `c788707acfa4` — merge normalized rate identities;
- `3ba4c36396f6` — stable TOTAL link;
- `0a9dd85097b1` + `c0a564583d1a` — header in hai tầng theo mẫu;
- `8bdf5dbb23d0` — cập nhật UI contract.

#### Tự kiểm tra trong môi trường hiện tại

Đã rà soát tĩnh các file V2-301 chính:

- ngoặc `{}`, `()`, `[]` cân bằng;
- không có `TODO/FIXME/NotImplementedException`;
- kiểm tra lại API MachineRateCatalog dùng bởi writer;
- kiểm tra selector/radio DG Cạn thực sự ảnh hưởng tập rate sinh;
- commit checkbox grid trước khi đọc lựa chọn;
- DG Biển không hiện khi không có rate biển;
- công thức DG dùng workbook Name từ VL-NC-M, không dùng số giá snapshot;
- header sheet đã sửa từ một tầng sang hai tầng theo file mẫu.

**Chưa chạy build VSTO/Excel thật và chưa chạy test console trong môi trường hiện tại. Không ghi PASS giả.**

#### Checklist runtime bắt buộc

- vào DG Cạn từ Tổng quan -> pane mở đúng, không form modal;
- số metric khớp rate plan;
- hai công tác cùng norm+variant -> chỉ một block DG;
- chọn/bỏ chọn rate -> Sinh mới chỉ sinh đúng tập mong muốn và giữ block đã có;
- Cập nhật từ VL-NC-M -> giá đổi thì DG đổi theo formula/link;
- DG Nước chỉ sinh khi có NORM-030;
- DG Biển chỉ sinh khi có NORM-040;
- workbook chỉ có biển -> từ DG Nước vẫn thấy điều hướng DG Biển;
- header in hai tầng đúng A:H;
- cột I trở đi hidden;
- không có `#NAME?`, `#REF!`, `#VALUE!`;
- TOTAL = VL + NC + M;
- save/close/open -> workbook Name rate vẫn còn;
- rate có adjustment/constraint phải hiện cảnh báo, không âm thầm coi như đã áp hệ số.

#### Việc tiếp theo

V2-301 đã chốt. V2-401 đã được triển khai ở section bên dưới.

### V2-401 — Gia DT TC + THKP-TC

**Trạng thái:** DONE - implementation / runtime pending

#### UI đã đối chiếu ảnh chuẩn

Đã đọc và bám trực tiếp:

- `/mnt/data/chuan_UI/07-THKP-TC-Kiem-tra.png`

View mới:

- `ExcelAddIn1/Winform/EstimateCostSummaryPaneView.cs`

Màn hình hiện có:

- 4 metric: Tổng công tác / Đủ đơn giá / Cảnh báo / Lỗi công thức;
- card xanh lớn `Cập nhật THKP-TC`;
- khu `Kết quả kiểm tra hồ sơ` với icon trạng thái và `Xem chi tiết`;
- 3 card chức năng: Kiểm tra hồ sơ / Xuất báo cáo / Mở thư mục hồ sơ;
- footer xanh dương đúng ngôn ngữ thiết kế của ảnh chuẩn;
- vẫn là CustomTaskPane bên phải, không mở form modal.

`Kiểm tra hồ sơ` trong V2-401 chỉ kiểm tra các dependency/link cơ bản cần cho bước tổng hợp. Validation đầy đủ vẫn để đúng task V2-501.

#### Đối chiếu file dự toán thực tế

Đã đọc cấu trúc các mẫu:

- `/mnt/data/Du toan RPBM HoaLuNamDinh_Ver1.xlsx`
- `/mnt/data/Du toan RPBM Pleiku_TP2_QuyNhon_Ver7.2.xlsx`

Các mẫu cho thấy:

- `Gia DT TC` hoặc các sheet `Gia DT TC_...` chứa công tác chi tiết;
- đơn giá VL / NC / M được dùng để tính thành tiền theo từng công tác;
- `THKP-TC` có các dòng chi phí trực tiếp `VL`, `NC`, `M`, `T`;
- các dòng sau đó (chi phí chung, TL, K1..Kn, VAT, tổng cuối...) phụ thuộc vào phần chi phí trực tiếp và khác nhau theo mẫu/pháp lý.

Vì vậy V2-401 **chỉ thay thế nguồn liên kết phần chi phí trực tiếp**, không phá công thức pháp lý phía sau của THKP-TC hiện hữu.

#### WorkItem -> RateId

Đã thêm core model:

- `ExcelAddIn1.Core/EstimateV2CostLinks.cs`

Mỗi WorkItem được resolve theo state bền vững:

```text
WorkItemId
  -> PackageIdentity + NormCode + VariantCode
  -> RateId
  -> workbook Name VL / NC / M
```

Không dùng RowIndex làm identity.

Hai WorkItem dùng chung một định mức + variant tiếp tục dùng chung một RateId nhưng mỗi dòng có khối lượng và thành tiền riêng.

#### Gia DT TC / bảng công tác chi tiết

Writer:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkWriter.cs`

Trên mỗi bảng công tác đã đăng ký, sáu cột kết quả được đặt ngay sau cột Khối lượng:

```text
Đơn giá VL
Đơn giá NC
Đơn giá M
Thành tiền VL
Thành tiền NC
Thành tiền M
```

Với layout chuẩn A:F, chúng tương ứng G:L.

Công thức:

```text
Đơn giá VL = RateId.VL
Đơn giá NC = RateId.NC
Đơn giá M  = RateId.M

Thành tiền VL = Khối lượng × Đơn giá VL
Thành tiền NC = Khối lượng × Đơn giá NC
Thành tiền M  = Khối lượng × Đơn giá M
```

Các đơn giá là **formula tham chiếu workbook Name**, không ghi số chết.

Nếu WorkItem chưa gắn định mức hoặc chưa có DG tương ứng, writer không tự bịa giá và để dòng đó chưa liên kết để pane tiếp tục cảnh báo.

#### Sửa xung đột metadata quan trọng

V2-101 trước đây có thể đặt:

```text
__TTB_ID
__TTB_NORM
__TTB_KIND
__TTB_HASH
```

ngay sau cột visible cuối cùng. Với bảng có Khối lượng ở F, metadata cũ có thể rơi vào G:J — đúng vùng V2-401 cần dùng cho đơn giá/thành tiền.

Đã bổ sung migration:

- `WorkbookEstimateV2RegistrationService.EnsureTechnicalColumnsAfter(...)`

Khi V2-401 chạy:

1. metadata cũ được copy nguyên trạng tới vùng an toàn;
2. tối thiểu nằm sau sáu cột kết quả (M:P với layout A:F), hoặc xa hơn nếu workbook đã dùng các cột đó;
3. vùng metadata mới được hide;
4. custom property của source được cập nhật;
5. cột cũ được giải phóng để G:L là dữ liệu visible;
6. WorkItemId/binding vẫn giữ nguyên.

Đây là migration cấu trúc, không yêu cầu người dùng quét/gắn lại từ đầu.

#### Tổng hợp nhiều bảng / nhiều khu vực

V2-401 không phụ thuộc dự án có 1, 2 hay 3 khu vực.

Mỗi registered source được tổng hợp bằng WorkItemId hợp lệ; helper THKP dùng công thức `SUMIF` trên cột `__TTB_ID` và cột thành tiền tương ứng. Vì vậy:

- dòng nhóm/tổng phụ không có WorkItemId không bị cộng hai lần;
- nhiều bảng/sheet được cộng chung;
- không hard-code tên khu vực VT/DN hay số lượng khu vực.

#### THKP-TC

Đã thêm stable workbook Names:

- `TTBMVN_V2_GIADT_VL`
- `TTBMVN_V2_GIADT_NC`
- `TTBMVN_V2_GIADT_M`
- `TTBMVN_V2_GIADT_TOTAL`

Helper tổng hợp nằm ở các cột ẩn ngoài vùng in của THKP-TC.

Nếu `THKP-TC` đã tồn tại, writer:

1. tự tìm cột `Ký hiệu` và `Thành tiền`;
2. tự tìm các dòng có ký hiệu `VL`, `NC`, `M`, `T`;
3. thay **chỉ** công thức thành tiền của bốn dòng này bằng workbook Name V2;
4. giữ nguyên toàn bộ công thức phía sau của mẫu hiện hữu.

Không hard-code các tỷ lệ minh họa 6,5%, 5,5%, 10% trong ảnh UI.

Nếu workbook chưa có `THKP-TC`, V2-401 tạo một mẫu tối thiểu an toàn chỉ cho phần chi phí trực tiếp và ghi rõ các khoản pháp lý khác chưa được tự suy đoán.

#### Preview / trạng thái

Đã thêm:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkService.cs`

Preview hiện tính:

- tổng WorkItem;
- số WorkItem đã có RateId và workbook Name đơn giá;
- số WorkItem cảnh báo (không double-count cùng một WorkItem);
- lỗi Excel cơ bản trong các ô output;
- THKP có thật sự tham chiếu bốn workbook Name VL/NC/M/T hay chưa.

Chỉ việc workbook Name tồn tại chưa được coi là THKP đã cập nhật; service còn kiểm tra công thức trong THKP có reference các Name đó.

#### Navigation

`EstimateTaskPaneControl` đã nối bước cuối:

```text
Tổng quan
  -> THKP-TC & Kiểm tra
```

Bước 6 chỉ hiển thị `Đã xong` khi THKP-TC thực sự có các liên kết V2, không chỉ vì sheet `THKP-TC` tồn tại.

#### Test code đã thêm

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2CostLinkPlan`

Test core kiểm tra:

- 2 WorkItem cùng rate -> cùng RateId;
- WorkItem chưa gắn -> `Unbound`;
- binding không resolve rate -> `RateNotResolved`;
- đếm Ready / Unbound / MissingRate;
- propagate cờ `RequiresConditionReview`.

#### File chính đã thêm/cập nhật

- `ExcelAddIn1.Core/EstimateV2CostLinks.cs`
- `ExcelAddIn1.Core/EstimateV2ExcelNames.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2RegistrationService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkWriter.cs`
- `ExcelAddIn1/Winform/EstimateCostSummaryPaneView.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneControl.cs`

#### Commit quan trọng

- `e01a8ae866e4` — WorkItem -> RateId link plan;
- `c9d099a88dec` — stable Gia DT TC total names;
- `10e2b4564a72` — di chuyển metadata ra sau vùng đơn giá/thành tiền;
- `308abc67b3ed` — preview service;
- `56ad90b950c7` — writer link WorkItem + THKP;
- `45705188e220` — UI THKP-TC & Kiểm tra;
- `2384fb09bab8` — navigation;
- `99f76f0f34db` — warning metric không double-count;
- `77fe3a8d949c` — core link-plan test;
- `f4306c83aeee` — xác nhận THKP thực sự reference Name V2;
- `7f3ed7ae27fe` — trạng thái bước THKP ở Tổng quan;
- `a6565ed3c03f` — cleanup COM range khi tạo THKP tối thiểu;
- `80d05d920b5c` — cập nhật UI contract.

#### Tự kiểm tra trong môi trường hiện tại

Đã rà soát tĩnh các file V2-401:

- ngoặc `{}`, `()`, `[]` cân bằng;
- không có conflict marker;
- không có `TODO/FIXME/NotImplementedException` trong các file V2-401 chính;
- project file chứa đúng các source mới;
- metadata migration không dùng RowIndex làm identity;
- các giá/ thành tiền là formula/link, không phải số snapshot;
- THKP chỉ cập nhật VL/NC/M/T và không ghi đè các công thức pháp lý phía sau;
- helper THKP nằm ngoài vùng in và bị ẩn.

Đã đối chiếu logic với hai workbook mẫu bằng công cụ spreadsheet, nhưng **chưa chạy build VSTO/Excel thật và chưa chạy test console trong môi trường hiện tại. Không ghi PASS giả.**

#### Checklist runtime bắt buộc

- workbook cũ có metadata ở G:J -> cập nhật -> metadata chuyển sang M:P hoặc vùng an toàn, ID/binding không mất;
- G:L hiện đúng 6 cột đơn giá/thành tiền;
- đơn giá G:I là workbook Name của RateId;
- J:L = Khối lượng × đơn giá;
- section/group/subtotal không có WorkItemId không bị tính trùng;
- dự án 1/2/3 khu vực -> tổng VL/NC/M đúng tổng các WorkItem;
- sửa khối lượng -> thành tiền và THKP cập nhật theo công thức;
- sửa giá VL-NC-M -> DG -> Gia DT TC -> THKP cập nhật theo chuỗi link;
- WorkItem chưa có DG -> không có giá giả, pane báo cảnh báo;
- THKP hiện hữu -> chỉ VL/NC/M/T bị thay source link; các dòng pháp lý phía sau giữ nguyên;
- THKP không tồn tại -> tạo mẫu tối thiểu trực tiếp, không tự gán tỷ lệ pháp lý;
- đóng/mở workbook -> workbook Names/link còn nguyên;
- pane 07 mở đúng, 4 metric và card/action không tràn ở DPI 100/125/150%;
- không có `#REF!`, `#NAME?`, `#VALUE!` sau cập nhật.

#### Việc tiếp theo

V2-401 đã chốt. V2-501 đã được triển khai ở section bên dưới.

### V2-501 — Validation / phục hồi / phát hiện lỗi

**Trạng thái:** DONE - implementation / runtime pending

#### UI giữ nguyên ảnh chuẩn 07

Đã đọc lại trực tiếp:

- `/mnt/data/chuan_UI/07-THKP-TC-Kiem-tra.png`

Không tạo màn hình validation mới. V2-501 mở rộng chính:

- `ExcelAddIn1/Winform/EstimateCostSummaryPaneView.cs`

Giữ nguyên bố cục đã chốt:

- 4 metric Tổng công tác / Đủ đơn giá / Cảnh báo / Lỗi công thức;
- card xanh `Cập nhật THKP-TC`;
- khu `Kết quả kiểm tra hồ sơ`;
- nút `Xem chi tiết`;
- 3 card chức năng phía dưới;
- footer thông tin;
- CustomTaskPane dock bên phải Excel.

Card `Kiểm tra hồ sơ` hiện chạy validation V2 đầy đủ. `Xem chi tiết` ưu tiên mở đúng sheet và select đúng ô lỗi khi validation xác định được địa chỉ.

#### Core rule mới

Đã thêm:

- `ExcelAddIn1.Core/EstimateV2Validation.cs`

Core rule chịu trách nhiệm:

- sinh công thức chuẩn G:L theo WorkItem/RateId;
- sinh bốn công thức chuẩn VL/NC/M/T của THKP;
- chuẩn hóa công thức để so sánh không bị nhiễu bởi `# TIẾN ĐỘ TRIỂN KHAI DỰ TOÁN RPBM V2

> Tài liệu này là nguồn trạng thái chính cho quá trình tiếp quản và phát triển module Dự toán V2.
> Mỗi task phải ghi rõ: mục tiêu, thay đổi đã làm, file liên quan, trạng thái kiểm tra và việc tiếp theo.
> Người tiếp quản sau chỉ cần đọc file này cùng README, ARCHITECTURE, MIGRATION-PLAN và UI-CONTRACT là biết dự án đang ở đâu.

## Trạng thái tổng quan

| Mã | Task | Trạng thái |
|---|---|---|
| V2-001 | Bỏ gate khởi động Dự toán | DONE - implementation / runtime pending |
| V2-002 | Task pane UI theo bộ ảnh đã chốt | DONE - shell + 3 màn hình đầu / runtime pending |
| V2-101 | WorkItemId + binding định mức bền vững | DONE - implementation / runtime pending |
| V2-201 | Tổng hợp và sinh VL-NC-M | DONE - implementation / runtime pending |
| V2-301 | Sinh DG Cạn / DG Nước / DG Biển | DONE - implementation / runtime pending |
| V2-401 | Link Gia DT TC + THKP-TC | DONE - implementation / runtime pending |
| V2-501 | Validation / phục hồi / phát hiện lỗi | DONE - implementation / runtime pending |
| V2-601 | Tương thích file cũ và migration | NEXT |

---

## UI CONTRACT — BẮT BUỘC

UI V2 phải bám theo bộ 10 ảnh người dùng đã chốt và được mô tả trong:

- `docs/du-toan-v2/UI-CONTRACT.md`

Nguyên tắc không được tự ý thay đổi:

- UI chính là **CustomTaskPane dock bên phải Excel**, khoảng 430 px.
- Không quay lại form Dự toán lớn/menu dọc của bản legacy.
- Excel luôn là vùng làm việc chính bên trái.
- Màu xanh/vàng/xanh dương/đỏ, card trắng và mật độ UI phải đồng nhất với ảnh chốt.
- Icon thiếu phải tự vẽ bằng renderer nội bộ; không dùng bộ icon lệch phong cách.
- Màn hình nào chưa có mockup chi tiết thì phải dựng tiếp theo đúng ngôn ngữ thiết kế hiện tại trước khi làm khác.

---

## V2-001 — Bỏ gate khởi động Dự toán

**Trạng thái:** DONE - implementation / runtime pending  
**Commit gốc:** `63c54d886e2d3b463044a98a6328d4c04c8fe0c9`  
**Luồng UI hiện tại:** CustomTaskPane, không còn mở form `Dutoan` làm UI chính.

### Mục tiêu

Sửa hành vi cũ:

```text
Bấm Dự toán
  -> FrmProjectSetup modal
  -> map đủ role
  -> resolve package
  -> hợp lệ mới mở Dự toán
```

thành:

```text
Bấm Dự toán
  -> mở Trợ lý Dự toán ngay
  -> từng chức năng tự kiểm tra dependency khi cần
```

### Đã làm

- Bỏ `FrmProjectSetup` khỏi startup gate.
- `Ribbon1.btnDutoan_Click` hiện gọi `EstimateTaskPaneManager.Show(currentWorkbook)`.
- Không bắt map `THKP-TC` hoặc đủ 7 worksheet role trước khi vào module.
- Missing package không được phép chặn task pane.
- Task pane được đóng theo workbook trong `ThisAddIn.WorkbookBeforeClose`.

### File chính

- `ExcelAddIn1/Ribbon1.cs`
- `ExcelAddIn1/ThisAddIn.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneManager.cs`

### Runtime test còn thiếu

Chưa chạy build VSTO/Excel thật trong môi trường hiện tại. Khi chạy local phải kiểm tra:

- workbook chưa có role -> pane vẫn mở;
- workbook chưa pin package -> pane vẫn mở;
- workbook đóng -> pane đóng;
- không xuất hiện `FrmProjectSetup` trước pane.

---

## V2-002 — Task pane UI theo bộ ảnh chốt

**Trạng thái:** DONE - shell + 3 màn hình đầu / runtime pending

### Đã làm

Đã chuyển UI chính sang VSTO CustomTaskPane và dựng theo ảnh chốt:

1. **Tổng quan** — `EstimateTaskPaneControl`
2. **Công tác** — `EstimateWorkItemsPaneView`
3. **Gắn định mức** — `EstimateNormBindingPaneView`

Đã có renderer icon riêng:

- `EstimateUiIcons.cs`

Các icon hiện tại được vẽ bằng GDI+ để giữ đúng phong cách và tránh phụ thuộc file icon bên ngoài.

### Navigation đã nối

```text
Tổng quan
  -> Công tác
       -> Gắn định mức
  -> Gắn định mức
```

Các màn hình VL-NC-M, DG Cạn, DG Nước, THKP/Kiểm tra, Thiết lập, Gói pháp lý, Báo cáo sẽ được bổ sung ở đúng task nghiệp vụ tương ứng nhưng phải giữ nguyên UI contract.

### File chính

- `ExcelAddIn1/Winform/EstimateUiIcons.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneControl.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneManager.cs`
- `ExcelAddIn1/Winform/EstimateWorkItemsPaneView.cs`
- `ExcelAddIn1/Winform/EstimateNormBindingPaneView.cs`
- `docs/du-toan-v2/UI-CONTRACT.md`

---

## V2-101 — WorkItemId + binding định mức bền vững

**Trạng thái:** DONE - implementation / runtime pending

### Mục tiêu đã đạt ở code

Không dùng RowIndex làm identity.

Mỗi công tác có state:

```text
WorkItemId
SourceKey
NormCode
VariantCode
PackageId
DataVersion
Kind
Fingerprint
IsOrphaned
```

Nguồn sự thật là **Custom XML Part** trong workbook.

Các cột kỹ thuật ẩn làm anchor:

```text
__TTB_ID
__TTB_NORM
__TTB_KIND
__TTB_HASH
```

Cột kỹ thuật được đặt ngoài vùng dữ liệu đang dùng, sau đó hide. Không thêm sheet metadata visible.

### Hành vi đã cài đặt

- Đăng ký vùng công tác từ selection.
- Tự nhận diện các cột:
  - Mã công tác
  - Định mức
  - Mô tả công việc
  - Đơn vị
  - Khối lượng
- Sinh GUID cho dòng công tác mới.
- Lưu binding định mức trong Custom XML.
- Lưu SourceKey để nhiều bảng/sheet không ghi đè trạng thái của nhau.
- Mở task pane -> tự reconcile state đã lưu.
- Xóa ô “Định mức” hiển thị -> binding thật vẫn còn; reconcile có thể phục hồi lại.
- Chèn/xóa dòng -> không phụ thuộc RowIndex.
- Copy/paste gây trùng WorkItemId -> cấp ID mới cho dòng copy, giữ binding.
- Sort làm ID ẩn lệch với dữ liệu hiển thị -> fingerprint được dùng để phục hồi khi có match duy nhất.
- Dòng biến mất -> đánh orphaned theo đúng source, không làm hỏng source khác.
- “Bỏ gắn” trong UI mới thực sự xóa binding.

### Core / service đã thêm

- `ExcelAddIn1.Core/EstimateV2State.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2StateService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2RegistrationService.cs`

### UI đã nối

- `EstimateWorkItemsPaneView`
- `EstimateNormBindingPaneView`

### Test code đã thêm

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2StateRoundTrip`
- `EstimateV2Fingerprint`

**Chưa chạy test thực tế trong môi trường hiện tại**, vì chưa có build/runtime VSTO + Excel. Không ghi PASS giả.

### Checklist runtime bắt buộc

- register vùng công tác -> save -> đóng -> mở lại -> binding còn;
- delete ô định mức -> mở pane/reconcile -> định mức phục hồi;
- insert row -> dòng cũ giữ binding, dòng mới có ID mới;
- delete row -> dòng khác không lệch binding;
- sort toàn vùng -> binding đi theo công tác;
- sort chỉ cột visible -> fingerprint recovery không gắn nhầm;
- copy/paste row -> duplicate ID được tách;
- nhiều bảng/sheet -> orphan state không ghi đè lẫn nhau;
- thiếu package -> Công tác vẫn dùng được, Gắn định mức chỉ cảnh báo dependency.

---

## V2-201 — Tổng hợp và sinh VL-NC-M

**Trạng thái:** DONE - implementation / runtime pending

### Mục tiêu đã đạt ở code

Từ các WorkItem đã gắn định mức, V2 hiện:

1. Load đúng package/version/checksum đã pin trong từng binding.
2. Chỉ gom **resource thực sự đang dùng**, không dump toàn bộ catalog.
3. Tách resource theo vật liệu / nhân công / máy thi công.
4. Với resource logic (`-OR-`, `.DIVING`), VL-NC-M sinh các **ứng viên giá vật lý** cần thiết; lựa chọn chính xác dùng ở công tác nào để dành cho bước đơn giá.
5. Bổ sung cả **nhân công điều khiển máy** từ `MachineRateCatalog`, kể cả khi loại nhân công đó không xuất hiện trực tiếp trong hao phí định mức.
6. Sinh/cập nhật `VL-NC-M` bằng writer riêng, không ghi kết quả tính toán chết.
7. Giá người dùng đã nhập được ưu tiên giữ lại khi refresh.
8. Nhân công và giá ca máy dùng **công thức/liên kết Excel**; các đầu vào được neo bằng workbook Name ổn định.
9. Metadata đặt ngoài vùng in và bị ẩn; không tạo sheet kỹ thuật visible.

### UI đã hoàn thiện theo ảnh chuẩn 04

Màn hình `EstimateResourcesPaneView` đã đối chiếu trực tiếp:

- `/mnt/data/chuan_UI/04-VL-NC-M.png`

Đã có:

- 4 metric: Số vật liệu / Số nhân công / Số máy / Thiếu giá;
- 3 card tổng quan theo nhóm;
- icon Vật liệu / Nhân công / Máy thi công tự vẽ bằng GDI+;
- danh sách thiếu giá có STT, mã hiệu, tên tài nguyên, đơn vị, nhóm và nút điều hướng;
- 4 chức năng chính;
- icon `fx` cho công thức nhân công và calculator cho giá ca máy;
- thông báo cuối pane về formula/link, không dùng số chết;
- giao diện vẫn là CustomTaskPane bên phải, không chiếm vùng Excel chính.

### Writer VL-NC-M

File chính:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2ResourceSheetWriter.cs`

Hành vi chính:

- giữ A:F là vùng biểu mẫu in;
- giữ lại giá vật liệu, giá nhiên liệu, đầu vào nhân công đã có khi refresh;
- tạo workbook Name bền vững cho giá resource và input;
- tính nhân công bằng công thức Excel;
- tính giá ca máy bằng công thức Excel: nhiên liệu/năng lượng, sửa chữa, khấu hao, chi phí khác, nhân công điều khiển máy;
- áp dụng hệ số nhiên liệu phụ;
- áp dụng hệ số môi trường ăn mòn cho ngữ cảnh nước/biển theo calculator hiện có;
- không xóa các cột unrelated ngoài vùng writer quản lý;
- tìm và tái sử dụng vị trí metadata cũ để tránh metadata trôi sang phải sau mỗi lần refresh;
- ẩn toàn bộ cột phụ ngoài A:F.

### Đối chiếu mẫu Excel thực tế

Đã đọc cấu trúc file mẫu:

- `/mnt/data/Du toan RPBM HoaLuNamDinh_Ver1.xlsx`
- sheet `VL-NC-M`

Writer đã được chỉnh theo đặc điểm in của mẫu:

- dòng 1–2 ẩn;
- block in chính bắt đầu từ mục `I. GIÁ NHÂN CÔNG`;
- header phần giá ca máy được lặp khi in;
- phần `III. GIÁ VẬT LIỆU` là block in riêng;
- vật liệu có ghi chú `ĐG thị trường`;
- hướng giấy dọc;
- in đen trắng;
- scale 98%;
- chỉ A:F thuộc vùng in;
- các cột phụ bị ẩn và nằm ngoài vùng in.

Mẫu hiện hành có PrintArea dạng hai vùng; writer V2 tạo lại cùng nguyên tắc theo số dòng sinh thực tế, không hard-code số dòng của file mẫu.

### Core / service / UI đã thêm hoặc cập nhật

- `ExcelAddIn1.Core/EstimateV2Resources.cs`
- `ExcelAddIn1.Core/EstimateV2ResourcePriceSheetProjector.cs`
- `ExcelAddIn1.Core/EstimateV2ResourceNames.cs`
- `ExcelAddIn1.Core/EstimateV2ExcelNames.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2ResourceService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2ResourceSheetWriter.cs`
- `ExcelAddIn1/Winform/EstimateResourcesPaneView.cs`
- `ExcelAddIn1/Winform/EstimateUiIcons.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneControl.cs`

### Commit quan trọng

- `3d4ae4f3cead` — unique resource aggregation plan
- `0c21071546f7` — logical resource price projector
- `21a2c64a8908` — writer dùng các ứng viên giá vật lý
- `dc63fff7bce6` — đưa writer VL-NC-M vào project
- `85dc8569f4f9` — sửa công thức hệ số nhiên liệu phụ
- `9d046c2d3fe5` — căn vùng in VL-NC-M theo file mẫu
- `bbd823b83463` — icon riêng theo UI chuẩn
- `f1449438a336` — chỉnh card VL-NC-M theo ảnh chuẩn
- `b44fbac63b2d` — đưa nhân công điều khiển máy vào preview/tổng hợp

### Test code đã có

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2ResourcePlan`
- `EstimateV2PriceSheetProjection`

Test projection kiểm tra:

- `M010.DIVING` -> các máy lặn vật lý;
- `M010.002-OR-M010.003` -> hai ứng viên;
- vật liệu `MAT-GASOLINE-OR-DIESEL` -> xăng + dầu;
- merge resource trùng;
- output price-sheet không còn resource logic.

### Kiểm tra đã thực hiện trong môi trường hiện tại

Đã tự rà soát tĩnh các file V2-201 sau thay đổi:

- cân bằng ngoặc `{}`, `()`, `[]`;
- không còn `TODO/FIXME/NotImplementedException` trong các file V2-201 chính;
- đã sửa lỗi kiểu dữ liệu `PriceProfile.TryFind`: API trả `PriceProfilePrice`, không phải `PriceProfileEntry`;
- đã sửa tham chiếu hidden parameter của hệ số nhiên liệu;
- đã sửa metadata không bị dịch sang phải qua nhiều lần refresh;
- đã sửa preview để tính cả nhân công điều khiển máy.

**Chưa chạy build VSTO/Excel thật và chưa chạy bộ test console trong môi trường hiện tại. Không ghi PASS giả.**

### Checklist runtime bắt buộc khi chạy trên máy có Excel/VSTO

- mở task pane -> vào VL-NC-M không lỗi;
- số VL / NC / Máy trên pane khớp resource writer;
- công tác có máy yêu cầu NC điều khiển -> NC đó xuất hiện trong pane và sheet;
- logical resource -> đủ ứng viên giá vật lý, không còn mã logic trong VL-NC-M;
- sinh mới VL-NC-M -> công thức NC/M hoạt động;
- nhập giá vật liệu -> refresh -> giá vẫn còn;
- nhập giá nhiên liệu / NC -> refresh -> giá ca máy cập nhật;
- refresh nhiều lần -> metadata không chạy sang phải;
- G:P hoặc các cột phụ legacy không xuất hiện trong vùng in;
- PrintArea có 2 block giống nguyên tắc file mẫu;
- dòng tiêu đề 1–2 ẩn; header máy lặp khi in;
- đóng/mở workbook -> workbook Name và giá nhập vẫn đọc lại được;
- không có `#REF!`, `#VALUE!`, `#NAME?` trong công thức sinh.

### Việc tiếp theo

V2-201 đã chốt. V2-301 đã được triển khai ở section bên dưới.

---

## Các task tiếp theo

### V2-301 — DG Cạn / DG Nước / DG Biển

**Trạng thái:** DONE - implementation / runtime pending

#### UI đã đối chiếu ảnh chuẩn

Đã đọc và bám trực tiếp:

- `/mnt/data/chuan_UI/05-DG-Can.png`
- `/mnt/data/chuan_UI/06-DG-Nuoc.png`

View mới:

- `ExcelAddIn1/Winform/EstimateUnitRatesPaneView.cs`

DG Cạn hiện có đúng cấu trúc đã chốt:

- 4 metric: Định mức cần sinh / Đã sinh / Thiếu giá / Liên kết công thức;
- bước 1 chọn định mức cần sinh;
- bước 2 tùy chọn Sinh mới / Cập nhật từ VL-NC-M / Chỉ sinh định mức đang dùng;
- bước 3 preview VL / NC / M / Tổng cộng;
- nút xanh `Sinh đơn giá`;
- footer giải thích mỗi định mức chỉ sinh một block dùng chung.

DG Nước hiện có:

- 4 metric: Định mức nước / Đã sinh / Thiếu giá / Cảnh báo;
- quy trình 4 bước;
- danh sách định mức công tác dưới nước;
- trạng thái từng định mức;
- footer formula/link.

DG Biển dùng cùng ngôn ngữ UI của DG Nước và **chỉ hiện điều hướng khi thực sự có định mức biển đang dùng**. Không tạo sheet DG Biển rác khi dự toán không có công tác biển.

#### Identity đơn giá

Đã thêm core model:

- `ExcelAddIn1.Core/EstimateV2Rates.cs`

Đơn giá unique theo:

```text
PackageIdentity + NormCode + VariantCode
        -> RateId ổn định
```

Hai hoặc nhiều WorkItem dùng cùng package + định mức + variant chỉ sinh **một block đơn giá** và tăng `UsageCount`; không nhân bản block theo dòng công tác.

Đã xử lý binding legacy có variant rỗng nhưng định mức chỉ có một variant: normalize về variant thật trước khi tạo RateId để không sinh trùng.

#### Phân loại Cạn / Nước / Biển

- `NORM-000.*`, `NORM-010.*`, `NORM-020.*` -> DG Cạn;
- `NORM-030.*` -> DG Nước;
- `NORM-040.*` -> DG Biển.

Sheet chỉ được sinh khi môi trường đó có rate đang dùng.

Tên sheet tương thích mẫu hiện tại:

- `DG Can` / `DG Cạn`;
- `DG Nuoc` / `DG Nước`;
- `DG Bien` / `DG Biển`.

#### Writer đơn giá

Đã thêm:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2RateService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2RateSheetWriter.cs`

Writer:

1. cập nhật `VL-NC-M` trước để bảo đảm workbook Name giá tồn tại;
2. sinh block đơn giá từ định mức/variant đang dùng;
3. hao phí định mức nằm ở cột D;
4. đơn giá resource ở cột E là **formula link** tới workbook Name của `VL-NC-M`;
5. thành tiền:
   - VL -> F;
   - NC -> G;
   - M -> H;
6. vật liệu phần trăm tính bằng công thức từ tổng vật liệu trực tiếp;
7. dòng `Cộng:` giữ riêng VL / NC / M;
8. thêm dòng `Tổng cộng đơn giá` và workbook Name `TOTAL`;
9. metadata nằm từ cột I trở đi và bị ẩn;
10. vùng in chỉ A:H.

Không copy snapshot giá chết vào block đơn giá.

#### Logical resource

Các resource logic như:

- `M010.002-OR-M010.003`;
- `M010.DIVING`;
- `MAT-GASOLINE-OR-DIESEL`;

được mở thành các price candidate vật lý dùng chung với VL-NC-M.

Khi chưa có UI chọn candidate riêng cho từng công tác, V2-301 dùng policy xác định được: **chọn candidate đầu tiên đang có giá > 0** bằng công thức Excel. Pane đồng thời đánh cảnh báo để người dùng biết rate có lựa chọn cần rà soát.

#### Điều kiện / hệ số định mức

State V2 hiện mới lưu NormCode + Variant, chưa lưu tập điều kiện như lưu tốc dòng chảy, đào có nước, độ dốc...

Vì vậy V2-301 hiện sinh **hao phí cơ sở của variant đã gắn**. Nếu `NormDefinition` có adjustment/constraint, rate được đánh cảnh báo `RequiresConditionReview`.

Không tự đoán điều kiện và không ghi hệ số giả vào workbook.

Đây là giới hạn đã biết của V2-301, cần được xử lý ở luồng thiết lập/validation sau; không được âm thầm tính sai.

#### Đối chiếu sheet mẫu thực tế

Đã đọc/rendere:

- `/mnt/data/Du toan RPBM HoaLuNamDinh_Ver1.xlsx`
- `DG Can`
- `DG Nuoc`

Ảnh render dùng để đối chiếu:

- `/mnt/data/dg_can_sample.png`
- `/mnt/data/dg_nuoc_sample.png`

Writer đã chỉnh theo cấu trúc in thực tế:

- Times New Roman;
- A:H;
- title block;
- tên công tác / số hiệu định mức / đơn vị;
- header **hai tầng**:
  - A:E merge dọc;
  - F:H merge ngang `Thành tiền (đồng)`;
  - hàng dưới F/G/H = Vật liệu / Nhân công / Máy;
- section I Vật liệu / II Nhân công / III Máy thi công;
- dòng `Cộng:`;
- Portrait;
- BlackAndWhite;
- FitToPagesWide = 1;
- helper/metadata từ I trở đi ẩn.

#### Navigation

`EstimateTaskPaneControl` đã nối:

```text
Tổng quan
  -> DG Cạn
  -> DG Nước
       -> DG Biển (chỉ khi có định mức biển)
```

Output tile DG Biển cũng chỉ xuất hiện nếu workbook thực sự có sheet biển.

#### Test code đã thêm

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2RatePlan`

Test kiểm tra:

- 2 WorkItem cùng Land norm + variant -> 1 rate, UsageCount = 2;
- tách đúng Land / InlandWater / Sea;
- RateId ổn định;
- `M010.DIVING` có candidate máy lặn cụ thể;
- rate nước có adjustment/constraint được đánh cần rà soát;
- rate biển giữ đúng logical machine để writer liên kết candidate.

#### Commit quan trọng

- `8b40d688ca6d` — expose resource candidate expansion;
- `7e311d8ed845` — core rate plan / RateId;
- `f9e90e8314e3` — preview service DG;
- `58b6120af5e2` — formula-linked DG writer;
- `66960b684b07` — UI DG Cạn/Nước/Biển;
- `8edc18670b88` — navigation vào task pane;
- `71e4764eba62` — core rate tests;
- `c788707acfa4` — merge normalized rate identities;
- `3ba4c36396f6` — stable TOTAL link;
- `0a9dd85097b1` + `c0a564583d1a` — header in hai tầng theo mẫu;
- `8bdf5dbb23d0` — cập nhật UI contract.

#### Tự kiểm tra trong môi trường hiện tại

Đã rà soát tĩnh các file V2-301 chính:

- ngoặc `{}`, `()`, `[]` cân bằng;
- không có `TODO/FIXME/NotImplementedException`;
- kiểm tra lại API MachineRateCatalog dùng bởi writer;
- kiểm tra selector/radio DG Cạn thực sự ảnh hưởng tập rate sinh;
- commit checkbox grid trước khi đọc lựa chọn;
- DG Biển không hiện khi không có rate biển;
- công thức DG dùng workbook Name từ VL-NC-M, không dùng số giá snapshot;
- header sheet đã sửa từ một tầng sang hai tầng theo file mẫu.

**Chưa chạy build VSTO/Excel thật và chưa chạy test console trong môi trường hiện tại. Không ghi PASS giả.**

#### Checklist runtime bắt buộc

- vào DG Cạn từ Tổng quan -> pane mở đúng, không form modal;
- số metric khớp rate plan;
- hai công tác cùng norm+variant -> chỉ một block DG;
- chọn/bỏ chọn rate -> Sinh mới chỉ sinh đúng tập mong muốn và giữ block đã có;
- Cập nhật từ VL-NC-M -> giá đổi thì DG đổi theo formula/link;
- DG Nước chỉ sinh khi có NORM-030;
- DG Biển chỉ sinh khi có NORM-040;
- workbook chỉ có biển -> từ DG Nước vẫn thấy điều hướng DG Biển;
- header in hai tầng đúng A:H;
- cột I trở đi hidden;
- không có `#NAME?`, `#REF!`, `#VALUE!`;
- TOTAL = VL + NC + M;
- save/close/open -> workbook Name rate vẫn còn;
- rate có adjustment/constraint phải hiện cảnh báo, không âm thầm coi như đã áp hệ số.

#### Việc tiếp theo

V2-301 đã chốt. V2-401 đã được triển khai ở section bên dưới.

### V2-401 — Gia DT TC + THKP-TC

**Trạng thái:** DONE - implementation / runtime pending

#### UI đã đối chiếu ảnh chuẩn

Đã đọc và bám trực tiếp:

- `/mnt/data/chuan_UI/07-THKP-TC-Kiem-tra.png`

View mới:

- `ExcelAddIn1/Winform/EstimateCostSummaryPaneView.cs`

Màn hình hiện có:

- 4 metric: Tổng công tác / Đủ đơn giá / Cảnh báo / Lỗi công thức;
- card xanh lớn `Cập nhật THKP-TC`;
- khu `Kết quả kiểm tra hồ sơ` với icon trạng thái và `Xem chi tiết`;
- 3 card chức năng: Kiểm tra hồ sơ / Xuất báo cáo / Mở thư mục hồ sơ;
- footer xanh dương đúng ngôn ngữ thiết kế của ảnh chuẩn;
- vẫn là CustomTaskPane bên phải, không mở form modal.

`Kiểm tra hồ sơ` trong V2-401 chỉ kiểm tra các dependency/link cơ bản cần cho bước tổng hợp. Validation đầy đủ vẫn để đúng task V2-501.

#### Đối chiếu file dự toán thực tế

Đã đọc cấu trúc các mẫu:

- `/mnt/data/Du toan RPBM HoaLuNamDinh_Ver1.xlsx`
- `/mnt/data/Du toan RPBM Pleiku_TP2_QuyNhon_Ver7.2.xlsx`

Các mẫu cho thấy:

- `Gia DT TC` hoặc các sheet `Gia DT TC_...` chứa công tác chi tiết;
- đơn giá VL / NC / M được dùng để tính thành tiền theo từng công tác;
- `THKP-TC` có các dòng chi phí trực tiếp `VL`, `NC`, `M`, `T`;
- các dòng sau đó (chi phí chung, TL, K1..Kn, VAT, tổng cuối...) phụ thuộc vào phần chi phí trực tiếp và khác nhau theo mẫu/pháp lý.

Vì vậy V2-401 **chỉ thay thế nguồn liên kết phần chi phí trực tiếp**, không phá công thức pháp lý phía sau của THKP-TC hiện hữu.

#### WorkItem -> RateId

Đã thêm core model:

- `ExcelAddIn1.Core/EstimateV2CostLinks.cs`

Mỗi WorkItem được resolve theo state bền vững:

```text
WorkItemId
  -> PackageIdentity + NormCode + VariantCode
  -> RateId
  -> workbook Name VL / NC / M
```

Không dùng RowIndex làm identity.

Hai WorkItem dùng chung một định mức + variant tiếp tục dùng chung một RateId nhưng mỗi dòng có khối lượng và thành tiền riêng.

#### Gia DT TC / bảng công tác chi tiết

Writer:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkWriter.cs`

Trên mỗi bảng công tác đã đăng ký, sáu cột kết quả được đặt ngay sau cột Khối lượng:

```text
Đơn giá VL
Đơn giá NC
Đơn giá M
Thành tiền VL
Thành tiền NC
Thành tiền M
```

Với layout chuẩn A:F, chúng tương ứng G:L.

Công thức:

```text
Đơn giá VL = RateId.VL
Đơn giá NC = RateId.NC
Đơn giá M  = RateId.M

Thành tiền VL = Khối lượng × Đơn giá VL
Thành tiền NC = Khối lượng × Đơn giá NC
Thành tiền M  = Khối lượng × Đơn giá M
```

Các đơn giá là **formula tham chiếu workbook Name**, không ghi số chết.

Nếu WorkItem chưa gắn định mức hoặc chưa có DG tương ứng, writer không tự bịa giá và để dòng đó chưa liên kết để pane tiếp tục cảnh báo.

#### Sửa xung đột metadata quan trọng

V2-101 trước đây có thể đặt:

```text
__TTB_ID
__TTB_NORM
__TTB_KIND
__TTB_HASH
```

ngay sau cột visible cuối cùng. Với bảng có Khối lượng ở F, metadata cũ có thể rơi vào G:J — đúng vùng V2-401 cần dùng cho đơn giá/thành tiền.

Đã bổ sung migration:

- `WorkbookEstimateV2RegistrationService.EnsureTechnicalColumnsAfter(...)`

Khi V2-401 chạy:

1. metadata cũ được copy nguyên trạng tới vùng an toàn;
2. tối thiểu nằm sau sáu cột kết quả (M:P với layout A:F), hoặc xa hơn nếu workbook đã dùng các cột đó;
3. vùng metadata mới được hide;
4. custom property của source được cập nhật;
5. cột cũ được giải phóng để G:L là dữ liệu visible;
6. WorkItemId/binding vẫn giữ nguyên.

Đây là migration cấu trúc, không yêu cầu người dùng quét/gắn lại từ đầu.

#### Tổng hợp nhiều bảng / nhiều khu vực

V2-401 không phụ thuộc dự án có 1, 2 hay 3 khu vực.

Mỗi registered source được tổng hợp bằng WorkItemId hợp lệ; helper THKP dùng công thức `SUMIF` trên cột `__TTB_ID` và cột thành tiền tương ứng. Vì vậy:

- dòng nhóm/tổng phụ không có WorkItemId không bị cộng hai lần;
- nhiều bảng/sheet được cộng chung;
- không hard-code tên khu vực VT/DN hay số lượng khu vực.

#### THKP-TC

Đã thêm stable workbook Names:

- `TTBMVN_V2_GIADT_VL`
- `TTBMVN_V2_GIADT_NC`
- `TTBMVN_V2_GIADT_M`
- `TTBMVN_V2_GIADT_TOTAL`

Helper tổng hợp nằm ở các cột ẩn ngoài vùng in của THKP-TC.

Nếu `THKP-TC` đã tồn tại, writer:

1. tự tìm cột `Ký hiệu` và `Thành tiền`;
2. tự tìm các dòng có ký hiệu `VL`, `NC`, `M`, `T`;
3. thay **chỉ** công thức thành tiền của bốn dòng này bằng workbook Name V2;
4. giữ nguyên toàn bộ công thức phía sau của mẫu hiện hữu.

Không hard-code các tỷ lệ minh họa 6,5%, 5,5%, 10% trong ảnh UI.

Nếu workbook chưa có `THKP-TC`, V2-401 tạo một mẫu tối thiểu an toàn chỉ cho phần chi phí trực tiếp và ghi rõ các khoản pháp lý khác chưa được tự suy đoán.

#### Preview / trạng thái

Đã thêm:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkService.cs`

Preview hiện tính:

- tổng WorkItem;
- số WorkItem đã có RateId và workbook Name đơn giá;
- số WorkItem cảnh báo (không double-count cùng một WorkItem);
- lỗi Excel cơ bản trong các ô output;
- THKP có thật sự tham chiếu bốn workbook Name VL/NC/M/T hay chưa.

Chỉ việc workbook Name tồn tại chưa được coi là THKP đã cập nhật; service còn kiểm tra công thức trong THKP có reference các Name đó.

#### Navigation

`EstimateTaskPaneControl` đã nối bước cuối:

```text
Tổng quan
  -> THKP-TC & Kiểm tra
```

Bước 6 chỉ hiển thị `Đã xong` khi THKP-TC thực sự có các liên kết V2, không chỉ vì sheet `THKP-TC` tồn tại.

#### Test code đã thêm

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2CostLinkPlan`

Test core kiểm tra:

- 2 WorkItem cùng rate -> cùng RateId;
- WorkItem chưa gắn -> `Unbound`;
- binding không resolve rate -> `RateNotResolved`;
- đếm Ready / Unbound / MissingRate;
- propagate cờ `RequiresConditionReview`.

#### File chính đã thêm/cập nhật

- `ExcelAddIn1.Core/EstimateV2CostLinks.cs`
- `ExcelAddIn1.Core/EstimateV2ExcelNames.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2RegistrationService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkWriter.cs`
- `ExcelAddIn1/Winform/EstimateCostSummaryPaneView.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneControl.cs`

#### Commit quan trọng

- `e01a8ae866e4` — WorkItem -> RateId link plan;
- `c9d099a88dec` — stable Gia DT TC total names;
- `10e2b4564a72` — di chuyển metadata ra sau vùng đơn giá/thành tiền;
- `308abc67b3ed` — preview service;
- `56ad90b950c7` — writer link WorkItem + THKP;
- `45705188e220` — UI THKP-TC & Kiểm tra;
- `2384fb09bab8` — navigation;
- `99f76f0f34db` — warning metric không double-count;
- `77fe3a8d949c` — core link-plan test;
- `f4306c83aeee` — xác nhận THKP thực sự reference Name V2;
- `7f3ed7ae27fe` — trạng thái bước THKP ở Tổng quan;
- `a6565ed3c03f` — cleanup COM range khi tạo THKP tối thiểu;
- `80d05d920b5c` — cập nhật UI contract.

#### Tự kiểm tra trong môi trường hiện tại

Đã rà soát tĩnh các file V2-401:

- ngoặc `{}`, `()`, `[]` cân bằng;
- không có conflict marker;
- không có `TODO/FIXME/NotImplementedException` trong các file V2-401 chính;
- project file chứa đúng các source mới;
- metadata migration không dùng RowIndex làm identity;
- các giá/ thành tiền là formula/link, không phải số snapshot;
- THKP chỉ cập nhật VL/NC/M/T và không ghi đè các công thức pháp lý phía sau;
- helper THKP nằm ngoài vùng in và bị ẩn.

Đã đối chiếu logic với hai workbook mẫu bằng công cụ spreadsheet, nhưng **chưa chạy build VSTO/Excel thật và chưa chạy test console trong môi trường hiện tại. Không ghi PASS giả.**

#### Checklist runtime bắt buộc

- workbook cũ có metadata ở G:J -> cập nhật -> metadata chuyển sang M:P hoặc vùng an toàn, ID/binding không mất;
- G:L hiện đúng 6 cột đơn giá/thành tiền;
- đơn giá G:I là workbook Name của RateId;
- J:L = Khối lượng × đơn giá;
- section/group/subtotal không có WorkItemId không bị tính trùng;
- dự án 1/2/3 khu vực -> tổng VL/NC/M đúng tổng các WorkItem;
- sửa khối lượng -> thành tiền và THKP cập nhật theo công thức;
- sửa giá VL-NC-M -> DG -> Gia DT TC -> THKP cập nhật theo chuỗi link;
- WorkItem chưa có DG -> không có giá giả, pane báo cảnh báo;
- THKP hiện hữu -> chỉ VL/NC/M/T bị thay source link; các dòng pháp lý phía sau giữ nguyên;
- THKP không tồn tại -> tạo mẫu tối thiểu trực tiếp, không tự gán tỷ lệ pháp lý;
- đóng/mở workbook -> workbook Names/link còn nguyên;
- pane 07 mở đúng, 4 metric và card/action không tràn ở DPI 100/125/150%;
- không có `#REF!`, `#NAME?`, `#VALUE!` sau cập nhật.

#### Việc tiếp theo

V2-401 đã chốt. V2-501 đã được triển khai ở section bên dưới.

, whitespace hoặc dấu phân cách;
- map mã lỗi Excel sang:
  - `#REF!`;
  - `#VALUE!`;
  - `#N/A`;
  - `#NAME?`;
  - `#DIV/0!`;
  - `#NUM!`;
  - các lỗi còn lại.

#### Validation service

Đã thêm:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2ValidationService.cs`

Report validation giữ:

- reconcile result;
- cost link plan;
- số công tác đủ đơn giá;
- số WorkItem cảnh báo theo ID duy nhất;
- số lỗi công thức;
- số công thức bị ghi đè;
- số công tác bị thiếu giá;
- trạng thái THKP link;
- package errors;
- danh sách finding có:
  - code;
  - severity;
  - title/detail;
  - WorkItemId liên quan;
  - worksheet;
  - address;
  - cờ recoverable.

#### Missing norm / missing rate / missing price

Validation phát hiện riêng:

- `MISSING_NORM`: WorkItem chưa có binding định mức;
- `RATE_NOT_RESOLVED`: binding có nhưng không resolve được RateId;
- `MISSING_RATE`: RateId hợp lệ nhưng chưa có block DG/workbook Name;
- `MISSING_PRICE`: rate có resource chưa có giá đầu vào > 0;
- `PACKAGE`: package/version/checksum hoặc module không resolve được;
- `CONDITION_REVIEW`: định mức có adjustment/constraint chưa được lưu điều kiện đầy đủ.

Thiếu giá được map ngược về các WorkItem bị ảnh hưởng để metric Cảnh báo không chỉ đếm số resource.

**Không tự điền giá thị trường bị thiếu.**

#### Formula overwritten

V2-501 kiểm tra công thức ở nhiều tầng.

1. **Gia DT TC / bảng công tác**

Với mỗi WorkItem đã có RateId:

```text
Đơn giá VL = RateId.VL
Đơn giá NC = RateId.NC
Đơn giá M  = RateId.M

Thành tiền VL = Khối lượng × Đơn giá VL
Thành tiền NC = Khối lượng × Đơn giá NC
Thành tiền M  = Khối lượng × Đơn giá M
```

G:L được so sánh với công thức chuẩn. Ghi giá chết hoặc thay bằng công thức khác đều bị phát hiện.

2. **DG Cạn / DG Nước / DG Biển**

Validation kiểm tra workbook Name:

- VL;
- NC;
- M;
- TOTAL.

Không chỉ kiểm tra Name tồn tại. Cell được Name trỏ tới phải còn công thức; với sheet V2 có metadata, validation dựng lại công thức expected cho TOTAL và các component để phát hiện công thức bị sửa chứ không chỉ phát hiện constant.

3. **VL-NC-M**

Vật liệu là input nên được phép là số người dùng nhập.

Giá **nhân công** và **máy** là output formula. Workbook Name của NC/M:

- bị xóa -> `RESOURCE_NAME_MISSING`;
- bị ghi đè bằng số -> `RESOURCE_FORMULA_OVERWRITTEN`;
- lỗi Excel -> finding formula error.

4. **THKP-TC**

Validation nhận diện layout qua cột `Ký hiệu` / `Thành tiền` và các dòng:

- VL;
- NC;
- M;
- T.

Bốn cell này phải link đúng bốn workbook Name V2. Nếu người dùng thay formula khác hoặc value chết -> `THKP_FORMULA_OVERWRITTEN`.

`ThkpLinked=true` chỉ khi liên kết thật còn đúng, không chỉ vì Name tồn tại đâu đó trong workbook.

#### Lỗi Excel

Validation scan các vùng V2 quản lý và các sheet:

- `VL-NC-M`;
- `DG Can / DG Cạn`;
- `DG Nuoc / DG Nước`;
- `DG Bien / DG Biển`;
- `THKP-TC`;
- G:L của các registered source.

Phát hiện `ErrorWrapper`/Excel CVErr và gom theo loại.

Yêu cầu bắt buộc đã có:

- `#REF!`;
- `#VALUE!`;
- `#N/A`.

Các lỗi khác như `#NAME?`, `#DIV/0!`, `#NUM!` cũng được báo.

Metric `Lỗi công thức` không còn bằng 0 khi lỗi nằm ở THKP hoặc sheet trung gian nhưng không map trực tiếp được về một WorkItem.

#### Orphan / duplicate / mismatch

Mỗi lần scan gọi reconcile an toàn trước.

Validation báo:

- `ORPHAN`: WorkItem đã biến mất khỏi bảng nhưng còn lịch sử trong Custom XML;
- `DUPLICATE_ID_RECOVERED`: duplicate ID do copy/paste đã được tách;
- `IDENTITY_RECOVERED`: ID lệch sau sort/chèn/xóa được phục hồi theo fingerprint duy nhất;
- `NORM_DISPLAY_RECOVERED`: ô hiển thị định mức bị xóa/sửa đã phục hồi từ binding;
- `DUPLICATE_ID`: duplicate vẫn còn sau reconcile;
- `IDENTITY_MISMATCH`: ID invalid, state không có ID đó hoặc ID thuộc source khác.

Không dùng RowIndex để quyết định identity.

#### Layout cũ

Validation phát hiện:

- `COST_LAYOUT_PENDING`

khi metadata V2-101 vẫn nằm trong vùng cần dành cho G:L.

Finding này được đánh recoverable; writer V2-401 sẽ di metadata tới M:P hoặc vùng an toàn xa hơn rồi tạo 6 cột đơn giá/thành tiền.

#### Safe recovery

Đã có:

- `WorkbookEstimateV2ValidationService.RepairRecoverable(...)`

Nguyên tắc:

- reconcile WorkItemId / duplicate / norm display là sửa chữa an toàn và được chạy trong scan;
- formula/link chỉ được phục hồi thông qua **writer sở hữu**:
  - VL-NC-M;
  - DG Cạn/Nước/Biển;
  - Gia DT TC;
  - THKP-TC;
- missing DG có thể được tái sinh;
- layout cũ có thể được migrate bởi writer;
- không ghi giá thị trường giả;
- không suy đoán condition/adjustment của định mức;
- không thay công thức pháp lý ngoài vùng V2 quản lý.

Nút `Kiểm tra hồ sơ` không âm thầm điền giá đầu vào. Main action `Cập nhật THKP-TC` vẫn là thao tác ghi rõ ràng cho liên kết Gia DT TC/THKP.

#### Navigation lỗi

`EstimateV2CostIssue` đã bổ sung:

- `WorksheetName`;
- `Address`;
- `Recoverable`.

Nếu finding có target, `Xem chi tiết`:

1. activate đúng worksheet;
2. select đúng address;
3. hiển thị hướng phục hồi trong footer.

Ví dụ thiếu giá mở `VL-NC-M`; công thức G:L hoặc THKP bị sửa mở thẳng cell đầu tiên bị lỗi.

#### Tổng quan / startup

Không chạy full validation khi mở Tổng quan để tránh quay lại lỗi kiến trúc cũ: mở module phải quét dữ liệu/package rồi mới cho dùng.

`EstimateTaskPaneControl` chỉ dùng kiểm tra nhẹ `HasDirectCostLinks` cho bước 6. Full validation chỉ chạy khi:

- mở màn hình `THKP-TC & Kiểm tra`;
- bấm `Kiểm tra hồ sơ`;
- hoặc gọi service validation/repair rõ ràng.

Như vậy missing package hoặc hồ sơ chưa hoàn chỉnh vẫn không khóa việc mở module.

#### Test code đã thêm

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2ValidationRules`

Test core kiểm tra:

- công thức G:L sinh đúng;
- formula normalization bỏ `# TIẾN ĐỘ TRIỂN KHAI DỰ TOÁN RPBM V2

> Tài liệu này là nguồn trạng thái chính cho quá trình tiếp quản và phát triển module Dự toán V2.
> Mỗi task phải ghi rõ: mục tiêu, thay đổi đã làm, file liên quan, trạng thái kiểm tra và việc tiếp theo.
> Người tiếp quản sau chỉ cần đọc file này cùng README, ARCHITECTURE, MIGRATION-PLAN và UI-CONTRACT là biết dự án đang ở đâu.

## Trạng thái tổng quan

| Mã | Task | Trạng thái |
|---|---|---|
| V2-001 | Bỏ gate khởi động Dự toán | DONE - implementation / runtime pending |
| V2-002 | Task pane UI theo bộ ảnh đã chốt | DONE - shell + 3 màn hình đầu / runtime pending |
| V2-101 | WorkItemId + binding định mức bền vững | DONE - implementation / runtime pending |
| V2-201 | Tổng hợp và sinh VL-NC-M | DONE - implementation / runtime pending |
| V2-301 | Sinh DG Cạn / DG Nước / DG Biển | DONE - implementation / runtime pending |
| V2-401 | Link Gia DT TC + THKP-TC | DONE - implementation / runtime pending |
| V2-501 | Validation / phục hồi / phát hiện lỗi | DONE - implementation / runtime pending |
| V2-601 | Tương thích file cũ và migration | NEXT |

---

## UI CONTRACT — BẮT BUỘC

UI V2 phải bám theo bộ 10 ảnh người dùng đã chốt và được mô tả trong:

- `docs/du-toan-v2/UI-CONTRACT.md`

Nguyên tắc không được tự ý thay đổi:

- UI chính là **CustomTaskPane dock bên phải Excel**, khoảng 430 px.
- Không quay lại form Dự toán lớn/menu dọc của bản legacy.
- Excel luôn là vùng làm việc chính bên trái.
- Màu xanh/vàng/xanh dương/đỏ, card trắng và mật độ UI phải đồng nhất với ảnh chốt.
- Icon thiếu phải tự vẽ bằng renderer nội bộ; không dùng bộ icon lệch phong cách.
- Màn hình nào chưa có mockup chi tiết thì phải dựng tiếp theo đúng ngôn ngữ thiết kế hiện tại trước khi làm khác.

---

## V2-001 — Bỏ gate khởi động Dự toán

**Trạng thái:** DONE - implementation / runtime pending  
**Commit gốc:** `63c54d886e2d3b463044a98a6328d4c04c8fe0c9`  
**Luồng UI hiện tại:** CustomTaskPane, không còn mở form `Dutoan` làm UI chính.

### Mục tiêu

Sửa hành vi cũ:

```text
Bấm Dự toán
  -> FrmProjectSetup modal
  -> map đủ role
  -> resolve package
  -> hợp lệ mới mở Dự toán
```

thành:

```text
Bấm Dự toán
  -> mở Trợ lý Dự toán ngay
  -> từng chức năng tự kiểm tra dependency khi cần
```

### Đã làm

- Bỏ `FrmProjectSetup` khỏi startup gate.
- `Ribbon1.btnDutoan_Click` hiện gọi `EstimateTaskPaneManager.Show(currentWorkbook)`.
- Không bắt map `THKP-TC` hoặc đủ 7 worksheet role trước khi vào module.
- Missing package không được phép chặn task pane.
- Task pane được đóng theo workbook trong `ThisAddIn.WorkbookBeforeClose`.

### File chính

- `ExcelAddIn1/Ribbon1.cs`
- `ExcelAddIn1/ThisAddIn.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneManager.cs`

### Runtime test còn thiếu

Chưa chạy build VSTO/Excel thật trong môi trường hiện tại. Khi chạy local phải kiểm tra:

- workbook chưa có role -> pane vẫn mở;
- workbook chưa pin package -> pane vẫn mở;
- workbook đóng -> pane đóng;
- không xuất hiện `FrmProjectSetup` trước pane.

---

## V2-002 — Task pane UI theo bộ ảnh chốt

**Trạng thái:** DONE - shell + 3 màn hình đầu / runtime pending

### Đã làm

Đã chuyển UI chính sang VSTO CustomTaskPane và dựng theo ảnh chốt:

1. **Tổng quan** — `EstimateTaskPaneControl`
2. **Công tác** — `EstimateWorkItemsPaneView`
3. **Gắn định mức** — `EstimateNormBindingPaneView`

Đã có renderer icon riêng:

- `EstimateUiIcons.cs`

Các icon hiện tại được vẽ bằng GDI+ để giữ đúng phong cách và tránh phụ thuộc file icon bên ngoài.

### Navigation đã nối

```text
Tổng quan
  -> Công tác
       -> Gắn định mức
  -> Gắn định mức
```

Các màn hình VL-NC-M, DG Cạn, DG Nước, THKP/Kiểm tra, Thiết lập, Gói pháp lý, Báo cáo sẽ được bổ sung ở đúng task nghiệp vụ tương ứng nhưng phải giữ nguyên UI contract.

### File chính

- `ExcelAddIn1/Winform/EstimateUiIcons.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneControl.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneManager.cs`
- `ExcelAddIn1/Winform/EstimateWorkItemsPaneView.cs`
- `ExcelAddIn1/Winform/EstimateNormBindingPaneView.cs`
- `docs/du-toan-v2/UI-CONTRACT.md`

---

## V2-101 — WorkItemId + binding định mức bền vững

**Trạng thái:** DONE - implementation / runtime pending

### Mục tiêu đã đạt ở code

Không dùng RowIndex làm identity.

Mỗi công tác có state:

```text
WorkItemId
SourceKey
NormCode
VariantCode
PackageId
DataVersion
Kind
Fingerprint
IsOrphaned
```

Nguồn sự thật là **Custom XML Part** trong workbook.

Các cột kỹ thuật ẩn làm anchor:

```text
__TTB_ID
__TTB_NORM
__TTB_KIND
__TTB_HASH
```

Cột kỹ thuật được đặt ngoài vùng dữ liệu đang dùng, sau đó hide. Không thêm sheet metadata visible.

### Hành vi đã cài đặt

- Đăng ký vùng công tác từ selection.
- Tự nhận diện các cột:
  - Mã công tác
  - Định mức
  - Mô tả công việc
  - Đơn vị
  - Khối lượng
- Sinh GUID cho dòng công tác mới.
- Lưu binding định mức trong Custom XML.
- Lưu SourceKey để nhiều bảng/sheet không ghi đè trạng thái của nhau.
- Mở task pane -> tự reconcile state đã lưu.
- Xóa ô “Định mức” hiển thị -> binding thật vẫn còn; reconcile có thể phục hồi lại.
- Chèn/xóa dòng -> không phụ thuộc RowIndex.
- Copy/paste gây trùng WorkItemId -> cấp ID mới cho dòng copy, giữ binding.
- Sort làm ID ẩn lệch với dữ liệu hiển thị -> fingerprint được dùng để phục hồi khi có match duy nhất.
- Dòng biến mất -> đánh orphaned theo đúng source, không làm hỏng source khác.
- “Bỏ gắn” trong UI mới thực sự xóa binding.

### Core / service đã thêm

- `ExcelAddIn1.Core/EstimateV2State.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2StateService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2RegistrationService.cs`

### UI đã nối

- `EstimateWorkItemsPaneView`
- `EstimateNormBindingPaneView`

### Test code đã thêm

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2StateRoundTrip`
- `EstimateV2Fingerprint`

**Chưa chạy test thực tế trong môi trường hiện tại**, vì chưa có build/runtime VSTO + Excel. Không ghi PASS giả.

### Checklist runtime bắt buộc

- register vùng công tác -> save -> đóng -> mở lại -> binding còn;
- delete ô định mức -> mở pane/reconcile -> định mức phục hồi;
- insert row -> dòng cũ giữ binding, dòng mới có ID mới;
- delete row -> dòng khác không lệch binding;
- sort toàn vùng -> binding đi theo công tác;
- sort chỉ cột visible -> fingerprint recovery không gắn nhầm;
- copy/paste row -> duplicate ID được tách;
- nhiều bảng/sheet -> orphan state không ghi đè lẫn nhau;
- thiếu package -> Công tác vẫn dùng được, Gắn định mức chỉ cảnh báo dependency.

---

## V2-201 — Tổng hợp và sinh VL-NC-M

**Trạng thái:** DONE - implementation / runtime pending

### Mục tiêu đã đạt ở code

Từ các WorkItem đã gắn định mức, V2 hiện:

1. Load đúng package/version/checksum đã pin trong từng binding.
2. Chỉ gom **resource thực sự đang dùng**, không dump toàn bộ catalog.
3. Tách resource theo vật liệu / nhân công / máy thi công.
4. Với resource logic (`-OR-`, `.DIVING`), VL-NC-M sinh các **ứng viên giá vật lý** cần thiết; lựa chọn chính xác dùng ở công tác nào để dành cho bước đơn giá.
5. Bổ sung cả **nhân công điều khiển máy** từ `MachineRateCatalog`, kể cả khi loại nhân công đó không xuất hiện trực tiếp trong hao phí định mức.
6. Sinh/cập nhật `VL-NC-M` bằng writer riêng, không ghi kết quả tính toán chết.
7. Giá người dùng đã nhập được ưu tiên giữ lại khi refresh.
8. Nhân công và giá ca máy dùng **công thức/liên kết Excel**; các đầu vào được neo bằng workbook Name ổn định.
9. Metadata đặt ngoài vùng in và bị ẩn; không tạo sheet kỹ thuật visible.

### UI đã hoàn thiện theo ảnh chuẩn 04

Màn hình `EstimateResourcesPaneView` đã đối chiếu trực tiếp:

- `/mnt/data/chuan_UI/04-VL-NC-M.png`

Đã có:

- 4 metric: Số vật liệu / Số nhân công / Số máy / Thiếu giá;
- 3 card tổng quan theo nhóm;
- icon Vật liệu / Nhân công / Máy thi công tự vẽ bằng GDI+;
- danh sách thiếu giá có STT, mã hiệu, tên tài nguyên, đơn vị, nhóm và nút điều hướng;
- 4 chức năng chính;
- icon `fx` cho công thức nhân công và calculator cho giá ca máy;
- thông báo cuối pane về formula/link, không dùng số chết;
- giao diện vẫn là CustomTaskPane bên phải, không chiếm vùng Excel chính.

### Writer VL-NC-M

File chính:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2ResourceSheetWriter.cs`

Hành vi chính:

- giữ A:F là vùng biểu mẫu in;
- giữ lại giá vật liệu, giá nhiên liệu, đầu vào nhân công đã có khi refresh;
- tạo workbook Name bền vững cho giá resource và input;
- tính nhân công bằng công thức Excel;
- tính giá ca máy bằng công thức Excel: nhiên liệu/năng lượng, sửa chữa, khấu hao, chi phí khác, nhân công điều khiển máy;
- áp dụng hệ số nhiên liệu phụ;
- áp dụng hệ số môi trường ăn mòn cho ngữ cảnh nước/biển theo calculator hiện có;
- không xóa các cột unrelated ngoài vùng writer quản lý;
- tìm và tái sử dụng vị trí metadata cũ để tránh metadata trôi sang phải sau mỗi lần refresh;
- ẩn toàn bộ cột phụ ngoài A:F.

### Đối chiếu mẫu Excel thực tế

Đã đọc cấu trúc file mẫu:

- `/mnt/data/Du toan RPBM HoaLuNamDinh_Ver1.xlsx`
- sheet `VL-NC-M`

Writer đã được chỉnh theo đặc điểm in của mẫu:

- dòng 1–2 ẩn;
- block in chính bắt đầu từ mục `I. GIÁ NHÂN CÔNG`;
- header phần giá ca máy được lặp khi in;
- phần `III. GIÁ VẬT LIỆU` là block in riêng;
- vật liệu có ghi chú `ĐG thị trường`;
- hướng giấy dọc;
- in đen trắng;
- scale 98%;
- chỉ A:F thuộc vùng in;
- các cột phụ bị ẩn và nằm ngoài vùng in.

Mẫu hiện hành có PrintArea dạng hai vùng; writer V2 tạo lại cùng nguyên tắc theo số dòng sinh thực tế, không hard-code số dòng của file mẫu.

### Core / service / UI đã thêm hoặc cập nhật

- `ExcelAddIn1.Core/EstimateV2Resources.cs`
- `ExcelAddIn1.Core/EstimateV2ResourcePriceSheetProjector.cs`
- `ExcelAddIn1.Core/EstimateV2ResourceNames.cs`
- `ExcelAddIn1.Core/EstimateV2ExcelNames.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2ResourceService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2ResourceSheetWriter.cs`
- `ExcelAddIn1/Winform/EstimateResourcesPaneView.cs`
- `ExcelAddIn1/Winform/EstimateUiIcons.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneControl.cs`

### Commit quan trọng

- `3d4ae4f3cead` — unique resource aggregation plan
- `0c21071546f7` — logical resource price projector
- `21a2c64a8908` — writer dùng các ứng viên giá vật lý
- `dc63fff7bce6` — đưa writer VL-NC-M vào project
- `85dc8569f4f9` — sửa công thức hệ số nhiên liệu phụ
- `9d046c2d3fe5` — căn vùng in VL-NC-M theo file mẫu
- `bbd823b83463` — icon riêng theo UI chuẩn
- `f1449438a336` — chỉnh card VL-NC-M theo ảnh chuẩn
- `b44fbac63b2d` — đưa nhân công điều khiển máy vào preview/tổng hợp

### Test code đã có

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2ResourcePlan`
- `EstimateV2PriceSheetProjection`

Test projection kiểm tra:

- `M010.DIVING` -> các máy lặn vật lý;
- `M010.002-OR-M010.003` -> hai ứng viên;
- vật liệu `MAT-GASOLINE-OR-DIESEL` -> xăng + dầu;
- merge resource trùng;
- output price-sheet không còn resource logic.

### Kiểm tra đã thực hiện trong môi trường hiện tại

Đã tự rà soát tĩnh các file V2-201 sau thay đổi:

- cân bằng ngoặc `{}`, `()`, `[]`;
- không còn `TODO/FIXME/NotImplementedException` trong các file V2-201 chính;
- đã sửa lỗi kiểu dữ liệu `PriceProfile.TryFind`: API trả `PriceProfilePrice`, không phải `PriceProfileEntry`;
- đã sửa tham chiếu hidden parameter của hệ số nhiên liệu;
- đã sửa metadata không bị dịch sang phải qua nhiều lần refresh;
- đã sửa preview để tính cả nhân công điều khiển máy.

**Chưa chạy build VSTO/Excel thật và chưa chạy bộ test console trong môi trường hiện tại. Không ghi PASS giả.**

### Checklist runtime bắt buộc khi chạy trên máy có Excel/VSTO

- mở task pane -> vào VL-NC-M không lỗi;
- số VL / NC / Máy trên pane khớp resource writer;
- công tác có máy yêu cầu NC điều khiển -> NC đó xuất hiện trong pane và sheet;
- logical resource -> đủ ứng viên giá vật lý, không còn mã logic trong VL-NC-M;
- sinh mới VL-NC-M -> công thức NC/M hoạt động;
- nhập giá vật liệu -> refresh -> giá vẫn còn;
- nhập giá nhiên liệu / NC -> refresh -> giá ca máy cập nhật;
- refresh nhiều lần -> metadata không chạy sang phải;
- G:P hoặc các cột phụ legacy không xuất hiện trong vùng in;
- PrintArea có 2 block giống nguyên tắc file mẫu;
- dòng tiêu đề 1–2 ẩn; header máy lặp khi in;
- đóng/mở workbook -> workbook Name và giá nhập vẫn đọc lại được;
- không có `#REF!`, `#VALUE!`, `#NAME?` trong công thức sinh.

### Việc tiếp theo

V2-201 đã chốt. V2-301 đã được triển khai ở section bên dưới.

---

## Các task tiếp theo

### V2-301 — DG Cạn / DG Nước / DG Biển

**Trạng thái:** DONE - implementation / runtime pending

#### UI đã đối chiếu ảnh chuẩn

Đã đọc và bám trực tiếp:

- `/mnt/data/chuan_UI/05-DG-Can.png`
- `/mnt/data/chuan_UI/06-DG-Nuoc.png`

View mới:

- `ExcelAddIn1/Winform/EstimateUnitRatesPaneView.cs`

DG Cạn hiện có đúng cấu trúc đã chốt:

- 4 metric: Định mức cần sinh / Đã sinh / Thiếu giá / Liên kết công thức;
- bước 1 chọn định mức cần sinh;
- bước 2 tùy chọn Sinh mới / Cập nhật từ VL-NC-M / Chỉ sinh định mức đang dùng;
- bước 3 preview VL / NC / M / Tổng cộng;
- nút xanh `Sinh đơn giá`;
- footer giải thích mỗi định mức chỉ sinh một block dùng chung.

DG Nước hiện có:

- 4 metric: Định mức nước / Đã sinh / Thiếu giá / Cảnh báo;
- quy trình 4 bước;
- danh sách định mức công tác dưới nước;
- trạng thái từng định mức;
- footer formula/link.

DG Biển dùng cùng ngôn ngữ UI của DG Nước và **chỉ hiện điều hướng khi thực sự có định mức biển đang dùng**. Không tạo sheet DG Biển rác khi dự toán không có công tác biển.

#### Identity đơn giá

Đã thêm core model:

- `ExcelAddIn1.Core/EstimateV2Rates.cs`

Đơn giá unique theo:

```text
PackageIdentity + NormCode + VariantCode
        -> RateId ổn định
```

Hai hoặc nhiều WorkItem dùng cùng package + định mức + variant chỉ sinh **một block đơn giá** và tăng `UsageCount`; không nhân bản block theo dòng công tác.

Đã xử lý binding legacy có variant rỗng nhưng định mức chỉ có một variant: normalize về variant thật trước khi tạo RateId để không sinh trùng.

#### Phân loại Cạn / Nước / Biển

- `NORM-000.*`, `NORM-010.*`, `NORM-020.*` -> DG Cạn;
- `NORM-030.*` -> DG Nước;
- `NORM-040.*` -> DG Biển.

Sheet chỉ được sinh khi môi trường đó có rate đang dùng.

Tên sheet tương thích mẫu hiện tại:

- `DG Can` / `DG Cạn`;
- `DG Nuoc` / `DG Nước`;
- `DG Bien` / `DG Biển`.

#### Writer đơn giá

Đã thêm:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2RateService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2RateSheetWriter.cs`

Writer:

1. cập nhật `VL-NC-M` trước để bảo đảm workbook Name giá tồn tại;
2. sinh block đơn giá từ định mức/variant đang dùng;
3. hao phí định mức nằm ở cột D;
4. đơn giá resource ở cột E là **formula link** tới workbook Name của `VL-NC-M`;
5. thành tiền:
   - VL -> F;
   - NC -> G;
   - M -> H;
6. vật liệu phần trăm tính bằng công thức từ tổng vật liệu trực tiếp;
7. dòng `Cộng:` giữ riêng VL / NC / M;
8. thêm dòng `Tổng cộng đơn giá` và workbook Name `TOTAL`;
9. metadata nằm từ cột I trở đi và bị ẩn;
10. vùng in chỉ A:H.

Không copy snapshot giá chết vào block đơn giá.

#### Logical resource

Các resource logic như:

- `M010.002-OR-M010.003`;
- `M010.DIVING`;
- `MAT-GASOLINE-OR-DIESEL`;

được mở thành các price candidate vật lý dùng chung với VL-NC-M.

Khi chưa có UI chọn candidate riêng cho từng công tác, V2-301 dùng policy xác định được: **chọn candidate đầu tiên đang có giá > 0** bằng công thức Excel. Pane đồng thời đánh cảnh báo để người dùng biết rate có lựa chọn cần rà soát.

#### Điều kiện / hệ số định mức

State V2 hiện mới lưu NormCode + Variant, chưa lưu tập điều kiện như lưu tốc dòng chảy, đào có nước, độ dốc...

Vì vậy V2-301 hiện sinh **hao phí cơ sở của variant đã gắn**. Nếu `NormDefinition` có adjustment/constraint, rate được đánh cảnh báo `RequiresConditionReview`.

Không tự đoán điều kiện và không ghi hệ số giả vào workbook.

Đây là giới hạn đã biết của V2-301, cần được xử lý ở luồng thiết lập/validation sau; không được âm thầm tính sai.

#### Đối chiếu sheet mẫu thực tế

Đã đọc/rendere:

- `/mnt/data/Du toan RPBM HoaLuNamDinh_Ver1.xlsx`
- `DG Can`
- `DG Nuoc`

Ảnh render dùng để đối chiếu:

- `/mnt/data/dg_can_sample.png`
- `/mnt/data/dg_nuoc_sample.png`

Writer đã chỉnh theo cấu trúc in thực tế:

- Times New Roman;
- A:H;
- title block;
- tên công tác / số hiệu định mức / đơn vị;
- header **hai tầng**:
  - A:E merge dọc;
  - F:H merge ngang `Thành tiền (đồng)`;
  - hàng dưới F/G/H = Vật liệu / Nhân công / Máy;
- section I Vật liệu / II Nhân công / III Máy thi công;
- dòng `Cộng:`;
- Portrait;
- BlackAndWhite;
- FitToPagesWide = 1;
- helper/metadata từ I trở đi ẩn.

#### Navigation

`EstimateTaskPaneControl` đã nối:

```text
Tổng quan
  -> DG Cạn
  -> DG Nước
       -> DG Biển (chỉ khi có định mức biển)
```

Output tile DG Biển cũng chỉ xuất hiện nếu workbook thực sự có sheet biển.

#### Test code đã thêm

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2RatePlan`

Test kiểm tra:

- 2 WorkItem cùng Land norm + variant -> 1 rate, UsageCount = 2;
- tách đúng Land / InlandWater / Sea;
- RateId ổn định;
- `M010.DIVING` có candidate máy lặn cụ thể;
- rate nước có adjustment/constraint được đánh cần rà soát;
- rate biển giữ đúng logical machine để writer liên kết candidate.

#### Commit quan trọng

- `8b40d688ca6d` — expose resource candidate expansion;
- `7e311d8ed845` — core rate plan / RateId;
- `f9e90e8314e3` — preview service DG;
- `58b6120af5e2` — formula-linked DG writer;
- `66960b684b07` — UI DG Cạn/Nước/Biển;
- `8edc18670b88` — navigation vào task pane;
- `71e4764eba62` — core rate tests;
- `c788707acfa4` — merge normalized rate identities;
- `3ba4c36396f6` — stable TOTAL link;
- `0a9dd85097b1` + `c0a564583d1a` — header in hai tầng theo mẫu;
- `8bdf5dbb23d0` — cập nhật UI contract.

#### Tự kiểm tra trong môi trường hiện tại

Đã rà soát tĩnh các file V2-301 chính:

- ngoặc `{}`, `()`, `[]` cân bằng;
- không có `TODO/FIXME/NotImplementedException`;
- kiểm tra lại API MachineRateCatalog dùng bởi writer;
- kiểm tra selector/radio DG Cạn thực sự ảnh hưởng tập rate sinh;
- commit checkbox grid trước khi đọc lựa chọn;
- DG Biển không hiện khi không có rate biển;
- công thức DG dùng workbook Name từ VL-NC-M, không dùng số giá snapshot;
- header sheet đã sửa từ một tầng sang hai tầng theo file mẫu.

**Chưa chạy build VSTO/Excel thật và chưa chạy test console trong môi trường hiện tại. Không ghi PASS giả.**

#### Checklist runtime bắt buộc

- vào DG Cạn từ Tổng quan -> pane mở đúng, không form modal;
- số metric khớp rate plan;
- hai công tác cùng norm+variant -> chỉ một block DG;
- chọn/bỏ chọn rate -> Sinh mới chỉ sinh đúng tập mong muốn và giữ block đã có;
- Cập nhật từ VL-NC-M -> giá đổi thì DG đổi theo formula/link;
- DG Nước chỉ sinh khi có NORM-030;
- DG Biển chỉ sinh khi có NORM-040;
- workbook chỉ có biển -> từ DG Nước vẫn thấy điều hướng DG Biển;
- header in hai tầng đúng A:H;
- cột I trở đi hidden;
- không có `#NAME?`, `#REF!`, `#VALUE!`;
- TOTAL = VL + NC + M;
- save/close/open -> workbook Name rate vẫn còn;
- rate có adjustment/constraint phải hiện cảnh báo, không âm thầm coi như đã áp hệ số.

#### Việc tiếp theo

V2-301 đã chốt. V2-401 đã được triển khai ở section bên dưới.

### V2-401 — Gia DT TC + THKP-TC

**Trạng thái:** DONE - implementation / runtime pending

#### UI đã đối chiếu ảnh chuẩn

Đã đọc và bám trực tiếp:

- `/mnt/data/chuan_UI/07-THKP-TC-Kiem-tra.png`

View mới:

- `ExcelAddIn1/Winform/EstimateCostSummaryPaneView.cs`

Màn hình hiện có:

- 4 metric: Tổng công tác / Đủ đơn giá / Cảnh báo / Lỗi công thức;
- card xanh lớn `Cập nhật THKP-TC`;
- khu `Kết quả kiểm tra hồ sơ` với icon trạng thái và `Xem chi tiết`;
- 3 card chức năng: Kiểm tra hồ sơ / Xuất báo cáo / Mở thư mục hồ sơ;
- footer xanh dương đúng ngôn ngữ thiết kế của ảnh chuẩn;
- vẫn là CustomTaskPane bên phải, không mở form modal.

`Kiểm tra hồ sơ` trong V2-401 chỉ kiểm tra các dependency/link cơ bản cần cho bước tổng hợp. Validation đầy đủ vẫn để đúng task V2-501.

#### Đối chiếu file dự toán thực tế

Đã đọc cấu trúc các mẫu:

- `/mnt/data/Du toan RPBM HoaLuNamDinh_Ver1.xlsx`
- `/mnt/data/Du toan RPBM Pleiku_TP2_QuyNhon_Ver7.2.xlsx`

Các mẫu cho thấy:

- `Gia DT TC` hoặc các sheet `Gia DT TC_...` chứa công tác chi tiết;
- đơn giá VL / NC / M được dùng để tính thành tiền theo từng công tác;
- `THKP-TC` có các dòng chi phí trực tiếp `VL`, `NC`, `M`, `T`;
- các dòng sau đó (chi phí chung, TL, K1..Kn, VAT, tổng cuối...) phụ thuộc vào phần chi phí trực tiếp và khác nhau theo mẫu/pháp lý.

Vì vậy V2-401 **chỉ thay thế nguồn liên kết phần chi phí trực tiếp**, không phá công thức pháp lý phía sau của THKP-TC hiện hữu.

#### WorkItem -> RateId

Đã thêm core model:

- `ExcelAddIn1.Core/EstimateV2CostLinks.cs`

Mỗi WorkItem được resolve theo state bền vững:

```text
WorkItemId
  -> PackageIdentity + NormCode + VariantCode
  -> RateId
  -> workbook Name VL / NC / M
```

Không dùng RowIndex làm identity.

Hai WorkItem dùng chung một định mức + variant tiếp tục dùng chung một RateId nhưng mỗi dòng có khối lượng và thành tiền riêng.

#### Gia DT TC / bảng công tác chi tiết

Writer:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkWriter.cs`

Trên mỗi bảng công tác đã đăng ký, sáu cột kết quả được đặt ngay sau cột Khối lượng:

```text
Đơn giá VL
Đơn giá NC
Đơn giá M
Thành tiền VL
Thành tiền NC
Thành tiền M
```

Với layout chuẩn A:F, chúng tương ứng G:L.

Công thức:

```text
Đơn giá VL = RateId.VL
Đơn giá NC = RateId.NC
Đơn giá M  = RateId.M

Thành tiền VL = Khối lượng × Đơn giá VL
Thành tiền NC = Khối lượng × Đơn giá NC
Thành tiền M  = Khối lượng × Đơn giá M
```

Các đơn giá là **formula tham chiếu workbook Name**, không ghi số chết.

Nếu WorkItem chưa gắn định mức hoặc chưa có DG tương ứng, writer không tự bịa giá và để dòng đó chưa liên kết để pane tiếp tục cảnh báo.

#### Sửa xung đột metadata quan trọng

V2-101 trước đây có thể đặt:

```text
__TTB_ID
__TTB_NORM
__TTB_KIND
__TTB_HASH
```

ngay sau cột visible cuối cùng. Với bảng có Khối lượng ở F, metadata cũ có thể rơi vào G:J — đúng vùng V2-401 cần dùng cho đơn giá/thành tiền.

Đã bổ sung migration:

- `WorkbookEstimateV2RegistrationService.EnsureTechnicalColumnsAfter(...)`

Khi V2-401 chạy:

1. metadata cũ được copy nguyên trạng tới vùng an toàn;
2. tối thiểu nằm sau sáu cột kết quả (M:P với layout A:F), hoặc xa hơn nếu workbook đã dùng các cột đó;
3. vùng metadata mới được hide;
4. custom property của source được cập nhật;
5. cột cũ được giải phóng để G:L là dữ liệu visible;
6. WorkItemId/binding vẫn giữ nguyên.

Đây là migration cấu trúc, không yêu cầu người dùng quét/gắn lại từ đầu.

#### Tổng hợp nhiều bảng / nhiều khu vực

V2-401 không phụ thuộc dự án có 1, 2 hay 3 khu vực.

Mỗi registered source được tổng hợp bằng WorkItemId hợp lệ; helper THKP dùng công thức `SUMIF` trên cột `__TTB_ID` và cột thành tiền tương ứng. Vì vậy:

- dòng nhóm/tổng phụ không có WorkItemId không bị cộng hai lần;
- nhiều bảng/sheet được cộng chung;
- không hard-code tên khu vực VT/DN hay số lượng khu vực.

#### THKP-TC

Đã thêm stable workbook Names:

- `TTBMVN_V2_GIADT_VL`
- `TTBMVN_V2_GIADT_NC`
- `TTBMVN_V2_GIADT_M`
- `TTBMVN_V2_GIADT_TOTAL`

Helper tổng hợp nằm ở các cột ẩn ngoài vùng in của THKP-TC.

Nếu `THKP-TC` đã tồn tại, writer:

1. tự tìm cột `Ký hiệu` và `Thành tiền`;
2. tự tìm các dòng có ký hiệu `VL`, `NC`, `M`, `T`;
3. thay **chỉ** công thức thành tiền của bốn dòng này bằng workbook Name V2;
4. giữ nguyên toàn bộ công thức phía sau của mẫu hiện hữu.

Không hard-code các tỷ lệ minh họa 6,5%, 5,5%, 10% trong ảnh UI.

Nếu workbook chưa có `THKP-TC`, V2-401 tạo một mẫu tối thiểu an toàn chỉ cho phần chi phí trực tiếp và ghi rõ các khoản pháp lý khác chưa được tự suy đoán.

#### Preview / trạng thái

Đã thêm:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkService.cs`

Preview hiện tính:

- tổng WorkItem;
- số WorkItem đã có RateId và workbook Name đơn giá;
- số WorkItem cảnh báo (không double-count cùng một WorkItem);
- lỗi Excel cơ bản trong các ô output;
- THKP có thật sự tham chiếu bốn workbook Name VL/NC/M/T hay chưa.

Chỉ việc workbook Name tồn tại chưa được coi là THKP đã cập nhật; service còn kiểm tra công thức trong THKP có reference các Name đó.

#### Navigation

`EstimateTaskPaneControl` đã nối bước cuối:

```text
Tổng quan
  -> THKP-TC & Kiểm tra
```

Bước 6 chỉ hiển thị `Đã xong` khi THKP-TC thực sự có các liên kết V2, không chỉ vì sheet `THKP-TC` tồn tại.

#### Test code đã thêm

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2CostLinkPlan`

Test core kiểm tra:

- 2 WorkItem cùng rate -> cùng RateId;
- WorkItem chưa gắn -> `Unbound`;
- binding không resolve rate -> `RateNotResolved`;
- đếm Ready / Unbound / MissingRate;
- propagate cờ `RequiresConditionReview`.

#### File chính đã thêm/cập nhật

- `ExcelAddIn1.Core/EstimateV2CostLinks.cs`
- `ExcelAddIn1.Core/EstimateV2ExcelNames.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2RegistrationService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkWriter.cs`
- `ExcelAddIn1/Winform/EstimateCostSummaryPaneView.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneControl.cs`

#### Commit quan trọng

- `e01a8ae866e4` — WorkItem -> RateId link plan;
- `c9d099a88dec` — stable Gia DT TC total names;
- `10e2b4564a72` — di chuyển metadata ra sau vùng đơn giá/thành tiền;
- `308abc67b3ed` — preview service;
- `56ad90b950c7` — writer link WorkItem + THKP;
- `45705188e220` — UI THKP-TC & Kiểm tra;
- `2384fb09bab8` — navigation;
- `99f76f0f34db` — warning metric không double-count;
- `77fe3a8d949c` — core link-plan test;
- `f4306c83aeee` — xác nhận THKP thực sự reference Name V2;
- `7f3ed7ae27fe` — trạng thái bước THKP ở Tổng quan;
- `a6565ed3c03f` — cleanup COM range khi tạo THKP tối thiểu;
- `80d05d920b5c` — cập nhật UI contract.

#### Tự kiểm tra trong môi trường hiện tại

Đã rà soát tĩnh các file V2-401:

- ngoặc `{}`, `()`, `[]` cân bằng;
- không có conflict marker;
- không có `TODO/FIXME/NotImplementedException` trong các file V2-401 chính;
- project file chứa đúng các source mới;
- metadata migration không dùng RowIndex làm identity;
- các giá/ thành tiền là formula/link, không phải số snapshot;
- THKP chỉ cập nhật VL/NC/M/T và không ghi đè các công thức pháp lý phía sau;
- helper THKP nằm ngoài vùng in và bị ẩn.

Đã đối chiếu logic với hai workbook mẫu bằng công cụ spreadsheet, nhưng **chưa chạy build VSTO/Excel thật và chưa chạy test console trong môi trường hiện tại. Không ghi PASS giả.**

#### Checklist runtime bắt buộc

- workbook cũ có metadata ở G:J -> cập nhật -> metadata chuyển sang M:P hoặc vùng an toàn, ID/binding không mất;
- G:L hiện đúng 6 cột đơn giá/thành tiền;
- đơn giá G:I là workbook Name của RateId;
- J:L = Khối lượng × đơn giá;
- section/group/subtotal không có WorkItemId không bị tính trùng;
- dự án 1/2/3 khu vực -> tổng VL/NC/M đúng tổng các WorkItem;
- sửa khối lượng -> thành tiền và THKP cập nhật theo công thức;
- sửa giá VL-NC-M -> DG -> Gia DT TC -> THKP cập nhật theo chuỗi link;
- WorkItem chưa có DG -> không có giá giả, pane báo cảnh báo;
- THKP hiện hữu -> chỉ VL/NC/M/T bị thay source link; các dòng pháp lý phía sau giữ nguyên;
- THKP không tồn tại -> tạo mẫu tối thiểu trực tiếp, không tự gán tỷ lệ pháp lý;
- đóng/mở workbook -> workbook Names/link còn nguyên;
- pane 07 mở đúng, 4 metric và card/action không tràn ở DPI 100/125/150%;
- không có `#REF!`, `#NAME?`, `#VALUE!` sau cập nhật.

#### Việc tiếp theo

V2-401 đã chốt. V2-501 đã được triển khai ở section bên dưới.

 và whitespace;
- dấu `;` / `,` được normalize;
- prefix `_xlfn.` được xử lý;
- công thức THKP dùng đúng stable Name;
- map lỗi `#VALUE!`, `#REF!`, `#N/A`.

#### File chính đã thêm/cập nhật

- `ExcelAddIn1.Core/EstimateV2Validation.cs`
- `ExcelAddIn1.Core/ExcelAddIn1.Core.csproj`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2ValidationService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkService.cs`
- `ExcelAddIn1/Winform/EstimateCostSummaryPaneView.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneControl.cs`
- `ExcelAddIn1.Tests/Program.cs`
- `docs/du-toan-v2/UI-CONTRACT.md`

#### Commit quan trọng

- `3eb8dc25be8b` — core validation/formula rules;
- `4f2190dca543` — workbook validation + safe repair service;
- `e1f886e4e3b0` — THKP preview dùng full validation report;
- `5545726bf359` — validation core tests;
- `fbe4dff199b2` — chạy scan và điều hướng lỗi trong pane 07;
- `ab0c04964d41` — siết trạng thái và repair scope;
- `10ce4cb3368d` — metric formula error nhận lỗi THKP không có WorkItem;
- `0ad838d06e4a` — phát hiện DG formula bị đổi sang formula khác;
- `769994bf3fc7` — kiểm tra formula giá NC/M trong VL-NC-M;
- `bb6770a9162c` — cảnh báo layout metadata cũ;
- `e64a6c46514b` — repair có thể sinh DG thiếu/migrate layout;
- `7513172b15c7` — giữ Tổng quan nhẹ, full validation chạy on-demand;
- `c251d0ddb0de` — tránh chạy full scan lặp lại khi bấm Kiểm tra hồ sơ;
- `2041577ab432` — cập nhật UI contract.

#### Tự kiểm tra trong môi trường hiện tại

Đã tự rà soát tĩnh:

- ngoặc `{}`, `()`, `[]` cân bằng ở các file V2-501 chính;
- không có conflict marker;
- không có `TODO/FIXME/NotImplementedException` trong code V2-501;
- core/service mới đều đã thêm vào đúng `.csproj`;
- validation không dùng RowIndex làm identity;
- vật liệu market price không bị coi là formula bắt buộc;
- NC/M trong VL-NC-M được coi là formula output;
- G:L và THKP dùng exact expected formulas;
- error scan có mapping `#REF!`, `#VALUE!`, `#N/A`;
- safe repair không tự điền missing price và chỉ gọi đúng writer sở hữu.

Đã thử chuẩn bị build ngoài repo nhưng môi trường container hiện tại không có .NET/MSBuild/VSTO runtime và không có network để clone/build repo. Vì vậy **chưa chạy build VSTO/Excel thật và chưa chạy test console; không ghi PASS giả**.

#### Checklist runtime bắt buộc

- xóa ô định mức hiển thị -> scan -> ô được phục hồi từ Custom XML;
- copy/paste tạo duplicate ID -> scan -> ID copy được tách, binding giữ;
- sort visible columns làm ID lệch -> fingerprint unique phục hồi đúng;
- xóa một dòng -> state orphan được báo nhưng không tham gia tổng;
- xóa binding thật -> `MISSING_NORM`;
- xóa một material price -> `MISSING_PRICE`, mở VL-NC-M;
- xóa workbook Name Rate -> `RATE_NAME_MISSING`;
- ghi số đè lên G:I -> `COST_FORMULA_OVERWRITTEN`;
- sửa J:L sang formula khác -> vẫn phát hiện overwrite;
- ghi số đè lên NC/M price cell trong VL-NC-M -> `RESOURCE_FORMULA_OVERWRITTEN`;
- sửa formula TOTAL trong DG -> `RATE_FORMULA_OVERWRITTEN`;
- sửa VL/NC/M/T trong THKP -> `THKP_FORMULA_OVERWRITTEN`;
- tạo `#REF!`, `#VALUE!`, `#N/A` -> đúng severity/address;
- `Xem chi tiết` -> activate đúng sheet/cell;
- RepairRecoverable -> không thay material market price đã nhập;
- RepairRecoverable -> tái sinh đúng DG/VL/Gia/THKP bị hỏng;
- workbook cũ metadata ở G:J -> báo `COST_LAYOUT_PENDING`, cập nhật -> migrate an toàn;
- pane 07 ở DPI 100/125/150% không tràn layout;
- save/close/open -> scan cho cùng kết quả khi workbook không bị sửa.

#### Việc tiếp theo

V2-501 dừng ở đây. Task kế tiếp là **V2-601 — Tương thích file cũ và migration**. Chưa triển khai V2-601 trong task này.

### V2-601 — Migration / compatibility

**Trạng thái:** DONE - implementation / runtime pending

#### UI đã đối chiếu lại

Đã đọc bộ chuẩn trong:

- `/mnt/data/chuan_UI`
- đặc biệt `01-Tong-quan.png`, `02-Cong-tac.png`, `08-Thiet-lap-chung.png` và `du-toan-v2-ui-spec.md`.

V2-601 **không tạo thêm wizard/modal migration**. Luồng giữ đúng thiết kế đã chốt:

```text
Bấm Dự toán
  -> Tổng quan mở ngay
  -> vào Công tác
       -> kiểm tra/migration legacy on-demand
       -> nếu không chắc VT/DN/bản copy thì không tự chọn
```

Không đưa package scan/converter nặng trở lại startup gate.

#### Đối chiếu hai workbook mẫu thực tế

Đã kiểm tra cấu trúc workbook:

- `/mnt/data/Du toan RPBM HoaLuNamDinh_Ver1.xlsx`
- `/mnt/data/Du toan RPBM Pleiku_TP2_QuyNhon_Ver7.2.xlsx`

Mẫu Hoa Lư/Nam Định có:

- `THKP-TC`;
- `Gia DT TC`;
- `DG Can`;
- `DG Nuoc`;
- `VL-NC-M`;
- thêm một số tab/copy legacy như `VL-NC-M_VT`;
- khu vực 1/2... nằm **trong các dòng nhóm của cùng bảng công tác**, không cần kiến trúc sheet riêng cho mỗi khu vực.

Mẫu Pleiku/Quy Nhơn có các alias/copy legacy:

- `Gia DT TC_DN`;
- `DG Can_VT`;
- `DG Nuoc_VT`;
- `DG Can_DN`;
- `DG Nuoc_DN`;
- `VL-NC-M_VT`;
- `VL-NC-M_DN`;
- `THKP-TC (2)`.

Đây là bằng chứng thực tế để V2-601 xử lý VT/DN và copy cũ theo nguyên tắc **không đoán khi có nhiều lựa chọn**.

#### Core compatibility rules

Đã thêm:

- `ExcelAddIn1.Core/EstimateV2Compatibility.cs`

Có classifier cho:

- `EstimateAppendix`;
- `CostSummary`;
- `ResourcePrices`;
- `UnitRateLand`;
- `UnitRateWater`;
- `UnitRateSea`.

Tên sheet được normalize bỏ dấu/ký tự phân cách để nhận các dạng:

```text
Gia DT TC
Gia DT TC_DN
THKP-TC
THKP-TC (2)
DG Can / DG Cạn
DG Can_VT / DG Can_DN
DG Nuoc / DG Nước
DG Nuoc_VT / DG Nuoc_DN
DG Bien / DG Biển
VL-NC-M
VL-NC-M_VT / VL-NC-M_DN
```

Classifier còn tách hint:

- VT / hưởng lương ngân sách;
- DN / không hưởng lương ngân sách/doanh nghiệp.

#### Không tự nhận text định mức legacy thành binding pháp lý

Đã sửa một lỗi logic quan trọng trong:

- `WorkbookEstimateV2StateService.Reconcile(...)`

Trước đây khi state chưa có binding, reconcile có thể lấy trực tiếp text ở ô `Định mức` đang hiển thị để điền vào `NormCode`.

V2-601 đã bỏ hành vi này.

Từ nay:

```text
ô Định mức visible
        = display/cache

binding thật
        = Custom XML state
          + PackageId
          + DataVersion
          + PackageChecksum
          + NormCode
          + VariantCode
```

Workbook cũ có text định mức không được tự biến thành căn cứ pháp lý hợp lệ. Người dùng vẫn phải gắn/xác nhận định mức qua luồng V2.

#### Migration bảng Gia DT TC legacy

Đã thêm:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2CompatibilityService.cs`

Khi vào màn hình `Công tác`, service kiểm tra on-demand.

Nếu chỉ có **một** bảng `Gia DT TC_*` legacy và layout khớp mẫu đã xác minh, converter chuyển:

```text
Legacy:
A TT
B Mô tả
C Đơn vị
D Khối lượng
E Nghiệm thu
F:H Đơn giá VL/NC/M
I:K Thành tiền VL/NC/M
L:N Thành tiền nghiệm thu ...
```

sang:

```text
V2:
A TT
B Mã công tác
C Định mức
D Mô tả công việc
E Đơn vị
F Khối lượng
G:I Đơn giá VL/NC/M
J:L Thành tiền VL/NC/M
M... vùng ẩn / dữ liệu legacy / metadata
```

Converter:

1. chỉ chạy với layout legacy nhận diện chắc chắn;
2. giữ dữ liệu nghiệm thu/tail cũ ở vùng cột ẩn;
3. freeze phần tail legacy thành value trước khi xóa cột Nghiệm thu để tránh helper cũ phát sinh `#REF!`;
4. tạo Mã công tác visible dạng `CT-Lxxxx` cho dòng công tác legacy để fingerprint có anchor ổn định;
5. không tạo mã cho dòng tiêu đề nhóm/khu vực;
6. đăng ký source V2 sau khi layout đã chuyển;
7. binding định mức vẫn để người dùng gắn thủ công.

Nếu layout khác mẫu đã biết, service **không tự dịch cột** mà để người dùng đăng ký thủ công.

#### 1 / 2 / 3 khu vực

Không thêm `KV1/KV2/KV3` vào identity.

Sau migration, dòng nhóm như:

```text
Các công trình thuộc mật độ khu vực 1
Các công trình thuộc mật độ khu vực 2
...
```

không có đơn vị/khối lượng nên không trở thành WorkItem.

Các dòng công tác thật tiếp tục có:

- WorkItemId riêng;
- binding định mức riêng;
- RateId theo định mức + variant;
- tổng hợp theo WorkItem.

Vì vậy dự án 1/2/3 khu vực dùng cùng một code path.

#### VT / DN

V2-601 tuyệt đối không tự chọn sai audience.

Nếu chỉ có một alias legacy phù hợp, có thể chuẩn hóa nó.

Nếu đồng thời có nhiều bản như:

```text
DG Can_VT
DG Can_DN
VL-NC-M_VT
VL-NC-M_DN
```

và chưa có output chuẩn/role duy nhất:

- không tự gán role;
- không tự đăng ký nhiều `Gia DT TC_*`;
- không cộng cả hai bộ vào dự toán;
- ghi status để người dùng xử lý/chọn nguồn phù hợp.

Sau khi writer V2 đã sinh/cập nhật output chuẩn thành công, các copy legacy cùng loại có thể bị **ẩn**, không xóa, nhằm giảm tab rác.

#### Đổi tên sheet

Identity output hiện ưu tiên:

```text
Worksheet Role
    -> CodeName
    -> custom environment metadata
    -> canonical/legacy alias fallback
```

Đã nối lại các writer/service:

- VL-NC-M: `ResourcePrices` role;
- DG Cạn: `UnitRateLand` role;
- DG Nước: `UnitRateWater` role;
- DG Biển: custom property `TTBMVN.EstimateV2.UnitRateEnvironment=Sea`;
- THKP-TC: `CostSummary` role;
- bảng công tác: registration theo worksheet CodeName.

Vì vậy rename tab không làm mất identity V2.

Đã cập nhật cả:

- validation;
- Tổng quan;
- mở VL-NC-M;
- mở THKP;
- `Xem chi tiết`;
- điều hướng tới bảng Gia DT TC đã đăng ký.

#### Metadata không đè vùng visible

`RegisterSelectedRange` không còn mặc định đặt metadata ngay sau cột visible cuối cùng.

V2-601 đặt technical columns tối thiểu:

```text
sau Khối lượng + 6 cột kết quả
và
sau UsedRange hiện hữu
```

nếu cần.

Do đó bảng A:F sẽ giữ G:L cho đơn giá/thành tiền, còn metadata nằm M trở đi hoặc xa hơn.

#### Package cũ / missing / corrupt

Compatibility service kiểm tra ProjectProfile/package pin theo hướng graceful degradation.

Trạng thái:

- `NotConfigured`;
- `Ready`;
- `MissingOrCorrupt`;
- `ProfileCorrupt`.

Nếu workbook đang pin:

```text
PackageId@DataVersion#Checksum
```

mà máy hiện tại thiếu/corrupt package:

- module vẫn mở;
- giữ nguyên identity đã pin;
- thông báo package đang thiếu;
- **không tự chuyển sang package latest**.

`EstimateTaskPaneControl.RefreshOverview` cũng đã được bọc lỗi khi state/profile cũ bị corrupt để dữ liệu legacy không làm task pane crash.

#### Workbook mở khi không có add-in

V2 không dùng UDF của add-in làm kết quả in.

Các output tính toán tiếp tục dùng:

- công thức Excel chuẩn;
- cell reference;
- workbook Name đã lưu trong file.

Custom XML / CustomProperties chỉ phục vụ add-in khi chỉnh sửa/validation.

Vì vậy khi máy không cài add-in:

- workbook vẫn mở được;
- các công thức đã sinh vẫn nằm trong workbook;
- các sheet in vẫn xem/in được;
- người dùng chỉ mất các chức năng UI/migration/gắn định mức của add-in.

#### UI

Không thêm pane mới.

`EstimateWorkItemsPaneView` giữ giao diện ảnh 02 và chỉ bổ sung status:

- số bảng legacy đã migrate;
- cảnh báo VT/DN;
- package pin thiếu/corrupt;
- profile corrupt.

`EstimateTaskPaneControl` giữ ảnh 01 nhưng sheet count/step status nhận được sheet đã đổi tên qua persistent identity.

#### Test code đã thêm

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2CompatibilityRules`

Test code kiểm tra:

- `DG Cạn` -> Land canonical;
- `DG Can_VT` -> Land + StateBudgetSalary;
- `VL-NC-M_DN` -> ResourcePrices + NonStateSalary;
- `THKP-TC (2)` -> CostSummary alias;
- `Gia DT TC_DN` -> EstimateAppendix + DN;
- nhận diện header V2;
- nhận diện header legacy;
- không nhầm bảng thiếu Mã công tác/Định mức thành V2.

#### File chính đã thêm/cập nhật

- `ExcelAddIn1.Core/EstimateV2Compatibility.cs`
- `ExcelAddIn1.Core/ExcelAddIn1.Core.csproj`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2CompatibilityService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2RegistrationService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2StateService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2ResourceSheetWriter.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2RateSheetWriter.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkWriter.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2CostLinkService.cs`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2ValidationService.cs`
- `ExcelAddIn1/Winform/EstimateWorkItemsPaneView.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneControl.cs`
- `ExcelAddIn1/Winform/EstimateResourcesPaneView.cs`
- `ExcelAddIn1/Winform/EstimateCostSummaryPaneView.cs`
- `ExcelAddIn1.Tests/Program.cs`
- `docs/du-toan-v2/UI-CONTRACT.md`

#### Commit quan trọng

- `3885f257210a` — classifier legacy/VT-DN;
- `221ba0f99967` — metadata không đè vùng G:L;
- `04fbf43984d3` — bỏ auto-binding từ text định mức legacy;
- `c9e2273edc7b` — compatibility/migration service;
- `bcd9c1ca7d99` — freeze tail legacy trước structural conversion;
- `764bb5dc2f0f` — migration on-demand khi vào Công tác;
- `f39455e52277` — chuẩn hóa VL-NC-M và ẩn copy cũ sau rebuild;
- `54b94ae2699b` — DG bền vững qua rename;
- `e879fa94e183` — THKP role identity;
- `b4c82d5d82a0` — validation theo persistent output identity;
- `5a2c7d114f6c` — Tổng quan không crash bởi profile/state cũ;
- `51c6a59a3f78` — core compatibility tests;
- `f905047915a9` — không auto-register nhiều Gia DT TC VT/DN;
- `61829470ec34` — DG Biển sống qua rename;
- `d8e9151bc2e0` — Mã công tác ổn định cho legacy;
- `5ba23c4f909d` — cập nhật UI contract V2-601.

#### Tự kiểm tra trong môi trường hiện tại

Đã rà soát tĩnh các file V2-601 chính:

- ngoặc `{}`, `()`, `[]` cân bằng ở các file code V2-601 chính;
- không có conflict marker;
- không có `TODO/FIXME/NotImplementedException`;
- compatibility core/service được include đúng một lần trong project tương ứng;
- migration không tự lấy text Định mức legacy làm legal binding;
- ambiguous VT/DN không được tự chọn;
- metadata mới không chồng G:L;
- rename output được resolve qua Role/CodeName/environment metadata;
- migration không tạo sheet technical visible;
- không thêm startup gate mới.

Đã đọc trực tiếp cấu trúc hai file XLSX mẫu để xác minh tên sheet/header/layout legacy. Spreadsheet rendering engine trong môi trường hiện tại không import được hai workbook mẫu, nên đối chiếu cấu trúc được thực hiện từ nội dung XLSX/XML; **không coi đây là Excel runtime test**.

Môi trường hiện tại vẫn không có Excel/VSTO runtime phù hợp để chạy end-to-end. Vì vậy **chưa ghi PASS runtime/build/test console giả**.

#### Checklist runtime bắt buộc

- workbook trống không có sheet dự toán -> bấm Dự toán vẫn mở Tổng quan;
- workbook Hoa Lư/Nam Định -> vào Công tác không mất dữ liệu và không tạo tab rác;
- workbook Pleiku/Quy Nhơn -> phát hiện VT/DN, không tự chọn đồng thời hai bộ;
- `Gia DT TC_DN` duy nhất -> migrate thành A:L V2, dữ liệu nghiệm thu legacy còn ở vùng ẩn;
- dòng nhóm khu vực 1/2/3 -> không tạo WorkItem;
- dòng công tác legacy -> có Mã công tác `CT-Lxxxx` + WorkItemId riêng;
- ô định mức legacy có text -> không tự trở thành binding package V2;
- tự gắn định mức lại -> save/close/open -> binding còn;
- rename `VL-NC-M` -> pane vẫn mở đúng sheet;
- rename `DG Can` / `DG Nuoc` -> writer cập nhật đúng sheet cũ, không tạo bản sao;
- rename `DG Bien` -> environment metadata resolve đúng;
- rename `THKP-TC` -> validation/update mở đúng sheet;
- rename bảng Gia DT TC đã đăng ký -> writer vẫn resolve theo CodeName;
- missing/corrupt pinned package -> task pane vẫn mở, không tự upgrade latest;
- sau writer V2 thành công -> alias VT/DN/copy legacy cùng loại bị ẩn chứ không bị xóa;
- workbook sau khi sinh xong mở trên máy không cài add-in -> công thức/sheet vẫn xem và in được;
- không phát sinh `#REF!` do structural migration;
- pane ảnh 01/02/04/05/06/07 vẫn không tràn ở DPI 100/125/150%.

#### Việc tiếp theo

V2-601 đã chốt ở mức implementation. Task sau đó **V2-701 — Thiết lập chung** đã được triển khai và được ghi riêng ở section bên dưới.

---

## V2-701 — Thiết lập chung theo workbook

**Trạng thái:** DONE - implementation / runtime pending

### UI đã bám ảnh chuẩn 08

Đã đối chiếu trực tiếp:

- `/mnt/data/chuan_UI/08-Thiet-lap-chung.png`

Không tạo form modal mới. Điểm vào là nút **Thiết lập Chung** trên Ribbon, sau đó mở đúng **CustomTaskPane hiện có**, rộng khoảng 430 px.

View chính:

- `ExcelAddIn1/Winform/EstimateSettingsPaneView.cs`

Bố cục hiện có theo ảnh chốt:

- header `Trợ lý Dự toán / Thiết lập chung`;
- 4 metric: Workbook / Vai trò sheet / Metadata / Tự động lưu;
- mục 1 `Sheet đầu ra` với 6 mapping;
- mục 2 `Hành vi cập nhật`;
- mục 3 `Bảo vệ dữ liệu`;
- ba nút cuối pane: Lưu thiết lập / Khôi phục mặc định / Mở thư mục cấu hình;
- icon Save / Lock / Cloud / Info dùng renderer nội bộ, không dùng emoji.

### Core settings policy

Đã thêm:

- `ExcelAddIn1.Core/EstimateV2Settings.cs`

Settings theo workbook gồm:

- `AutoRestoreNormDisplay`;
- `ValidateOnOpen`;
- `FormulaLinksRequired`;
- `AutoSyncRows`;
- `UseCustomXml`;
- `HideTechnicalColumns`;
- `WarnOnMappingLoss`;
- `AutoSaveEnabled`;
- `AutoSaveMinutes`.

Ba invariant bắt buộc luôn được policy ép ON:

1. kết quả tính phải là formula/link, không số chết;
2. binding thật tiếp tục dùng Custom XML;
3. technical metadata phải ẩn khỏi hồ sơ in.

Auto-save được normalize trong khoảng 1–60 phút, mặc định 5 phút.

### Lưu settings và mapping output

Đã thêm:

- `ExcelAddIn1/Funtion/WorkbookEstimateV2SettingsService.cs`

Settings được lưu bằng **Custom Document Properties** trong chính workbook, không tạo sheet settings visible.

Mapping output gồm 6 slot:

- `CostSummary` -> THKP-TC;
- `EstimateAppendix` -> Gia DT TC;
- `UnitRateLand` -> DG Cạn;
- `UnitRateWater` -> DG Nước;
- `ResourcePrices` -> VL-NC-M;
- `UnitRateSea` -> DG Biển.

Identity không phụ thuộc tên tab:

- các output chính dùng Worksheet Role / CodeName;
- DG Biển dùng custom environment metadata;
- settings còn lưu expected output map để phát hiện mapping đã bị xóa/đổi sai.

Không cho phép một sheet được gán đồng thời cho nhiều output V2.

Nếu SaveConfiguration lỗi giữa chừng, service rollback:

- worksheet roles;
- DG Biển environment mapping;
- expected output maps;
- settings trước đó.

### Khôi phục mặc định

`Khôi phục mặc định` chỉ reset **behavior settings** về policy mặc định.

Không được:

- xóa Custom XML binding;
- xóa sheet mapping hiện có;
- xóa output sheet;
- reset package pháp lý.

### Runtime behavior đã nối

`EstimateTaskPaneControl` đọc settings theo workbook khi pane được tạo.

Đã nối:

- `ValidateOnOpen`: chỉ reconcile/kiểm tra cấu trúc nhẹ; không gọi full package/validation nặng ở startup;
- `AutoRestoreNormDisplay`: cho phép bật/tắt việc tự ghi lại ô định mức display bị xóa, nhưng binding thật trong Custom XML không mất;
- `AutoSyncRows`: nghe `Workbook.SheetChange`, nhưng chỉ debounce khi sheet thay đổi là **registered work source**;
- debounce row sync: khoảng 650 ms;
- `WarnOnMappingLoss`: hiện cảnh báo mềm trên Tổng quan/Settings, không khóa module;
- `AutoSaveEnabled` + phút: dùng timer trong task pane;
- auto-save chỉ gọi `workbook.Save()` khi workbook đã có path, không read-only và đang dirty; không tự bật Save As cho file mới.

Khi đóng workbook, `EstimateTaskPaneManager.CloseForWorkbook` dispose pane nên timer/event theo pane cũng dừng theo lifecycle workbook.

### Ribbon / navigation

Đã thêm nút:

- `Thiết lập Chung`

trong Ribbon Dự toán.

Luồng:

```text
Ribbon
  -> Thiết lập Chung
       -> EstimateTaskPaneManager.ShowSettings(workbook)
       -> mở cùng CustomTaskPane
       -> EstimateSettingsPaneView
```

Không mở form modal và không tạo task pane thứ hai cho cùng workbook.

### Mapping loss

Service lưu expected mapping theo slot.

Nếu người dùng:

- xóa sheet;
- gỡ role;
- đổi mapping sang sheet khác ngoài settings;
- làm mất environment metadata DG Biển;

thì `MappingLossCount` tăng.

Hành vi hiện tại:

- cảnh báo mềm;
- module vẫn mở và các chức năng khác vẫn dùng được;
- người dùng sửa lại trong Thiết lập chung;
- không tự gán đại một sheet khác.

### Test code đã thêm

Trong `ExcelAddIn1.Tests/Program.cs`:

- `EstimateV2SettingsPolicy`

Test core kiểm tra:

- default settings;
- các invariant không thể bị tắt qua Normalize;
- auto-save minutes được clamp đúng min/max.

### File chính đã thêm/cập nhật

- `ExcelAddIn1.Core/EstimateV2Settings.cs`
- `ExcelAddIn1.Core/ExcelAddIn1.Core.csproj`
- `ExcelAddIn1/Funtion/WorkbookEstimateV2SettingsService.cs`
- `ExcelAddIn1/Winform/EstimateSettingsPaneView.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneControl.cs`
- `ExcelAddIn1/Winform/EstimateTaskPaneManager.cs`
- `ExcelAddIn1/Winform/EstimateUiIcons.cs`
- `ExcelAddIn1/Ribbon1.cs`
- `ExcelAddIn1/Ribbon1.Designer.cs`
- `ExcelAddIn1.Tests/Program.cs`
- `docs/du-toan-v2/UI-CONTRACT.md`

### Commit quan trọng

- `e8299a079c56` — core settings policy;
- `0761c829530c` — persist workbook settings/output mapping;
- `688a53a89ce3` — include settings service;
- `798af7bb3d64` — configurable restore norm display;
- `81ea48df5f80` — configurable reconcile runtime;
- `3dd6057a1b3a` — settings pane theo UI 08;
- `961318845a9c` — áp settings vào task pane runtime;
- `ac0d73cfdbef` — navigation settings từ task pane manager;
- `15c653df5d9e` — settings policy tests;
- `0e9f79a379ef` — rollback settings + sheet mappings;
- `b7274db94931` — debounce sync chỉ registered work sheet;
- `c21b429ba119` / `dded04ebf9d6` — Ribbon Thiết lập Chung;
- `14d867142cd1` — giữ Tổng quan đúng ảnh, settings mở từ Ribbon;
- `b20160b35b7f` — UI contract V2-701;
- `43e8a6d6e294` — persist expected output maps + detect mapping loss;
- `840c02e9702e` / `5ea6e007e7f2` — surface mapping loss mềm trong settings/runtime.

### Tự kiểm tra đã làm

Đã rà soát tĩnh các file V2-701 chính:

- ngoặc code cân bằng ở core/service/view/manager/ribbon;
- không có conflict marker;
- không có `TODO/FIXME/NotImplementedException` trong các file V2-701 chính;
- `EstimateV2Settings.cs`, service và pane đều đã include đúng project;
- auto-sync chỉ schedule trên registered source;
- auto-save không Save As workbook mới;
- mapping loss là warning mềm;
- reset mặc định không xóa mapping;
- ba invariant formula/Custom XML/hidden technical luôn ON qua policy.

**Chưa chạy build VSTO/Excel thật và chưa chạy test console trong môi trường hiện tại. Không ghi PASS giả.**

### Checklist runtime bắt buộc

- mở workbook -> Tổng quan vẫn mở ngay, không có settings gate;
- Ribbon `Thiết lập Chung` -> mở đúng pane 08;
- pane 08 ở DPI 100/125/150% không tràn;
- lưu mapping 6 output -> đóng/mở workbook -> mapping còn;
- rename output sheet -> mapping theo CodeName/Role vẫn đúng;
- xóa output sheet đã map -> hiện warning mapping loss, module không crash;
- chọn cùng một sheet cho 2 output -> Save bị từ chối rõ ràng;
- lỗi giữa SaveConfiguration -> rollback mapping/settings cũ;
- reset default -> behavior về mặc định nhưng mapping giữ nguyên;
- tắt AutoRestoreNormDisplay -> xóa ô display không tự viết lại, binding Custom XML vẫn còn;
- bật AutoSyncRows -> insert/delete/copy/sort trên registered source được reconcile sau debounce;
- thay đổi sheet không đăng ký -> không chạy reconcile toàn workbook;
- auto-save bật 5 phút -> chỉ save workbook có path và dirty;
- workbook mới chưa Save As -> timer không bật hộp Save As;
- mapping loss warning có thể tắt bằng setting;
- đóng workbook -> pane/timer/event được dispose sạch;
- mở workbook không có add-in sau khi đã sinh output -> vẫn xem/tính/in bằng công thức Excel.

### Việc tiếp theo

Task kế tiếp là **V2-801 — Gói pháp lý & Dữ liệu**.

Yêu cầu bắt buộc trước khi code UI:

- đọc `docs/du-toan-v2/HANDOVER.md`;
- đọc `UI-CONTRACT.md`;
- kiểm tra `/mnt/data/chuan_UI`.
- tại thời điểm bàn giao, thư mục chuẩn hiện có ảnh **01–08**; chưa thấy file ảnh chuẩn 09–10 trong thư mục này. Nếu vẫn thiếu khi bắt đầu V2-801, phải dựng/chốt mockup 09 theo đúng ngôn ngữ UI hiện tại trước hoặc bám đặc tả UI contract đã có; không tự chuyển sang form/modal kiểu khác.

V2-801 phải tiếp tục nguyên tắc: package pháp lý là dependency on-demand, **không được quay lại gate startup** và **không tự nâng package pinned lên latest**.

---

## Roadmap còn lại

### V2-801 — Gói pháp lý & Dữ liệu

**Trạng thái:** NEXT

Mục tiêu:

- UI package/pháp lý theo CustomTaskPane;
- hiển thị rõ package đang pin: PackageId / DataVersion / checksum;
- phân biệt Installed / Missing / Corrupt / Available;
- cài/chọn package on-demand;
- dùng lại package store/offline update/effective-date/transition/migration engine hiện có;
- nếu đổi package đã pin phải có preview/confirm và rollback an toàn;
- không thay binding pháp lý âm thầm;
- không tự chọn `latest` khi package workbook đang pin bị thiếu.

### V2-901 — Báo cáo & Xuất in

**Trạng thái:** TODO

Mục tiêu:

- UI theo phong cách ảnh 10;
- chọn các sheet hồ sơ chính cần in/xuất;
- kiểm tra PrintArea/print setup trước khi xuất;
- không xuất các cột metadata/technical;
- hỗ trợ bộ output thực tế Cạn/Nước/Biển;
- không tạo tab báo cáo rác chỉ để xuất;
- phần export phải dùng workbook hiện tại làm nguồn sự thật.

---

## Quy tắc cập nhật file tiến độ

Khi hoàn thành một task:

1. Đổi trạng thái task thành `DONE`.
2. Ghi commit SHA quan trọng.
3. Ghi file đã thay đổi.
4. Ghi acceptance criteria đã đạt.
5. Ghi test đã chạy và **kết quả thật**.
6. Chuyển task kế tiếp thành `NEXT`.
7. Không ghi `PASS` nếu chưa thực sự chạy test tương ứng.
8. Mọi thay đổi UI phải đối chiếu `UI-CONTRACT.md` trước khi commit.
