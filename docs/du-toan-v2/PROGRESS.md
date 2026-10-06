# TIẾN ĐỘ TRIỂN KHAI DỰ TOÁN RPBM V2

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
| V2-401 | Link Gia DT TC + THKP-TC | NEXT |
| V2-501 | Validation / phục hồi / phát hiện lỗi | TODO |
| V2-601 | Tương thích file cũ và migration | TODO |

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

V2-301 dừng ở đây. Task kế tiếp là **V2-401 — Link Gia DT TC + THKP-TC**. Chưa triển khai V2-401 trong task này.

### V2-401 — Gia DT TC + THKP-TC

- link đơn giá vào từng WorkItem;
- thành tiền = khối lượng × đơn giá;
- THKP lấy tổng từ Gia DT TC;
- không hard-code số kết quả.

### V2-501 — Validation

- missing norm;
- missing price;
- formula overwritten;
- #REF!, #VALUE!, #N/A;
- orphan/duplicate/mismatch;
- UI theo ảnh chốt 07.

### V2-601 — Migration / compatibility

- 1/2/3 khu vực;
- VT/DN;
- file cũ;
- đổi tên sheet;
- package cũ/missing;
- workbook mở được khi add-in không cài.

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
