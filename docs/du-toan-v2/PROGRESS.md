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
| V2-201 | Tổng hợp và sinh VL-NC-M | NEXT - đang triển khai |
| V2-301 | Sinh DG Cạn / DG Nước / DG Biển | TODO |
| V2-401 | Link Gia DT TC + THKP-TC | TODO |
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

**Trạng thái:** NEXT - đang triển khai

### Mục tiêu

Từ các WorkItem đã gắn định mức:

1. Load đúng package/version của binding.
2. Gom **resource unique** đang thực sự dùng:
   - vật liệu;
   - nhân công;
   - máy thi công.
3. Không dump toàn bộ catalog vào workbook.
4. Sinh/cập nhật sheet `VL-NC-M` theo đúng mẫu in hiện hành.
5. Vùng in giữ nguyên biểu mẫu; metadata phụ đặt ở cột ẩn.
6. Giá người dùng nhập phải được giữ lại khi refresh.
7. Giá nhân công và giá ca máy phải là **formula/link Excel**, không ghi kết quả chết.
8. UI phải khớp màn hình `04-VL-NC-M.png` trong UI contract:
   - 4 metric;
   - tổng quan theo nhóm;
   - danh sách thiếu giá;
   - 4 chức năng chính;
   - thông báo formula/link ở cuối.

### Hướng triển khai ngay

- Tận dụng `NormCatalog`, `MachineRateCatalog`, `PriceProfile` hiện có.
- Thêm Core model `EstimateV2ResourcePlan` để gom resource độc lập với Interop.
- Thêm workbook writer riêng cho `VL-NC-M`, không tái sử dụng writer nào ghi snapshot số chết.
- Tạo `EstimateResourcesPaneView` theo ảnh chốt trước khi nối writer.
- Package chỉ được load khi V2-201 cần; missing package chỉ làm màn hình này báo dependency, không đóng module.

---

## Các task tiếp theo

### V2-301 — DG Cạn / DG Nước / DG Biển

- unique theo NormCode + Variant;
- sinh một block đơn giá dùng chung;
- công thức link về VL-NC-M;
- chỉ tạo sheet khi có công tác tương ứng;
- UI theo ảnh chốt 05/06.

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
