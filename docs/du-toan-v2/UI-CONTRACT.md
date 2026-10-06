# UI CONTRACT — DỰ TOÁN RPBM V2

> Đây là hợp đồng giao diện bắt buộc của module Dự toán V2.
> Các màn hình đã được người dùng chốt. Khi triển khai code, không tự ý đổi bố cục, màu sắc, cách chiếm không gian Excel hoặc chuyển sang form toàn màn hình.

## 1. Nguyên tắc bắt buộc

1. UI chính là **CustomTaskPane dock bên phải Excel**.
2. Pane mặc định rộng khoảng **430 px** và không che phần lớn worksheet.
3. Excel vẫn là vùng làm việc chính bên trái.
4. Không dùng form modal để vào module Dự toán.
5. Không dùng form lớn có menu dọc như module legacy làm UI chính.
6. Màu chủ đạo:
   - xanh lá: hành động chính / trạng thái đạt;
   - vàng: cảnh báo / sheet đơn giá;
   - xanh dương: thông tin / sheet nước / dữ liệu;
   - đỏ: lỗi;
   - nền trắng, card viền xám rất nhạt.
7. Font ưu tiên Segoe UI, mật độ thông tin gọn như UI Excel hiện đại.
8. Icon thiếu phải tự vẽ/renderer nội bộ; không thay bằng bộ icon có phong cách khác.
9. Các sheet Excel bên trái giữ nguyên hình thức in thực tế. Dữ liệu kỹ thuật chỉ nằm ở metadata/cột ẩn.
10. Các màn hình sau phải giữ cùng một ngôn ngữ thiết kế, không mỗi tab một kiểu.

## 2. Bộ ảnh chuẩn đã chốt

Bộ ảnh chuẩn trong quá trình thiết kế:

1. `01-Tong-quan.png`
2. `02-Cong-tac.png`
3. `03-Gan-dinh-muc.png`
4. `04-VL-NC-M.png`
5. `05-DG-Can.png`
6. `06-DG-Nuoc.png`
7. `07-THKP-TC-Kiem-tra.png`
8. `08-Thiet-lap-chung.png`
9. `09-Goi-phap-ly-du-lieu.png`
10. `10-Bao-cao-xuat-in.png`

Các ảnh này là **reference bắt buộc**, không phải ý tưởng tham khảo.

## 3. Mapping màn hình -> code

| Ảnh chuẩn | View / code V2 |
|---|---|
| 01 Tổng quan | `EstimateTaskPaneControl` |
| 02 Công tác | `EstimateWorkItemsPaneView` |
| 03 Gắn định mức | `EstimateNormBindingPaneView` |
| 04 VL-NC-M | `EstimateResourcesPaneView` |
| 05 DG Cạn | `EstimateUnitRatesPaneView` (`Land`) |
| 06 DG Nước | `EstimateUnitRatesPaneView` (`InlandWater`) |
| 07 THKP-TC & Kiểm tra | `EstimateCostSummaryPaneView` (V2-401; V2-501 mở rộng validation) |
| 08 Thiết lập chung | task settings sau V2-101 |
| 09 Gói pháp lý & Dữ liệu | package UI on-demand |
| 10 Báo cáo & Xuất in | report/export phase |

## 4. Tổng quan

Màn hình phải có:

- tiêu đề `Tổng quan`;
- card dự án;
- 4 card metric;
- danh sách quy trình 6 bước;
- card các sheet đầu ra;
- thông báo cuối pane;
- không bắt THKP/package trước khi mở.

## 5. Công tác

Màn hình phải có:

- tiêu đề `Trợ lý Dự toán / Công tác`;
- 3 card metric: Tổng công tác / Đã nhận diện / Chưa gắn;
- bước 1: Chọn bảng công tác;
- bước 2: Nhận diện cột dữ liệu;
- các field: Mã công tác / Định mức / Mô tả / Đơn vị / Khối lượng;
- nút Quét vùng chọn / Nhận diện lại / Đăng ký bảng;
- nút xanh lớn `Tiếp tục gắn định mức`;
- ghi chú rằng module dùng được dù chưa có THKP.

## 6. Gắn định mức

Màn hình phải có:

- metric Đã gắn / Chưa gắn / Định mức đã dùng;
- ô tìm mã / tên định mức;
- filter Tất cả / Trên cạn / Dưới nước / Trên biển;
- bảng mã định mức / tên / đơn vị;
- phần chi tiết hao phí theo tab Vật liệu / Nhân công / Máy;
- nút Gắn dòng này / Gắn các dòng đã chọn / Đổi định mức / Bỏ gắn;
- thông báo rõ: xóa ô hiển thị định mức trên sheet **không làm mất binding**.

## 7. Đơn giá Cạn / Nước / Biển

Implementation V2-301 phải bám trực tiếp `05-DG-Can.png` và `06-DG-Nuoc.png`:

- DG Cạn: 4 metric, 3 bước, bảng chọn định mức, tùy chọn sinh, preview VL/NC/M/Tổng, nút xanh `Sinh đơn giá`, ghi chú formula/link.
- DG Nước: 4 metric, quy trình 4 bước, danh sách định mức nước, trạng thái và ghi chú liên kết.
- DG Biển: chỉ xuất hiện khi có định mức biển đang dùng; dùng cùng ngôn ngữ thiết kế của DG Nước.
- mỗi `package + NormCode + VariantCode` chỉ có một RateId và một block đơn giá dùng chung;
- các cột kết quả phải liên kết tới `VL-NC-M` bằng workbook Name/công thức, không chép giá chết;
- sheet in dùng A:H; metadata/helper đặt từ I trở đi và phải ẩn;
- header bảng đơn giá dùng hai tầng như file mẫu: A:E gộp dọc, F:H có `Thành tiền (đồng)` phía trên `Vật liệu / Nhân công / Máy`.

## 8. THKP-TC & Kiểm tra

Implementation V2-401 phải bám trực tiếp `07-THKP-TC-Kiem-tra.png`:

- tiêu đề `THKP-TC & Kiểm tra` và subtitle ngắn;
- 4 metric: Tổng công tác / Đủ đơn giá / Cảnh báo / Lỗi công thức;
- card xanh lớn `Cập nhật THKP-TC`;
- khu `Kết quả kiểm tra hồ sơ` với icon trạng thái và nút `Xem chi tiết`;
- khu `Chức năng khác` gồm Kiểm tra hồ sơ / Xuất báo cáo / Mở thư mục hồ sơ;
- footer xanh dương nhắc module vẫn mở được khi chưa hoàn thiện THKP.

Quy tắc dữ liệu V2-401:

- trên bảng công tác chuẩn, sau cột Khối lượng phải có 6 cột kết quả: Đơn giá VL / NC / M và Thành tiền VL / NC / M;
- với mẫu chuẩn A:F thì các cột kết quả là G:L; metadata WorkItem phải được chuyển ra M trở đi hoặc xa hơn nếu workbook đã dùng các cột đó;
- Đơn giá VL / NC / M phải là công thức tham chiếu workbook Name của RateId; Thành tiền = Khối lượng × Đơn giá, không ghi số kết quả chết;
- các dòng nhóm/tổng phụ không có WorkItemId không được cộng trùng vào tổng;
- nhiều bảng/khu vực đã đăng ký được tổng hợp theo WorkItemId, không phụ thuộc số khu vực;
- `THKP-TC` hiện hữu chỉ được cập nhật các dòng chi phí trực tiếp VL / NC / M / T; các công thức chi phí chung, thu nhập chịu thuế, K1..Kn, VAT và tổng cuối đang có phải được giữ nguyên;
- không lấy các tỷ lệ minh họa trên ảnh UI làm số pháp lý hard-code;
- nếu workbook chưa có `THKP-TC`, V2-401 chỉ tạo mẫu an toàn cho phần chi phí trực tiếp và ghi rõ các chi phí pháp lý khác chưa được tự suy đoán;
- helper tổng hợp của THKP phải đặt ở cột ẩn ngoài vùng in và được neo bằng workbook Name ổn định.

V2-401 chỉ cung cấp kiểm tra liên kết cơ bản để phục vụ màn hình này. Validation đầy đủ (#REF!, #VALUE!, override công thức, orphan/mismatch...) thuộc V2-501.

## 9. Quy tắc mở lại workbook

Khi mở pane:

- tự đọc Custom XML;
- tự đọc metadata worksheet;
- tự reconcile WorkItemId;
- tự sửa lại cột hiển thị định mức nếu bị người dùng xóa nhầm;
- không bắt quét lại bảng đã đăng ký;
- không dùng RowIndex làm identity.

## 10. Trạng thái triển khai hiện tại

Đã có:

- CustomTaskPane dock bên phải.
- Renderer icon nội bộ `EstimateUiIcons`.
- Tổng quan.
- Công tác.
- Gắn định mức.
- WorkItemId / Custom XML / cột kỹ thuật ẩn.
- Reconcile khi mở pane.
- Fingerprint recovery khi chèn/xóa/sort/copy.
- Missing package không chặn task pane.
- VL-NC-M.
- DG Cạn / DG Nước; DG Biển dùng cùng view và chỉ hiện khi có định mức biển.
- THKP-TC & Kiểm tra theo ảnh 07; V2-401 đã nối chi phí trực tiếp, V2-501 sẽ mở rộng validation.

Các màn hình còn lại phải tiếp tục cùng phong cách và kích thước này.
