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
| 07 THKP-TC & Kiểm tra | V2-401 / V2-501 |
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

## 8. Quy tắc mở lại workbook

Khi mở pane:

- tự đọc Custom XML;
- tự đọc metadata worksheet;
- tự reconcile WorkItemId;
- tự sửa lại cột hiển thị định mức nếu bị người dùng xóa nhầm;
- không bắt quét lại bảng đã đăng ký;
- không dùng RowIndex làm identity.

## 9. Trạng thái triển khai hiện tại

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

Các màn hình còn lại phải tiếp tục cùng phong cách và kích thước này.
