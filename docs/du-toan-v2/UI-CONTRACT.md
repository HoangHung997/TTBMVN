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

**Trạng thái file tham chiếu tại thời điểm bàn giao:** `/mnt/data/chuan_UI` hiện có file ảnh 01–08. Chưa thấy file ảnh chuẩn 09–10 trong folder này. Khi bắt đầu màn hình 09 hoặc 10, nếu ảnh vẫn thiếu thì phải dựng/chốt mockup cùng đúng ngôn ngữ của 01–08 hoặc bám đặc tả đã ghi trong contract; không tự chuyển sang form/modal hay phong cách UI khác.

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
| 08 Thiết lập chung | `EstimateSettingsPaneView` (V2-701) |
| 09 Gói pháp lý & Dữ liệu | `EstimatePackagesPaneView`, đặc tả bố cục trong V2-801.md |
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

V2-501 mở rộng chính màn hình 07, **không tạo thêm một kiểu UI khác**:

- nút/card `Kiểm tra hồ sơ` chạy validation đầy đủ nhưng vẫn giữ đúng bố cục ảnh 07;
- metric `Cảnh báo` đếm WorkItem bị ảnh hưởng theo ID duy nhất, không double-count một công tác vì nhiều lỗi;
- metric `Lỗi công thức` phải lên đỏ khi có lỗi công thức V2, kể cả lỗi tổng hợp THKP không gắn trực tiếp với một WorkItem;
- kết quả kiểm tra phải phát hiện tối thiểu:
  - công tác chưa gắn định mức;
  - thiếu giá VL/NC/M;
  - RateId / workbook Name bị thiếu;
  - công thức V2 bị ghi đè;
  - `#REF!`, `#VALUE!`, `#N/A` và các lỗi Excel khác trong vùng V2 quản lý;
  - orphan WorkItem;
  - WorkItemId trùng;
  - metadata/state mismatch;
  - layout cũ có metadata chồng vào vùng G:L;
- `Xem chi tiết` phải ưu tiên mở đúng sheet/ô lỗi khi xác định được địa chỉ;
- scan được phép tự reconcile các sửa chữa **an toàn và xác định được** như:
  - phục hồi WorkItemId theo fingerprint duy nhất;
  - tách duplicate ID do copy/paste;
  - phục hồi ô hiển thị định mức từ binding Custom XML;
- tuyệt đối không tự điền giá thị trường bị thiếu, không tự suy đoán điều kiện định mức và không tự thay công thức pháp lý ngoài vùng V2 quản lý;
- engine có `RepairRecoverable` để tái sinh công thức/link V2 qua đúng writer sở hữu (VL-NC-M / DG / Gia DT TC / THKP-TC); thao tác kiểm tra thông thường không âm thầm ghi lại giá đầu vào của người dùng;
- Tổng quan phải tiếp tục nhẹ và không chạy full validation/package scan ở startup; bước 6 chỉ phản ánh nhanh trạng thái THKP link, còn full validation chạy khi người dùng mở `THKP-TC & Kiểm tra` hoặc bấm `Kiểm tra hồ sơ`.

## 9. Tương thích workbook cũ / migration V2-601

V2-601 **không tạo thêm một kiểu UI mới**. Các trạng thái tương thích được đưa vào đúng các màn hình đã chốt, chủ yếu là `Tổng quan` và `Công tác`.

Nguyên tắc bắt buộc:

- mở module vẫn phải nhẹ; không chạy converter/package scan nặng như một gate khởi động;
- migration legacy chỉ chạy on-demand khi vào `Công tác`, nơi người dùng đang thao tác với bảng công tác;
- workbook chưa có bất kỳ sheet dự toán nào vẫn mở được bình thường;
- workbook có 1/2/3 khu vực dùng cùng một kiến trúc WorkItem; dòng tiêu đề khu vực không được coi là công tác;
- các tên legacy `Gia DT TC_DN`, `DG Can_VT`, `DG Can_DN`, `DG Nuoc_VT`, `DG Nuoc_DN`, `VL-NC-M_VT`, `VL-NC-M_DN`, `THKP-TC (2)` phải được nhận diện là alias/copy cũ;
- nếu có đồng thời VT/DN hoặc nhiều bản copy cùng vai trò mà không có một lựa chọn duy nhất, add-in **không được tự đoán**; phải giữ nguyên và yêu cầu người dùng chọn đúng nguồn khi cần;
- nếu chỉ có một `Gia DT TC_*` legacy phù hợp, converter được phép chuẩn hóa về layout V2:
  - A TT;
  - B Mã công tác;
  - C Định mức;
  - D Mô tả công việc;
  - E Đơn vị;
  - F Khối lượng;
  - G:I Đơn giá VL/NC/M;
  - J:L Thành tiền VL/NC/M;
  - dữ liệu nghiệm thu/tail legacy được giữ ở vùng cột ẩn ngoài vùng in;
- mã công tác được tạo cho dòng legacy chỉ nhằm tạo anchor visible ổn định; **không** được tự biến nội dung ô Định mức cũ thành binding pháp lý;
- binding định mức chỉ hợp lệ khi được gắn qua state/package V2; text legacy ở ô hiển thị chỉ là display;
- technical metadata mới phải đặt sau vùng visible/output đang dùng, không chèn đè G:L;
- sheet output V2 dùng Worksheet Role / CodeName / custom environment metadata làm identity bền vững; đổi tên tab không được làm mất liên kết;
- sau khi writer V2 đã tạo/cập nhật thành công sheet chuẩn, các bản output legacy VT/DN/copy thừa cùng loại có thể bị ẩn để giảm tab rác; không xóa dữ liệu cũ một cách âm thầm;
- project profile/package pin bị thiếu/corrupt không được làm task pane crash; không tự chuyển sang package latest;
- workbook generated phải dùng công thức Excel + workbook Name chuẩn, không dùng UDF của add-in cho kết quả in, để vẫn xem/tính/in được khi máy không cài add-in.

Về giao diện:

- `Công tác` giữ đúng ảnh 02; chỉ bổ sung thông điệp trạng thái migration trong footer/status;
- `Tổng quan` giữ đúng ảnh 01; trạng thái sheet phải nhận được cả sheet đã đổi tên qua identity bền vững;
- các pane VL-NC-M, DG và THKP phải mở đúng sheet hiện hành kể cả khi người dùng đã đổi tên tab;
- không thêm wizard/modal bắt buộc trước khi vào module.

## 10. Thiết lập chung

Implementation V2-701 phải bám trực tiếp `08-Thiet-lap-chung.png`.

Bố cục bắt buộc:

- header `Trợ lý Dự toán / Thiết lập chung` với icon bánh răng;
- subtitle `Cấu hình workbook và hành vi của Trợ lý Dự toán`;
- 4 card trạng thái:
  - Workbook;
  - Vai trò sheet;
  - Metadata;
  - Tự động lưu;
- mục 1 `Sheet đầu ra` gồm 6 mapping:
  - THKP-TC / Tổng hợp chi phí;
  - Gia DT TC / Bảng dự toán;
  - DG Cạn / Đơn giá cạn;
  - DG Nước / Đơn giá nước;
  - VL-NC-M / Vật liệu - Nhân công - Máy;
  - DG Biển / Đơn giá biển;
- mục 2 `Hành vi cập nhật`:
  - Tự phục hồi ô định mức hiển thị;
  - Kiểm tra khi mở file;
  - Tạo công thức thay vì số chết;
  - Tự động đồng bộ khi chèn/xóa dòng;
- mục 3 `Bảo vệ dữ liệu`:
  - Sử dụng vùng Custom XML;
  - Cột kỹ thuật ẩn;
  - Cảnh báo khi phát hiện xóa mapping;
- ba nút cuối pane:
  - `Lưu thiết lập`;
  - `Khôi phục mặc định`;
  - `Mở thư mục cấu hình`.

Quy tắc dữ liệu/behavior:

- settings là **theo workbook**, lưu trong Custom Document Properties; không tạo sheet settings visible;
- mapping sheet dùng Worksheet Role / CodeName / environment metadata, không phụ thuộc tên tab;
- một sheet không được gán đồng thời cho nhiều output V2;
- `Khôi phục mặc định` chỉ reset hành vi, **không xóa mapping sheet**;
- `Tạo công thức thay vì số chết` là invariant bắt buộc và luôn ON;
- `Sử dụng vùng Custom XML` là invariant bắt buộc và luôn ON;
- `Cột kỹ thuật ẩn` là invariant bắt buộc và luôn ON;
- các invariant bắt buộc có thể hiển thị như switch/lock theo ảnh nhưng không cho phép tắt;
- `Tự phục hồi ô định mức hiển thị` điều khiển việc reconcile có ghi lại ô display bị xóa hay không; binding thật trong Custom XML không bị xóa;
- `Kiểm tra khi mở file` chỉ chạy kiểm tra/reconcile cấu trúc nhẹ; không được kéo full package/validation nặng trở lại startup gate;
- `Tự động đồng bộ khi chèn/xóa dòng` dùng event + debounce trên các bảng công tác đã đăng ký, không poll toàn workbook liên tục;
- tự động lưu mặc định 5 phút; chỉ Save workbook đã có đường dẫn, không tự bật Save As cho workbook mới;
- cảnh báo mapping chỉ là status mềm, không khóa module;
- nút `Mở thư mục cấu hình` mở thư mục ứng dụng TTBMVN trong LocalAppData;
- icon Save / Lock / Cloud / Info phải dùng renderer nội bộ, không dùng emoji/icon lệch phong cách.

Điểm vào UI:

- Ribbon có nút `Thiết lập Chung`;
- nút mở cùng CustomTaskPane hiện tại, không mở form modal mới;
- quay lại Tổng quan vẫn giữ pane rộng khoảng 430 px.

## 11. Quy tắc mở lại workbook

Khi mở pane:

- tự đọc Custom XML;
- tự đọc metadata worksheet;
- tự reconcile WorkItemId;
- tự sửa lại cột hiển thị định mức nếu bị người dùng xóa nhầm;
- không bắt quét lại bảng đã đăng ký;
- không dùng RowIndex làm identity.

## 12. Trạng thái triển khai hiện tại

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
- THKP-TC & Kiểm tra theo ảnh 07; V2-401 đã nối chi phí trực tiếp và V2-501 đã mở rộng validation/phục hồi an toàn.
- V2-601 tương thích workbook legacy/VT-DN/đổi tên sheet theo identity bền vững mà không thêm startup gate hay UI modal mới.
- V2-701 Thiết lập chung theo ảnh 08: mapping output, hành vi runtime, bảo vệ metadata và auto-save theo workbook.

Các màn hình 09–10 phải tiếp tục cùng phong cách và kích thước này.
