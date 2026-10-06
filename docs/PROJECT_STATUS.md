# Trạng thái dự án

Cập nhật: 2026-10-06.

## Bản hiện tại

- Kênh: `PILOT x64`.
- Release gần nhất đã đóng gói: `1.0.4`.
- Core test: `71/71 PASS` sau khi chuyển license sang ECDSA P-256.
- Lộ trình dự toán: `DT-101` đến `DT-701` đã `DONE`.
- Workbook chuẩn và checkpoint: có trong `Dutoanmau/`; checksum ở `du-toan/BASELINE.md` và
  `du-toan/HANDOFF.md`.
- Repository GitHub: `HoangHung997/TTBMVN`, nhánh chính `main`.

## Đã hoàn thành

- Hố đào 3 m/5 m, cấu hình H/V hai chiều, xử lý song song theo logical processor và batch write.
- Mapping dữ liệu nguồn/kết quả, sheet mới/sheet hiện có, định dạng, STT và link ngược tổng hợp.
- Role mapping sheet, ProjectProfile, RegulationPackage và migration có rollback.
- Bộ dữ liệu BQP 2021/2025, tra cứu định mức, chi phí, PriceProfile, đơn giá, phụ lục và THKP.
- Estimate Workspace modeless, nhóm đơn giá theo định mức chi tiết/ngữ cảnh và đối tượng lương.
- Runtime diagnostics, support package đã lọc, update offline có chữ ký và release gate.
- Bộ cài đặt xử lý đăng ký VSTO cũ cho các bản PILOT gần nhất.

## Việc phải chốt trước thương mại hóa

1. QA duyệt toàn bộ mục P0/P1 trong `QA_FUNCTIONAL_SPEC_1.0.4.md`, đặc biệt luồng ghi workbook,
   công thức tổng hợp, quyền override và hành vi khi thiếu dữ liệu.
2. Cấp lại các license HMAC cũ bằng key `TTB26` vì cơ chế cũ đã bị loại khỏi mã nguồn công khai.
3. Chạy lại acceptance installer `1.0.4` sau thay đổi chữ ký license và tạo artifact mới.
4. Xác minh trên ít nhất một máy sạch ngoài máy phát triển, bao gồm nâng cấp từ bản đã cài.
5. Chốt chính sách bản quyền repository, kênh báo lỗi bảo mật và nơi lưu private key dự phòng.
6. Bổ sung CI build/test khi có runner Windows chứa VSTO build tools; không đặt khóa ký trong CI log.

## Điểm tiếp tục cho người khác

Không có task đang `IN_PROGRESS`. Công việc tiếp theo nên được tạo thành task mới sau khi QA chốt
logic. Trước khi làm:

1. Đọc [QA_FUNCTIONAL_SPEC_1.0.4.md](QA_FUNCTIONAL_SPEC_1.0.4.md).
2. Đọc [du-toan/HANDOFF.md](du-toan/HANDOFF.md) và [du-toan/PROGRESS.md](du-toan/PROGRESS.md).
3. Xác nhận workbook/checksum trong [du-toan/BASELINE.md](du-toan/BASELINE.md).
4. Tạo task có điều kiện nghiệm thu trong `du-toan/TASKS.md`.
5. Chỉ đánh dấu `DONE` sau khi test của task và regression liên quan đều đạt.

## Rủi ro đã biết

- VSTO identity có thể xung đột với bản debug/stale ClickOnce trên máy từng cài nhiều kênh.
- Excel COM vẫn có thể trả `0x800A03EC`; log correlation ID là nguồn chẩn đoán chính.
- Các workbook khách hàng có merge, filter, protection hoặc công thức đặc thù cần test riêng.
- Dữ liệu pháp lý phải được rà soát lại khi có thông tư mới; không sửa package đã phát hành tại chỗ.
- License key `TTB25`/legacy không được cơ chế mới chấp nhận.
