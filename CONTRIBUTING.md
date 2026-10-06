# Đóng góp vào TTBMVN Excel Tools

## Trước khi sửa

- Đọc `README.md`, `docs/PROJECT_STATUS.md` và `docs/INDEX.md`.
- Với logic dự toán, đọc thêm `docs/du-toan/HANDOFF.md`, `TASKS.md` và `TESTING.md`.
- Không tự suy diễn các mục `LR-*` chưa được QA chốt.
- Một thời điểm chỉ có một task chính ở trạng thái `IN_PROGRESS`.

## Phạm vi thay đổi

- Giữ engine nghiệp vụ trong `ExcelAddIn1.Core` khi không cần Excel COM.
- Không gọi Excel COM từ worker thread.
- Không đổi workbook baseline, package quy định đã phát hành hoặc checksum lịch sử nếu task không
  yêu cầu rõ.
- Không đưa refactor không liên quan vào cùng commit sửa lỗi.

## Kiểm thử bắt buộc

1. Build `Release|x64` bằng `scripts/build-release.ps1`.
2. Chạy toàn bộ `ExcelAddIn1.Tests`.
3. Chạy nhóm live Excel test tương ứng trong `docs/du-toan/TESTING.md`.
4. Với thay đổi ghi workbook, kiểm tra rollback và mở lại file.
5. Cập nhật tài liệu trạng thái và bằng chứng test.

## Commit và pull request

- Commit message ngắn, mô tả hành vi thay đổi.
- Pull request phải nêu lý do, phạm vi, test đã chạy và ảnh hưởng workbook/package/license.
- Không commit `bin`, `obj`, `artifacts`, `tmp`, private key, PFX, log khách hàng hay gói hỗ trợ.
- Thay đổi pháp lý phải trỏ tới nguồn trong `docs/du-toan/SOURCES.md` và có audit record.
