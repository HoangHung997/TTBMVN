# Chỉ mục tài liệu

File này là bản đồ đọc tài liệu của TTBMVN Excel Tools. Không cần đọc toàn bộ repository theo
thứ tự tên file; hãy đi theo vai trò dưới đây.

## Người mới tiếp nhận dự án

1. [PROJECT_OVERVIEW.md](PROJECT_OVERVIEW.md) để hiểu sản phẩm và các module.
2. [PROJECT_STATUS.md](PROJECT_STATUS.md) để biết trạng thái chính xác và việc tiếp theo.
3. [DEVELOPMENT_GUIDE.md](DEVELOPMENT_GUIDE.md) để dựng môi trường và chạy kiểm thử.
4. [du-toan/HANDOFF.md](du-toan/HANDOFF.md) để xem bằng chứng kỹ thuật và checkpoint gần nhất.
5. [du-toan/PROGRESS.md](du-toan/PROGRESS.md) để xem bảng task chính thức.

## QA và phân tích nghiệp vụ

- [QA_FUNCTIONAL_SPEC_1.0.4.md](QA_FUNCTIONAL_SPEC_1.0.4.md): đặc tả kiểm thử toàn ứng dụng,
  ma trận rủi ro và danh sách quyết định logic.
- [du-toan/TESTING.md](du-toan/TESTING.md): lệnh test, live Excel test và tiêu chí gate.
- [du-toan/BASELINE.md](du-toan/BASELINE.md): workbook chuẩn, checksum và kết quả kỳ vọng.
- [du-toan/RELEASE-ACCEPTANCE.md](du-toan/RELEASE-ACCEPTANCE.md): biên bản acceptance hiện có.
- [HO_DAO_USER_GUIDE.md](HO_DAO_USER_GUIDE.md): luồng sử dụng form hố đào.

## Lập trình viên

- [PROJECT_OVERVIEW.md](PROJECT_OVERVIEW.md): ranh giới module và luồng dữ liệu.
- [DEVELOPMENT_GUIDE.md](DEVELOPMENT_GUIDE.md): build, test, debug và quy ước thay đổi.
- [du-toan/ARCHITECTURE.md](du-toan/ARCHITECTURE.md): quyết định kiến trúc chi tiết.
- [du-toan/TASKS.md](du-toan/TASKS.md): yêu cầu và điều kiện nghiệm thu từng task.
- [du-toan/WORKLOG.md](du-toan/WORKLOG.md): lịch sử triển khai và bằng chứng test.
- [CONTRIBUTING.md](../CONTRIBUTING.md): quy trình đóng góp vào repository.

## Build, cài đặt và phát hành

- [VSTO_DEPLOYMENT.md](VSTO_DEPLOYMENT.md): triển khai VSTO.
- [du-toan/INSTALLATION.md](du-toan/INSTALLATION.md): cài, nâng cấp và gỡ cài đặt.
- [du-toan/RELEASE-GATE.md](du-toan/RELEASE-GATE.md): điều kiện chặn phát hành.
- [du-toan/RELEASE-NOTES-1.0.4.md](du-toan/RELEASE-NOTES-1.0.4.md): thay đổi bản PILOT hiện tại.
- [du-toan/OFFLINE-UPDATE.md](du-toan/OFFLINE-UPDATE.md): gói cập nhật dữ liệu quy định.
- [LICENSE_TOOL.md](LICENSE_TOOL.md): cấp key offline và quản lý private key phát hành.
- [SECURITY.md](../SECURITY.md): dữ liệu nhạy cảm và quy trình báo cáo bảo mật.

## Căn cứ pháp lý và dữ liệu

- [du-toan/SOURCES.md](du-toan/SOURCES.md): danh mục nguồn và phạm vi sử dụng.
- `data/regulations/raw/`: PDF nguồn.
- `data/regulations/text/`: lớp text trích xuất phục vụ đối chiếu.
- `data/regulations/packages/`: package dữ liệu dùng bởi engine.

Khi thêm tài liệu mới, phải cập nhật file này và liên kết từ `README.md` nếu tài liệu đó là điểm
vào chính của dự án.
