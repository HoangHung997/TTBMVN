# Chính sách bảo mật

## Báo cáo lỗ hổng

Không đăng công khai license key, private key, workbook khách hàng, log runtime hoặc gói chẩn đoán
trong GitHub issue. Hãy liên hệ trực tiếp chủ repository để thống nhất kênh truyền riêng trước khi
gửi bằng chứng nhạy cảm.

## Dữ liệu không được commit

- Private key license ECDSA và private key ký gói cập nhật.
- PFX/P12/PVK/PEM/KEY/SNK, mật khẩu, token truy cập và certificate có private key.
- File license của người dùng, Machine ID dạng rõ và thông tin phần cứng nhận dạng máy.
- Log, support package, workbook hoặc đường dẫn chứa dữ liệu khách hàng.
- Artifact phát hành chưa qua release gate.

`.gitignore` chặn các định dạng phổ biến nhưng không thay thế việc rà soát trước commit.

## Mô hình ký license

Client chỉ chứa public key ECDSA P-256. LicenseTool ký bằng private key nằm ngoài repository tại
`%LOCALAPPDATA%\TTBMVN Excel Tools\Publisher\license-signing-key-v2.bin` hoặc đường dẫn do biến
`TTBMVN_LICENSE_SIGNING_KEY` chỉ định.

Mất private key đồng nghĩa không thể cấp key tương thích với public key đang phát hành. Rò rỉ
private key yêu cầu thu hồi, thay public key trong ứng dụng, phát hành phiên bản mới và cấp lại
license. Không gửi private key qua email/chat và không lưu bản sao không mã hóa trên cloud.

## Gói cập nhật và chẩn đoán

- Gói cập nhật offline phải qua checksum, chữ ký, version policy và tamper test.
- Support package phải token hóa Machine ID, workbook/project ID và loại dữ liệu nội dung.
- Mọi thay đổi lớp lọc log phải chạy lại nhóm test privacy trong `docs/du-toan/TESTING.md`.

## Repository public

Repository là public. Mọi nội dung đã push phải được coi là đã công khai vĩnh viễn, kể cả khi xóa
ở commit sau. Nếu phát hiện secret trong lịch sử, phải xoay khóa ngay; chỉ rewrite lịch sử là chưa
đủ.
