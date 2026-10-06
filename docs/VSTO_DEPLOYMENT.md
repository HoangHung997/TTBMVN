# VSTO Deployment Notes

## Build Chuẩn

Add-in nên build/publish cùng platform với Office trên máy người dùng.

- Office 64-bit: dùng `x64`.
- Office 32-bit: dùng `Any CPU` hoặc tạo thêm cấu hình `x86` nếu cần.

Lệnh kiểm tra nhanh:

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" .\ExcelAddIn1.sln /t:Build /p:Configuration=Release /p:Platform=x64
```

## Publish Ổn Định

Không cài trực tiếp từ nhiều thư mục `bin\...` khác nhau. VSTO ghi nhớ vị trí cài đặt, nên nếu lần trước cài từ `bin\Debug` rồi lần sau cài từ `bin\x64\Debug`, người dùng có thể gặp lỗi:

```text
another version is currently installed and cannot be upgraded from this location
```

Quy trình phát hành khuyến nghị:

1. Build `Release|x64`.
2. Publish ra một thư mục cố định, ví dụ `publish\x64`.
3. Mỗi lần update chỉ publish đè lên đúng thư mục đó.
4. Nếu đổi location publish, cần gỡ bản cũ trong `Control Panel > Programs and Features` trước.

## Ký Code

Hiện dự án đã đổi user-facing product name thành `TTBMVN Excel Tools`, nhưng vẫn cần chứng chỉ ký code thật trước khi bán.

Việc cần làm trước release thương mại:

- Dùng certificate thuộc công ty/cá nhân phát hành.
- Không dùng temporary/self-signed certificate cho khách hàng.
- Version hóa theo từng bản phát hành.
- Lưu lại đúng publish URL/location để update không bị conflict.
