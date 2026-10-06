# Đặc tả chức năng và hồ sơ rà soát logic QA

## 1. Kiểm soát tài liệu

| Thuộc tính | Giá trị |
|---|---|
| Sản phẩm | TTBMVN Excel Tools |
| Bản được đặc tả | 1.0.4.0, kênh PILOT, x64 |
| Ngày rà soát nguồn | 04/08/2026 |
| Workbook chuẩn | `Dutoanmau/00. Du toan TP 3.xlsm` |
| Package mặc định hiện tại | `BQP-RPBM-2025@2.0.1` |
| Mục đích | Cho QA và chủ sản phẩm xác nhận logic trước khi tiếp tục sửa code |
| Phạm vi | Hố đào, Dự toán RPBM, cập nhật, hỗ trợ, kích hoạt và phát hành |

Tài liệu này mô tả **hành vi thực tế của code hiện tại**, không mặc nhiên coi hành vi đó
là yêu cầu đúng. Những mục có nhãn `CẦN CHỐT` là nơi QA/chủ sản phẩm phải chọn một
trong bốn kết luận: `Giữ nguyên`, `Sửa`, `Loại bỏ`, hoặc `Chưa đủ căn cứ`.

### 1.1. Quy ước trạng thái

| Trạng thái | Ý nghĩa |
|---|---|
| `ĐÃ CÓ` | Có luồng code và đã có kiểm thử tương ứng |
| `MỘT PHẦN` | Có giao diện hoặc engine nhưng chưa nối kín toàn bộ luồng |
| `LEGACY` | Chức năng cũ còn tồn tại, chưa phù hợp kiến trúc mới |
| `CHƯA CÓ` | Chưa có hành vi trong code hiện tại |
| `CẦN CHỐT` | Logic nghiệp vụ chưa được chủ sản phẩm phê duyệt rõ ràng |

### 1.2. Nguyên tắc QA

1. QA không chỉ xác nhận “không báo lỗi”; phải so sánh kết quả Excel, công thức, liên kết,
   định dạng, nguồn pháp lý và dữ liệu lưu sau khi đóng/mở lại Excel.
2. Mọi test ghi Excel phải chạy trên bản sao workbook, không ghi đè baseline.
3. Với phép tính pháp lý, kết quả hiển thị và công thức Excel phải cùng đúng. Không dùng
   engine render Open XML làm oracle tính toán thay Microsoft Excel.
4. Với Hố đào, không so sánh bộ số ngẫu nhiên cụ thể; so sánh bất biến, cận, tổng và số dòng.
5. Test modeless phải thử chuyển qua lại giữa ít nhất hai workbook đang mở.

## 2. Phạm vi sản phẩm hiện tại

### 2.1. Các miền chức năng

| Miền | Điểm vào | Trạng thái |
|---|---|---|
| Hố đào 3 m/5 m | Ribbon group `Hố đào` và form `Xử lý hố đào` | `ĐÃ CÓ`, còn rủi ro logic |
| Dự toán RPBM | Nút `Dự toán` trên Ribbon | `MỘT PHẦN`, module riêng đã có nhưng luồng mới chưa nối kín |
| Hỗ trợ/chẩn đoán | Nút `Hỗ trợ` trong hai form | `ĐÃ CÓ` |
| Kích hoạt offline | Form `Hỗ trợ và kích hoạt` | `MỘT PHẦN`, có nhập key nhưng chưa khóa chức năng |
| Kích hoạt online | Form hỗ trợ | `CHƯA CÓ`, nút bị vô hiệu hóa |
| Cập nhật package pháp lý offline | Màn `Cập nhật` | `ĐÃ CÓ` |
| Cài đặt/cập nhật add-in | Bộ cài VSTO ClickOnce | `ĐÃ CÓ` cho Office x64 |

### 2.2. Nền tảng được công bố

- Windows 11 + Microsoft 365 x64 đã có bằng chứng chạy thật.
- Build `Release|x64` là cấu hình phát hành chính.
- Office x86, Office 2016/2019/2021 và Windows 10 chưa được coi là PASS nếu chưa có máy thật.
- .NET Framework tối thiểu của release: `4.8.1`.
- Dấu thập phân, phân nhóm và dấu phân cách công thức phải lấy theo Excel.

## 3. Kiến trúc dữ liệu chung

### 3.1. Workbook là hồ sơ nghiệp vụ

Workbook giữ các dữ liệu sau:

- `ProjectProfile`: mã dự án, ngày lập, ngày phê duyệt, ngày đánh giá, mốc giá,
  PriceProfile và package pháp lý đã pin.
- Vai trò sheet trong `Worksheet.CustomProperties`, khóa `TTBMVN.WorksheetRole`.
- PriceProfile hoặc portfolio gồm hai đối tượng lương.
- Estimate Workspace của Phụ lục dự toán.
- Audit trail kết quả của luồng dự toán legacy.
- Metadata package, migration và validation.

Tên tab chỉ là tên hiển thị. Việc đổi tên sheet không được làm mất vai trò nghiệp vụ.

### 3.2. Bảy vai trò sheet bắt buộc

| ID ổn định | Ý nghĩa |
|---|---|
| `ResourcePrices` | Giá vật liệu, nhân công, máy |
| `UnitRateLand` | Đơn giá trên cạn |
| `UnitRateWater` | Đơn giá dưới nước |
| `EstimateAppendix` | Phụ lục dự toán |
| `CostSummary` | Tổng hợp kinh phí |
| `NormLookupView` | View tra cứu định mức |
| `CostRuleView` | View quy tắc chi phí |

Mỗi vai trò phải được gán đúng một sheet và một sheet không được nhận hai vai trò. Workbook
thiếu/trùng role không được đi tiếp mà phải mở wizard ánh xạ.

### 3.3. Package và hồ sơ giá

- Package pháp lý là bất biến, cài song song theo phiên bản, không ghi đè “bản mới nhất”.
- Workbook đã dùng package nào thì pin package đó; không tự động nâng cấp ngầm.
- PriceProfile tách khỏi package pháp lý và có phiên bản/checksum riêng.
- Ghi đè giá phải giữ giá cũ, giá mới, lý do, nguồn, người sửa và thời gian.

## 4. Hành vi chung của giao diện

### 4.1. Form Dự toán

- Mỗi workbook có tối đa một form `Dự toán rà phá bom mìn`.
- Form chính là modeless: người dùng vẫn chọn ô, nhập công thức và sửa Excel khi form mở.
- Bấm lại nút Dự toán sẽ đưa form đã mở lên trước, không tạo bản thứ hai.
- Đóng workbook phải đóng form gắn với workbook đó.
- Các dialog chọn vùng, ánh xạ, xác nhận và hỗ trợ vẫn có thể là modal.

### 4.2. Form Hố đào

- Form Hố đào cũng mở modeless và ghi nhớ workbook đang active lúc tạo form để cập nhật
  danh sách sheet.
- Danh sách sheet tự cập nhật khi thêm, xóa, đổi tên hoặc kích hoạt lại form.
- Chọn định dạng Excel tạo tạm một sheet, mở dialog định dạng số, xóa sheet tạm và cố gắng
  phục hồi sheet/selection trước đó.

`CẦN CHỐT`: code tính/chọn vùng Hố đào hiện vẫn dùng `ActiveWorkbook` tại thời điểm bấm
chạy, không bắt buộc dùng workbook đã gắn với form. Xem quyết định `LR-01`.

## 5. Đặc tả chức năng Hố đào

### 5.1. Công thức và cận

Công thức thể tích:

```text
A1 = d1 * r1
A2 = d2 * r2
V  = (A1 + A2 + sqrt(A1 * A2)) * H / 3
H  = 3 * V / (A1 + A2 + sqrt(A1 * A2))
```

Các cận phải lớn hơn `0` và `min <= max`.

| Loại | d1 | r1 | d2 | r2 | H | V có thể tạo mặc định |
|---|---:|---:|---:|---:|---:|---:|
| 3 m | 1,7..2,3 | 1,2..1,8 | 1,1..1,5 | 0,6..1,0 | 0,1..10 | 0,128678..27,106624 |
| 5 m | 2,7..3,3 | 2,2..2,8 | 1,3..1,7 | 0,8..1,2 | 0,1..10 | 0,315516..52,072042 |

### 5.2. Đồng bộ H và V

1. Khi mở form lần đầu, `V` là phía được giữ; app tính ngược `H` từ `V` và `d/r`.
2. Nếu người dùng sửa `H min/max`, app đánh dấu `H` là nguồn và tính lại `V`.
3. Nếu người dùng sửa `V min/max`, app đánh dấu `V` là nguồn và tính lại `H`.
4. Sau đó sửa một cận `d1/r1/d2/r2` sẽ giữ phía được sửa gần nhất và tính phía còn lại.
5. Chuyển radio 3 m/5 m chỉ nạp bộ thông số của loại hố tương ứng; các trường Table Data,
   mapping, sheet và hàng xuất không được nạp lại.
6. H cho phép tối đa 4 chữ số thập phân. d/r/V cho phép tối đa 2 chữ số thập phân.
7. `SmartNumericUpDown` đổi bước tăng theo số chữ số thập phân người dùng đang nhập.

### 5.3. Số lần thử và số luồng

- `Số lần thử tối đa` mặc định `10000`, phải lớn hơn `0`.
- `Số luồng tối đa = 0` nghĩa là tự động.
- Tự động hiện dùng `Environment.ProcessorCount - 1` nếu máy có trên 2 logical processor;
  máy có 1-2 logical processor dùng toàn bộ số đó.
- ComboBox hiển thị từ `0 - Tự động` đến số logical processor của máy.

Thuật toán hiện tại:

1. Kiểm tra trung bình `Tổng V / N` nằm trong `[Vmin, Vmax]`.
2. Sinh `N` phần V khác nhau sau khi làm tròn 2 chữ số, scale về tổng đích và làm tròn
   mỗi phần 3 chữ số.
3. Với từng phần V, random d1/r1/d2/r2 đến 2 chữ số và tính H; H phải nằm trong cận,
   sau đó làm tròn H đến 4 chữ số.
4. Tổng V tính lại sau khi làm tròn phải khớp tổng đích với sai số `<= 0,001`.
5. Không được bỏ qua tín hiệu không sinh được. Hết số lần thử phải dừng toàn bộ và báo gợi ý.

`CẦN CHỐT`: yêu cầu các phần V phải đôi một khác nhau không phải yêu cầu nghiệp vụ đã được
xác nhận. Vòng lặp hiện lồng tối đa `MaxAttempts * N * MaxAttempts`, có thể chạy rất lâu.

### 5.4. Hai chế độ đầu vào

#### A. Chế độ Table Data

Được chọn khi ô `Table Data` có địa chỉ range.

- Range có thể ở một sheet khác sheet kết quả.
- Mỗi hàng nguồn đọc tổng số tín hiệu và tổng thể tích.
- Ô `TH3;V3;TH5;V5` nhận cú pháp:
  - Dự án chỉ 3 m: `A;B`.
  - Dự án cả 3 m và 5 m: `A;B;C;D`.
  - Cú pháp cũ có tên như `TH3m=A;V3m=B;...` vẫn được parser chấp nhận.
- Hàng có cả TH3 và TH5 bằng `0` được bỏ qua.
- `TH > 0` nhưng `V <= 0` là lỗi chặn.

#### B. Chế độ vùng đang chọn

Được chọn khi `Table Data` để trống.

- Người dùng phải chọn các ô tổng V trong đúng một cột.
- Ô bên trái mỗi ô tổng V là N.
- Ô rỗng, tổng V bằng `0`, N rỗng hoặc N không dương được bỏ qua.
- Kết quả mặc định ghi dưới từng ô tổng V, hoặc bắt đầu từ một cell riêng nếu run option
  legacy có `UseOutputStartCell`.

`CẦN CHỐT`: UI hiện tại ưu tiên Table Data nhưng vẫn giữ chế độ vùng chọn legacy. QA cần
xác nhận có tiếp tục hỗ trợ cả hai hay loại bỏ một chế độ để giảm sai khác hành vi.

### 5.5. Loại dự án

| Lựa chọn | Hành vi |
|---|---|
| `Chỉ có 3m` | Chỉ đọc TH3/V3 và chỉ sinh cột kết quả 3 m |
| `Cả 3m và 5m` | Đọc cả TH3/V3/TH5/V5; mỗi hàng nguồn tạo một block chung |

Với block chung, hàng tổng của 3 m và 5 m nằm cùng hàng. Chi tiết 3 m và 5 m cùng bắt đầu
ở hàng kế tiếp. Số hàng chi tiết bằng `max(TH3, TH5)`, do đó loại có ít tín hiệu hơn để
trống phần cuối block.

### 5.6. Mapping DataGridView

Các key mặc định, không cho đổi tên và không cho xóa:

```text
TH3; V3; D1_3; R1_3; D2_3; R2_3; H_3
TH5; V5; D1_5; R1_5; D2_5; R2_5; H_5
```

| Cột | Quy tắc |
|---|---|
| `Tên dữ liệu` | Key mặc định hoặc tên tùy ý của người dùng |
| `Cột trong Data` | Chỉ chữ cái cột nguồn, hoặc `key_STT` |
| `Cột trong đào đắp` | Chỉ chữ cái cột đích; để trống nghĩa là không ghi loại dữ liệu đó |
| `Định dạng` | NumberFormatLocal; double-click để chọn bằng dialog Excel |

Quy tắc ghi:

- Key chuẩn TH/V/D/R/H là dữ liệu được app sinh; chỉ cần `Cột trong đào đắp`.
- Key tùy ý có cả cột nguồn và cột đích sẽ tạo công thức link từ cell nguồn cùng hàng
  dữ liệu sang **hàng tổng** của block kết quả.
- `key_STT` ở `Cột trong Data` ghi số thứ tự block `1, 2, 3...` vào hàng tổng.
- Dòng có `Cột trong đào đắp` trống không ghi gì vào Excel.
- Header đầu tiên tại vùng xuất dùng chính `Tên dữ liệu`; kết quả bắt đầu ở hàng dưới.
- Delete trên cell chỉ xóa giá trị được chọn; Delete trên row mặc định chỉ xóa mapping,
  không xóa dòng; Backspace/xóa row tùy chỉnh xóa cả dòng.

### 5.7. Sheet và hàng xuất

| Điều khiển | Hành vi hiện tại |
|---|---|
| `New sheet` | Table Data tạo sheet trắng sau sheet nguồn; tên mặc định `HoDao`, tự thêm hậu tố nếu trùng |
| `Sheet hiện có` | Ghi vào sheet được chọn trong ComboBox |
| `Hàng mặc định` bật | Dùng số hàng nhập ở `Hàng trong đào đắp` |
| `Hàng mặc định` tắt | Khi chạy, yêu cầu chọn cell; app lấy hàng của cell đó |
| `Link ngược` | Link tổng TH/V ở hàng nguồn về hàng tổng tương ứng trên sheet kết quả |
| `Xuống dòng` | Control vẫn hiển thị nhưng code form luôn đặt `InsertRows = true` |

Khi ghi vào sheet hiện có, app luôn chèn đủ số dòng block cộng một dòng header tại hàng
bắt đầu. Khi tạo sheet mới từ Table Data, app không cần chèn vì sheet đang trống.

`CẦN CHỐT`: phải chặn chọn sheet kết quả trùng sheet nguồn hoặc vùng xuất giao vùng nguồn.
Code hiện chưa chặn và việc chèn dòng có thể làm sai `SourceRow` đã lập kế hoạch.

### 5.8. Ghi kết quả và link ngược

Mỗi block có:

1. Hàng tổng: tổng TH, tổng V, STT và các cột phụ được link.
2. Các hàng chi tiết: TH bằng `1`; V là công thức nếu đủ D1/R1/D2/R2/H, nếu thiếu cột
   thành phần thì V được ghi giá trị; các cột d/r/H là số.
3. Tổng TH và V phải bằng dữ liệu nguồn trong sai số cho phép.
4. Link ngược ghi công thức vào chính các cell TH/V nguồn, thay thế giá trị ban đầu.

### 5.9. Preview, lỗi và khôi phục

Preview phải hiển thị nguồn, sheet đích, hàng bắt đầu, số block, số dòng, tổng TH/V,
chế độ dự án và link ngược. Người dùng phải xác nhận trước khi ghi.

- Sai cận hoặc tổng ngoài khoảng khả thi: dừng trước khi ghi.
- Không sinh được sau N lần: dừng toàn bộ, không bỏ qua dòng.
- Lỗi khi đang tạo sheet mới: cố gắng xóa sheet vừa tạo.
- Lỗi khi ghi sheet hiện có: hiện **không có rollback toàn vùng**; dòng đã chèn và dữ liệu
  đã ghi có thể còn lại.
- `ExcelWriteContext` chỉ bảo đảm phục hồi ScreenUpdating, Calculation, Events và
  DisplayAlerts; không phải transaction dữ liệu.

### 5.10. Lưu, Export và Import

- Mở form tự nạp lần cài đặt gần nhất từ user settings.
- Lưu riêng bộ thông số 3 m và 5 m.
- Export dùng `.daodatsettings`, schema JSON 2 và nhớ thư mục dialog gần nhất.
- Import sao lưu setting hiện tại vào `%LOCALAPPDATA%\TTBMVNExcelTools\SettingsBackups`.
- Export/Import có thông số, số lần thử, số luồng, loại dự án, mapping và định dạng.
- Không Export/Import các trường phụ thuộc workbook: Table Data, hàng xuất, A;B;C;D,
  chế độ/tên sheet, sheet hiện có và link ngược. Import giữ nguyên các trường này của form.

## 6. Đặc tả chức năng Dự toán RPBM

### 6.1. Luồng nghiệp vụ mục tiêu

```text
ProjectProfile + RegulationPackage + PriceProfile
                         |
                         v
Quét Phụ lục DT -> Gắn định mức/ngữ cảnh
                         |
                         v
VL-NC-M theo HLNS/KHLNS -> ĐG Cạn/Nước theo HLNS/KHLNS
                         |
                         v
Link đơn giá/thành tiền về Phụ lục DT -> Tổng hợp kinh phí
                         |
                         v
Truy vết + Kiểm tra + Chuyển package + Gói hỗ trợ
```

Phần từ Phụ lục DT đến các sheet giá/đơn giá mới đã có. Phần nối từ Workspace mới sang
Tổng hợp kinh phí, Truy vết, Kiểm tra và Chuyển package chưa hoàn chỉnh.

### 6.2. Mở hồ sơ và thiết lập dự án

Khi chưa có form đang mở, nút Dự toán mở wizard `Thiết lập dự án dự toán` trước:

| Trường | Quy tắc hiện tại |
|---|---|
| Mã dự án | Bắt buộc; mặc định tên file không phần mở rộng |
| Ngày lập | Mặc định ngày hiện tại hoặc giá trị đã lưu |
| Ngày phê duyệt | Chỉ dùng khi tick `Đã phê duyệt` |
| Ngày đánh giá | Dùng để resolver package theo hiệu lực/chuyển tiếp |
| Mốc giá | Ngày của PriceProfile |
| Hồ sơ giá | ID PriceProfile |
| Căn cứ | Package được resolver tự chọn và hiển thị lý do |
| Ánh xạ sheet | Bắt buộc đủ bảy role |

Chỉ bấm `Lưu và mở dự toán` mới ghi workbook. Commit profile + role nằm trong cùng luồng
rollback logic; hủy wizard không được ghi.

`CẦN CHỐT`: wizard mở lại mỗi lần form đã đóng rồi mở lại, kể cả workbook đã hợp lệ.
Nên xác nhận có cần chế độ “mở ngay” và nút sửa thiết lập riêng hay không.

### 6.3. Menu chính

| Mục | Chức năng hiện tại | Trạng thái |
|---|---|---|
| `Thông tin chung` | Màn legacy gắn tag cell và ComboBox | `LEGACY` |
| `Phụ lục DT` | Quét range, gắn định mức, tính thử, sinh đơn giá | `ĐÃ CÓ` |
| `Bảng giá` | Quản lý PriceProfile HLNS/KHLNS | `ĐÃ CÓ`, cần chốt nghiệp vụ giá |
| `Đơn giá` | Danh sách đơn giá duy nhất của Workspace, read-only, sinh lại sheet | `ĐÃ CÓ` |
| `Tổng hợp KP` | Tính và ghi THKP-TC theo layout legacy | `MỘT PHẦN` |
| `Tra cứu định mức` | Tìm kiếm catalog đã pin | `ĐÃ CÓ` |
| `Chi phí` | Calculator độc lập K1-K6, không ghi workbook | `MỘT PHẦN`/trùng vai trò |
| `Truy vết` | Đọc audit theo ô đang chọn | `ĐÃ CÓ` cho luồng legacy |
| `Kiểm tra` | Quét role/profile/giá/công thức/tổng | `ĐÃ CÓ` cho luồng legacy |
| `Chuyển gói` | Preview, backup, apply/rollback package | `ĐÃ CÓ` cho luồng legacy |
| `Cập nhật` | Cài package `.ttbupdate` offline | `ĐÃ CÓ` |
| `Hỗ trợ` | Thông tin app/license và gói chẩn đoán | `ĐÃ CÓ` |

### 6.4. Thông tin chung legacy

Màn hiện có các ComboBox: loại dự toán, loại địa hình, khối lượng hủy nổ, loại công trình
lán trại, loại công trình giám sát; sáu nút để gắn tag cell và hai ô giá trị.

Hành vi thực tế:

- Nút tag cho phép chọn một cell và đặt/di chuyển shape tag.
- Chỉ thay đổi `Khối lượng hủy nổ` đang ghi giá trị vào cell tag `KhoiLuongHuyNo`.
- Handler cho loại dự toán, địa hình, lán trại, giám sát và cập nhật giá trị đã bị comment.
- Form không nhận workbook từ shell mà dùng Excel active toàn cục.

Màn này chưa đủ điều kiện nghiệm thu thương mại. QA phải xem nó là chức năng chưa hoàn thiện,
không phải nguồn thiết lập ProjectProfile mới.

### 6.5. Bảng giá và hai đối tượng lương

Mỗi PriceProfile có:

- Mã hồ sơ, phiên bản, tên hồ sơ, địa điểm, ngày giá.
- Đối tượng `Hưởng lương NSNN` hoặc `Không hưởng lương NSNN`.
- Danh sách mã giá, loại VL/NC/M, tên, đơn vị, giá gốc, giá áp dụng, nguồn, alias,
  tên lookup legacy và override.
- Checksum deterministic.

Luồng hiện tại:

1. Nếu workbook có portfolio, nạp cả hai profile.
2. Nếu chỉ có profile cũ, nạp profile đó.
3. Nếu chưa có, import bảng `VL-NC-M` legacy thành profile KHLNS.
4. Chuyển sang đối tượng chưa có sẽ hỏi tạo bản nháp bằng cách clone toàn bộ profile hiện tại.
5. Override bật theo từng dòng; phải có giá áp dụng và lý do/nguồn theo validator.
6. `Lưu hồ sơ` ghi store máy, portfolio workbook và pin profile đang active.
7. Khi lưu profile KHLNS, app cập nhật giá vật liệu trực tiếp về bảng legacy; profile HLNS
   không thực hiện bước này.
8. Import/Export profile dùng file riêng của PriceProfile.

`CẦN CHỐT`: chưa có workflow bắt buộc người dùng chọn nguồn giá nhân công/máy theo văn bản,
hệ số hoặc nhập tay có lý do. Clone HLNS hiện sao chép cả giá NC/M của KHLNS và chỉ nhắc
người dùng rà soát, không áp quy tắc tự động.

### 6.6. Quét vùng Phụ lục dự toán

Người dùng bấm `Quét vùng Excel`, chọn một vùng liên tục và ánh xạ:

| Trường | Bắt buộc | Hành vi |
|---|---|---|
| Số hiệu/Mã công tác | Không | Đọc text |
| Nội dung công tác | Có | Đọc text |
| Đơn vị | Không | Đọc text |
| Khối lượng | Có | Giữ `FormulaLocal` nếu là công thức, đồng thời đọc giá trị hiện tại |
| Khối lượng nghiệm thu | Không | Giữ công thức; lỗi/text hiện được coi là 0 |
| Đơn giá VL/NC/M | Không | Là cột app sẽ ghi công thức link |
| Thành tiền VL/NC/M | Không | Là cột app sẽ ghi `Khối lượng * Đơn giá` |

- Có tùy chọn dòng đầu là tiêu đề, mặc định bật.
- Mặc định môi trường `Cạn`, đối tượng `KHLNS`.
- Dialog tự gợi ý cột theo tên header không dấu; người dùng có thể đổi thủ công.
- Chỉ nhận vùng thuộc workbook đã gắn với form.
- Bỏ qua hàng hoàn toàn rỗng ở các trường code/description/unit/quantity.
- Re-scan cùng nguồn cố gắng giữ RowId và cấu hình định mức theo hàng đã binding.
- Workspace lưu trong workbook; đổi tên sheet và chèn dòng được theo dõi bằng source key
  và defined name ẩn cho từng dòng.

### 6.7. Gắn định mức cho từng hàng

Panel chi tiết cho phép chọn:

- Định mức.
- Mã chi tiết/biến thể.
- Môi trường Cạn/Nước.
- Đối tượng HLNS/KHLNS.
- Danh sách điều kiện/hệ số.
- Binding mã hao phí logic sang mã giá cụ thể và lý do chọn.

`Gắn định mức` cập nhật hàng đang chọn. `Dòng văn bản` xóa norm/variant/điều kiện/binding;
hàng đó không tham gia tính. Grid chính là read-only, người dùng không sửa mô tả/khối lượng
trên app mà sửa ở Excel rồi quét lại.

### 6.8. Quy tắc dùng chung hoặc tách đơn giá

Hai hàng dùng chung một đơn giá khi và chỉ khi toàn bộ khóa sau giống nhau:

1. Package ID, version và checksum.
2. Norm key.
3. Mã chi tiết/variant.
4. Cạn hoặc Nước.
5. HLNS hoặc KHLNS.
6. PriceProfile ID, version và checksum của đúng đối tượng lương.
7. Tập điều kiện/hệ số.
8. Tập binding `mã hao phí -> mã giá -> lý do`.

Mã công tác hiển thị, mô tả, đơn vị, khối lượng và vị trí hàng không tham gia khóa dùng chung.
Vì vậy:

- Cùng `DM-001.1` và cùng ngữ cảnh: một bảng đơn giá.
- `DM-001.1` khác `DM-001.2`: hai bảng.
- Cùng `.1` nhưng một dòng HLNS và một dòng KHLNS: hai bảng.
- Cùng `.1` nhưng khác mã giá thiết bị logic: hai bảng.

Rate ID có dạng `DG-` cộng 16 ký tự đầu của SHA-256 khóa canonical.

### 6.9. Tính đơn giá

- Hao phí lấy từ NormCatalog của package đã pin và variant đã chọn.
- Điều kiện/hệ số điều chỉnh hao phí theo catalog.
- Giá lấy từ đúng PriceProfile HLNS/KHLNS.
- Nguồn lực logic bắt buộc có binding giá; thiếu binding hoặc thiếu giá là lỗi.
- Kiểm tra loại tài nguyên và đơn vị giữa định mức với mã giá.
- Thành tiền nguồn lực = hao phí * giá áp dụng, dùng decimal và không làm tròn trung gian.
- Vật liệu khác theo tỷ lệ phần trăm tính trên tổng vật liệu thường.
- Tổng đơn giá = VL + NC + M.

Màn `Đơn giá` chỉ đọc preview Workspace; không cho sửa trực tiếp. Chọn một đơn giá hiển thị
các hao phí, mã giá, nguồn, giá và override.

### 6.10. Sinh sheet giá và đơn giá

App chỉ tạo các sheet thực sự được dùng:

| Nhóm | Tên đề xuất |
|---|---|
| Giá HLNS | `VL-NC-M - HLNS` |
| Giá KHLNS | `VL-NC-M - KHLNS` |
| Đơn giá cạn HLNS | `ĐG Cạn - HLNS` |
| Đơn giá nước HLNS | `ĐG Nước - HLNS` |
| Đơn giá cạn KHLNS | `ĐG Cạn - KHLNS` |
| Đơn giá nước KHLNS | `ĐG Nước - KHLNS` |

Sheet sinh được nhận diện bằng custom property `TTBMVN.GeneratedEstimateRole`, nên đổi tên
tab rồi chạy lại vẫn dùng đúng sheet. Nếu chưa có, app tạo sau sheet cuối và đặt tên không trùng.

#### Sheet giá

- Chỉ chứa các mã giá đang được các đơn giá dùng.
- Cột: Mã, Loại, Tên dữ liệu, ĐVT, Giá gốc, Giá áp dụng, Nguồn, Lý do ghi đè.
- Tạo workbook defined name cho từng giá áp dụng.

#### Sheet đơn giá

- Mỗi đơn giá là một block có Rate ID, norm/variant, tên và số công tác dùng chung.
- Dòng hao phí gồm loại, mã, tên, đơn vị, hao phí, công thức link giá, thành tiền.
- Cuối block có tổng VL, NC, M và tổng đơn giá.
- Tạo workbook defined name cho bốn tổng của mỗi Rate ID.

#### Link về Phụ lục DT

- Đơn giá VL/NC/M được ghi bằng công thức tham chiếu defined name.
- Thành tiền VL/NC/M = cột khối lượng đã map * cột đơn giá tương ứng.
- Hàng văn bản không bị ghi.
- Chỉ các cột output đã map bị thay đổi; các ô khác trong bounding range được clone giữ lại.

Chạy lại không tạo sheet trùng, nhưng nội dung `A:H` của sheet app quản lý bị ghi lại. Sheet
generated cũ không còn được dùng và defined name cũ hiện chưa được tự dọn.

### 6.11. Tổng hợp kinh phí

Màn hiện tại đọc VL/NC/M từ các ô cố định `I27:K27` trên sheet role `EstimateAppendix`.
Các input gồm biểu mẫu, địa hình, diện tích, loại dự án, loại công trình, khối lượng hủy nổ,
TL, VAT và checkbox K1..K6.

Quy tắc tính hiện tại:

```text
VL, NC, M = làm tròn VND, AwayFromZero
T  = VL + NC + M
C  = 40% * NC
TL = (T + C) * tỷ lệ TL, chỉ khi biểu mẫu 04
Z  = T + C + TL
K2 tính trên T
K1, K3, K4, K5, K6 tính trên Z
K  = tổng các thành phần được tick
Q  = Z + K
VAT base = Q - K3 - K4, chỉ khi biểu mẫu 04
H  = Q + VAT
Kết quả cuối = làm tròn H đến 1.000 đồng, AwayFromZero
```

Bốn biểu mẫu:

1. Điều tra, khảo sát.
2. Dự án độc lập vốn Nhà nước.
3. Hạng mục vốn Nhà nước.
4. Nguồn vốn khác.

Chỉ biểu mẫu 04 hiện bật TL và VAT. Người dùng có thể bật/tắt từng K1..K6. `Ghi vào THKP-TC`
ghi cố định vùng `D10:E27` và `A28`, dùng công thức Excel và tiền bằng chữ.

`CẦN CHỐT`: đây chưa phải cơ cấu linh hoạt “có/không VAT, có/không giám sát” theo yêu cầu
người dùng. Cần phê duyệt matrix biểu mẫu - thành phần - cơ sở tính - thuế trước khi coi đúng.

### 6.12. Màn Chi phí

Đây là calculator độc lập với dữ liệu mẫu ban đầu VL `300 triệu`, NC `200 triệu`, máy
`100 triệu`. Người dùng chọn K1..K6 và có thể override từng khoản kèm lý do. Màn chỉ tính
và hiển thị; không lưu cấu hình, không ghi THKP và không nối trực tiếp Workspace.

`CẦN CHỐT`: giữ màn này như công cụ thử, gộp vào Tổng hợp KP, hay loại khỏi bản thương mại.

### 6.13. Tra cứu định mức

- Tìm exact/fuzzy, có hỗ trợ chuỗi không dấu.
- Lọc theo môi trường, độ sâu và loại tài nguyên.
- Hiển thị norm, biến thể, bảng hao phí, điều kiện/hệ số và locator văn bản nguồn.
- Chỉ dùng package đã pin trong workbook.
- Không sửa catalog từ màn tra cứu.

### 6.14. Truy vết

- Người dùng chọn một ô kết quả rồi bấm `Đọc ô đang chọn`.
- Hiển thị vị trí, package, PriceProfile, norm/variant và các nguồn pháp lý/giá.
- Audit bền qua đổi tên sheet, save và reopen cho luồng legacy `Gia DT TC`/`THKP-TC`.

`CẦN CHỐT`: writer Workspace mới không tạo audit entry tương ứng cho sheet giá, đơn giá và
công thức link mới; truy vết end-to-end của luồng mới hiện chưa được đảm bảo.

### 6.15. Kiểm tra workbook

Scanner hiện kiểm tra role, profile/package, thiếu giá, sai tổng, công thức bị sửa, audit stale
và defined name app. Có điều hướng đến ô lỗi. Các defined name hỏng kế thừa được gom thành
cảnh báo thay vì chặn.

`CẦN CHỐT`: scanner hiện chưa có rule đầy đủ cho Estimate Workspace, generated role, Rate ID,
portfolio hai đối tượng, link Phụ lục mới và generated sheet thừa.

### 6.16. Chuyển package

- Preview package nguồn/đích và diff module/record/kết quả.
- Preview không ghi workbook.
- Apply tạo backup, ghi profile/package/kết quả, validate, save; lỗi giữa chừng phải rollback.
- Không tự động nâng package khi chỉ cài package mới.

`CẦN CHỐT`: migration hiện dựa trên plan/kết quả legacy. Workspace và generated sheet mới
không được tái tính/ghi lại như một phần bắt buộc của migration.

### 6.17. Cập nhật package

- Chọn file `.ttbupdate`, kiểm tra manifest, hash, chữ ký RSA, minimum app version,
  downgrade và xung đột version.
- Có thể cài và đặt package mặc định trên máy.
- Đặt package mặc định không tự đổi workbook đang mở.
- Nút `Kiểm tra online` bị vô hiệu hóa; hiện chỉ hỗ trợ offline.

## 7. Hỗ trợ, log và kích hoạt

### 7.1. Log runtime

- File: `%LOCALAPPDATA%\TTBMVNExcelTools\Logs\runtime.log`.
- Mỗi lỗi nghiệp vụ quan trọng ghi correlation ID, UTC, task/phase, workbook token,
  worksheet role, package ID/version/checksum, exception type và HResult.
- Dữ liệu nhạy cảm phải được token hóa hoặc loại bỏ khỏi snapshot hỗ trợ.

### 7.2. Gói chẩn đoán

Gói Dự toán có thể gồm:

- `app-info.txt`.
- `workbook-info.txt` đã sanitize.
- `daodat-settings.json` không chứa field phụ thuộc workbook.
- `runtime.log` đã sanitize.
- `project-profile.txt` dùng token thay ID nhạy cảm.
- `package-manifest.ttbmanifest`.
- Báo cáo validation/package liên quan.
- `manifest.sha256` kiểm tra toàn vẹn.

Không được chứa đường dẫn workbook, tên sheet, địa chỉ cell, Machine ID thô, license key,
Project ID thô, PriceProfile ID thô hoặc nội dung override nhạy cảm.

### 7.3. Kích hoạt offline

- Form hiển thị Product, Version, Machine ID, trạng thái và hạn dùng.
- Key gắn với Machine ID và ngày hết hạn; ngày hết hạn nằm trong payload đã ký, người dùng
  không thể sửa chuỗi ngày mà vẫn giữ chữ ký hợp lệ.
- Có 30 ngày dùng thử tính từ file `trial.dat` trên máy.
- Online activation bị tắt.

`CẦN CHỐT` nghiêm trọng:

1. Trạng thái trial/license hiện chỉ hiển thị và log; các nút nghiệp vụ không kiểm tra quyền,
   nên hết hạn vẫn có thể chạy chức năng.
2. Bí mật ký key đối xứng nằm trong binary add-in và tool sinh key. Đây không phải mô hình
   an toàn để thương mại hóa; người phân tích binary có thể tự sinh key.
3. Chưa có chính sách đổi máy, thu hồi key, chống lùi ngày hệ thống hoặc cấp lại license.

## 8. Cài đặt và phát hành

- Người dùng phải đóng Excel trước khi cài/gỡ thật.
- `Cai-dat-TTBMVN.cmd` gọi PowerShell installer; không mở `.vsto` trực tiếp.
- Installer xử lý registration Visual Studio `|vstolocal`, backup key registry và gỡ đúng
  ClickOnce subscription cũ trước khi cài.
- Bản hiện tại ghi `LoadBehavior=3` và đăng ký theo user.
- Artifact x64 có setup/manifest ký bằng certificate PILOT; chưa timestamp.
- Build từ Visual Studio có thể đăng ký lại manifest `bin\x64\Release|vstolocal`; sau build
  trên máy dev phải cài lại artifact nếu muốn test đúng bản publish.

## 9. Yêu cầu phi chức năng

### 9.1. Tính toàn vẹn Excel

1. Mọi luồng ghi phải phục hồi ScreenUpdating, Calculation, EnableEvents và DisplayAlerts.
2. Không được ghi nhầm workbook khi nhiều workbook cùng mở.
3. Không được làm mất công thức/format ngoài vùng app sở hữu.
4. Lỗi giữa chừng phải rollback hoặc có checkpoint/phục hồi rõ ràng.
5. Mọi sheet app tạo phải có identity bền qua đổi tên.

### 9.2. Hiệu năng

- Ghi Excel dùng batch, tránh ghi từng cell.
- Không chạy Excel Interop từ worker thread.
- Tính Hố đào phải có thời hạn kết thúc dự đoán được và cho phép hủy.
- Auto worker dựa trên logical processor nhưng không được làm treo Excel/UI hoặc chiếm toàn bộ
  máy trong thời gian dài.

### 9.3. Locale

- NumericUpDown và preview dùng culture lấy theo Excel.
- Formula/NumberFormat phải chạy với cả dấu thập phân `,` và `.`.
- Import/export nội bộ dùng invariant hoặc schema rõ ràng, không phụ thuộc locale máy tạo file.

### 9.4. Bảo mật và riêng tư

- Package cập nhật phải xác minh chữ ký trước khi cài.
- Không ghi license key hoặc Machine ID thô vào log/support package.
- Không nhúng secret có khả năng sinh license hợp lệ trong client thương mại.
- Không xuất nội dung workbook vào support package nếu chưa có đồng ý rõ ràng.

## 10. Danh sách quyết định logic cần chốt

| ID | Mức | Vấn đề/hành vi hiện tại | Khuyến nghị ban đầu |
|---|---|---|---|
| LR-01 | P0 | Form Hố đào gắn danh sách với workbook A nhưng khi chạy dùng ActiveWorkbook B | Khóa form với workbook A; nếu A không active thì hỏi chuyển hoặc chặn |
| LR-02 | P0 | Sheet kết quả Hố đào có thể trùng sheet/range nguồn | Chặn tuyệt đối giao nhau trước preview |
| LR-03 | P0 | Ghi Hố đào vào sheet hiện có không rollback khi COM lỗi | Dùng transaction/snapshot vùng và rollback cả dòng chèn |
| LR-04 | P0 | Vòng lặp random lồng rất lớn, không có nút hủy/time budget | Thiết kế solver có giới hạn tổng và cancellation |
| LR-05 | P1 | Thuật toán bắt N giá trị V phải khác nhau | Bỏ ràng buộc nếu nghiệp vụ không yêu cầu |
| LR-06 | P1 | Checkbox `Xuống dòng` hiển thị nhưng luôn bị ép `true` | Ẩn/bỏ checkbox hoặc giao chức năng mới rõ ràng |
| LR-07 | P1 | New sheet của Table Data là sheet trắng; chế độ vùng chọn lại copy sheet nguồn | Chọn một quy tắc nhất quán và ghi rõ |
| LR-08 | P1 | Link ngược thay giá trị nguồn bằng công thức, không có undo/checkpoint | Preview nêu cell bị thay; tạo transaction/undo |
| LR-09 | P0 | Màn Thông tin chung legacy chỉ hoạt động một phần và dùng ActiveWorkbook | Thay bằng ProjectProfile mới hoặc loại khỏi menu thương mại |
| LR-10 | P0 | License hết hạn không khóa chức năng | Thêm policy gate tập trung tại Ribbon/service |
| LR-11 | P0 | Secret sinh license nằm trong client | Chuyển sang chữ ký bất đối xứng; client chỉ giữ public key |
| LR-12 | P0 | Tổng hợp KP luôn đọc `I27:K27`, không tổng hợp generic từ Workspace | Tính trực tiếp từ Workspace hoặc cấu hình output total binding |
| LR-13 | P0 | THKP writer cố định `D10:E27;A28` | Định nghĩa template/layout version hoặc mapping output |
| LR-14 | P0 | Workspace mới chưa tạo audit entry | Bổ sung audit cho giá, đơn giá, link và tổng hợp |
| LR-15 | P0 | Validation chưa kiểm tra đầy đủ Workspace/generated sheets | Thêm rule và reconciliation end-to-end |
| LR-16 | P0 | Migration package chưa tái sinh Workspace/generated sheets | Migration phải preview và regenerate toàn bộ luồng mới |
| LR-17 | P0 | Chỉ biểu mẫu 04 tính TL/VAT | Chủ nghiệp vụ chốt matrix 4 biểu mẫu |
| LR-18 | P0 | K1..K6 mặc định đều bật; người dùng tự bỏ tùy ý | Chốt component bắt buộc/tùy chọn theo biểu mẫu |
| LR-19 | P1 | Chưa có lựa chọn riêng “có giám sát/không giám sát” | Đưa thành component/template rõ ràng, không dựa checkbox K mơ hồ |
| LR-20 | P0 | Profile HLNS được clone từ KHLNS, chưa có nguồn/hệ số bắt buộc | Wizard nguồn giá NC/M, nhập tay bắt buộc lý do |
| LR-21 | P1 | KHLNS save mới cập nhật vật liệu vào VL-NC-M legacy | Chốt dữ liệu legacy còn là output hay phải bỏ đồng bộ |
| LR-22 | P1 | Màn Chi phí và Tổng hợp KP trùng logic nhưng khác khả năng ghi/override | Gộp thành một luồng hoặc ghi rõ calculator chỉ tham khảo |
| LR-23 | P1 | Dòng chưa gắn định mức mặc nhiên là văn bản và preview có thể vẫn hợp lệ | Yêu cầu người dùng đánh dấu văn bản chủ động hoặc cảnh báo hàng có KL |
| LR-24 | P1 | Generated sheet cũ và defined name cũ không tự dọn | Có garbage collection với preview/xác nhận |
| LR-25 | P1 | Chạy lại generated sheet ghi đè A:H và format | Xác nhận sheet app sở hữu hoàn toàn; khóa/vạch vùng rõ ràng |
| LR-26 | P1 | Wizard Project Setup luôn mở lại sau khi đóng form | Nếu hồ sơ hợp lệ, mở thẳng; cung cấp nút `Thiết lập dự án` riêng |
| LR-27 | P1 | Accepted quantity được quét/lưu nhưng không dùng khi tính thành tiền | Chốt mục đích: dự toán, nghiệm thu hay cả hai chế độ |

## 11. Ma trận kiểm thử QA tối thiểu

### 11.1. Global/workbook

| Test ID | Ca kiểm thử | Kết quả bắt buộc |
|---|---|---|
| GBL-01 | Mở hai workbook, mở một form Dự toán cho mỗi file | Hai form giữ đúng dữ liệu, không ghi chéo |
| GBL-02 | Đổi tên cả bảy sheet khi form mở | Role còn đúng, ComboBox cập nhật, field đang nhập không reset |
| GBL-03 | Xóa một sheet role | Validation chặn và yêu cầu ánh xạ lại |
| GBL-04 | Thêm/xóa/rename sheet liên tục | Không NRE/COM error, không chọn toàn bộ grid ngoài ý muốn |
| GBL-05 | Đóng workbook có form modeless | Chỉ form của workbook đó đóng |
| GBL-06 | Thử locale Excel `,` và `.` | Số, format và formula cùng đúng |

### 11.2. Hố đào

| Test ID | Ca kiểm thử | Kết quả bắt buộc |
|---|---|---|
| HD-01 | H mặc định, sửa d/r | Giữ V và tính lại H |
| HD-02 | Sửa H rồi sửa d/r | Giữ H và tính lại V |
| HD-03 | Sửa V rồi sửa d/r | Giữ V và tính lại H |
| HD-04 | Chuyển 3 m/5 m khi đang nhập Table Data/grid | Chỉ thông số đổi; Table Data/grid/sheet/hàng giữ nguyên |
| HD-05 | Chuyển loại dự án | Không reset Table Data, hàng, grid, format hoặc sheet |
| HD-06 | TH3/V3 trung bình dưới Vmin | Preview chặn, gợi ý giảm cận/tăng N |
| HD-07 | TH3/V3 trung bình trên Vmax | Preview chặn, gợi ý tăng cận/giảm N |
| HD-08 | Hết MaxAttempts | Không ghi dòng nào, báo đủ N/tổng/cận/gợi ý |
| HD-09 | Cả 3 m/5 m có N khác nhau | Một block chung, số detail = max(N3,N5) |
| HD-10 | Cột output của một key để trống | Không đụng cột đó |
| HD-11 | Custom mapping và `key_STT` | Link đúng hàng tổng; STT tăng theo block |
| HD-12 | Chọn format custom chưa có trong file khác | Excel nhận format và ghi đúng |
| HD-13 | Link ngược trên new sheet | Cell TH/V nguồn link đúng hàng tổng từng block |
| HD-14 | Link ngược trên existing sheet khác nguồn | Link đúng dù khoảng cách block khác N |
| HD-15 | Output sheet trùng source | Phải chặn theo quyết định LR-02 |
| HD-16 | COM fault sau khi chèn dòng | Workbook được rollback theo quyết định LR-03 |
| HD-17 | Mở hai workbook rồi đổi active trước khi chạy | Không ghi nhầm workbook |
| HD-18 | Export/import | Không đổi Table Data, hàng, A;B;C;D, sheet mode/link back |
| HD-19 | Đóng/mở Excel | Nạp đúng setting lần cuối |
| HD-20 | Auto workers trên máy 2/4/16 logical processor | Giới hạn hợp lý, UI không treo vô hạn |

### 11.3. Project Setup và role

| Test ID | Ca kiểm thử | Kết quả bắt buộc |
|---|---|---|
| PS-01 | Workbook chưa tag | Wizard gợi ý nhưng chỉ lưu khi xác nhận |
| PS-02 | Một sheet gán hai role | Không cho lưu |
| PS-03 | Hai sheet cùng một role | Không cho mở nghiệp vụ |
| PS-04 | Hủy wizard | Workbook fingerprint không đổi |
| PS-05 | Lỗi giữa profile và role save | Khôi phục cả profile và role |
| PS-06 | Đổi tên/save/reopen | Resolver tìm đúng role |

### 11.4. Bảng giá

| Test ID | Ca kiểm thử | Kết quả bắt buộc |
|---|---|---|
| PRICE-01 | Import legacy lần đầu | Tạo KHLNS đúng mã/giá/nguồn |
| PRICE-02 | Tạo HLNS từ KHLNS | Có cảnh báo phải rà soát; kết quả theo quyết định LR-20 |
| PRICE-03 | Override không lý do | Không cho lưu |
| PRICE-04 | Hai profile cùng audience | Validation chặn |
| PRICE-05 | Save/reopen | Hai profile và checksum giữ nguyên |
| PRICE-06 | Import/export profile | Round-trip không đổi canonical checksum |
| PRICE-07 | Thiếu mã giá đang dùng | Preview đơn giá chỉ đúng hàng/mã lỗi |

### 11.5. Phụ lục và đơn giá

| Test ID | Ca kiểm thử | Kết quả bắt buộc |
|---|---|---|
| EST-01 | Quét bảng một header | Đọc đúng dòng/công thức/giá trị |
| EST-02 | Quét bảng hai tầng header | Bỏ header phụ, không tạo công tác giả |
| EST-03 | Công thức KL phức tạp | Giữ FormulaLocal và giá trị Excel hiện tại |
| EST-04 | Dòng text không norm | Không tính và không ghi output |
| EST-05 | Hàng có KL nhưng chưa norm | Hành vi theo quyết định LR-23 |
| EST-06 | Cùng norm/variant/ngữ cảnh | Một Rate ID, usage count đúng |
| EST-07 | Khác variant `.1/.2` | Hai Rate ID và hai block |
| EST-08 | Khác HLNS/KHLNS | Tách profile, sheet giá và sheet đơn giá |
| EST-09 | Khác cạn/nước | Tách sheet đơn giá |
| EST-10 | Khác condition/binding | Tách Rate ID |
| EST-11 | Logical resource không binding | Preview lỗi, không sinh sheet |
| EST-12 | Chạy writer hai lần | Không tạo sheet trùng, link vẫn đúng |
| EST-13 | Rename generated sheets | Chạy lại dùng đúng sheet qua custom role |
| EST-14 | Chèn dòng trong vùng nguồn | Binding row đi theo dòng thật |
| EST-15 | Save/reopen | Workspace, profile, Rate ID và công thức còn đúng |
| EST-16 | Xóa một công tác rồi regenerate | Xử lý sheet/name thừa theo LR-24 |
| EST-17 | Sửa giá override rồi regenerate | Đơn giá và thành tiền cập nhật đúng |
| EST-18 | Mapping không có cột đơn giá nhưng có thành tiền | Formula dùng defined name trực tiếp |
| EST-19 | Chọn accepted quantity mode | Kết quả theo quyết định LR-27 |

### 11.6. Tổng hợp, audit, validation và migration

| Test ID | Ca kiểm thử | Kết quả bắt buộc |
|---|---|---|
| SUM-01 | Đối chiếu T/C/TL/Z/K/Q/VAT/H | Khớp calculator và công thức Excel |
| SUM-02 | Bốn biểu mẫu | Thành phần/thuế đúng matrix đã phê duyệt |
| SUM-03 | Bật/tắt K1..K6 | Chỉ component hợp lệ được phép bỏ |
| SUM-04 | K3 min/max, K5 nội suy, K6 biên 1000 kg | Đúng package và báo biên chưa quy định |
| SUM-05 | Workspace bố trí khác legacy | Tổng hợp vẫn lấy đúng VL/NC/M, không phụ thuộc I27:K27 |
| SUM-06 | Ghi THKP layout khác | Hành vi theo template/mapping đã duyệt |
| AUD-01 | Chọn cell giá/đơn giá/link mới | Truy được package/norm/profile/source |
| VAL-01 | Sửa một công thức generated | Scanner chỉ ra đúng cell/Rate ID |
| VAL-02 | Xóa generated sheet hoặc name | Scanner báo đúng missing dependency |
| MIG-01 | Chuyển package có Workspace | Preview toàn bộ Rate ID và tổng trước/sau |
| MIG-02 | Fault từng phase migration | Backup/rollback giữ nguyên workspace và generated output |

### 11.7. License, support và cài đặt

| Test ID | Ca kiểm thử | Kết quả bắt buộc |
|---|---|---|
| LIC-01 | Key đúng máy/còn hạn | Activated, hiện đúng hạn |
| LIC-02 | Key máy khác | Từ chối |
| LIC-03 | Sửa ngày/ký tự trong key | Từ chối |
| LIC-04 | Hết trial/license | Chức năng bị khóa theo policy đã duyệt |
| LIC-05 | Lùi ngày hệ thống | Hành vi theo policy chống rollback |
| SUP-01 | Tạo gói hỗ trợ | Manifest hash đúng, đủ file quy định |
| SUP-02 | Quét privacy | Không có Machine ID/key/path/sheet/cell/ID nhạy cảm |
| INS-01 | Máy chưa cài | Cài 1.0.4 thành công, LoadBehavior 3 |
| INS-02 | Có bản ClickOnce cũ | Gỡ đúng subscription rồi cài mới |
| INS-03 | Có `vstolocal` stale | Backup registry, không lỗi -401 |
| INS-04 | Excel đang mở | Preflight chặn mutation |
| INS-05 | Gỡ cài đặt | Xóa đúng customization, không xóa dữ liệu người dùng ngoài policy |

## 12. Tiêu chí chấp nhận bản thương mại

Bản thương mại chưa được coi là đạt nếu còn một trong các điều kiện sau:

1. Có khả năng ghi nhầm workbook hoặc ghi chồng vùng nguồn.
2. Lỗi giữa chừng để workbook ở trạng thái chèn/ghi một phần mà không có phục hồi.
3. License hết hạn vẫn chạy được hoặc client chứa khóa bí mật có thể sinh license.
4. Matrix biểu mẫu/thuế/giám sát/K1..K6 chưa được chủ nghiệp vụ ký duyệt.
5. Workspace mới chưa nối với Tổng hợp KP, Truy vết, Validation và Migration.
6. Giá HLNS/KHLNS, nhân công và máy chưa có workflow nguồn/hệ số/override được duyệt.
7. Chưa chạy ma trận thật trên các Office/Windows nằm trong phạm vi bán.
8. Artifact chưa có certificate phát hành chính thức và timestamp phù hợp.

## 13. Truy xuất nguồn và bằng chứng hiện có

Các file code chính dùng để lập đặc tả:

- `ExcelAddIn1/Winform/FrmDaodat.cs`.
- `ExcelAddIn1/RandomDaodat*.cs`.
- `ExcelAddIn1/Winform/Dutoan.cs`.
- `ExcelAddIn1/Winform/FrmProjectSetup.cs`.
- `ExcelAddIn1/Winform/PriceProfileControl.cs`.
- `ExcelAddIn1/Winform/EstimateAppendixControl.cs`.
- `ExcelAddIn1/Funtion/WorkbookGeneratedEstimateWriter.cs`.
- `ExcelAddIn1/Winform/CostSummaryControl.cs`.
- `ExcelAddIn1.Core/CostSummaryCalculator.cs`.
- `ExcelAddIn1/Funtion/LicenseManager.cs`.
- `ExcelAddIn1/Funtion/SupportPackage.cs`.

Bằng chứng test hiện có:

- Core: `71/71` tại mốc DT-701.
- Release acceptance legacy: `14/14 PASS` tại DT-605.
- Estimate Workspace: scan, grouping, 5 sheet thực dùng, writer idempotent,
  rename/save/reopen đã PASS.
- Modeless: Excel vẫn chọn/sửa cell khi form mở và form đóng theo workbook đã PASS.
- Release gate 1.0.4: build/sign/hash/tamper/preflight đã PASS.

Các bằng chứng trên không thay thế việc phê duyệt 27 quyết định logic ở mục 10.

## 14. Phiếu phản hồi QA/chủ sản phẩm

QA có thể sao chép bảng dưới đây vào biên bản review. Không đánh dấu `Đạt` chung cho toàn
bộ sản phẩm khi vẫn còn quyết định P0 chưa có kết luận.

| ID | Quyết định (`Giữ/Sửa/Loại bỏ/Chưa đủ căn cứ`) | Mô tả logic mong muốn | Người chốt | Ngày |
|---|---|---|---|---|
| LR-01 |  |  |  |  |
| LR-02 |  |  |  |  |
| LR-03 |  |  |  |  |
| LR-04 |  |  |  |  |
| LR-05 |  |  |  |  |
| LR-06 |  |  |  |  |
| LR-07 |  |  |  |  |
| LR-08 |  |  |  |  |
| LR-09 |  |  |  |  |
| LR-10 |  |  |  |  |
| LR-11 |  |  |  |  |
| LR-12 |  |  |  |  |
| LR-13 |  |  |  |  |
| LR-14 |  |  |  |  |
| LR-15 |  |  |  |  |
| LR-16 |  |  |  |  |
| LR-17 |  |  |  |  |
| LR-18 |  |  |  |  |
| LR-19 |  |  |  |  |
| LR-20 |  |  |  |  |
| LR-21 |  |  |  |  |
| LR-22 |  |  |  |  |
| LR-23 |  |  |  |  |
| LR-24 |  |  |  |  |
| LR-25 |  |  |  |  |
| LR-26 |  |  |  |  |
| LR-27 |  |  |  |  |

### Kết luận review

```text
Phạm vi đã review:
Các quyết định P0 đã chốt:
Các quyết định còn treo:
Test cần bổ sung:
Cho phép sửa code: Có / Không
Điều kiện trước khi phát hành tiếp:
```
