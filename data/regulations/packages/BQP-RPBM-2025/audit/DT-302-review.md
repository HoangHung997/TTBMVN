# DT-302 - Biên bản đối chiếu gói BQP 2025

Ngày kiểm tra: 2026-08-03.

## Phạm vi

- Snapshot đã hợp nhất có hiệu lực từ 28/10/2025.
- 60 record quy trình, 42 định mức, 37 chi phí, 33 dữ liệu máy, 69 địa bàn và
  7 record compliance; tổng 248 record.
- Tất cả record là `VerifiedAgainstOfficialSource`, có văn bản, trang và mục/bảng nguồn.
- Ma trận `change-reasons.tsv` phủ đủ 248 record trong diff; không có record 2021 bị xóa.

## Nguồn chính thức

| Văn bản | Trang | SHA-256 PDF |
|---|---:|---|
| TT101/2025/TT-BQP | 69 | `E445F63CFDB515CB187780247423610A5A29DD0BBFDB231EBC21D858516DD240` |
| 94/VBHN-BQP | 30 | `D375AAC8F0C0B016D0C8B1E8CC65140560DF23D13AB3B499DE581DC5276F8935` |
| 95/VBHN-BQP | 99 | `2D5207B562EC63A24BD0B9D38C86E748AF6F7A4009195250517B83EFE22A1094` |
| 96/VBHN-BQP | 35 | `10A25B90BC95F8EB8330C9CD6D2423AD0DF46746D95DB7D33C7A2157228AE0B1` |
| 97/VBHN-BQP | 52 | `92B6C42996E554A908661B0395CE7286D41DDD69C17859B40993DC46300B69E0` |
| 98/VBHN-BQP | 76 | `62CC32BDA9ADED4310AB183180D6D4D861E41FEA94852848425EA0C19BAFE0A2` |

Lớp text được trích theo từng trang và có SHA-256 tại
`data/regulations/text/2025/text-manifest.tsv`. Các trang không có text của VBHN95
(43, 61) và VBHN98 (33, 38, 55) đã được render kiểm tra, đúng là trang watermark trống.

## Lượt A - TT101 độc lập

- Điều 2: xác nhận sửa Điều 5, 12, 14, 55 của quy trình và thay Mẫu RPBM-13.
- Phụ lục III trang PDF 35: 020.0500 có NC `6,40;7,05;7,76;8,54`, Vallon
  `4,27;4,70;5,17;-`, Vet1 `-;-;-;5,69`.
- Trang 36-37: tám suất tạm tính theo khu vực và nguồn lương khớp record.
- Trang 37: K2 là `2,2;2,0;1,9;1,8;1,7` và `1,1;1,0;0,95;0,9;0,85`;
  hai mức tối thiểu đều `18 x 350.000 = 6.300.000` đồng.
- Trang 38: Biểu mẫu 05 có đủ ba phần thuyết minh.

Kết quả: `PASS`.

## Lượt B - Văn bản hợp nhất

- 95/VBHN-BQP: 59 điều có page locator theo tiêu đề và phạm vi đến điều kế tiếp;
  Mẫu RPBM-13 tại trang PDF 89-90.
- 96/VBHN-BQP trang 15-19: rà trực quan đủ 33 dòng Bảng 01. Các giá thay đổi,
  M010.020 không có nhân công điều khiển, M010.026 là Máy điểm hỏa và M010.027 dùng
  60 lít xăng E5 RON 92-II đều khớp source.
- 97/VBHN-BQP trang 5-14: đủ 34 tỉnh/thành phố sau sắp xếp; mỗi địa bàn có một key,
  danh sách zone không trùng và có quy tắc ưu tiên rõ. Danh sách khu vực 2/3 dưới biển
  khớp trang 14.
- 97/VBHN-BQP trang 42, 44-46 và 52: suất tạm tính, K1/K4, K2, K3, K5, K6,
  hai mức tối thiểu và Biểu mẫu 05 khớp TT101.

Kết quả: `PASS`.

## Test tự động

- Build `Release|x64`: sạch, 29/29 console tests đạt.
- Bundle tái tạo giữ checksum; 6 PDF và text manifest đúng hash/số trang.
- Diff cấp record: `248 changes = 50 added + 0 removed + 198 changed`.
- `scripts/test-bqp-2025-package.ps1 -RegenerateSource`: `PASS`.
- B1 regression: role rename, profile reopen, sheet coordinator và ba workbook variant đều đạt.

Package checksum:
`F4618DF4D256378C59A5719912CFF4D9EED410B52FCF157AEAD1D4EB6EC49A4D`.

Day la checksum tai thoi diem dong DT-302. DT-303 hoan thien bien the Bang 03 trong cung
snapshot development; checksum hien tai xem tai `../README.md` va bien ban DT-303.

## Checkpoint workbook

- Start: `DT-302-start.xlsm`, 670221 byte, SHA-256
  `0AD1CCD5BD610268BF303E71982CF83A93AF2435994CD2A146DCBC8FD83AAF89`.
- End: `DT-302-end.xlsm`, 670220 byte, SHA-256
  `1AB5B6487827DFDD55803BCF602C880598414B1C42F17B9AAFA891710EB79272`.
- Cả hai: profile schema 2, package pin Available, role PASS, 92 lỗi công thức kế thừa,
  `THKP-TC!E27=1198731000`; DT-302 không thay đổi dữ liệu nghiệp vụ workbook.
