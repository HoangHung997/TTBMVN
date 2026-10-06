# TTBMVN Excel Tools 1.0.2 PILOT x64

## Sua loi

- Sua `NullReferenceException` khi DataGridView phat SelectionChanged trong luc quet va nap dong.
- Bao ve BuildWorkspace khoi hang DataGridView dang duoc tao nhung chua co RowDraft.
- Tu bo qua tang tieu de phu khong co ma/noi dung/don vi/khoi luong trong bang Excel nhieu header.
- Cot khoi luong nghiem thu tuy chon coi dau gach, text va ma loi Excel la chua co du lieu.
- Cot khoi luong chinh van fail ro rang neu cong thuc Excel bi loi.

## Kiem thu

- Release x64 build sach; 71/71 core test PASS.
- Quet truc tiep `Gia DT TC!A9:L27` tren ban sao workbook chuan PASS.
- Nap lai grid khi control dang hien thi PASS.
- O nghiem thu `#N/A`, rename sheet, save/reopen va writer idempotent PASS.

## Cai dat

- Dong tat ca cua so Excel.
- Giai nen toan bo ZIP, sau do nhap dup `Cai-dat-TTBMVN.cmd`.
- Khong mo truc tiep `.vsto` neu may dang co ban cu tu vi tri khac.
