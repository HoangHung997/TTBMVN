# Danh sach task chi tiet

Moi task chi duoc danh dau `DONE` khi hoan thanh deliverable, test va cap nhat tai lieu ban giao.

## Khoi dong

### DT-000 - Tao ho so dieu phoi

- Muc tieu: tao roadmap, task list, progress, test protocol, kien truc, nguon va handoff.
- Dau ra: toan bo file trong `docs/du-toan`.
- Nghiem thu: cac file lien ket cheo dung; build hien tai va test hien co duoc ghi lai.
- Trang thai: `DONE`.

### DT-101 - Chot workbook chuan va ket qua ky vong

- Muc tieu: bien workbook tren man hinh thanh chuan kiem thu tai lap duoc.
- Dau vao: `Dutoanmau/00. Du toan TP 3.xlsm` dang mo trong Excel.
- Thuc hien: luu workbook dang mo; tao ban baseline bat bien; tao checkpoint dau task; ghi danh sach sheet, named range quan trong, cell/formula dau vao va dau ra cua luong du toan.
- Dau ra: baseline, checksum, bang expected output va quy tac phuc hoi working workbook.
- Test: mo baseline; xac nhan VBA/format/formula con nguyen; chay recalculate; so sanh cac cell da chot; dong/mo lai khong thay doi ket qua.
- Dat khi: co checksum, khong con expected value trong trang thai `TBD`, va co the khoi phuc workbook de chay lai.

## B1 - Workbook identity va project profile

### DT-102 - Dinh nghia va luu vai tro sheet

- Muc tieu: nhan dien sheet theo ID ngu nghia, khong theo ten hien thi.
- Vai tro ban dau: `ResourcePrices`, `UnitRateLand`, `UnitRateWater`, `EstimateAppendix`, `CostSummary`, `NormLookupView`, `CostRuleView`.
- Thuc hien: schema role; doc/ghi `Worksheet.CustomProperties`; validate trung role va role thieu.
- Test: unit test serializer/resolver; gan role tren workbook chuan; doi ten sheet va xac nhan van resolve dung.
- Dat khi: ten sheet co the doi tu do ma luong nghiep vu khong mat tham chieu.

### DT-103 - Project profile trong workbook

- Muc tieu: luu can cu du an doc lap voi may va tai khoan nguoi dung.
- Du lieu: project ID, ngay lap, ngay phe duyet, ngay gia, package ID, price profile ID, schema version va override summary.
- Thuc hien: model co version; doc/ghi qua workbook custom properties; migration schema.
- Test: round-trip unit test; save/close/open workbook; du lieu khong mat va khong bi trung.
- Dat khi: workbook tu mo ta duoc goi phap ly va bo gia no dang dung.

### DT-104 - Dong bo danh sach sheet khi workbook thay doi

- Muc tieu: form dang mo phan anh dung them, xoa, doi ten va activate sheet.
- Thuc hien: quan ly su kien workbook; refresh co debounce; khong reset cac gia tri form dang nhap.
- Test: them, xoa, doi ten sheet trong luc form mo; danh sach cap nhat; DataGridView va field khac khong bi reload.
- Dat khi: khong con danh sach sheet cu va khong mat du lieu nguoi dung chua luu.

### DT-105 - Wizard anh xa va gate B1

- Muc tieu: xu ly file chua duoc danh dau hoac bi mat sheet.
- Thuc hien: wizard chon sheet cho tung role; kiem tra mot sheet khong nhan role xung dot; luu mapping.
- Test: ban sao workbook bo tag; workbook thieu sheet; workbook doi ten tat ca sheet; save/reopen.
- Dat khi: toan bo test B1 dat va workbook chuan van cho ket qua baseline.
- Trang thai: `DONE`.

## B2 - Regulation package engine

### DT-201 - Mo hinh RegulationPackage

- Muc tieu: mo hinh hoa goi phap ly bat bien.
- Du lieu: package ID, schema/data version, hieu luc, trang thai, van ban nguon, checksum, transition note va cac module catalog.
- Test: serialization round-trip, validation ID/version/date, reject package sai.
- Dat khi: schema khong phu thuoc WinForms hoac Excel Interop.
- Trang thai: `DONE`.

### DT-202 - Kho va import package

- Muc tieu: cai nhieu phien ban du lieu song song.
- Thuc hien: manifest, noi dung, checksum, staging import, atomic install, rollback khi import loi.
- Test: package dung/sai checksum, trung version, mat file, schema cu.
- Dat khi: cap nhat khong ghi de package da dung trong ho so cu.
- Trang thai: `DONE`.

### DT-203 - Resolver ngay hieu luc va chuyen tiep

- Muc tieu: de xuat dung package theo ngay lap/phe duyet va quy dinh chuyen tiep.
- Test: truoc 05/11/2021, tu 05/11/2021, truoc/sau 28/10/2025, ho so da phe duyet truoc ngay hieu luc.
- Dat khi: ket qua resolver co ly do va can cu, khong chi tra ve package ID.
- Trang thai: `DONE`.

### DT-204 - Pin package vao workbook

- Muc tieu: workbook cu luon mo voi dung package da ap dung.
- Thuc hien: pin package/version/checksum; canh bao package thieu; cho phep cai lai tu goi offline.
- Test: mo workbook tren may khac; co/khong co package; app co package moi hon.
- Dat khi: khong tu dong doi ket qua ho so cu.
- Trang thai: `DONE`.

### DT-205 - Preview migration va gate B2

- Muc tieu: chuyen goi co kiem soat.
- Thuc hien: tao checkpoint/ban sao; diff ma dinh muc, hao phi, he so, chi phi va may; chi apply sau xac nhan.
- Test: cancel khong doi workbook; apply thanh cong; loi giua chung rollback.
- Dat khi: toan bo test B2 dat va baseline workbook khong bi thay doi ngam.
- Trang thai: `DONE`.

## B3 - Du lieu BQP

### DT-301 - Goi BQP 2021

- Muc tieu: tao package hieu luc tu 05/11/2021 tu TT121, TT122, TT123.
- Thuc hien: nhap quy trinh, dinh muc, chi phi va du lieu may thanh cac module rieng.
- Test: dem record, key unique, kieu don vi, tong/quan he hao phi va doi chieu mau ngau nhien hai lan.
- Dat khi: moi record co source locator va khong con gia tri chua xac minh.
- Trang thai: `DONE`.

### DT-302 - Goi BQP 2025

- Muc tieu: tao snapshot da resolve hieu luc 28/10/2025.
- Nguon chinh: TT101/2025 va VBHN 94-98/BQP.
- Test: diff voi goi 2021; kiem tra tat ca noi dung sua/thay the; dia ban va vung khong trung/xung dot.
- Dat khi: diff giai thich duoc tung thay doi.
- Trang thai: `DONE`.

### DT-303 - Catalog may va gia ca may

- Muc tieu: trien khai du lieu/phuong phap TT122 va Phu luc IV TT101.
- Test: mau may can/nuoc; khau hao, sua chua, nhien lieu, nhan cong dieu khien, chi phi khac; so sanh workbook chuan.
- Dat khi: ket qua doc lap voi locale va co precision ro rang.
- Trang thai: `DONE`.

### DT-304 - Catalog dinh muc va chi phi

- Muc tieu: trien khai Phu luc I/II TT123 va thay doi tai TT101.
- Test: hao phi VL/NC/M; chi phi truc tiep/chung/thue/khac; dieu kien va noi suy neu co.
- Dat khi: moi quy tac co input, output, rounding va source locator.
- Trang thai: `DONE`.

### DT-305 - Rasoat doc lap va gate B3

- Muc tieu: kiem tra du lieu bang nguoi/luong doc lap truoc khi dung tinh tien.
- Test: bao cao missing/duplicate/outlier, sample doi chieu van ban, package checksum.
- Dat khi: khong con loi nghiem trong va co bien ban phe duyet du lieu.
- Trang thai: `DONE`.

## B4 - Luong lap du toan

### DT-401 - Project setup va sheet mapping

- Muc tieu: chon can cu, moc gia va anh xa sheet trong mot luong ro rang.
- Test: file chuan, file doi ten sheet, file chua tag, cancel/restore.
- Trang thai: `DONE`.

### DT-402 - Tra cuu dinh muc

- Muc tieu: thay nguon sheet `Tracuu` bang catalog trong app.
- Chuc nang: tim ma/ten/tu khoa, loc can-nuoc/do sau/vung, xem hao phi va can cu.
- Test: exact/fuzzy search, ma khong ton tai, hai package co cung ma.
- Trang thai: `DONE`.

### DT-403 - Rule engine ChiPhi

- Muc tieu: thay bang hard-code va sheet `ChiPhi` bang rule co phien ban.
- Test: can bien, noi suy, dieu kien ap dung, rounding, override co ly do.
- Trang thai: `DONE`.

### DT-404 - VL-NC-M va PriceProfile

- Muc tieu: quan ly gia VL, NC, nhien lieu/may theo dia diem va thoi diem, tach khoi package phap ly.
- Test: import/nhap tay, thay moc gia, missing price, save/reopen, link vao don gia.
- Trang thai: `DONE`.

### DT-405 - DG Can va DG Nuoc

- Muc tieu: tinh don gia chi tiet tu dinh muc va PriceProfile; quan ly tuy chon lay he so.
- Test: ma mau can/nuoc, bat/tat he so, gia thieu, so sanh baseline.
- Trang thai: `DONE`.

### DT-406 - Gia DT TC

- Muc tieu: ap don gia vao phu luc/gia du toan thi cong.
- Test: khoi luong, don gia, thanh tien, round, link nguoc/nguon va thay doi dau vao.
- Trang thai: `DONE`.

### DT-407 - THKP-TC va gate B4

- Muc tieu: tong hop kinh phi tron luong va ho tro cac loai du toan da chot.
- Test: end-to-end tu VL-NC-M den THKP-TC; recalculate; save/reopen; compare expected output.
- Dat khi: tat ca ket qua trong baseline dat dung sai va khong co `#REF!`/loi Excel moi.
- Trang thai: `DONE`.

## B5 - Excel integration va audit

### DT-501 - Ghi Excel theo batch

- Muc tieu: giam COM round-trip va bao toan trang thai Excel.
- Thuc hien: mo rong `ExcelWriteContext`; ghi mang `Value2`/formula theo block; release COM hop ly.
- Test: performance, exception injection, ScreenUpdating/Calculation/Events/DisplayAlerts duoc khoi phuc.
- Trang thai: `DONE`.

### DT-502 - Truy vet can cu

- Muc tieu: moi block ket qua biet package, price profile, ma dinh muc va source locator.
- Test: chon mot ket qua bat ky va truy ve dung van ban/phu luc/bang.
- Trang thai: `DONE`.

### DT-503 - Validation va reconciliation

- Muc tieu: phat hien thieu gia, sai mapping, sai tong, stale formula va broken name.
- Test: bo co chu dich tung dau vao va xac nhan thong bao chi dung vi tri.
- Trang thai: `DONE`.

### DT-504 - Chuyen package va rollback

- Muc tieu: migration ho so co preview va phuc hoi.
- Test: 2021 -> 2025, cancel, partial failure, restore checkpoint, compare report.
- Trang thai: `DONE`.

### DT-505 - Runtime log va gate B5

- Muc tieu: log co correlation ID, task/phase, workbook/sheet role, exception va package version.
- Test: gay loi COM co kiem soat; tao support package; khong ghi du lieu nhay cam khong can thiet.
- Trang thai: `DONE`.

## B6 - Cap nhat va thuong mai hoa

### DT-601 - Dinh dang goi cap nhat offline

- Muc tieu: goi du lieu co manifest, checksum va chu ky so.
- Test: package hop le, bi sua, sai chu ky, downgrade, trung version.
- Trang thai: `DONE`.

### DT-602 - Update center

- Muc tieu: danh sach package, xem diff, cai, rollback; de san interface online nhung khong bat buoc server.
- Test: offline-only, mat mang, package moi/cu, restart app.
- Trang thai: `DONE`.

### DT-603 - Ma tran tuong thich

- Muc tieu: Windows 10/11, Office 2016/2019/2021/365, x86/x64 va decimal/list separator theo Excel.
- Test: theo ma tran trong `TESTING.md`; ghi ro cau hinh chua co may test.
- Trang thai: `DONE`.

### DT-604 - Support va chan doan

- Muc tieu: mo rong nut Ho tro cho module du toan.
- Dau ra: log, app-info, activation status, project profile, package manifest va validation report da loc du lieu nhay cam.
- Trang thai: `DONE`.

### DT-605 - Bo acceptance test phat hanh

- Muc tieu: tu dong hoa regression core va danh sach live Excel test co bang chung.
- Test: clean install, upgrade, open old workbook, migration va uninstall/reinstall.
- Trang thai: `DONE`.

### DT-606 - Gate phat hanh

- Muc tieu: build Release, ky code, publish, tai lieu cai/go cai/cap nhat va release notes.
- Dat khi: khong con blocker P0/P1; acceptance dat; package va workbook mau duoc luu tru.
- Trang thai: `DONE`.

## B7 - Luong du toan theo Phu luc DT

### DT-701 - Estimate Workspace va don gia theo ngu canh

- Muc tieu: lay Phu luc DT do nguoi dung bo tri tren Excel lam nguon cong tac, gan dinh muc
  chi tiet trong app va chi tinh cac dong da gan dinh muc.
- Dinh danh don gia: package + dinh muc + ma chi tiet + can/nuoc + HLNS/KHLNS + ho so gia
  + dieu kien/he so + binding thiet bi logic. So hieu cong tac va khoi luong khong nam trong key.
- Sheet sinh: chi tao cac nhom dang dung trong ho so; tach DG Can/DG Nuoc va HLNS/KHLNS,
  tach VL-NC-M theo HLNS/KHLNS.
- Test: hai dong cung key dung chung mot don gia; khac ma chi tiet/doi tuong/moi truong/binding
  phai tach; dong van ban khong tinh; sheet nguon doi ten van resolve; writer chay lai khong tao
  trung; save/reopen va hai control UI nap lai dung.
- Trang thai: `DONE`.
