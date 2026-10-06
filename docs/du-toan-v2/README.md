# Du toan RPBM V2 - huong trien khai moi

Thu muc nay la ho so rieng cho huong thiet ke moi cua module du toan RPBM sau khi tiep quan lai du an hien co.

> Muc tieu: giu Excel la ho so du toan chinh, dung add-in de gan dinh muc, sinh bang, sinh cong thuc/link va kiem tra. Khong bien add-in thanh mot phan mem dong tinh xong roi do so chet vao Excel.

## 1. Da doc va nam duoc add-in hien tai

Cac README da duoc ra soat truoc khi lap ho so V2:

- `/README.md`
- `/docs/du-toan/README.md`
- `/data/regulations/packages/BQP-RPBM-2021/README.md`
- `/data/regulations/packages/BQP-RPBM-2025/README.md`
- `/data/regulations/raw/2026/README.md`

Dong thoi da doi chieu cac tai lieu ban giao quan trong:

- `docs/PROJECT_STATUS.md`
- `docs/QA_FUNCTIONAL_SPEC_1.0.4.md`
- `docs/du-toan/PROGRESS.md`
- `docs/du-toan/HANDOFF.md`
- `docs/du-toan/ARCHITECTURE.md`

Add-in hien tai da co nen tang tot gom: VSTO x64, Core tach khoi Excel, worksheet role, ProjectProfile, package phap ly 2021/2025, lookup dinh muc, CostRule, PriceProfile, don gia can/nuoc, phu luc du toan, THKP, validation, offline update, diagnostics va license.

V2 **khong viet lai tat ca tu dau**. V2 se giu lai nhung thanh phan dung va thay cach to chuc luong du toan de bam sat ho so Excel thuc te.

## 2. Quyet dinh nghiep vu da chot cho V2

### Excel la ho so, database la thu vien, add-in la engine tu dong hoa

- Moi workbook Excel la mot du an.
- Database khong luu khoi luong du an, tong tien hay ket qua du toan.
- Database chi luu thu vien dinh muc, chi tiet hao phi, danh muc VL/NC/M, thong so may va version/can cu phap ly.
- Nguoi dung lap khoi luong trong Excel truoc.
- Add-in quet cac bang/vung cong tac da duoc nguoi dung chi dinh.
- Nguoi dung **gan dinh muc thu cong** cho tung dong hoac nhom dong.
- Add-in dua vao danh sach dinh muc da gan de sinh/cap nhat cac bang con lai.

### Khong co so chet doi voi ket qua tinh

Chi cho phep gia tri nhap truc tiep doi voi:

1. Khoi luong/thong so dau vao cua du an.
2. Gia thi truong, gia nhien lieu, tham so nguoi dung nhap.
3. Hao phi dinh muc la du lieu goc tu van ban/package.

Moi ket qua phat sinh phai la cong thuc/link Excel:

`Gia dau -> Gia ca may -> Don gia cong tac -> Thanh tien -> THKP`

`Gia nhan cong -> Don gia cong tac -> Thanh tien -> THKP`

`Gia vat lieu -> Don gia cong tac -> Thanh tien -> THKP`

### Khong lay so khu vuc lam bien dieu khien chinh

Du an co 1, 2 hay 3 khu vuc khong lam thay doi kien truc.

Nguoi lap gan dung dinh muc/variant tuong ung cho tung cong tac; add-in tin vao dinh muc duoc gan. Khu vuc/mat do chi la metadata ho tro tra cuu, khong phai wizard bat buoc cua du an.

### Bo sheet in chinh

Mau hien hanh duoc chon lam huong chinh:

- `THKP-TC`
- `Gia DT TC`
- `DG Can`
- `DG Nuoc`
- `VL-NC-M`

Neu thuc te co cong tac bien thi co the sinh `DG Bien`. Khong tao sheet rong chi de du cau truc.

Truong hop cu co hai doi tuong luong VT/DN van duoc xem la yeu cau mo rong can ho tro, nhung khong lam phuc tap luong V2 co ban.

## 3. Ba mau du toan da doi chieu

1. `Du toan RPBM Pleiku_TP2_QuyNhon_Ver7.2.xlsx` - mau tham khao luong bang va don gia.
2. `DU TOAN RPBM-VD4- Ver1.2.xls` - mau 2023 co ca huong luong NSNN va khong huong luong NSNN.
3. `Du toan RPBM HoaLuNamDinh_Ver1.xlsx` - mau moi nhat, chon lam huong chinh cho V2; du an co 2 khu vuc nhung kien truc khong phu thuoc so khu vuc.

## 4. Anh tham chieu du an

Cac anh duoi day duoc tach rieng vao thu muc `images/` de dung khi doi chieu giao dien va mau in:

- [THKP-TC](images/THKP-TC.png)
- [Gia DT TC](images/Gia-DT-TC.png)
- [DG Can](images/DG-Can.png)
- [DG Nuoc](images/DG-Nuoc.png)
- [VL-NC-M](images/VL-NC-M.png)

## 5. Tai lieu trong thu muc nay

- [ARCHITECTURE.md](ARCHITECTURE.md): kien truc dich V2.
- [MIGRATION-PLAN.md](MIGRATION-PLAN.md): ke hoach tiep quan, phan nao giu/phan nao sua va cac moc trien khai.
- [PROGRESS.md](PROGRESS.md): trang thai task dang lam, task da xong, commit va viec tiep theo.
- [UI-CONTRACT.md](UI-CONTRACT.md): bo UI bat buoc phai bam theo anh da chot.
- [HANDOVER.md](HANDOVER.md): diem vao ban giao cho nguoi tiep quan, kien truc bat buoc, rui ro runtime va task tiep theo.

## 6. Trang thai

V2 da trien khai qua cac moc chinh:

- `V2-001`: bo gate khoi dong, bam Du toan mo ngay CustomTaskPane.
- `V2-002`: shell UI theo bo anh chot.
- `V2-101`: WorkItemId + Custom XML + fingerprint/reconcile.
- `V2-201`: VL-NC-M bang formula/link.
- `V2-301`: DG Can / DG Nuoc / DG Bien theo RateId unique.
- `V2-401`: link Gia DT TC + THKP-TC.
- `V2-501`: validation / safe repair.
- `V2-601`: legacy migration, VT/DN, rename sheet, package missing graceful degradation.
- `V2-701`: Thiet lap chung theo workbook, output mapping, auto-sync, auto-save va mapping-loss warning.

Task tiep theo:

- **`V2-801` — Goi phap ly & Du lieu**.
- Sau do `V2-901` — Bao cao & Xuat in.

**Luu y:** cac task tren moi o muc implementation/runtime pending. Moi truong phat trien hien tai chua co bang chung build VSTO/Excel end-to-end, vi vay khong duoc coi la runtime PASS.

Nguoi tiep quan bat dau tai [HANDOVER.md](HANDOVER.md), sau do xem [PROGRESS.md](PROGRESS.md) de biet commit, checklist va trang thai chi tiet.
