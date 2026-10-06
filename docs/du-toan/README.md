# Ho so dieu phoi module Du toan RPBM

Tai lieu trong thu muc nay la nguon dieu phoi chinh cho viec phat trien module du toan ra pha bom min vat no. Bat ky nguoi hay may nao tiep tuc cong viec deu phai bat dau tu file nay.

## Tai lieu QA hien tai

- Ban dac ta hanh vi va ho so ra soat logic cho release 1.0.4:
  [`../QA_FUNCTIONAL_SPEC_1.0.4.md`](../QA_FUNCTIONAL_SPEC_1.0.4.md).
- Truoc khi sua tiep logic nghiep vu, phai doc va chot cac muc `LR-01` den `LR-27`,
  dac biet cac muc P0. Test pass khong duoc coi la thay the phe duyet logic.

## Thu tu doc bat buoc

1. [PROGRESS.md](PROGRESS.md): xem task dang lam, task ke tiep va blocker.
2. [HANDOFF.md](HANDOFF.md): doc boi canh hien tai va hanh dong tiep theo.
3. [TASKS.md](TASKS.md): doc day du yeu cau va dieu kien nghiem thu cua task hien tai.
4. [TESTING.md](TESTING.md): chay dung quy trinh kiem thu va workbook chuan.
5. [BASELINE.md](BASELINE.md): checksum, expected output va loi ke thua cua workbook chuan.
6. [ARCHITECTURE.md](ARCHITECTURE.md): tuan thu cac quyet dinh kien truc da chot.
7. [SOURCES.md](SOURCES.md): kiem tra can cu phap ly va nguon tham khao.
8. [WORKLOG.md](WORKLOG.md): xem ket qua cac lan lam viec truoc.

## Quy tac lam viec

- Chi mot task duoc co trang thai `IN_PROGRESS` tai mot thoi diem.
- Khong bat dau task moi khi task hien tai chua dat tat ca dieu kien nghiem thu.
- Build thanh cong chua du de hoan thanh task. Phai co unit test hoac integration test va test tren workbook chuan tuy theo pham vi.
- Neu test that bai, task giu nguyen `IN_PROGRESS` hoac chuyen `BLOCKED`; khong danh dau `DONE`.
- Moi thay doi phai bao toan cac chuc nang ho dao da co, tru khi task noi ro thay doi behavior.
- Khong tu dong nang cap ho so cu sang bo phap ly moi.
- Khong hard-code ten sheet de nhan dien vai tro nghiep vu.
- Khong sua du lieu nguon phap ly da phat hanh. Dieu chinh cua nguoi dung phai nam o lop override rieng va co nhat ky.

## Ket thuc mot task

Trong cung lan lam viec, nguoi thuc hien phai:

1. Chay day du test trong [TESTING.md](TESTING.md).
2. Ghi bang chung vao [WORKLOG.md](WORKLOG.md).
3. Cap nhat trang thai task trong [PROGRESS.md](PROGRESS.md).
4. Cap nhat [HANDOFF.md](HANDOFF.md) de chi ra dung task tiep theo.
5. Cap nhat [ARCHITECTURE.md](ARCHITECTURE.md) neu co quyet dinh moi.
6. Chi sau do moi duoc chuyen task tiep theo sang `IN_PROGRESS`.

## Trang thai du an hien tai

- Task tai lieu hoa `DT-000`: `DONE`.
- Gate B1 da dat; checkpoint moi nhat: `Dutoanmau/Checkpoints/DT-105-end.xlsm`.
- Toan bo `DT-101` den `DT-606` da `DONE`; Gate B6 dat cho ban `1.0.0 PILOT x64`.
- Artifact: `artifacts/TTBMVN-Excel-Tools-1.0.0-PILOT-x64.zip`.
- Workbook lam viec duoc phep sua: `Dutoanmau/00. Du toan TP 3.xlsm`.
- Baseline `DT-101` da dat; expected output nam trong [BASELINE.md](BASELINE.md).
