# DT-303 - Review catalog may va gia ca may

Ngay review: 2026-08-03.

## Pham vi

- Phuong phap: VBHN96-2025-BQP, Phu luc I, trang 5-14.
- Du lieu doi tuong huong luong ngan sach: Bang 01/02, trang 15-23.
- Du lieu doi tuong khong huong luong ngan sach: Bang 03/04, trang 24-33.
- Luong thuy thu va nhan cong: trang 34-35.
- Oracle Excel: `Dutoanmau/00. Du toan TP 3.xlsm`, sheet `VL-NC-M`.

## Quyet dinh tinh toan

- `Ccm = Ckh + Csc + Cnl + Cnc + Ccpk`.
- Khau hao dung `(G - Gth) * Dkh / Nca`; sua chua va chi phi khac dung
  `G * rate / Nca`; cac rate phan tram chia 100.
- Moi thanh phan duoc tinh bang `decimal`, sau do lam tron 1 VND bang
  `MidpointRounding.AwayFromZero`. Tong bang tong nam thanh phan da lam tron.
- Moi truong nuoc man/lo/an mon cao nhan 1,05 vao dinh muc khau hao va sua chua.
- Nhien lieu phu mac dinh: xang 1,02; diesel 1,03; dien 1,05; pin 1,00. Price profile
  co the ghi de he so theo dieu kien du an.
- Neu nhien lieu da nam trong chi phi vat lieu, thanh phan nhien lieu trong gia ca may bang 0.
- Gia ca may cho: `50% Ckh + 50% Cnc + Ccpk`. Gia gio phan bo raw total theo so gio/ca.
- Nguyen gia ghi de vuot/khong vuot 30 trieu tu dong chon gia tri thu hoi 10%/0%, tru khi
  co override co truy vet.

## Sua du lieu nguon

- Giu 33 key logic `MACHINE-M010.xxx`, bo sung ma phap ly theo doi tuong:
  `M010.xxx` va `M011.xxx`.
- Cac chuoi `6 x 20`, `4 x 16`, `4 x 14`, `3 x 8`, `2 x 6`, `1 x 2` la thanh phan
  si quan + thuy thu, khong phai bac tho. Da ma hoa thanh hai hao phi nhan cong.
- M011.020 co them `1-bac-5-10`; M010.020 khong co nhan cong dieu khien.
- M011.024 dung nguyen gia 350.000 VND; M010.024 dung 1.350.000 VND.
- Locator moi bao phu ca Bang 01 va Bang 03 trong VBHN96.

## Sai khac phat hanh duoc ghi nhan

| Ma | Cong thuc tu du lieu | Bang 04 in | Ket luan |
|---|---:|---:|---|
| M011.024 - sua chua | 53 | 44 | Bang 04 lech cong thuc (3); engine giu 53 |
| M011.024 - tong | 495.298 | 495.289 | He qua cua sai khac sua chua |
| M011.009 - tong | 518.226 | 518.227 | Bang 04 lech 1 VND; workbook va engine dung 518.226 |
| M011.010 - tong | 554.101 | 554.102 | Bang 04 lech 1 VND; workbook va engine dung 554.101 |

Khong hard-code cac tong in sai vao engine. Audit van giu ca ket qua cong thuc va gia tri in.

## Evidence

- Release build: 32/32 test dat, khong warning.
- `test-bqp-2025-package.ps1`: 248 record va PDF hash/page gate dat.
- `test-machine-rate-catalog.ps1`: 33 may, 3 core gate va 10 tong `VL-NC-M` dat.
- Package checksum sau bo sung hai doi tuong:
  `D58CAC64422E14151FD68F1C92944FAB7A9859908165C59F92089E7F13998B7C`.
- MachineRate module checksum:
  `6352542DC1F488FEA3D6B5D05D16E37D80464FEC91C09A0486FA4E7839BE8CBE`.
