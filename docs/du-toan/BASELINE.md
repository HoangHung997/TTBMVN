# Baseline workbook du toan

Tai lieu nay la bang chung bat bien cua `DT-101`. Moi task co tac dong vao workbook phai
so sanh voi cac moc trong file nay va chi chap nhan sai lech khi task co yeu cau nghiep vu
ro rang.

## Dinh danh file

- Working: `Dutoanmau/00. Du toan TP 3.xlsm`.
- Baseline bat bien: `Dutoanmau/Baseline/00. Du toan TP 3.baseline.xlsm`.
- Checkpoint dau task: `Dutoanmau/Checkpoints/DT-101-start.xlsm`.
- Kich thuoc baseline: `667469` byte.
- Modified time luc capture: `2026-08-03 09:21:14`.
- SHA-256 baseline va checkpoint dau task:
  `E4BFDBD5FFA3AB47835A59D794F8E907633820F91E2FDF18EC9696559C89EADA`.
- Open XML package co `73` entry. File mang duoi `.xlsm` nhung khong co
  `xl/vbaProject.bin`; vi vay baseline nay khong chua VBA payload.
- Anh preview: `Dutoanmau/Baseline/previews`.

Baseline khong duoc ghi de. Neu working workbook bi hong, khoi phuc tu checkpoint `end`
gan nhat da dat, sau do sao chep thanh working workbook moi.

## Danh sach sheet

| Thu tu | Sheet | Used range tai baseline |
|---:|---|---|
| 1 | PL01 | A1:AE47 |
| 2 | THKL | A1:U36 |
| 3 | DosauHLAT | Trong |
| 4 | KLCT1 | A1:T52 |
| 5 | Tracuu | A1:G41 |
| 6 | THKP-TC | A1:L35 |
| 7 | Sheet2 | A1:E23 |
| 8 | Gia DT TC | A1:U27 |
| 9 | DG Can | A1:R145 |
| 10 | DG Nuoc | A1:Q138 |
| 11 | VL-NC-M | A1:J101 |
| 12 | ChiPhi | A1:Q65 |
| 13 | THNK_C | A1:W18 |
| 14 | Thoi gian | A1:H12 |
| 15 | THNK_N | A1:S17 |
| 16 | PTVT | A1:Q63 |

Workbook co `4826` defined name ke thua: `4222` workbook-scope va `604`
sheet-scope. Trong do `4034` name co `#REF!` va `120` name co `#N/A`. Day la du lieu
rac ke thua cua template, khong duoc tu dong sua/xoa hang loat trong task nhan dien sheet.

## Expected output - THKL

Gia tri duoi day la cached value do Excel luu trong baseline:

| Cell | Expected | Formula/nguon |
|---|---:|---|
| C18 | 38450.84395 | `='PL01'!M47` |
| D18 | 3.85 | `=ROUND(C18/10^4,2)` |
| E18 | 0.04 | `=ROUND(4*100/10^4,2)` |
| F18 | 3.81 | `=D18-E18` |
| C19 | 25085.91395 | `='PL01'!N47` |
| D19 | 2.51 | Lam tron hai chu so |
| E19 | 0.04 | Khoi luong loai tru |
| F19 | 2.47 | Khoi luong tinh du toan |
| C21 | 67346.99925 | `='PL01'!O47` |
| D21 | 6.73 | Lam tron hai chu so |
| E21 | 0.08 | Khoi luong loai tru |
| F21 | 6.65 | Khoi luong tinh du toan |
| C22 | 66626.59925 | `='PL01'!P47` |
| D22 | 6.66 | Cached value baseline |
| E22 | 0.08 | `=E21` |
| F22 | 6.58 | `=D22-E22` |

## Expected output - VL-NC-M

| Nhom | Cell | Expected |
|---|---|---:|
| Nhan cong bac 5/10 | F7 | 450000 |
| Nhan cong bac 7/10 | F11 | 495000 |
| Nhan cong bac 8/10 | F15 | 517500 |
| Luong toi thieu mau | D8, D12, D16 | 2340000 |
| May do min | F23 | 762996 |
| May do bom tren can | F29 | 1070324 |
| May do bom duoi nuoc | F35 | 1587824 |
| Thuyen cao su nho | F41 | 518226 |
| Thuyen cao su trung | F46 | 554101 |
| Thuyen composite | F51 | 995354 |
| Thiet bi soi va hut bun cat | F56 | 2946909 |
| Thiet bi lan 0.5m-3m | F62 | 1105582 |
| Thiet bi lan 3m-6m | F67 | 1108263 |
| Thiet bi lan 6m-12m | F72 | 1112914 |
| Coc be tong | F81 | 100000 |
| Coc go | F82 | 8000 |
| Co duoi nheo | F83 | 2500 |

## Expected output - DG Can

Cac dong tong hop cua tung block don gia, theo thu tu `VL / NC / M`:

| Row | F | G | H |
|---:|---:|---:|---:|
| 11 | 0 | 33165000 | 0 |
| 29 | 851550 | 8593200 | 8827863.72 |
| 46 | 0 | 0 | 0 |
| 63 | 0 | 0 | 0 |
| 75 | 0 | 31050 | 10681.944 |
| 93 | 640500 | 3489750 | 5030522.8 |
| 111 | 0 | 0 | 0 |
| 127 | 19108.73944 | 589950 | 8562.592 |
| 145 | 0 | 0 | 0 |

## Expected output - DG Nuoc

| Row | F | G | H |
|---:|---:|---:|---:|
| 11 | 0 | 33165000 | 0 |
| 34 | 572387.2 | 11790900 | 39574428.28 |
| 57 | 360287.2 | 4855950 | 19792190.91 |
| 74 | 4393.5 | 27720 | 29986.95 |
| 88 | 0 | 113850 | 441645.14 |
| 103 | 0 | 113850 | 764584.713 |
| 121 | 535040 | 495000 | 1107985.333 |
| 138 | 0 | 0 | 0 |

## Expected output - Gia DT TC truoc DT-406

| Cell | Expected | Formula/nguon |
|---|---:|---|
| I11 | 5208997.4635888 | `=SUM(I12:I18)` |
| J11 | 124640248.5 | `=SUM(J12:J18)` |
| K11 | 49083054.22904 | `=SUM(K12:K18)` |
| I19 | 7331588.156 | Tong VL duoi nuoc |
| J19 | 143774571.6 | Tong NC duoi nuoc |
| K19 | 509112566.75818 | Tong M duoi nuoc |
| J26 | 13070415.6 | Phan nuoc chay he so 1.1 |
| K26 | 46282960.61438 | Phan nuoc chay he so 1.1 |
| I27 | 12540585.6195888 | `=I19+I11` |
| J27 | 268414820.1 | `=J19+J11` |
| K27 | 558195620.98722 | Cached total baseline |

## Expected output - THKP-TC truoc DT-406

Gia tri chuan la cached value cua Excel va ket qua tren workbook dang mo luc capture:

| Cell | Expected | Formula |
|---|---:|---|
| E10 | 12540585.6195888 | `=TRANSPOSE('Gia DT TC'!I27:K27)` |
| E11 | 268414820.1 | Spill/cached tu `Gia DT TC` |
| E12 | 558195620.98722 | Spill/cached tu `Gia DT TC` |
| E13 | 839151026.7068089 | `=E12+E11+E10` |
| E14 | 107365928.04 | `=40%*E11` |
| E15 | 52058432.51107449 | `=$D15%*(E13+E14)` |
| E16 | 998575387.2578834 | `=E13+E14+E15` |
| E17 | 112470408.69288826 | `=SUM(E18:E23)` |
| E18 | 24964384.68144709 | `=$D$18%*E16` |
| E19 | 18461322.5875498 | `=$D$19%*E13` |
| E20 | 4992876.936289418 | Dieu kien min/max hoac `%*E16` |
| E21 | 9985753.872578835 | `=$D21%*E$16` |
| E22 | 24108808.99728662 | `=$D$22%*E13` |
| E23 | 29957261.617736503 | `=$D23%*E$16` |
| E24 | 1111045795.9507718 | `=E17+E16` |
| E25 | 87685373.21135229 | `=8%*(E24-E21-E20)` |
| E26 | 1198731169.1621242 | `=E24+E25` |
| E27 | 1198731000 | `=ROUND(E26,-3)` |

Text bang chu tai dong 28 phai noi dung tuong ung voi
`Mot ty, mot tram chin muoi tam trieu, bay tram ba muoi mot nghin dong`.

## Oracle phap ly tu DT-406

DT-406 phat hien workbook mau dang dung hao phi may `0,014` cua bien the nuoc 0,5-3 m
cho dong 22 co mo ta va khoi luong nuoc 3-12 m. VBHN97/2025-BQP trang 32-33 quy dinh
bien the 3-12 m la `0,016` cho ca `M010.008` va `M010.010`. Ket qua sau day thay the
oracle cu cho DT-406 tro di; sai lech la sua loi nghiep vu co can cu, khong phai regression.

| Cell | Expected tu DT-406 | Ghi chu |
|---|---:|---|
| Gia DT TC!H22 | 34270.8 | Don gia may dung bien the `water-3-12` |
| Gia DT TC!K19 | 509776991.89318 | Tong may nuoc, gom dong he so |
| Gia DT TC!K26 | 46343362.89938 | Chenh lech he so 1,10 cua may nuoc |
| Gia DT TC!I27 | 12540585.6195888 | Tong VL |
| Gia DT TC!J27 | 268414820.1 | Tong NC |
| Gia DT TC!K27 | 558860046.12222 | Tong M phap ly |
| THKP-TC!E13 | 839815451.841809 | Chi phi truc tiep |
| THKP-TC!E16 | 999276355.775308 | Tong sau chi phi chung |
| THKP-TC!E24 | 1111829538.55151 | Tong truoc thue |
| THKP-TC!E26 | 1199576770.00871 | Tong sau thue |
| THKP-TC!E27 | 1199577000 | `=ROUND(E26,-3)` |

Checkpoint oracle: `Dutoanmau/Checkpoints/DT-406-end.xlsm`, SHA-256
`288324623F921259C5633A4180FC9F8062AD138DA0A37718982C48F09C247B22`.

## Oracle phap ly tu DT-407

DT-407 thay cong thuc tong hop legacy bang Bieu 04 VBHN97/2025 trang 51. K5 phai tinh
tren Z, khong phai T; ty le nong nghiep va moi truong cho Z duoi 10 ty theo bang hien hanh
trong package 2.0.1 la 2.598%, khong phai 2.873% cu.

| Cell/khoan | Expected tu DT-407 | Ghi chu |
|---|---:|---|
| THKP-TC!E13 (T) | 839815452 | VL/NC/M lam tron VND |
| THKP-TC!E14 (C) | 107365928 | 40% NC |
| THKP-TC!E15 (TL) | 52094976 | 5.5% cua T+C |
| THKP-TC!E16 (Z) | 999276356 | T+C+TL |
| THKP-TC!D22 | 2.598 | K5 nong nghiep duoi 10 ty |
| THKP-TC!E22 | 25961200 | 2.598% cua Z |
| THKP-TC!E17 (K) | 114386486 | Tong K1-K6 |
| THKP-TC!E24 (Q) | 1113662842 | Z+K |
| THKP-TC!E25 (VAT) | 87893896 | 8% cua Q-K3-K4 |
| THKP-TC!E26 (H) | 1201556738 | Q+VAT |
| THKP-TC!E27 | 1201557000 | Lam tron 1.000 VND AwayFromZero |

Checkpoint oracle: `Dutoanmau/Checkpoints/DT-407-end.xlsm`, SHA-256
`1D878CA03CFA0B623FBA75C59C9B7E1D9E7F9163EDEEC19BBFA295FB551D2325`.

Checkpoint audit DT-502: `Dutoanmau/Checkpoints/DT-502-end.xlsm`, 679410 byte,
SHA-256 `60AFD5F96B102BE99B38E7169B2423D7EE387FA409CCC1BB7DC014BA452F5704`.
Workbook co 22 entry audit, 92 loi ke thua va `THKP-TC!E27=1201557000`.

## Expected output - Ho dao

Luong ho dao la random co rang buoc, vi vay khong so sanh tung so ngau nhien. Baseline test
so sanh invariant sau:

- Cong thuc: `V=((d1*r1)+(d2*r2)+SQRT((d1*r1)*(d2*r2)))*H/3`.
- Mau kiem tra core: `d1=2`, `r1=1.5`, `d2=1.2`, `r2=0.8`, `H=3` thi
  `V=5.657056...`, dung sai `0.0001` voi expected `5.6571`.
- 3m mac dinh: d1 `1.7..2.3`, r1 `1.2..1.8`, d2 `1.1..1.5`, r2
  `0.6..1.0`, H `0.1..10`; V co the tao `0.1286781592..27.1066238629`.
- 5m mac dinh: d1 `2.7..3.3`, r1 `2.2..2.8`, d2 `1.3..1.7`, r2
  `0.8..1.2`, H `0.1..10`; V co the tao `0.3155159274..52.0720420121`.
- `MaxAttempts=10000`; `MaxParallelWorkers=0` nghia la tu dong theo logical processor.
- Moi tin hieu phai co ket qua; neu tong dich nam ngoai `N*Vmin..N*Vmax` thi dung truoc
  khi ghi Excel va bao khoang de xuat. Khong duoc bo qua dong sau khi het lan thu.
- Output co mot dong header; ket qua bat dau ngay ben duoi. Cac key dac biet
  `TH3,V3,D1_3,R1_3,D2_3,R2_3,H_3,TH5,V5,D1_5,R1_5,D2_5,R2_5,H_5`
  chi ghi khi `Cot trong dao dap` co dia chi. Mapping phu link cell cung tin hieu tu cot
  data sang cot ket qua. `KEY_STT` danh so thu tu.

## Loi ton tai truoc DT-101

Cac loi sau da co san trong baseline. Gate regression chi that bai neu co loi moi ngoai
danh sach nay hoac gia tri expected o tren thay doi:

- `PL01!T47`: `#REF!`.
- `THKL!N17`, `THKL!N20`: `#REF!`.
- `KLCT1!R28`, `R31`, `R35`, `R38`: `#REF!`.
- `Gia DT TC!A3`: `#REF!`.
- `DG Can!K35`, `DG Can!K52`: cong thuc co `#REF!`.
- `DG Nuoc!K6`, `DG Nuoc!K9`: cong thuc/ket qua co `#REF!`.
- `ChiPhi!O10:O12`: `#N/A`; `J20`, `I22`, `G30`, `H30`: `#REF!`; `J22`: `#N/A`.
- `THNK_C!V7:W10`: `#REF!`.
- `Thoi gian!H5:H7`, `Thoi gian!E11`: `#REF!`.
- `PTVT`: nhieu cell ke thua co `#REF!`; chua thuoc pham vi luong du toan giai doan B1.
- Defined names: xem thong ke o phan danh sach sheet; khong coi cac name rac la loi moi.

## Quy tac recalculate va cong cu

- Expected value trong tai lieu nay la cached value do Excel ghi vao file va gia tri hien
  tren workbook live luc capture.
- Thu vien render Open XML da tu tinh lai mot so cong thuc dong va cho `THKP-TC!E27`
  thanh `1172694000`, khong khop Excel live/cached `1198731000`. Ket qua cua render engine
  chi dung de xem bo cuc, khong dung lam oracle tinh toan.
- Task co thay doi cong thuc phai duoc recalculate trong Microsoft Excel, luu, mo lai va
  so sanh cached value sau cung.
- Khi so sanh so double khong lam tron: dung sai tuyet doi `1e-6` neu task khong quy dinh
  khac. Cell tien hien thi so nguyen phai khop gia tri lam tron cua workbook.
