#!/usr/bin/env python3
"""Build the resolved BQP-RPBM-2025 source snapshot from the verified 2021 package."""

from __future__ import annotations

import argparse
import csv
import shutil
from pathlib import Path

from norm_catalog_data import NORM_DATA


RECORD_COLUMNS = [
    "key",
    "recordType",
    "unit",
    "title",
    "data",
    "sourceDocumentId",
    "pageFrom",
    "pageTo",
    "section",
    "verification",
]

SOURCE_COLUMNS = [
    "documentId",
    "title",
    "publisher",
    "issuedDate",
    "effectiveFrom",
    "effectiveTo",
    "officialUri",
    "contentChecksum",
]

PROCESS_START_PAGES = {
    1: 3, 2: 3, 3: 3, 4: 3, 5: 4, 6: 5, 7: 7, 8: 7, 9: 8,
    10: 9, 11: 10, 12: 10, 13: 13, 14: 13, 15: 15, 16: 16,
    17: 16, 18: 17, 19: 17, 20: 19, 21: 19, 22: 20, 23: 20,
    24: 21, 25: 21, 26: 22, 27: 22, 28: 23, 29: 24, 30: 24,
    31: 25, 32: 25, 33: 25, 34: 25, 35: 26, 36: 26, 37: 27,
    38: 27, 39: 28, 40: 29, 41: 29, 42: 30, 43: 30, 44: 31,
    45: 31, 46: 31, 47: 32, 48: 32, 49: 33, 50: 34, 51: 35,
    52: 35, 53: 35, 54: 36, 55: 36, 56: 37, 57: 37, 58: 38,
    59: 38,
}

NORM_START_PAGES = {
    "NORM-000.0100": 18,
    "NORM-000.0200": 19,
    "NORM-000.0300": 19,
    "NORM-000.0400": 20,
    "NORM-010.0100": 20,
    "NORM-010.0200": 21,
    "NORM-010.0300": 21,
    "NORM-010.0400": 21,
    "NORM-020.0100": 22,
    "NORM-020.0200": 23,
    "NORM-020.0300": 23,
    "NORM-020.0400": 24,
    "NORM-020.0500": 24,
    "NORM-020.0600": 25,
    "NORM-020.0700": 26,
    "NORM-020.0800": 26,
    "NORM-020.0900": 27,
    "NORM-020.1000": 28,
    "NORM-020.1100": 29,
    "NORM-020.1200": 29,
    "NORM-030.0100": 30,
    "NORM-030.0200": 31,
    "NORM-030.0300": 32,
    "NORM-030.0400": 33,
    "NORM-030.0500": 33,
    "NORM-030.0600": 34,
    "NORM-030.0700": 35,
    "NORM-030.0800": 36,
    "NORM-040.0100": 37,
    "NORM-040.0200": 37,
    "NORM-040.0300": 38,
    "NORM-040.0400": 39,
    "NORM-040.0500": 39,
    "NORM-040.0600": 40,
}

MACHINE_PAGES = {
    **{index: 15 for index in range(1, 3)},
    **{index: 16 for index in range(3, 11)},
    **{index: 17 for index in range(11, 19)},
    **{index: 18 for index in range(19, 26)},
    **{index: 19 for index in range(26, 34)},
}

MACHINE_NONSTATE_PAGES = {
    **{index: 24 for index in range(1, 5)},
    **{index: 25 for index in range(5, 12)},
    **{index: 26 for index in range(12, 19)},
    **{index: 27 for index in range(19, 24)},
    **{index: 28 for index in range(24, 31)},
    **{index: 29 for index in range(31, 34)},
}

MACHINE_PRICES = {
    1: "129877200",
    2: "303046800",
    3: "595270500",
    4: "613644600",
    7: "178707000",
    8: "303046800",
    16: "4115480000",
    17: "6136523020",
    19: "511242972",
}

MACHINE_CREWS = {
    11: "6-si-quan+20-thuy-thu",
    12: "4-si-quan+16-thuy-thu",
    13: "4-si-quan+14-thuy-thu",
    14: "3-si-quan+8-thuy-thu",
    15: "2-si-quan+6-thuy-thu",
    22: "1-si-quan+2-thuy-thu",
}

PROVINCES = [
    (1, "LANG-SON", "Lạng Sơn", "4,3,1", 1, 5, 5),
    (2, "CAO-BANG", "Cao Bằng", "4,3,1", 1, 5, 5),
    (3, "BAC-NINH", "Bắc Ninh", "2,1", 1, 5, 5),
    (4, "THAI-NGUYEN", "Thái Nguyên", "2,1", 1, 5, 6),
    (5, "TUYEN-QUANG", "Tuyên Quang", "4,3,1", 1, 6, 6),
    (6, "LAO-CAI", "Lào Cai", "3,1", 1, 6, 6),
    (7, "LAI-CHAU", "Lai Châu", "3,1", 1, 6, 6),
    (8, "DIEN-BIEN", "Điện Biên", "3,1", 1, 6, 6),
    (9, "SON-LA", "Sơn La", "1", 1, 6, 6),
    (10, "PHU-THO", "Phú Thọ", "1", 1, 6, 6),
    (11, "QUANG-NINH", "Quảng Ninh", "3,2,1", 1, 6, 7),
    (12, "HAI-PHONG", "Thành phố Hải Phòng", "2,1", 1, 7, 7),
    (13, "HUNG-YEN", "Hưng Yên", "2,1", 1, 7, 7),
    (14, "NINH-BINH", "Ninh Bình", "2,1", 1, 7, 7),
    (15, "THANH-HOA", "Thanh Hóa", "2,1", 1, 7, 8),
    (16, "NGHE-AN", "Nghệ An", "2,1", 1, 8, 8),
    (17, "HA-TINH", "Hà Tĩnh", "2,1", 1, 8, 9),
    (18, "QUANG-TRI", "Quảng Trị", "4,3,2", 3, 9, 9),
    (19, "HUE", "Thành phố Huế", "4,2,1", 1, 9, 9),
    (20, "DA-NANG", "Thành phố Đà Nẵng", "3,2,1", 1, 9, 10),
    (21, "QUANG-NGAI", "Quảng Ngãi", "3,2,1", 1, 10, 10),
    (22, "GIA-LAI", "Gia Lai", "3,2,1", 1, 10, 11),
    (23, "DAK-LAK", "Đắk Lắk", "3,2,1", 1, 11, 11),
    (24, "KHANH-HOA", "Khánh Hòa", "3,2,1", 1, 11, 11),
    (25, "LAM-DONG", "Lâm Đồng", "2,1", 1, 11, 11),
    (26, "DONG-NAI", "Đồng Nai", "2,1", 1, 12, 12),
    (27, "TAY-NINH", "Tây Ninh", "3,2,1", 1, 12, 12),
    (28, "HO-CHI-MINH", "Thành phố Hồ Chí Minh", "3,2,1", 1, 12, 12),
    (29, "DONG-THAP", "Đồng Tháp", "3,2,1", 1, 12, 13),
    (30, "CAN-THO", "Cần Thơ", "3,2,1", 1, 13, 13),
    (31, "VINH-LONG", "Vĩnh Long", "2,1", 1, 13, 13),
    (32, "AN-GIANG", "An Giang", "2,1", 1, 13, 13),
    (33, "CA-MAU", "Cà Mau", "2,1", 1, 14, 14),
    (34, "HA-NOI", "Thành phố Hà Nội", "1", 1, 14, 14),
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repository-root", type=Path, default=Path(__file__).resolve().parents[2])
    return parser.parse_args()


def read_tsv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        return list(csv.DictReader(stream, delimiter="\t"))


def write_tsv(path: Path, columns: list[str], rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=columns, delimiter="\t", lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def parse_data(value: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for part in value.split(";"):
        key, separator, item = part.partition("=")
        if not separator or not key or key in result:
            raise ValueError(f"Invalid record data: {value}")
        result[key] = item
    return result


def format_data(values: dict[str, str]) -> str:
    return ";".join(f"{key}={value}" for key, value in values.items())


def update_process(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    result = []
    for row in rows:
        item = dict(row)
        article = int(item["key"].rsplit("-", 1)[1])
        item["sourceDocumentId"] = "VBHN95-2025-BQP"
        item["pageFrom"] = str(PROCESS_START_PAGES[article])
        item["pageTo"] = str(PROCESS_START_PAGES.get(article + 1, PROCESS_START_PAGES[article]))
        if article in {5, 12, 14, 55}:
            data = parse_data(item["data"])
            data["amendedBy"] = "TT101-2025-BQP-Article-2"
            data["amendedPart"] = {
                5: "paragraphs-1-and-4",
                12: "paragraph-3-point-l-title",
                14: "paragraph-6-point-a",
                55: "paragraph-4",
            }[article]
            item["data"] = format_data(data)
        result.append(item)
    result.append(new_record(
        "PROC-TT121-FORM-RPBM-13",
        "TechnicalForm",
        "",
        "Mẫu RPBM-13 - Biên bản bàn giao mặt bằng đã thi công rà phá bom mìn vật nổ",
        "form=RPBM-13;replacedBy=TT101-2025-BQP-Article-2.5",
        "VBHN95-2025-BQP",
        89,
        90,
        "Phụ lục III, Mẫu RPBM-13",
    ))
    return result


def update_norms(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    ordered_keys = [row["key"] for row in rows]
    result = []
    for index, row in enumerate(rows):
        item = dict(row)
        key = item["key"]
        item["sourceDocumentId"] = "VBHN97-2025-BQP"
        item["pageFrom"] = str(NORM_START_PAGES[key])
        next_page = NORM_START_PAGES[ordered_keys[index + 1]] if index + 1 < len(ordered_keys) else 42
        item["pageTo"] = str(max(NORM_START_PAGES[key], next_page))
        item["data"] = NORM_DATA[key]
        if key == "NORM-020.0500":
            item["title"] = (
                "Rà phá bom mìn vật nổ bằng máy dò bom ở độ sâu đến 0,5 m, "
                "đến 1 m, đến 3 m, đến 5 m hoặc đến 10 m"
            )
        result.append(item)

    rates = [
        ("STATE", "state-budget", 1, "51.2", "75.2"),
        ("STATE", "state-budget", 2, "57.2", "80.4"),
        ("STATE", "state-budget", 3, "65.3", "88.8"),
        ("STATE", "state-budget", 4, "81.1", "92.4"),
        ("NONSTATE", "non-state-budget", 1, "68.8", "106.2"),
        ("NONSTATE", "non-state-budget", 2, "75.2", "115.1"),
        ("NONSTATE", "non-state-budget", 3, "83.1", "126.7"),
        ("NONSTATE", "non-state-budget", 4, "95.1", "132.1"),
    ]
    for key_group, funding, zone, land, water in rates:
        result.append(new_record(
            f"NORM-PROVISIONAL-{key_group}-ZONE-{zone}",
            "ProvisionalEstimateRate",
            "million-vnd-per-ha",
            f"Suất dự toán tạm tính khu vực {zone} - {funding}",
            f"funding={funding};zone={zone};landRate={land};underwaterTo12mRate={water}",
            "VBHN97-2025-BQP",
            42,
            42,
            "Phụ lục I, suất dự toán tạm tính theo héc ta",
        ))
    return result


def update_costs(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    k5_rates = {
        "COST-K5-CIVIL": "3.285,2.853,2.435,1.845,1.546,1.188,0.797,0.694,0.620,0.530,0.478",
        "COST-K5-INDUSTRIAL": "3.508,3.137,2.559,2.074,1.604,1.301,0.823,0.716,0.640,0.550,0.493",
        "COST-K5-TRANSPORT": "3.203,2.700,2.356,1.714,1.272,1.003,0.731,0.636,0.550,0.480,0.438",
        "COST-K5-AGRICULTURE": "2.598,2.292,2.075,1.545,1.189,0.950,0.631,0.550,0.490,0.420,0.378",
        "COST-K5-INFRASTRUCTURE": "2.566,2.256,1.984,1.461,1.142,0.912,0.584,0.509,0.452,0.390,0.350",
    }
    result = []
    for row in rows:
        item = dict(row)
        key = item["key"]
        item["sourceDocumentId"] = "VBHN97-2025-BQP"
        if key.startswith("COST-DIRECT-"):
            page_from, page_to = 43, 43
        elif key == "COST-COMMON" or key.startswith("COST-K1K4-"):
            page_from, page_to = 44, 44
        elif key.startswith("COST-K2-") or key.startswith("COST-K3-") or key.startswith("COST-K5-"):
            page_from, page_to = 45, 45
        elif key.startswith("COST-TEMPLATE-"):
            template = int(key[-2:])
            page_from = page_to = 47 + template
        else:
            page_from, page_to = 46, 46
        item["pageFrom"], item["pageTo"] = str(page_from), str(page_to)
        if key == "COST-K2-LINEAR":
            item["data"] = (
                "code=K2;base=T;thresholdBillion=15,100,500,1000;"
                "ratesPercent=2.2,2.0,1.9,1.8,1.7;boundary=upper-inclusive;"
                "interpolation=step;rounding=nearest-vnd-away-from-zero;"
                "currentExternalBasis=TT36-2026-BXD-PL3-table-3.7"
            )
        elif key == "COST-K2-OTHER":
            item["data"] = (
                "code=K2;base=T;thresholdBillion=15,100,500,1000;"
                "ratesPercent=1.1,1.0,0.95,0.9,0.85;boundary=upper-inclusive;"
                "interpolation=step;rounding=nearest-vnd-away-from-zero;"
                "currentExternalBasis=TT36-2026-BXD-PL3-table-3.7"
            )
        elif key in k5_rates:
            data = parse_data(item["data"])
            data["thresholdBillion"] = "10,20,50,100,200,500,1000,2000,5000,8000,10000"
            data["ratesPercent"] = k5_rates[key]
            data["boundary"] = "first-upper-inclusive"
            data["interpolation"] = "linear"
            data["aboveMaximum"] = "external-estimate"
            data["rounding"] = "nearest-vnd-away-from-zero"
            data["currentExternalBasis"] = "TT38-2026-BXD-PL8-table-2.24"
            item["data"] = format_data(data)
        elif item["recordType"] in {"CostFormula", "CostRate", "CostMinimum"}:
            data = parse_data(item["data"])
            data["rounding"] = "nearest-vnd-away-from-zero"
            item["data"] = format_data(data)
        result.append(item)

    result.extend([
        new_record(
            "COST-MIN-SURVEY-DESIGN-UNDER-4HA",
            "CostMinimum",
            "vnd",
            "Chi phí tối thiểu khảo sát, lập phương án kỹ thuật thi công và dự toán đến 4 ha",
            "maxAreaHa=4;teamMembers=18;allowancePerPersonDayVnd=350000;minimumVnd=6300000",
            "VBHN97-2025-BQP",
            44,
            44,
            "Phụ lục II, Phần I, mục III.1, mức tối thiểu",
        ),
        new_record(
            "COST-MIN-DISPOSAL-UNDER-4HA",
            "CostMinimum",
            "vnd",
            "Chi phí tối thiểu hủy bom mìn vật nổ sau dò tìm đến 4 ha",
            "maxAreaHa=4;teamMembers=18;allowancePerPersonDayVnd=350000;minimumVnd=6300000",
            "VBHN97-2025-BQP",
            46,
            46,
            "Phụ lục II, Phần I, mục III.5, mức tối thiểu",
        ),
        new_record(
            "COST-TEMPLATE-05",
            "EstimateTemplate",
            "vnd",
            "Thuyết minh dự toán rà phá bom mìn vật nổ",
            "template=05;parts=legal-basis,estimate-basis,estimate-value;attachments=summary-and-detail",
            "VBHN97-2025-BQP",
            52,
            52,
            "Phụ lục II, Phần II, Biểu mẫu 05",
        ),
    ])
    return result


def update_machines(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    result = []
    for row in rows:
        item = dict(row)
        number = int(item["key"].rsplit(".", 1)[1])
        item["sourceDocumentId"] = "VBHN96-2025-BQP"
        item["pageFrom"] = str(MACHINE_PAGES[number])
        item["pageTo"] = str(MACHINE_NONSTATE_PAGES[number])
        item["section"] = (
            f"Phụ lục II, Bảng 01/03, M010.{number:03d}/M011.{number:03d}"
        )
        data = parse_data(item["data"])
        if number in MACHINE_PRICES:
            data["referencePriceVnd"] = MACHINE_PRICES[number]
        if number in MACHINE_CREWS:
            data["operators"] = MACHINE_CREWS[number]
        if number == 20:
            data["operators"] = "none"
            data["nonStateOperators"] = "1-bac-5-10"
        if number == 24:
            data["nonStateReferencePriceVnd"] = "350000"
        if number == 27:
            data["fuel"] = "60-lit-gasoline-E5-RON-92-II"
        data["stateCode"] = f"M010.{number:03d}"
        data["nonStateCode"] = f"M011.{number:03d}"
        item["data"] = format_data(data)
        if number == 7:
            item["title"] = "Máy dò mìn dưới nước, MW 1630B là đại diện"
        if number == 26:
            item["title"] = "Máy điểm hỏa"
        result.append(item)
    return result


def update_geography(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    result = []
    for row in rows:
        item = dict(row)
        key = item["key"]
        item["sourceDocumentId"] = "VBHN97-2025-BQP"
        if key.startswith("GEO-DENSITY-"):
            page_from, page_to = 5, 5
        elif key.startswith("GEO-LAND-ZONE-"):
            page_from, page_to = 5, 14
        elif key.startswith("GEO-SEA-ZONE-"):
            page_from, page_to = 14, 14
        elif key.startswith("GEO-LAND-SIGNAL-"):
            page_from, page_to = 14, 14
        elif key.startswith("GEO-WATER-SIGNAL-") or key.startswith("GEO-SEA-SIGNAL-"):
            page_from, page_to = 15, 15
        elif key.startswith("GEO-FOREST-"):
            page_from, page_to = 15, 16
        else:
            page_from, page_to = 16, 18
        item["pageFrom"], item["pageTo"] = str(page_from), str(page_to)
        if key == "GEO-SEA-ZONE-2":
            item["data"] = (
                "priority=2;provinces=Thanh-Hoa,Nghe-An,Hue,Gia-Lai,Da-Nang,"
                "Ho-Chi-Minh,Ca-Mau,An-Giang"
            )
        elif key == "GEO-SEA-ZONE-3":
            item["data"] = "priority=3;provinces=Quang-Ninh,Hai-Phong,Ha-Tinh,Quang-Tri"
        result.append(item)

    for number, slug, province, zones, fallback, page_from, page_to in PROVINCES:
        result.append(new_record(
            f"GEO-LOCALITY-{slug}",
            "LocalityZoneAssignment",
            "locality",
            f"Phân vùng địa bàn {province}",
            (
                f"tableRow={number};province={province};zones={zones};fallbackZone={fallback};"
                "resolution=explicit-locality-first-then-fallback;conflictRule=highest-zone-priority"
            ),
            "VBHN97-2025-BQP",
            page_from,
            page_to,
            f"Phụ lục I, Phần I, mục 4a, dòng {number} - {province}",
        ))
    return result


def build_compliance() -> list[dict[str, str]]:
    return [
        new_record(
            "COMPLIANCE-TT121-EFFECTIVE", "EffectiveRule", "date",
            "Hiệu lực quy trình kỹ thuật đã hợp nhất",
            "baseEffectiveFrom=2021-11-05;amendmentEffectiveFrom=2025-10-28;resolvedDocument=VBHN95-2025-BQP",
            "VBHN95-2025-BQP", 1, 2, "Phần mở đầu văn bản hợp nhất 95/VBHN-BQP",
        ),
        new_record(
            "COMPLIANCE-TT122-EFFECTIVE", "EffectiveRule", "date",
            "Hiệu lực đơn giá ca máy đã hợp nhất",
            "baseEffectiveFrom=2021-11-05;amendmentEffectiveFrom=2025-10-28;resolvedDocument=VBHN96-2025-BQP",
            "VBHN96-2025-BQP", 1, 2, "Phần mở đầu văn bản hợp nhất 96/VBHN-BQP",
        ),
        new_record(
            "COMPLIANCE-TT123-EFFECTIVE", "EffectiveRule", "date",
            "Hiệu lực định mức và quản lý chi phí đã hợp nhất",
            "baseEffectiveFrom=2021-11-05;amendmentEffectiveFrom=2025-10-28;resolvedDocument=VBHN97-2025-BQP",
            "VBHN97-2025-BQP", 1, 2, "Phần mở đầu văn bản hợp nhất 97/VBHN-BQP",
        ),
        new_record(
            "COMPLIANCE-TT101-TRANSITION", "TransitionRule", "date",
            "Chuyển tiếp phương án kỹ thuật thi công và dự toán đã phê duyệt",
            "cutoverDate=2025-10-28;approvedBeforeCutover=keep-approved-design-and-estimate;ruleId=BQP-TT101-2025-ARTICLE-6",
            "TT101-2025-BQP", 17, 18, "Điều 6",
        ),
        new_record(
            "COMPLIANCE-TT101-EFFECTIVE", "EffectiveRule", "date",
            "Hiệu lực Thông tư 101/2025/TT-BQP",
            "effectiveFrom=2025-10-28;status=effective",
            "TT101-2025-BQP", 18, 18, "Điều 7",
        ),
        new_record(
            "COMPLIANCE-VBHN94-GOVERNANCE", "GovernanceRule", "document",
            "Quản lý hoạt động khắc phục hậu quả bom mìn vật nổ đã hợp nhất",
            "resolvedDocument=VBHN94-2025-BQP;baseDocument=TT195-2019-BQP;amendedBy=TT101-2025-BQP",
            "VBHN94-2025-BQP", 1, 3, "Phần mở đầu và Điều 1",
        ),
        new_record(
            "COMPLIANCE-QCVN01-2022-RESOLVED", "TechnicalComplianceRule", "document",
            "QCVN 01:2022/BQP đã hợp nhất với sửa đổi năm 2025",
            "resolvedDocument=VBHN98-2025-BQP;baseStandard=QCVN-01-2022-BQP;amendedBy=TT101-2025-BQP-Article-5",
            "VBHN98-2025-BQP", 1, 6, "Phần mở đầu và QCVN 01:2022/BQP",
        ),
    ]


def new_record(
    key: str,
    record_type: str,
    unit: str,
    title: str,
    data: str,
    source_document_id: str,
    page_from: int,
    page_to: int,
    section: str,
) -> dict[str, str]:
    return {
        "key": key,
        "recordType": record_type,
        "unit": unit,
        "title": title,
        "data": data,
        "sourceDocumentId": source_document_id,
        "pageFrom": str(page_from),
        "pageTo": str(page_to),
        "section": section,
        "verification": "VerifiedAgainstOfficialSource",
    }


def build_sources() -> list[dict[str, str]]:
    common = {"publisher": "Bộ Quốc phòng", "effectiveFrom": "2025-10-28", "effectiveTo": ""}
    return [
        {
            **common,
            "documentId": "TT101-2025-BQP",
            "title": "Thông tư 101/2025/TT-BQP sửa đổi, bổ sung các Thông tư về điều tra, khảo sát, rà phá bom mìn vật nổ",
            "issuedDate": "2025-09-13",
            "officialUri": "https://mod.gov.vn/wcm/connect/ead80087-4eed-4bdd-b739-0961956cf8e2/TT101.2025BQP.pdf?CACHEID=ROOTWORKSPACE-ead80087-4eed-4bdd-b739-0961956cf8e2-pBxSjyH&MOD=AJPERES",
            "contentChecksum": "E445F63CFDB515CB187780247423610A5A29DD0BBFDB231EBC21D858516DD240",
        },
        *[
            {
                **common,
                "documentId": f"VBHN{number}-2025-BQP",
                "title": title,
                "issuedDate": "2025-11-25",
                "officialUri": f"https://vbpl.vn/TW/Pages/vbpq-thuoctinh-hopnhat.aspx?ItemID={item_id}",
                "contentChecksum": checksum,
            }
            for number, item_id, title, checksum in [
                (94, 184657, "Văn bản hợp nhất 94/VBHN-BQP về quản lý hoạt động khắc phục hậu quả bom mìn vật nổ", "D375AAC8F0C0B016D0C8B1E8CC65140560DF23D13AB3B499DE581DC5276F8935"),
                (95, 184660, "Văn bản hợp nhất 95/VBHN-BQP về Quy trình kỹ thuật điều tra, khảo sát, rà phá bom mìn vật nổ", "2D5207B562EC63A24BD0B9D38C86E748AF6F7A4009195250517B83EFE22A1094"),
                (96, 184661, "Văn bản hợp nhất 96/VBHN-BQP về đơn giá ca máy và thiết bị thi công rà phá bom mìn vật nổ", "10A25B90BC95F8EB8330C9CD6D2423AD0DF46746D95DB7D33C7A2157228AE0B1"),
                (97, 184663, "Văn bản hợp nhất 97/VBHN-BQP về định mức dự toán và quản lý chi phí rà phá bom mìn vật nổ", "92B6C42996E554A908661B0395CE7286D41DDD69C17859B40993DC46300B69E0"),
                (98, 184665, "Văn bản hợp nhất 98/VBHN-BQP về QCVN 01:2022/BQP", "62CC32BDA9ADED4310AB183180D6D4D861E41FEA94852848425EA0C19BAFE0A2"),
            ]
        ],
    ]


def reason_for(module: str, old: dict[str, str] | None, new: dict[str, str]) -> str:
    key = new["key"]
    specific = {
        "PROC-TT121-FORM-RPBM-13": "TT101 Điều 2 khoản 5 thay Mẫu RPBM-13.",
        "NORM-020.0500": "TT101 Phụ lục III khoản 2 sửa chiều sâu và hao phí định mức 020.0500.",
        "COST-K2-LINEAR": "TT101 Phụ lục III khoản 5 giảm tỷ lệ K2 cho dự án theo tuyến.",
        "COST-K2-OTHER": "TT101 Phụ lục III khoản 5 giảm tỷ lệ K2 cho các dự án còn lại.",
        "COST-MIN-SURVEY-DESIGN-UNDER-4HA": "TT101 Phụ lục III khoản 4 bổ sung mức tối thiểu 6.300.000 đồng.",
        "COST-MIN-DISPOSAL-UNDER-4HA": "TT101 Phụ lục III khoản 6 bổ sung mức tối thiểu 6.300.000 đồng.",
        "COST-TEMPLATE-05": "TT101 Phụ lục III khoản 7 bổ sung Biểu mẫu 05.",
        "GEO-SEA-ZONE-2": "TT101 Phụ lục III khoản 1 thay danh sách tỉnh, thành phố khu vực 2 dưới biển.",
        "GEO-SEA-ZONE-3": "TT101 Phụ lục III khoản 1 thay danh sách tỉnh, thành phố khu vực 3 dưới biển.",
        "MACHINE-M010.026": "Phụ lục IV TT101 xác nhận tên thiết bị M010.026 là Máy điểm hỏa.",
        "COMPLIANCE-TT101-TRANSITION": "TT101 Điều 6 bổ sung quy tắc chuyển tiếp.",
        "COMPLIANCE-TT101-EFFECTIVE": "TT101 Điều 7 xác lập ngày hiệu lực 28/10/2025.",
        "COMPLIANCE-TT121-EFFECTIVE": "95/VBHN-BQP hợp nhất hiệu lực nền TT121 với sửa đổi TT101 từ 28/10/2025.",
        "COMPLIANCE-TT122-EFFECTIVE": "96/VBHN-BQP hợp nhất hiệu lực nền TT122 với sửa đổi TT101 từ 28/10/2025.",
        "COMPLIANCE-TT123-EFFECTIVE": "97/VBHN-BQP hợp nhất hiệu lực nền TT123 với sửa đổi TT101 từ 28/10/2025.",
        "COMPLIANCE-VBHN94-GOVERNANCE": "Bổ sung căn cứ quản lý đã hợp nhất tại 94/VBHN-BQP.",
        "COMPLIANCE-QCVN01-2022-RESOLVED": "TT101 Điều 5 và 98/VBHN-BQP hợp nhất sửa đổi QCVN 01:2022/BQP.",
    }
    if key in specific:
        return specific[key]
    if key.startswith("NORM-PROVISIONAL-"):
        return "TT101 Phụ lục III khoản 3 bổ sung suất dự toán tạm tính theo héc ta."
    if key.startswith("GEO-LOCALITY-"):
        return "TT101 Phụ lục III khoản 1 thay toàn bộ bảng phân vùng địa bàn sau sắp xếp hành chính."
    if key.startswith("MACHINE-M010.") and old and old["data"] != new["data"]:
        return "TT101 Phụ lục IV thay dữ liệu cơ bản xác định giá ca máy."
    if key.startswith("PROC-TT121-DIEU-") and old and old["data"] != new["data"]:
        return "TT101 Điều 2 sửa nội dung điều tương ứng trong Quy trình kỹ thuật."
    if old is None:
        return "Bổ sung bản ghi từ snapshot pháp lý hợp nhất hiệu lực 28/10/2025."
    return f"Đổi nguồn truy vết {module} sang văn bản hợp nhất hiệu lực 28/10/2025."


def build_audit(
    old_modules: dict[str, list[dict[str, str]]],
    new_modules: dict[str, list[dict[str, str]]],
) -> list[dict[str, str]]:
    rows = []
    compare_fields = [column for column in RECORD_COLUMNS if column != "key"]
    for module in sorted(new_modules):
        old_by_key = {row["key"]: row for row in old_modules[module]}
        for new in sorted(new_modules[module], key=lambda row: row["key"]):
            old = old_by_key.get(new["key"])
            changed = old is None or any(old[field] != new[field] for field in compare_fields)
            if not changed:
                continue
            changed_fields = "added" if old is None else ",".join(
                field for field in compare_fields if old[field] != new[field]
            )
            rows.append({
                "module": module,
                "key": new["key"],
                "changeType": "Added" if old is None else "Changed",
                "changedFields": changed_fields,
                "reason": reason_for(module, old, new),
                "sourceDocumentId": new["sourceDocumentId"],
                "pageFrom": new["pageFrom"],
                "pageTo": new["pageTo"],
                "reviewStatus": "VerifiedAgainstOfficialSource",
            })
    return rows


def main() -> int:
    args = parse_args()
    root = args.repository_root.resolve()
    old_source = root / "data/regulations/packages/BQP-RPBM-2021/source"
    package_root = root / "data/regulations/packages/BQP-RPBM-2025"
    source_root = package_root / "source"
    old_properties = old_source / "package.properties"
    if not old_properties.is_file() or "packageId=BQP-RPBM-2021" not in old_properties.read_text(encoding="utf-8"):
        raise RuntimeError(f"Invalid repository root or 2021 source: {root}")
    expected_source = (root / "data/regulations/packages/BQP-RPBM-2025/source").resolve()
    if source_root.resolve() != expected_source or source_root.name != "source":
        raise RuntimeError(f"Unsafe output path: {source_root}")
    if source_root.exists():
        shutil.rmtree(source_root)
    module_source = old_source / "modules"
    old_modules = {
        path.stem: read_tsv(path)
        for path in sorted(module_source.glob("*.tsv"))
    }
    new_modules = {
        "TechnicalProcess": update_process(old_modules["TechnicalProcess"]),
        "Norm": update_norms(old_modules["Norm"]),
        "CostRule": update_costs(old_modules["CostRule"]),
        "MachineRate": update_machines(old_modules["MachineRate"]),
        "Geography": update_geography(old_modules["Geography"]),
        "Compliance": build_compliance(),
    }

    source_root.mkdir(parents=True, exist_ok=True)
    (source_root / "package.properties").write_text(
        "\n".join([
            "packageId=BQP-RPBM-2025",
            "dataVersion=2.0.1",
            "effectiveFrom=2025-10-28",
            "effectiveTo=",
            "status=Published",
            "transitionNote=Ho so co phuong an ky thuat thi cong va du toan duoc phe duyet truoc 28/10/2025 tiep tuc theo ban da phe duyet; ho so con lai ap dung snapshot hop nhat 2025.",
            "",
        ]),
        encoding="utf-8",
    )
    write_tsv(source_root / "sources.tsv", SOURCE_COLUMNS, build_sources())
    for module, rows in new_modules.items():
        write_tsv(source_root / "modules" / f"{module}.tsv", RECORD_COLUMNS, rows)

    audit_columns = [
        "module", "key", "changeType", "changedFields", "reason",
        "sourceDocumentId", "pageFrom", "pageTo", "reviewStatus",
    ]
    write_tsv(
        package_root / "audit" / "change-reasons.tsv",
        audit_columns,
        build_audit(old_modules, new_modules),
    )
    counts = ", ".join(f"{module}={len(rows)}" for module, rows in sorted(new_modules.items()))
    print(f"Built {source_root}")
    print(counts)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
