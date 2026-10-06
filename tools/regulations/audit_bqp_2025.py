#!/usr/bin/env python3
"""Independent source-level audit for the BQP-RPBM-2025 package."""

from __future__ import annotations

import csv
import hashlib
import json
import re
import sys
import unicodedata
from collections import Counter
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]
PACKAGE_ROOT = REPO_ROOT / "data" / "regulations" / "packages" / "BQP-RPBM-2025"
SOURCE_ROOT = PACKAGE_ROOT / "source"
MODULE_ROOT = SOURCE_ROOT / "modules"
RAW_ROOT = REPO_ROOT / "data" / "regulations" / "raw" / "2025"
TEXT_ROOT = REPO_ROOT / "data" / "regulations" / "text" / "2025"
REPORT_PATH = PACKAGE_ROOT / "audit" / "DT-305-independent-audit.json"

EXPECTED_PACKAGE_CHECKSUM = "E2537C08E7156414585BEFC446E74DD51E79AB8D67A3774B1F5250E7EECCDA90"
EXPECTED_HEADER = [
    "key", "recordType", "unit", "title", "data", "sourceDocumentId",
    "pageFrom", "pageTo", "section", "verification",
]
EXPECTED_COUNTS = {
    "TechnicalProcess": 60,
    "Norm": 42,
    "CostRule": 37,
    "MachineRate": 33,
    "Geography": 69,
    "Compliance": 7,
}
EXPECTED_TYPES = {
    "TechnicalProcess": {"Article": 59, "TechnicalForm": 1},
    "Norm": {"NormCatalog": 34, "ProvisionalEstimateRate": 8},
    "CostRule": {
        "CostComponent": 4, "CostFormula": 2, "CostRate": 13,
        "CostRateTable": 7, "CostLimit": 1, "ExternalRule": 3,
        "EstimateTemplate": 5, "CostMinimum": 2,
    },
    "MachineRate": {"MachineBaseData": 33},
    "Geography": {
        "SignalDensityClass": 5, "GeographyZone": 8,
        "SignalDensityTable": 14, "TerrainClass": 8,
        "LocalityZoneAssignment": 34,
    },
    "Compliance": {
        "EffectiveRule": 4, "TransitionRule": 1, "GovernanceRule": 1,
        "TechnicalComplianceRule": 1,
    },
}
EXPECTED_NORM_KEYS = {
    *(f"NORM-000.{n:04d}" for n in range(100, 401, 100)),
    *(f"NORM-010.{n:04d}" for n in range(100, 401, 100)),
    *(f"NORM-020.{n:04d}" for n in range(100, 1201, 100)),
    *(f"NORM-030.{n:04d}" for n in range(100, 801, 100)),
    *(f"NORM-040.{n:04d}" for n in range(100, 601, 100)),
}
EXPECTED_COST_KEYS = {
    "COST-DIRECT-MATERIAL", "COST-DIRECT-LABOR-STATE",
    "COST-DIRECT-LABOR-NONSTATE", "COST-DIRECT-MACHINE",
    "COST-DIRECT-TOTAL", "COST-COMMON",
    *(f"COST-K1K4-TERRAIN-{n:02d}" for n in range(1, 9)),
    "COST-K2-LINEAR", "COST-K2-OTHER", "COST-K3-UNDER-1",
    "COST-K3-1-TO-5", "COST-K3-FROM-5", "COST-K3-LIMITS",
    "COST-K5-CIVIL", "COST-K5-INDUSTRIAL", "COST-K5-TRANSPORT",
    "COST-K5-AGRICULTURE", "COST-K5-INFRASTRUCTURE",
    "COST-K6-UNDER-1000", "COST-K6-OVER-1000", "COST-K7-K10",
    "COST-SPECIAL-CONDITION", "COST-TAX",
    *(f"COST-TEMPLATE-{n:02d}" for n in range(1, 6)),
    "COST-MIN-SURVEY-DESIGN-UNDER-4HA", "COST-MIN-DISPOSAL-UNDER-4HA",
}


class Audit:
    def __init__(self) -> None:
        self.errors: list[str] = []
        self.warnings: list[str] = []
        self.checks = 0

    def check(self, condition: bool, message: str) -> None:
        self.checks += 1
        if not condition:
            self.errors.append(message)

    def warn(self, condition: bool, message: str) -> None:
        self.checks += 1
        if not condition:
            self.warnings.append(message)


def read_tsv(path: Path, expected_header: list[str] | None, audit: Audit) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream, delimiter="\t")
        if expected_header is not None:
            audit.check(reader.fieldnames == expected_header, f"Unexpected header: {path}")
        rows = list(reader)
    audit.check(all(None not in row for row in rows), f"Malformed TSV row: {path}")
    return rows


def parse_fields(value: str, context: str, audit: Audit) -> dict[str, str]:
    result: dict[str, str] = {}
    for item in value.split(";"):
        if "=" not in item:
            audit.check(False, f"Malformed data field {context}: {item}")
            continue
        key, field_value = item.split("=", 1)
        audit.check(bool(key) and bool(field_value), f"Blank data field {context}: {item}")
        audit.check(key not in result, f"Duplicate data field {context}: {key}")
        result[key] = field_value
    return result


def decimal_value(value: str, context: str, audit: Audit) -> Decimal:
    try:
        return Decimal(value)
    except InvalidOperation:
        audit.check(False, f"Invalid invariant decimal {context}: {value}")
        return Decimal(0)


def decimal_list(value: str, context: str, audit: Audit) -> list[Decimal]:
    return [decimal_value(item, context, audit) for item in value.split(",")]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def pdf_code_visible(text: str, code: str) -> bool:
    ascii_text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    compact = re.sub(r"\s+", "", ascii_text.upper()).replace("MO10", "M010")
    compact = re.sub(r"M010[A-Z]+(?=\d{3})", "M010", compact)
    compact = re.sub(r"[^A-Z0-9]", "", compact)
    return re.sub(r"[^A-Z0-9]", "", code.upper()) in compact


def load_manifest_checksum() -> str:
    manifest = PACKAGE_ROOT / "bundle" / "manifest.ttbmanifest"
    for line in manifest.read_text(encoding="utf-8").splitlines():
        if line.startswith("checksum="):
            return line.split("=", 1)[1]
    return ""


def audit_norm(rows: list[dict[str, str]], audit: Audit) -> None:
    catalog = [row for row in rows if row["recordType"] == "NormCatalog"]
    audit.check({row["key"] for row in catalog} == EXPECTED_NORM_KEYS,
                "Norm legal key set is incomplete or unexpected")
    for row in catalog:
        fields = parse_fields(row["data"], row["key"], audit)
        audit.check(set(fields).issubset({"variantCodes", "rates", "adjustments", "constraints"}),
                    f"Unknown norm field: {row['key']}")
        variants = fields.get("variantCodes", "").split(",")
        audit.check(bool(variants) and len(set(variants)) == len(variants),
                    f"Invalid norm variants: {row['key']}")
        resources: set[tuple[str, str, str]] = set()
        for rate in fields.get("rates", "").split("|"):
            parts = rate.split(":", 3)
            audit.check(len(parts) == 4, f"Malformed norm rate: {row['key']}:{rate}")
            if len(parts) != 4:
                continue
            identity = tuple(parts[:3])
            audit.check(identity not in resources, f"Duplicate norm resource: {row['key']}:{identity}")
            resources.add(identity)
            quantities = decimal_list(parts[3], row["key"], audit)
            audit.check(len(quantities) == len(variants), f"Norm quantity count mismatch: {row['key']}")
            audit.check(all(quantity >= 0 for quantity in quantities),
                        f"Negative norm quantity: {row['key']}")
        audit.check(bool(resources), f"Norm has no resource: {row['key']}")

        code = row["key"].removeprefix("NORM-")
        page_text = "\n".join(
            (TEXT_ROOT / row["sourceDocumentId"] / "pages" / f"page-{page:04d}.txt")
            .read_text(encoding="utf-8-sig")
            for page in range(int(row["pageFrom"]), int(row["pageTo"]) + 1)
        )
        audit.check(code in page_text, f"Norm code not found in cited pages: {row['key']}")


def audit_cost(rows: list[dict[str, str]], audit: Audit) -> None:
    by_key = {row["key"]: row for row in rows}
    audit.check(set(by_key) == EXPECTED_COST_KEYS, "Cost legal key set is incomplete or unexpected")
    fields = {key: parse_fields(row["data"], key, audit) for key, row in by_key.items()}

    for index in range(1, 9):
        values = fields[f"COST-K1K4-TERRAIN-{index:02d}"]
        k1 = decimal_value(values["k1Percent"], "K1", audit)
        k4 = decimal_value(values["k4Percent"], "K4", audit)
        total = decimal_value(values["totalPercent"], "K1K4 total", audit)
        audit.check(k1 + k4 == total, f"K1+K4 mismatch at terrain {index}")

    for key in ("COST-K2-LINEAR", "COST-K2-OTHER"):
        values = fields[key]
        thresholds = decimal_list(values["thresholdBillion"], key, audit)
        rates = decimal_list(values["ratesPercent"], key, audit)
        audit.check(thresholds == sorted(set(thresholds)), f"K2 thresholds not increasing: {key}")
        audit.check(len(rates) == len(thresholds) + 1, f"K2 rate count mismatch: {key}")
        audit.check(values.get("boundary") == "upper-inclusive" and values.get("interpolation") == "step",
                    f"K2 boundary/interpolation mismatch: {key}")
        audit.check(values.get("currentExternalBasis") == "TT36-2026-BXD-PL3-table-3.7",
                    f"K2 current basis missing: {key}")

    for key in (
        "COST-K5-CIVIL", "COST-K5-INDUSTRIAL", "COST-K5-TRANSPORT",
        "COST-K5-AGRICULTURE", "COST-K5-INFRASTRUCTURE",
    ):
        values = fields[key]
        thresholds = decimal_list(values["thresholdBillion"], key, audit)
        rates = decimal_list(values["ratesPercent"], key, audit)
        audit.check(thresholds == [Decimal(v) for v in (10, 20, 50, 100, 200, 500, 1000, 2000, 5000, 8000, 10000)],
                    f"K5 thresholds mismatch: {key}")
        audit.check(len(rates) == len(thresholds), f"K5 rate count mismatch: {key}")
        audit.check(all(rates[i] > rates[i + 1] for i in range(len(rates) - 1)),
                    f"K5 rates are not strictly decreasing: {key}")
        audit.check(values.get("interpolation") == "linear" and values.get("aboveMaximum") == "external-estimate",
                    f"K5 interpolation/maximum mismatch: {key}")
        audit.check(values.get("currentExternalBasis") == "TT38-2026-BXD-PL8-table-2.24",
                    f"K5 current basis missing: {key}")

    for key in ("COST-MIN-SURVEY-DESIGN-UNDER-4HA", "COST-MIN-DISPOSAL-UNDER-4HA"):
        values = fields[key]
        members = decimal_value(values["teamMembers"], key, audit)
        allowance = decimal_value(values["allowancePerPersonDayVnd"], key, audit)
        minimum = decimal_value(values["minimumVnd"], key, audit)
        audit.check(members * allowance == minimum, f"Minimum cost arithmetic mismatch: {key}")


def audit_machine(rows: list[dict[str, str]], audit: Audit) -> None:
    expected = {f"MACHINE-M010.{index:03d}" for index in range(1, 34)}
    audit.check({row["key"] for row in rows} == expected, "Machine legal key set is incomplete")
    for row in rows:
        fields = parse_fields(row["data"], row["key"], audit)
        suffix = row["key"].split(".")[-1]
        required = {
            "annualShifts", "depreciationPercent", "repairPercent", "otherPercent",
            "fuel", "operators", "referencePriceVnd", "recoverableVatPercent",
            "stateCode", "nonStateCode",
        }
        audit.check(required.issubset(fields), f"Machine required field missing: {row['key']}")
        audit.check(fields.get("stateCode") == f"M010.{suffix}", f"Machine state code mismatch: {row['key']}")
        audit.check(fields.get("nonStateCode") == f"M011.{suffix}", f"Machine non-state code mismatch: {row['key']}")
        shifts = decimal_value(fields.get("annualShifts", "0"), row["key"], audit)
        price = decimal_value(fields.get("referencePriceVnd", "0"), row["key"], audit)
        audit.check(Decimal(1) <= shifts <= Decimal(366), f"Machine annual shifts outlier: {row['key']}")
        audit.check(price > 0, f"Machine price must be positive: {row['key']}")
        for name in ("depreciationPercent", "repairPercent", "otherPercent", "recoverableVatPercent"):
            value = decimal_value(fields.get(name, "-1"), f"{row['key']}:{name}", audit)
            audit.check(Decimal(0) <= value <= Decimal(100), f"Machine percentage outlier: {row['key']}:{name}")

        state_code = fields.get("stateCode", "")
        page_text = "\n".join(
            (TEXT_ROOT / row["sourceDocumentId"] / "pages" / f"page-{page:04d}.txt")
            .read_text(encoding="utf-8-sig")
            for page in range(int(row["pageFrom"]), int(row["pageTo"]) + 1)
        )
        audit.check(pdf_code_visible(page_text, state_code),
                    f"Machine code not found in cited pages: {row['key']}")


def audit_geography(rows: list[dict[str, str]], audit: Audit) -> None:
    assignments = [row for row in rows if row["recordType"] == "LocalityZoneAssignment"]
    table_rows: list[int] = []
    provinces: list[str] = []
    for row in assignments:
        fields = parse_fields(row["data"], row["key"], audit)
        table_rows.append(int(fields.get("tableRow", "0")))
        provinces.append(fields.get("province", ""))
        zones = [int(value) for value in fields.get("zones", "").split(",") if value]
        fallback = int(fields.get("fallbackZone", "0"))
        audit.check(len(zones) == len(set(zones)) and all(1 <= zone <= 4 for zone in zones),
                    f"Invalid locality zones: {row['key']}")
        audit.check(fallback in zones, f"Fallback zone not listed: {row['key']}")
    audit.check(sorted(table_rows) == list(range(1, 35)), "Locality table rows are not 1..34")
    audit.check(len(set(provinces)) == 34 and all(provinces), "Locality provinces are blank/duplicated")


def audit_pdf_sources(sources: list[dict[str, str]], page_counts: dict[str, int], audit: Audit) -> None:
    expected_docs = {
        "TT101-2025-BQP", "VBHN94-2025-BQP", "VBHN95-2025-BQP",
        "VBHN96-2025-BQP", "VBHN97-2025-BQP", "VBHN98-2025-BQP",
    }
    audit.check({row["documentId"] for row in sources} == expected_docs,
                "Official source document set mismatch")
    for row in sources:
        document_id = row["documentId"]
        pdf = RAW_ROOT / f"{document_id}.pdf"
        audit.check(pdf.exists(), f"Official PDF missing: {document_id}")
        if pdf.exists():
            audit.check(sha256(pdf) == row["contentChecksum"], f"Official PDF checksum mismatch: {document_id}")
        pages = TEXT_ROOT / document_id / "pages"
        actual_pages = len(list(pages.glob("page-*.txt")))
        audit.check(actual_pages == page_counts.get(document_id), f"Text page count mismatch: {document_id}")


def main() -> int:
    audit = Audit()
    source_header = [
        "documentId", "title", "publisher", "issuedDate", "effectiveFrom",
        "effectiveTo", "officialUri", "contentChecksum",
    ]
    sources = read_tsv(SOURCE_ROOT / "sources.tsv", source_header, audit)
    source_ids = {row["documentId"] for row in sources}

    manifest_rows = read_tsv(TEXT_ROOT / "text-manifest.tsv", None, audit)
    page_counts = {Path(row["fileName"]).stem: int(row["pageCount"]) for row in manifest_rows}
    audit_pdf_sources(sources, page_counts, audit)

    all_records: list[tuple[str, dict[str, str]]] = []
    module_rows: dict[str, list[dict[str, str]]] = {}
    for module, expected_count in EXPECTED_COUNTS.items():
        rows = read_tsv(MODULE_ROOT / f"{module}.tsv", EXPECTED_HEADER, audit)
        module_rows[module] = rows
        audit.check(len(rows) == expected_count, f"Record count mismatch: {module}")
        audit.check(Counter(row["recordType"] for row in rows) == Counter(EXPECTED_TYPES[module]),
                    f"Record type distribution mismatch: {module}")
        for row in rows:
            all_records.append((module, row))
            for field in ("key", "recordType", "title", "data", "sourceDocumentId", "pageFrom", "pageTo", "section"):
                audit.check(bool(row.get(field)), f"Blank mandatory field {module}:{row.get('key')}:{field}")
            audit.check(row.get("verification") == "VerifiedAgainstOfficialSource",
                        f"Unverified record: {module}:{row.get('key')}")
            audit.check(row.get("sourceDocumentId") in source_ids,
                        f"Unknown source document: {module}:{row.get('key')}")
            try:
                page_from = int(row["pageFrom"])
                page_to = int(row["pageTo"])
                page_limit = page_counts[row["sourceDocumentId"]]
                audit.check(1 <= page_from <= page_to <= page_limit,
                            f"Invalid page locator: {module}:{row['key']}")
            except (KeyError, ValueError):
                audit.check(False, f"Malformed page locator: {module}:{row.get('key')}")

    identities = [(module, row["key"]) for module, row in all_records]
    global_keys = [row["key"] for _, row in all_records]
    audit.check(len(identities) == len(set(identities)), "Duplicate module/key identity")
    audit.check(len(global_keys) == len(set(global_keys)), "Duplicate global record key")
    audit.check(len(all_records) == 248, "Package total record count is not 248")

    audit_norm(module_rows["Norm"], audit)
    audit_cost(module_rows["CostRule"], audit)
    audit_machine(module_rows["MachineRate"], audit)
    audit_geography(module_rows["Geography"], audit)

    change_header = [
        "module", "key", "changeType", "changedFields", "reason",
        "sourceDocumentId", "pageFrom", "pageTo", "reviewStatus",
    ]
    changes = read_tsv(PACKAGE_ROOT / "audit" / "change-reasons.tsv", change_header, audit)
    change_ids = [(row["module"], row["key"]) for row in changes]
    audit.check(len(changes) == 248 and len(set(change_ids)) == 248,
                "Change-reason matrix must contain 248 unique rows")
    audit.check(set(change_ids) == set(identities), "Change-reason matrix does not cover every package record")
    for row in changes:
        audit.check(row["reviewStatus"] == "VerifiedAgainstOfficialSource",
                    f"Change reason is not verified: {row['module']}:{row['key']}")
        audit.check(bool(row["reason"] and row["changedFields"]),
                    f"Change reason is incomplete: {row['module']}:{row['key']}")

    bundle_checksum = load_manifest_checksum()
    audit.check(bundle_checksum == EXPECTED_PACKAGE_CHECKSUM, "Frozen BQP 2025 package checksum changed")

    report = {
        "audit": "DT-305-independent-source-review",
        "reviewDate": date.today().isoformat(),
        "status": "PASS" if not audit.errors else "FAIL",
        "checks": audit.checks,
        "recordCount": len(all_records),
        "moduleCounts": {name: len(rows) for name, rows in module_rows.items()},
        "sourcePdfCount": len(sources),
        "changeReasonCount": len(changes),
        "packageChecksum": bundle_checksum,
        "errors": audit.errors,
        "warnings": audit.warnings,
        "knownLegalDecisions": [
            "K6 rejects exactly 1000 kg because the source defines only below/above 1000 kg.",
            "Logical machine resources remain unresolved until the unit-rate workflow supplies the method/depth choice.",
            "Workbook ChiPhi legacy labels/formulas do not override current legal tables.",
        ],
    }
    REPORT_PATH.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=True, indent=2))
    return 0 if not audit.errors else 1


if __name__ == "__main__":
    sys.exit(main())
