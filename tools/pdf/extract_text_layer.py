#!/usr/bin/env python3
"""Extract page-addressable UTF-8 text from searchable regulation PDFs."""

from __future__ import annotations

import argparse
import hashlib
from pathlib import Path

from pypdf import PdfReader


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def normalize_text(value: str) -> str:
    return value.replace("\r\n", "\n").replace("\r", "\n")


def process_pdf(pdf_path: Path, output_root: Path) -> list[str]:
    reader = PdfReader(str(pdf_path))
    if reader.is_encrypted:
        raise ValueError(f"Encrypted PDF is not supported: {pdf_path}")

    document_dir = output_root / pdf_path.stem
    page_dir = document_dir / "pages"
    page_dir.mkdir(parents=True, exist_ok=True)

    page_texts: list[str] = []
    blank_pages: list[int] = []
    for page_number, page in enumerate(reader.pages, start=1):
        text = normalize_text(page.extract_text() or "")
        if not text.strip():
            blank_pages.append(page_number)
        page_texts.append(text)
        (page_dir / f"page-{page_number:04d}.txt").write_text(
            text, encoding="utf-8", newline="\n"
        )

    expected_pages = {
        f"page-{page_number:04d}.txt"
        for page_number in range(1, len(page_texts) + 1)
    }
    actual_pages = {path.name for path in page_dir.glob("page-*.txt")}
    if actual_pages != expected_pages:
        extras = sorted(actual_pages - expected_pages)
        raise ValueError(
            f"Stale extracted pages in {page_dir}: {', '.join(extras)}"
        )

    combined_path = document_dir / "document.txt"
    with combined_path.open("w", encoding="utf-8", newline="\n") as stream:
        for page_number, text in enumerate(page_texts, start=1):
            stream.write(f"[[PAGE {page_number}]]\n")
            stream.write(text)
            if not text.endswith("\n"):
                stream.write("\n")

    return [
        pdf_path.name,
        str(len(page_texts)),
        sha256(pdf_path),
        sha256(combined_path),
        str(sum(len(text) for text in page_texts)),
        ",".join(str(page) for page in blank_pages),
        "pypdf-text-layer",
    ]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--match", default="*.pdf")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if not args.input.is_dir():
        raise NotADirectoryError(args.input)
    pdf_paths = sorted(args.input.glob(args.match))
    if not pdf_paths:
        raise FileNotFoundError(f"No PDF files in {args.input}")

    args.output.mkdir(parents=True, exist_ok=True)
    rows = [process_pdf(pdf_path, args.output) for pdf_path in pdf_paths]
    manifest = args.output / "text-manifest.tsv"
    with manifest.open("w", encoding="utf-8", newline="\n") as stream:
        stream.write(
            "fileName\tpageCount\tpdfSha256\ttextSha256\t"
            "totalCharacters\tblankPages\textractor\n"
        )
        for row in rows:
            stream.write("\t".join(row) + "\n")

    for row in rows:
        blank_label = row[5] or "none"
        print(
            f"{row[0]}: {row[1]} pages, {row[4]} characters, "
            f"blank pages: {blank_label}"
        )
    print(f"Manifest: {manifest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
