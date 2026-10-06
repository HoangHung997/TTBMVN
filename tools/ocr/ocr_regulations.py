#!/usr/bin/env python3
"""OCR scanned regulation PDFs into page-addressable UTF-8 evidence files."""

from __future__ import annotations

import argparse
import concurrent.futures
import hashlib
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile

from PIL import Image
from pypdf import PdfReader


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def run_text(command: list[str]) -> str:
    result = subprocess.run(
        command,
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    return result.stdout.decode("utf-8", errors="replace")


def ocr_page(
    pdf_path: Path,
    page_number: int,
    page_dir: Path,
    temp_root: Path,
    pdftoppm: Path,
    tesseract: Path,
    tessdata: Path,
    dpi: int,
    psm: int,
    auto_orient: bool,
    force: bool,
) -> tuple[int, int]:
    output_path = page_dir / f"page-{page_number:04d}.txt"
    if not force and output_path.exists() and output_path.stat().st_size > 0:
        return page_number, output_path.stat().st_size

    page_temp = Path(tempfile.mkdtemp(prefix=f"p{page_number:04d}-", dir=temp_root))
    try:
        image_prefix = page_temp / "page"
        subprocess.run(
            [
                str(pdftoppm),
                "-f",
                str(page_number),
                "-l",
                str(page_number),
                "-r",
                str(dpi),
                "-png",
                "-singlefile",
                str(pdf_path),
                str(image_prefix),
            ],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
        )
        image_path = image_prefix.with_suffix(".png")
        if auto_orient:
            osd_result = subprocess.run(
                [
                    str(tesseract),
                    str(image_path),
                    "stdout",
                    "--tessdata-dir",
                    str(tessdata),
                    "-l",
                    "osd",
                    "--psm",
                    "0",
                ],
                check=False,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            osd = (osd_result.stdout + osd_result.stderr).decode(
                "utf-8", errors="replace"
            )
            rotate_match = re.search(r"^Rotate:\s*(\d+)\s*$", osd, re.MULTILINE)
            confidence_match = re.search(
                r"^Orientation confidence:\s*([0-9.]+)\s*$", osd, re.MULTILINE
            )
            rotate = int(rotate_match.group(1)) if rotate_match else 0
            confidence = float(confidence_match.group(1)) if confidence_match else 0.0
            if rotate in (90, 180, 270) and confidence >= 5.0:
                with Image.open(image_path) as image:
                    image.rotate(-rotate, expand=True).save(image_path)

        text = run_text(
            [
                str(tesseract),
                str(image_path),
                "stdout",
                "--tessdata-dir",
                str(tessdata),
                "-l",
                "vie+eng",
                "--psm",
                str(psm),
                "-c",
                "preserve_interword_spaces=1",
            ]
        )
        output_path.write_text(text.replace("\r\n", "\n"), encoding="utf-8")
        return page_number, output_path.stat().st_size
    finally:
        shutil.rmtree(page_temp, ignore_errors=True)


def process_pdf(
    pdf_path: Path,
    output_root: Path,
    temp_root: Path,
    pdftoppm: Path,
    tesseract: Path,
    tessdata: Path,
    dpi: int,
    psm: int,
    workers: int,
    auto_orient: bool,
    force: bool,
) -> tuple[str, int, str, str]:
    page_count = len(PdfReader(str(pdf_path)).pages)
    document_dir = output_root / pdf_path.stem
    page_dir = document_dir / "pages"
    page_dir.mkdir(parents=True, exist_ok=True)

    print(f"OCR {pdf_path.name}: {page_count} pages, {workers} workers", flush=True)
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
        futures = [
            executor.submit(
                ocr_page,
                pdf_path,
                page_number,
                page_dir,
                temp_root,
                pdftoppm,
                tesseract,
                tessdata,
                dpi,
                psm,
                auto_orient,
                force,
            )
            for page_number in range(1, page_count + 1)
        ]
        completed = 0
        for future in concurrent.futures.as_completed(futures):
            future.result()
            completed += 1
            if completed == page_count or completed % 10 == 0:
                print(f"  {completed}/{page_count}", flush=True)

    combined_path = document_dir / "document.txt"
    with combined_path.open("w", encoding="utf-8", newline="\n") as stream:
        for page_number in range(1, page_count + 1):
            page_path = page_dir / f"page-{page_number:04d}.txt"
            page_text = page_path.read_text(encoding="utf-8")
            stream.write(f"[[PAGE {page_number}]]\n")
            stream.write(page_text)
            if not page_text.endswith("\n"):
                stream.write("\n")

    return pdf_path.name, page_count, sha256(pdf_path), sha256(combined_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--pdftoppm", required=True, type=Path)
    parser.add_argument("--tesseract", required=True, type=Path)
    parser.add_argument("--tessdata", required=True, type=Path)
    parser.add_argument("--dpi", type=int, default=300)
    parser.add_argument("--psm", type=int, default=3)
    parser.add_argument("--workers", type=int, default=0)
    parser.add_argument("--match", default="*.pdf")
    parser.add_argument("--auto-orient", action="store_true")
    parser.add_argument("--force", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    workers = args.workers or max(1, min(6, os.cpu_count() or 1))
    if workers < 1:
        raise ValueError("workers must be at least 1")
    if args.dpi < 150 or args.dpi > 600:
        raise ValueError("dpi must be between 150 and 600")
    if args.psm < 1 or args.psm > 13:
        raise ValueError("psm must be between 1 and 13")

    pdf_paths = sorted(args.input.glob(args.match))
    if not pdf_paths:
        raise FileNotFoundError(f"No PDF files in {args.input}")
    for required in (args.pdftoppm, args.tesseract):
        if not required.is_file():
            raise FileNotFoundError(required)
    for language in ("vie.traineddata", "eng.traineddata"):
        if not (args.tessdata / language).is_file():
            raise FileNotFoundError(args.tessdata / language)

    args.output.mkdir(parents=True, exist_ok=True)
    temp_root = Path(tempfile.mkdtemp(prefix=".ocr-run-", dir=args.output))
    rows = []
    try:
        for pdf_path in pdf_paths:
            rows.append(
                process_pdf(
                    pdf_path,
                    args.output,
                    temp_root,
                    args.pdftoppm,
                    args.tesseract,
                    args.tessdata,
                    args.dpi,
                    args.psm,
                    workers,
                    args.auto_orient,
                    args.force,
                )
            )
    finally:
        shutil.rmtree(temp_root, ignore_errors=True)

    tesseract_version = run_text([str(args.tesseract), "--version"]).splitlines()[0]
    manifest_path = args.output / "ocr-manifest.tsv"
    manifest_rows: dict[str, list[str]] = {}
    if manifest_path.exists():
        for line in manifest_path.read_text(encoding="utf-8").splitlines()[1:]:
            fields = line.split("\t")
            if len(fields) >= 6:
                if len(fields) == 6:
                    fields.extend(["6", "false"])
                elif len(fields) == 7:
                    fields.insert(6, "6")
                manifest_rows[fields[0]] = fields
    for name, pages, pdf_hash, text_hash in rows:
        manifest_rows[name] = [
            name,
            str(pages),
            pdf_hash,
            text_hash,
            str(args.dpi),
            tesseract_version,
            str(args.psm),
            str(args.auto_orient).lower(),
        ]
    lines = [
        "document\tpages\tpdf_sha256\ttext_sha256\tdpi\tengine\tpsm\tauto_orient",
        *["\t".join(manifest_rows[name]) for name in sorted(manifest_rows)],
    ]
    manifest_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Manifest: {manifest_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
