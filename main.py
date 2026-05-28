#!/usr/bin/env python3
"""Run the full EPUB -> TXT -> Excel word-frequency workflow."""

from __future__ import annotations

import argparse
from pathlib import Path

from epub_to_txt import convert_epub_to_text
from txt_to_excel import convert_txt_to_excel


BASE_DIR = Path(__file__).resolve().parent
BOOKS_DIR = BASE_DIR / "books"
TXT_DIR = BASE_DIR / "txtdir"
XLSX_DIR = BASE_DIR / "xlsxdir"


def _parse_percent_input(prompt: str, default: float) -> float:
    value = input(prompt).strip()
    return default if value == "" else float(value)


def _validate_percent_range(start: float, end: float) -> None:
    if not (0 <= start < end <= 100):
        raise ValueError("Percent range must satisfy 0 <= start < end <= 100")


def _resolve_epub_path(epub_input: str | Path) -> Path:
    epub_path = Path(epub_input).expanduser()
    if epub_path.exists():
        return epub_path

    if not epub_path.is_absolute():
        candidate = BOOKS_DIR / epub_path
        if candidate.exists():
            return candidate
        if epub_path.suffix == "":
            candidate = BOOKS_DIR / f"{epub_path}.epub"
            if candidate.exists():
                return candidate

    raise FileNotFoundError(f"EPUB not found: {epub_input}")


def _default_txt_path(epub_path: Path) -> Path:
    return TXT_DIR / f"{epub_path.stem}.txt"


def _default_xlsx_path(epub_path: Path) -> Path:
    return XLSX_DIR / f"{epub_path.stem}.xlsx"


def run_workflow(
    epub_path: str | Path,
    txt_path: str | Path,
    xlsx_path: str | Path,
    start_percent: float,
    end_percent: float,
    workers: int | None = None,
) -> tuple[int, int]:
    _validate_percent_range(start_percent, end_percent)
    TXT_DIR.mkdir(parents=True, exist_ok=True)
    XLSX_DIR.mkdir(parents=True, exist_ok=True)
    char_count = convert_epub_to_text(epub_path, txt_path, start_percent, end_percent, workers)
    unique_count = convert_txt_to_excel(txt_path, xlsx_path, workers)
    return char_count, unique_count


def interactive_main() -> int:
    epub_input = input("EPUB path or filename in books/: ").strip()
    if not epub_input:
        raise ValueError("EPUB path is required")
    epub_path = _resolve_epub_path(epub_input)

    txt_input = input(f"TXT output path [default { _default_txt_path(epub_path) }]: ").strip()
    xlsx_input = input(f"Excel output path [default { _default_xlsx_path(epub_path) }]: ").strip()

    txt_path = Path(txt_input).expanduser() if txt_input else _default_txt_path(epub_path)
    xlsx_path = Path(xlsx_input).expanduser() if xlsx_input else _default_xlsx_path(epub_path)
    start = _parse_percent_input("Start percent [default 0]: ", 0)
    end = _parse_percent_input("End percent [default 100]: ", 100)
    workers_input = input("Workers [default CPU-based]: ").strip()
    workers = int(workers_input) if workers_input else None

    char_count, unique_count = run_workflow(epub_path, txt_path, xlsx_path, start, end, workers)
    print(f"TXT: {txt_path}")
    print(f"Excel: {xlsx_path}")
    print(f"Characters: {char_count}")
    print(f"Unique lemmas: {unique_count}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Convert EPUB to TXT and word-frequency XLSX.")
    parser.add_argument("epub", nargs="?", help="EPUB path, or filename/stem inside books/.")
    parser.add_argument("--txt", help="Path to the output TXT file. Defaults to txtdir/<book>.txt.")
    parser.add_argument("--xlsx", help="Path to the output XLSX file. Defaults to xlsxdir/<book>.xlsx.")
    parser.add_argument("--start", type=float, default=0, help="Start percent. Defaults to 0.")
    parser.add_argument("--end", type=float, default=100, help="End percent. Defaults to 100.")
    parser.add_argument("--workers", type=int, default=None, help="Concurrent worker count. Defaults to CPU-based value.")
    args = parser.parse_args(argv)

    if not args.epub:
        return interactive_main()

    epub_path = _resolve_epub_path(args.epub)
    txt_path = Path(args.txt).expanduser() if args.txt else _default_txt_path(epub_path)
    xlsx_path = Path(args.xlsx).expanduser() if args.xlsx else _default_xlsx_path(epub_path)

    char_count, unique_count = run_workflow(epub_path, txt_path, xlsx_path, args.start, args.end, args.workers)
    print(f"TXT: {txt_path}")
    print(f"Excel: {xlsx_path}")
    print(f"Characters: {char_count}")
    print(f"Unique lemmas: {unique_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
