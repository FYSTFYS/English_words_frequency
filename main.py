#!/usr/bin/env python3
"""Run the full EPUB -> TXT -> Excel word-frequency workflow."""

from __future__ import annotations

import argparse
import shlex
from pathlib import Path

from ankicheck import DEFAULT_ANKI_CONNECT_URL
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


def _default_combined_xlsx_path(epub_paths: list[Path]) -> Path:
    if not epub_paths:
        raise ValueError("At least one EPUB is required")
    return XLSX_DIR / f"{epub_paths[0].stem}_combined.xlsx"


def _print_text_result(epub_path: Path, txt_path: Path, char_count: int) -> None:
    print(f"EPUB: {epub_path}")
    print(f"TXT: {txt_path}")
    print(f"Characters: {char_count}")


def _print_excel_result(xlsx_path: Path, unique_count: int) -> None:
    print(f"Excel: {xlsx_path}")
    print(f"Unique lemmas: {unique_count}")


def run_workflow(
    epub_path: str | Path,
    txt_path: str | Path,
    xlsx_path: str | Path,
    start_percent: float,
    end_percent: float,
    workers: int | None = None,
    anki_check: bool = True,
    anki_url: str = DEFAULT_ANKI_CONNECT_URL,
) -> tuple[int, int]:
    _validate_percent_range(start_percent, end_percent)
    TXT_DIR.mkdir(parents=True, exist_ok=True)
    XLSX_DIR.mkdir(parents=True, exist_ok=True)
    char_count = convert_epub_to_text(epub_path, txt_path, start_percent, end_percent, workers)
    unique_count = convert_txt_to_excel(
        txt_path,
        xlsx_path,
        workers,
        anki_check=anki_check,
        anki_url=anki_url,
    )
    return char_count, unique_count


def run_workflows(
    epub_paths: list[Path],
    start_percent: float,
    end_percent: float,
    workers: int | None = None,
    txt_path: Path | None = None,
    xlsx_path: Path | None = None,
    anki_check: bool = True,
    anki_url: str = DEFAULT_ANKI_CONNECT_URL,
) -> tuple[list[tuple[Path, Path, int]], Path, int]:
    if not epub_paths:
        raise ValueError("At least one EPUB is required")
    if len(epub_paths) > 1 and txt_path is not None:
        raise ValueError("--txt can only be used with a single EPUB input")

    TXT_DIR.mkdir(parents=True, exist_ok=True)
    XLSX_DIR.mkdir(parents=True, exist_ok=True)
    text_results = []
    txt_paths = []
    for epub_path in epub_paths:
        resolved_txt = txt_path if txt_path is not None else _default_txt_path(epub_path)
        char_count = convert_epub_to_text(
            epub_path,
            resolved_txt,
            start_percent,
            end_percent,
            workers,
        )
        text_results.append((epub_path, resolved_txt, char_count))
        txt_paths.append(resolved_txt)

    resolved_xlsx = xlsx_path if xlsx_path is not None else (
        _default_xlsx_path(epub_paths[0]) if len(epub_paths) == 1 else _default_combined_xlsx_path(epub_paths)
    )
    unique_count = convert_txt_to_excel(
        txt_paths,
        resolved_xlsx,
        workers,
        anki_check=anki_check,
        anki_url=anki_url,
    )
    return text_results, resolved_xlsx, unique_count


def interactive_main() -> int:
    epub_input = input("EPUB paths or filenames in books/ (space-separated, quote names with spaces): ").strip()
    if not epub_input:
        raise ValueError("EPUB path is required")
    epub_paths = [_resolve_epub_path(item) for item in shlex.split(epub_input)]

    if len(epub_paths) == 1:
        txt_input = input(f"TXT output path [default { _default_txt_path(epub_paths[0]) }]: ").strip()
        xlsx_input = input(f"Excel output path [default { _default_xlsx_path(epub_paths[0]) }]: ").strip()
        txt_path = Path(txt_input).expanduser() if txt_input else None
        xlsx_path = Path(xlsx_input).expanduser() if xlsx_input else None
    else:
        xlsx_input = input(f"Excel output path [default { _default_combined_xlsx_path(epub_paths) }]: ").strip()
        txt_path = None
        xlsx_path = Path(xlsx_input).expanduser() if xlsx_input else None

    start = _parse_percent_input("Start percent [default 0]: ", 0)
    end = _parse_percent_input("End percent [default 100]: ", 100)
    workers_input = input("Workers [default CPU-based]: ").strip()
    workers = int(workers_input) if workers_input else None

    text_results, resolved_xlsx, unique_count = run_workflows(epub_paths, start, end, workers, txt_path, xlsx_path)
    for result in text_results:
        _print_text_result(*result)
    _print_excel_result(resolved_xlsx, unique_count)
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Convert EPUB to TXT and word-frequency XLSX.")
    parser.add_argument("epub", nargs="*", help="One or more EPUB paths, filenames, or stems inside books/.")
    parser.add_argument("--txt", help="Path to the output TXT file. Defaults to txtdir/<book>.txt.")
    parser.add_argument("--xlsx", help="Path to the output XLSX file. For multiple EPUBs, defaults to xlsxdir/<first_book>_combined.xlsx.")
    parser.add_argument("--start", type=float, default=0, help="Start percent. Defaults to 0.")
    parser.add_argument("--end", type=float, default=100, help="End percent. Defaults to 100.")
    parser.add_argument("--workers", type=int, default=None, help="Concurrent worker count. Defaults to CPU-based value.")
    parser.add_argument("--anki-check", dest="anki_check", action="store_true", default=True, help="Check local Anki through AnkiConnect. Enabled by default.")
    parser.add_argument("--no-anki-check", dest="anki_check", action="store_false", help="Disable local Anki check.")
    parser.add_argument("--anki-url", default=DEFAULT_ANKI_CONNECT_URL, help="AnkiConnect URL. Defaults to http://127.0.0.1:8765.")
    args = parser.parse_args(argv)

    if not args.epub:
        return interactive_main()

    epub_paths = [_resolve_epub_path(epub_input) for epub_input in args.epub]
    txt_path = Path(args.txt).expanduser() if args.txt else None
    xlsx_path = Path(args.xlsx).expanduser() if args.xlsx else None

    try:
        text_results, resolved_xlsx, unique_count = run_workflows(
            epub_paths,
            args.start,
            args.end,
            args.workers,
            txt_path,
            xlsx_path,
            args.anki_check,
            args.anki_url,
        )
    except ValueError as exc:
        parser.error(str(exc))

    for result in text_results:
        _print_text_result(*result)
    _print_excel_result(resolved_xlsx, unique_count)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
