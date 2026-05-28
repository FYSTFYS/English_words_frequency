#!/usr/bin/env python3
"""Count word frequencies from TXT and write an Excel workbook."""

from __future__ import annotations

import argparse
import html
import os
import re
import zipfile
from collections import Counter
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path
from typing import Iterable

from epub_to_txt import is_hyphenated_token, tokenize_text


DEFAULT_CHUNK_SIZE = 4 * 1024 * 1024
TOKEN_RE = re.compile(r"[A-Za-z]+(?:-[A-Za-z]+)*")
BASE_DIR = Path(__file__).resolve().parent
TXT_DIR = BASE_DIR / "txtdir"
XLSX_DIR = BASE_DIR / "xlsxdir"
VOCABULARY_EXCLUDE_DIR = BASE_DIR / "vocabulary_exclude"


def default_worker_count() -> int:
    return os.cpu_count() or 1


def resolve_txt_path(txt_input: str | Path) -> Path:
    txt_path = Path(txt_input).expanduser()
    if txt_path.exists():
        return txt_path

    if not txt_path.is_absolute():
        candidate = TXT_DIR / txt_path
        if candidate.exists():
            return candidate
        if txt_path.suffix == "":
            candidate = TXT_DIR / f"{txt_path}.txt"
            if candidate.exists():
                return candidate

    raise FileNotFoundError(f"TXT not found: {txt_input}")


def default_xlsx_path(txt_path: str | Path) -> Path:
    return XLSX_DIR / f"{Path(txt_path).stem}.xlsx"


def load_excluded_words(vocabulary_dir: str | Path = VOCABULARY_EXCLUDE_DIR) -> set[str]:
    vocabulary_path = Path(vocabulary_dir)
    if not vocabulary_path.exists():
        return set()

    excluded: set[str] = set()
    for path in vocabulary_path.rglob("*"):
        if not path.is_file():
            continue
        with path.open("r", encoding="utf-8", errors="replace") as source:
            for line in source:
                for match in TOKEN_RE.finditer(line):
                    word = match.group(0)
                    excluded.add(word)
                    excluded.add(word.lower())
    return excluded


def _normalize_existing_token(token: str) -> str:
    return token if is_hyphenated_token(token) else token.lower()


def iter_txt_tokens(txt_path: str | Path, raw_text: bool = False) -> Iterable[str]:
    with Path(txt_path).open("r", encoding="utf-8", errors="replace") as source:
        for line in source:
            if raw_text:
                yield from tokenize_text(line)
            else:
                for match in TOKEN_RE.finditer(line):
                    yield _normalize_existing_token(match.group(0))


def _count_text_chunk(args: tuple[str, bool]) -> Counter[str]:
    text, raw_text = args
    if raw_text:
        return Counter(tokenize_text(text))
    return Counter(_normalize_existing_token(match.group(0)) for match in TOKEN_RE.finditer(text))


def _iter_text_chunks(txt_path: str | Path, chunk_size: int = DEFAULT_CHUNK_SIZE) -> Iterable[str]:
    buffer: list[str] = []
    buffered_size = 0
    with Path(txt_path).open("r", encoding="utf-8", errors="replace") as source:
        for line in source:
            buffer.append(line)
            buffered_size += len(line)
            if buffered_size >= chunk_size:
                yield "".join(buffer)
                buffer.clear()
                buffered_size = 0
    if buffer:
        yield "".join(buffer)


def count_txt_words(
    txt_path: str | Path,
    workers: int | None = None,
    chunk_size: int = DEFAULT_CHUNK_SIZE,
    raw_text: bool = False,
) -> Counter[str]:
    workers = default_worker_count() if workers is None else max(1, workers)
    chunks = _iter_text_chunks(txt_path, chunk_size)

    if workers <= 1:
        total: Counter[str] = Counter()
        for chunk in chunks:
            total.update(_count_text_chunk((chunk, raw_text)))
        return total

    total: Counter[str] = Counter()
    with ProcessPoolExecutor(max_workers=workers) as executor:
        for partial in executor.map(_count_text_chunk, ((chunk, raw_text) for chunk in chunks)):
            total.update(partial)
    return total


def _should_keep_word(word: str, excluded_words: set[str]) -> bool:
    if len(word) == 1:
        return False
    return word not in excluded_words and word.lower() not in excluded_words


def filter_counts(counts: Counter[str], excluded_words: set[str]) -> Counter[str]:
    return Counter({word: count for word, count in counts.items() if _should_keep_word(word, excluded_words)})


def _excel_col_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _sheet_xml(rows: list[list[str | int]]) -> str:
    xml_rows = []
    for row_index, row in enumerate(rows, start=1):
        cells = []
        for col_index, value in enumerate(row, start=1):
            cell_ref = f"{_excel_col_name(col_index)}{row_index}"
            if isinstance(value, int):
                cells.append(f'<c r="{cell_ref}"><v>{value}</v></c>')
            else:
                escaped = html.escape(value, quote=False)
                cells.append(f'<c r="{cell_ref}" t="inlineStr"><is><t>{escaped}</t></is></c>')
        xml_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols><col min="1" max="1" width="28" customWidth="1"/>'
        '<col min="2" max="2" width="12" customWidth="1"/>'
        '<col min="3" max="3" width="14" customWidth="1"/></cols>'
        f'<sheetData>{"".join(xml_rows)}</sheetData>'
        '<autoFilter ref="A1:C1"/>'
        '</worksheet>'
    )


def _write_xlsx(rows: list[list[str | int]], xlsx_path: str | Path) -> None:
    xlsx_path = Path(xlsx_path)
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    files = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/docProps/core.xml" '
            'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            '</Types>'
        ),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" '
            'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
            'Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
            'Target="docProps/app.xml"/>'
            '</Relationships>'
        ),
        "xl/workbook.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="word_frequency" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>'
        ),
        "xl/_rels/workbook.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/>'
            '</Relationships>'
        ),
        "xl/worksheets/sheet1.xml": _sheet_xml(rows),
        "docProps/core.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/">'
            '<dc:creator>wordcount</dc:creator>'
            '</cp:coreProperties>'
        ),
        "docProps/app.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>wordcount</Application>'
            '</Properties>'
        ),
    }

    with zipfile.ZipFile(xlsx_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, content in files.items():
            archive.writestr(name, content)


def convert_txt_to_excel(
    txt_path: str | Path,
    xlsx_path: str | Path,
    workers: int | None = None,
    raw_text: bool = False,
    vocabulary_dir: str | Path = VOCABULARY_EXCLUDE_DIR,
) -> int:
    counts = count_txt_words(txt_path, workers, raw_text=raw_text)
    counts = filter_counts(counts, load_excluded_words(vocabulary_dir))
    sorted_counts = sorted(counts.items(), key=lambda item: (-item[1], item[0].casefold(), item[0]))

    rows: list[list[str | int]] = [["word", "count", "hyphenated"]]
    rows.extend([word, count, "yes" if is_hyphenated_token(word) else ""] for word, count in sorted_counts)
    _write_xlsx(rows, xlsx_path)
    return len(sorted_counts)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Count TXT word frequencies and write an XLSX file.")
    parser.add_argument("txt", help="TXT path, or filename/stem inside txtdir/.")
    parser.add_argument("xlsx", nargs="?", help="Output XLSX path. Defaults to xlsxdir/<txt>.xlsx.")
    parser.add_argument("--workers", type=int, default=None, help="Concurrent worker count. Defaults to CPU count.")
    parser.add_argument(
        "--vocabulary-dir",
        default=str(VOCABULARY_EXCLUDE_DIR),
        help="Directory of vocabulary files to exclude. Defaults to ./vocabulary_exclude.",
    )
    parser.add_argument(
        "--raw-text",
        action="store_true",
        help="Treat TXT as original prose and expand hyphenated words. Default expects extracted token TXT.",
    )
    args = parser.parse_args(argv)

    txt_path = resolve_txt_path(args.txt)
    xlsx_path = Path(args.xlsx).expanduser() if args.xlsx else default_xlsx_path(txt_path)

    unique_count = convert_txt_to_excel(txt_path, xlsx_path, args.workers, args.raw_text, args.vocabulary_dir)
    print(f"Wrote {unique_count} unique words to {xlsx_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
