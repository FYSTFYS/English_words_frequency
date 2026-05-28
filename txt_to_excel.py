#!/usr/bin/env python3
"""Extract lemma frequencies from TXT and write an Excel workbook."""

from __future__ import annotations

import argparse
import html
import os
import re
import warnings
import zipfile
from collections import Counter, defaultdict
from functools import lru_cache
from pathlib import Path
from typing import Iterable


DEFAULT_CHUNK_SIZE = 120_000
TOKEN_RE = re.compile(r"[A-Za-z]+(?:-[A-Za-z]+)*")
SIMPLE_WORD_RE = re.compile(r"^[a-z]+$")
BASE_DIR = Path(__file__).resolve().parent
TXT_DIR = BASE_DIR / "txtdir"
XLSX_DIR = BASE_DIR / "xlsxdir"
VOCABULARY_EXCLUDE_DIR = BASE_DIR / "vocabulary_exclude"
WORDLIST_DIR = BASE_DIR / "wordlist"


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
                    excluded.add(match.group(0).lower())
    return excluded


def load_wordlist_tags(wordlist_dir: str | Path = WORDLIST_DIR) -> dict[str, str]:
    base = Path(wordlist_dir)
    readme = base / "readme.txt"
    if not readme.exists():
        return {}

    tags: dict[str, str] = {}
    for line in readme.read_text(encoding="utf-8", errors="replace").splitlines():
        filename = line.strip()
        if not filename:
            continue
        wordlist_path = base / filename
        if not wordlist_path.is_file():
            continue
        tag = wordlist_path.stem
        with wordlist_path.open("r", encoding="utf-8", errors="replace") as source:
            for word_line in source:
                word = word_line.strip().lower()
                if SIMPLE_WORD_RE.fullmatch(word) and word not in tags:
                    tags[word] = tag
    return tags


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


def _load_spacy_model(model_name: str):
    warnings.filterwarnings("ignore", message="urllib3 v2 only supports OpenSSL.*")

    try:
        import spacy
    except ImportError as exc:
        raise RuntimeError(
            "spaCy is required. Install it with: pyenvdir; python -m pip install spacy"
        ) from exc

    try:
        return spacy.load(model_name, disable=["ner"])
    except OSError as exc:
        raise RuntimeError(
            f"spaCy model '{model_name}' is required. Install it with: "
            f"pyenvdir; python -m spacy download {model_name}"
        ) from exc


@lru_cache(maxsize=50_000)
def _fallback_lemma(form: str, pos: str) -> str | None:
    pos_map = {
        "ADJ": "ADJ",
        "ADV": "ADV",
        "AUX": "VERB",
        "NOUN": "NOUN",
        "VERB": "VERB",
    }
    lemma_pos = pos_map.get(pos)
    if lemma_pos is None:
        return None

    try:
        from lemminflect import getAllLemmas
    except ImportError:
        return None

    lemmas = getAllLemmas(form).get(lemma_pos, ())
    for lemma in lemmas:
        lemma = lemma.lower()
        if lemma != form and SIMPLE_WORD_RE.fullmatch(lemma):
            return lemma
    return None


def _valid_lemma_token(token) -> tuple[str, str] | None:
    if not token.is_alpha:
        return None
    form = token.text.lower()
    if len(form) <= 1:
        return None

    lemma = token.lemma_.lower().strip()
    if lemma == form:
        lemma = _fallback_lemma(form, token.pos_) or lemma
    if lemma == form and form.endswith("s"):
        lemma = _fallback_lemma(form, "NOUN") or lemma
    if not lemma or lemma == "-pron-" or len(lemma) <= 1:
        return None
    if not SIMPLE_WORD_RE.fullmatch(lemma):
        return None
    return lemma, form


def collect_lemma_frequencies(
    txt_path: str | Path,
    workers: int | None = None,
    model_name: str = "en_core_web_sm",
    chunk_size: int = DEFAULT_CHUNK_SIZE,
) -> dict[str, Counter[str]]:
    nlp = _load_spacy_model(model_name)
    workers = default_worker_count() if workers is None else max(1, workers)
    n_process = workers if workers > 1 else 1
    lemma_forms: dict[str, Counter[str]] = defaultdict(Counter)

    for doc in nlp.pipe(_iter_text_chunks(txt_path, chunk_size), batch_size=8, n_process=n_process):
        for token in doc:
            parsed = _valid_lemma_token(token)
            if parsed is None:
                continue
            lemma, form = parsed
            lemma_forms[lemma][form] += 1
    return lemma_forms


def _forms_text(forms: Counter[str]) -> str:
    return ", ".join(
        f"{form}: {count}"
        for form, count in sorted(forms.items(), key=lambda item: (-item[1], item[0]))
    )


def _filtered_rows(
    lemma_forms: dict[str, Counter[str]],
    excluded_words: set[str],
    wordlist_tags: dict[str, str],
) -> list[list[str | int]]:
    rows: list[list[str | int]] = [["lemma", "count", "forms", "wordlist_tag"]]
    sortable = []
    for lemma, forms in lemma_forms.items():
        if lemma in excluded_words:
            continue
        count = sum(forms.values())
        sortable.append((lemma, count, forms))

    for lemma, count, forms in sorted(sortable, key=lambda item: (-item[1], item[0])):
        rows.append([lemma, count, _forms_text(forms), wordlist_tags.get(lemma, "")])
    return rows


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
        '<col min="3" max="3" width="64" customWidth="1"/>'
        '<col min="4" max="4" width="20" customWidth="1"/></cols>'
        f'<sheetData>{"".join(xml_rows)}</sheetData>'
        '<autoFilter ref="A1:D1"/>'
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
            '<sheets><sheet name="lemma_frequency" sheetId="1" r:id="rId1"/></sheets>'
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
    vocabulary_dir: str | Path = VOCABULARY_EXCLUDE_DIR,
    wordlist_dir: str | Path = WORDLIST_DIR,
    model_name: str = "en_core_web_sm",
) -> int:
    lemma_forms = collect_lemma_frequencies(txt_path, workers, model_name)
    rows = _filtered_rows(
        lemma_forms,
        load_excluded_words(vocabulary_dir),
        load_wordlist_tags(wordlist_dir),
    )
    _write_xlsx(rows, xlsx_path)
    return len(rows) - 1


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Count TXT lemma frequencies and write an XLSX file.")
    parser.add_argument("txt", help="TXT path, or filename/stem inside txtdir/.")
    parser.add_argument("xlsx", nargs="?", help="Output XLSX path. Defaults to xlsxdir/<txt>.xlsx.")
    parser.add_argument("--workers", type=int, default=None, help="spaCy n_process worker count. Defaults to CPU count.")
    parser.add_argument(
        "--vocabulary-dir",
        default=str(VOCABULARY_EXCLUDE_DIR),
        help="Directory of vocabulary files to exclude. Defaults to ./vocabulary_exclude.",
    )
    parser.add_argument(
        "--wordlist-dir",
        default=str(WORDLIST_DIR),
        help="Directory containing readme.txt and prioritized wordlist files. Defaults to ./wordlist.",
    )
    parser.add_argument("--model", default="en_core_web_sm", help="spaCy model name. Defaults to en_core_web_sm.")
    args = parser.parse_args(argv)

    txt_path = resolve_txt_path(args.txt)
    xlsx_path = Path(args.xlsx).expanduser() if args.xlsx else default_xlsx_path(txt_path)

    unique_count = convert_txt_to_excel(
        txt_path,
        xlsx_path,
        args.workers,
        args.vocabulary_dir,
        args.wordlist_dir,
        args.model,
    )
    print(f"Wrote {unique_count} lemmas to {xlsx_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
