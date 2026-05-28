#!/usr/bin/env python3
"""Extract word tokens from an EPUB into a TXT file."""

from __future__ import annotations

import argparse
import html
import os
import posixpath
import re
import zipfile
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable


WORD_RE = re.compile(r"[A-Za-z]+(?:-[A-Za-z]+)*")
TAG_RE = re.compile(r"<\s*([A-Za-z_][\w:.-]*)([^>]*)>", re.DOTALL)
ATTR_RE = re.compile(r"([A-Za-z_][\w:.-]*)\s*=\s*(['\"])(.*?)\2", re.DOTALL)
BASE_DIR = Path(__file__).resolve().parent
BOOKS_DIR = BASE_DIR / "books"
TXT_DIR = BASE_DIR / "txtdir"


def tokenize_text(text: str) -> list[str]:
    """Tokenize English text.

    Plain words are lowercased. Hyphenated words keep their original casing as
    full tokens and also emit lowercased component words.
    """
    tokens: list[str] = []
    for match in WORD_RE.finditer(text):
        word = match.group(0)
        if "-" in word:
            tokens.extend(part.lower() for part in word.split("-") if part)
            tokens.append(word)
        else:
            tokens.append(word.lower())
    return tokens


def is_hyphenated_token(token: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z]+(?:-[A-Za-z]+)+", token))


class _HTMLTextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._parts: list[str] = []
        self._skip_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() in {"script", "style", "noscript"}:
            self._skip_depth += 1
        elif self._skip_depth == 0 and tag.lower() in {
            "br",
            "p",
            "div",
            "section",
            "article",
            "header",
            "footer",
            "li",
            "tr",
            "td",
            "th",
            "h1",
            "h2",
            "h3",
            "h4",
            "h5",
            "h6",
        }:
            self._parts.append(" ")

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() in {"script", "style", "noscript"} and self._skip_depth:
            self._skip_depth -= 1
        elif self._skip_depth == 0:
            self._parts.append(" ")

    def handle_data(self, data: str) -> None:
        if self._skip_depth == 0:
            self._parts.append(data)

    def text(self) -> str:
        return re.sub(r"\s+", " ", html.unescape(" ".join(self._parts))).strip()


@dataclass(frozen=True)
class SpineItem:
    href: str
    zip_path: str


@dataclass(frozen=True)
class RangeChunk:
    zip_path: str
    start: int
    end: int


def default_worker_count() -> int:
    return min(32, (os.cpu_count() or 1) + 4)


def _local_name(name: str) -> str:
    return name.rsplit(":", 1)[-1]


def _read_xml_text(epub: zipfile.ZipFile, path: str) -> str:
    try:
        data = epub.read(path)
    except KeyError as exc:
        raise ValueError(f"EPUB is missing required file: {path}") from exc
    return data.decode("utf-8", errors="replace")


def _iter_xml_tags(xml_text: str, local_name: str) -> Iterable[dict[str, str]]:
    for match in TAG_RE.finditer(xml_text):
        name, attrs_text = match.groups()
        if _local_name(name) != local_name:
            continue
        attrs = {
            _local_name(attr_name): html.unescape(value)
            for attr_name, _quote, value in ATTR_RE.findall(attrs_text)
        }
        yield attrs


def _container_opf_path(epub: zipfile.ZipFile) -> str:
    xml_text = _read_xml_text(epub, "META-INF/container.xml")
    for attrs in _iter_xml_tags(xml_text, "rootfile"):
        full_path = attrs.get("full-path")
        if full_path:
            return full_path
    raise ValueError("EPUB container.xml does not contain a rootfile entry")


def _spine_items(epub: zipfile.ZipFile) -> list[SpineItem]:
    opf_path = _container_opf_path(epub)
    opf_text = _read_xml_text(epub, opf_path)
    opf_dir = posixpath.dirname(opf_path)

    id_to_href: dict[str, str] = {}
    for attrs in _iter_xml_tags(opf_text, "item"):
        item_id = attrs.get("id")
        href = attrs.get("href")
        media_type = attrs.get("media-type", "")
        if item_id and href and media_type in {
            "application/xhtml+xml",
            "text/html",
            "application/xml",
        }:
            id_to_href[item_id] = href

    items: list[SpineItem] = []
    for attrs in _iter_xml_tags(opf_text, "itemref"):
        idref = attrs.get("idref")
        href = id_to_href.get(idref or "")
        if not href:
            continue
        zip_path = posixpath.normpath(posixpath.join(opf_dir, href))
        items.append(SpineItem(href=href, zip_path=zip_path))

    if not items:
        raise ValueError("EPUB spine does not reference any HTML/XHTML documents")
    return items


def _extract_html_text(raw: bytes) -> str:
    parser = _HTMLTextExtractor()
    parser.feed(raw.decode("utf-8", errors="replace"))
    parser.close()
    return parser.text()


def _document_text(epub: zipfile.ZipFile, item: SpineItem) -> str:
    try:
        return _extract_html_text(epub.read(item.zip_path))
    except KeyError as exc:
        raise ValueError(f"EPUB spine references missing document: {item.zip_path}") from exc


def _document_text_from_path(args: tuple[str, str]) -> str:
    epub_path, zip_path = args
    with zipfile.ZipFile(epub_path) as epub:
        try:
            return _extract_html_text(epub.read(zip_path))
        except KeyError as exc:
            raise ValueError(f"EPUB spine references missing document: {zip_path}") from exc


def _document_length_from_path(args: tuple[str, str]) -> int:
    return len(_document_text_from_path(args))


def _range_chunk_tokens(args: tuple[str, RangeChunk]) -> list[str]:
    epub_path, chunk = args
    text = _document_text_from_path((epub_path, chunk.zip_path))
    return tokenize_text(text[chunk.start : chunk.end])


def _validate_percent_range(start: float, end: float) -> None:
    if not (0 <= start < end <= 100):
        raise ValueError("Percent range must satisfy 0 <= start < end <= 100")


def resolve_epub_path(epub_input: str | Path) -> Path:
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


def default_txt_path(epub_path: str | Path) -> Path:
    return TXT_DIR / f"{Path(epub_path).stem}.txt"


def _map_ordered(function, values, workers: int):
    if workers <= 1 or len(values) <= 1:
        return [function(value) for value in values]
    with ThreadPoolExecutor(max_workers=workers) as executor:
        return list(executor.map(function, values))


def _epub_range_chunks(
    epub_path: str | Path,
    start_percent: float,
    end_percent: float,
    workers: int | None = None,
) -> list[RangeChunk]:
    _validate_percent_range(start_percent, end_percent)
    epub_path = str(epub_path)
    workers = default_worker_count() if workers is None else max(1, workers)

    with zipfile.ZipFile(epub_path) as epub:
        items = _spine_items(epub)

    paths = [item.zip_path for item in items]
    lengths = _map_ordered(_document_length_from_path, [(epub_path, path) for path in paths], workers)
    total_length = sum(lengths)
    if total_length == 0:
        return []

    start_char = int(total_length * (start_percent / 100))
    end_char = int(total_length * (end_percent / 100))
    end_char = max(start_char + 1, end_char)

    chunks: list[RangeChunk] = []
    offset = 0
    for path, length in zip(paths, lengths):
        doc_start = offset
        doc_end = offset + length
        offset = doc_end

        if doc_end <= start_char:
            continue
        if doc_start >= end_char:
            break

        local_start = max(0, start_char - doc_start)
        local_end = min(length, end_char - doc_start)
        if local_start < local_end:
            chunks.append(RangeChunk(path, local_start, local_end))
    return chunks


def iter_epub_range_text(
    epub_path: str | Path,
    start_percent: float,
    end_percent: float,
    workers: int | None = None,
) -> Iterable[str]:
    """Yield text chunks within the requested reading-progress percentage range."""
    epub_path = str(epub_path)
    workers = default_worker_count() if workers is None else max(1, workers)
    chunks = _epub_range_chunks(epub_path, start_percent, end_percent, workers)
    texts = _map_ordered(
        lambda value: _document_text_from_path(value),
        [(epub_path, chunk.zip_path) for chunk in chunks],
        workers,
    )
    for text, chunk in zip(texts, chunks):
        yield text[chunk.start : chunk.end]


def convert_epub_to_txt(
    epub_path: str | Path,
    txt_path: str | Path,
    start_percent: float = 0,
    end_percent: float = 100,
    workers: int | None = None,
) -> int:
    """Convert a percentage range of an EPUB to tokenized TXT.

    Returns the number of tokens written.
    """
    epub_path = str(epub_path)
    workers = default_worker_count() if workers is None else max(1, workers)
    chunks = _epub_range_chunks(epub_path, start_percent, end_percent, workers)
    token_count = 0
    line_tokens: list[str] = []
    txt_path = Path(txt_path)
    txt_path.parent.mkdir(parents=True, exist_ok=True)

    with txt_path.open("w", encoding="utf-8") as output:
        token_batches = _map_ordered(
            _range_chunk_tokens,
            [(epub_path, chunk) for chunk in chunks],
            workers,
        )
        for tokens in token_batches:
            for token in tokens:
                line_tokens.append(token)
                token_count += 1
                if len(line_tokens) >= 1000:
                    output.write(" ".join(line_tokens))
                    output.write("\n")
                    line_tokens.clear()
        if line_tokens:
            output.write(" ".join(line_tokens))
            output.write("\n")

    return token_count


def _parse_percent(value: str, default: float) -> float:
    value = value.strip()
    return default if value == "" else float(value)


def prompt_percent_range() -> tuple[float, float]:
    start = _parse_percent(input("Start percent [default 0]: "), 0)
    end = _parse_percent(input("End percent [default 100]: "), 100)
    _validate_percent_range(start, end)
    return start, end


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Extract EPUB words into a TXT file.")
    parser.add_argument("epub", help="EPUB path, or filename/stem inside books/.")
    parser.add_argument("txt", nargs="?", help="Output TXT path. Defaults to txtdir/<book>.txt.")
    parser.add_argument("--start", type=float, default=None, help="Start percent. Defaults to prompt.")
    parser.add_argument("--end", type=float, default=None, help="End percent. Defaults to prompt.")
    parser.add_argument("--workers", type=int, default=None, help="Concurrent worker count. Defaults to CPU-based value.")
    args = parser.parse_args(argv)

    if args.start is None or args.end is None:
        start, end = prompt_percent_range()
    else:
        start, end = args.start, args.end
        _validate_percent_range(start, end)

    epub_path = resolve_epub_path(args.epub)
    txt_path = Path(args.txt).expanduser() if args.txt else default_txt_path(epub_path)

    count = convert_epub_to_txt(epub_path, txt_path, start, end, args.workers)
    print(f"Wrote {count} tokens to {txt_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
