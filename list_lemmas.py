#!/usr/bin/env python3
"""List the unique lemmas found in each EPUB/TXT file or folder."""

from __future__ import annotations

import argparse
import sys
import unicodedata
from collections import Counter
from pathlib import Path
from typing import Iterable

from ankicheck import DEFAULT_ANKI_CONNECT_URL, check_lemmas
from epub_to_txt import iter_epub_range_text, resolve_epub_path
from txt_to_excel import (
    VOCABULARY_EXCLUDE_DIR,
    collect_grouped_lemma_frequencies,
    iter_text_chunks,
    load_excluded_words,
    load_spacy_model,
    resolve_txt_path,
)


# When an input path is a folder, scan it for files with these suffixes (case-insensitive).
# Edit this list to change what a folder picks up, e.g. add ".txt" for extracted text files.
FOLDER_SCAN_SUFFIXES = (".epub",)

# Default frequency thresholds: each x produces a Freq>=x / Freq<x column pair.
DEFAULT_FREQ_THRESHOLDS = (10, 5)


def parse_freq_thresholds(raw: str) -> list[int]:
    """Parse a comma-separated threshold list like '10,5' into ordered ints."""
    thresholds: list[int] = []
    seen: set[int] = set()
    for part in raw.split(","):
        part = part.strip()
        if not part:
            raise ValueError(f"Empty threshold in '{raw}'")
        try:
            value = int(part)
        except ValueError:
            raise ValueError(f"Invalid frequency threshold: '{part}'")
        if value < 1:
            raise ValueError(f"Frequency threshold must be >= 1: {value}")
        if value not in seen:
            seen.add(value)
            thresholds.append(value)
    if not thresholds:
        raise ValueError(f"No valid frequency thresholds in '{raw}'")
    return thresholds


def frequency_band_members(
    lemma_forms: dict[str, Counter[str]],
    excluded_words: set[str],
    anki_known: dict[str, bool],
    thresholds: list[int],
) -> list[tuple[set[str], set[str]]]:
    """Return per-threshold (Freq>=x lemmas, Freq<x lemmas) after filtering.

    Follows the same filtering rules as the Excel workflow: vocabulary_exclude/
    words are dropped first, then a lemma is treated as known (removed) when
    AnkiConnect reports the lemma itself or any of its inflected forms in Anki.
    """
    members = [(set(), set()) for _ in thresholds]
    for lemma, forms in lemma_forms.items():
        if lemma in excluded_words or anki_known.get(lemma, False):
            continue
        count = sum(forms.values())
        for index, threshold in enumerate(thresholds):
            ge, lt = members[index]
            if count >= threshold:
                ge.add(lemma)
            else:
                lt.add(lemma)
    return members


def resolve_single_input(raw_input: str) -> Path:
    path = Path(raw_input).expanduser()
    if path.is_file():
        return path

    for resolver in (resolve_epub_path, resolve_txt_path):
        try:
            return resolver(raw_input)
        except FileNotFoundError:
            continue
    raise FileNotFoundError(f"Input not found (also checked books/ and txtdir/): {raw_input}")


def resolve_input_paths(raw_inputs: Iterable[str]) -> list[Path]:
    """Expand each input into files; folders are scanned for FOLDER_SCAN_SUFFIXES."""
    suffixes = {suffix.lower() for suffix in FOLDER_SCAN_SUFFIXES}
    resolved: list[Path] = []
    seen: set[Path] = set()
    for raw_input in raw_inputs:
        path = Path(raw_input).expanduser()
        if path.is_dir():
            found = sorted(
                child
                for child in path.iterdir()
                if child.is_file() and child.suffix.lower() in suffixes
            )
            if not found:
                raise FileNotFoundError(
                    f"No files with suffixes {', '.join(sorted(suffixes))} found in folder: {raw_input}"
                )
        else:
            found = [resolve_single_input(raw_input)]

        for candidate in found:
            if candidate not in seen:
                seen.add(candidate)
                resolved.append(candidate)
    return resolved


def _file_text_chunks(path: Path, start_percent: float, end_percent: float) -> Iterable[str]:
    if path.suffix.lower() == ".epub":
        return iter_epub_range_text(path, start_percent, end_percent)
    return iter_text_chunks(path)


def sorted_lemmas(
    lemma_forms: dict,
    excluded_words: set[str],
    sort_order: str,
) -> list[str]:
    entries = [
        (lemma, sum(forms.values()))
        for lemma, forms in lemma_forms.items()
        if lemma not in excluded_words
    ]
    if sort_order == "count":
        entries.sort(key=lambda item: (-item[1], item[0]))
        return [lemma for lemma, _count in entries]
    return sorted(lemma for lemma, _count in entries)


def _lemma_output_path(out_dir: Path, path: Path, used_names: set[str]) -> Path:
    for name in (f"{path.stem}.txt", f"{path.name}.txt"):
        if name not in used_names and not (out_dir / name).exists():
            used_names.add(name)
            return out_dir / name
    index = 2
    while True:
        name = f"{path.stem}_{index}.txt"
        if name not in used_names and not (out_dir / name).exists():
            used_names.add(name)
            return out_dir / name
        index += 1


def _counted_chunks(text_chunks: Iterable[str], char_total: list[int]) -> Iterable[str]:
    """Pass chunks through while accumulating their character count in char_total[0]."""
    for chunk in text_chunks:
        char_total[0] += len(chunk)
        yield chunk


def _display_width(text: str) -> int:
    """Terminal display width: CJK full-width characters count as two columns."""
    return sum(2 if unicodedata.east_asian_width(char) in ("W", "F") else 1 for char in text)


def _ljust_display(text: str, width: int) -> str:
    return text + " " * max(0, width - _display_width(text))


def _print_table(results: list[tuple], thresholds: list[int], totals: list, distincts: list) -> None:
    name_width = max(
        _display_width("Book"),
        _display_width("去重合计"),
        *(_display_width(path.name) for path, *_ in results),
    )
    char_width = max(len("Characters"), *(len(str(row[1])) for row in results))
    lemma_width = max(len("Unique lemmas"), *(len(str(row[2])) for row in results))
    band_headers = [header for x in thresholds for header in (f"Freq>={x}", f"Freq<{x}")]
    band_widths = [
        max(len(header), *(len(str(row[3 + index])) for row in results))
        for index, header in enumerate(band_headers)
    ]
    header_line = f"{_ljust_display('Book', name_width)}  {'Characters':>{char_width}}  {'Unique lemmas':>{lemma_width}}"
    for header, width in zip(band_headers, band_widths):
        header_line += f"  {header:>{width}}"
    rule = "-" * (name_width + char_width + lemma_width + sum(band_widths) + 2 * (2 + len(band_widths)))
    print(header_line)
    print(rule)
    for row in results:
        line = f"{_ljust_display(row[0].name, name_width)}  {row[1]:>{char_width}}  {row[2]:>{lemma_width}}"
        for index, width in enumerate(band_widths):
            line += f"  {row[3 + index]:>{width}}"
        print(line)
    print(rule)
    for label, row in (("合计", totals), ("去重合计", distincts)):
        line = f"{_ljust_display(label, name_width)}  {row[1]:>{char_width}}  {row[2]:>{lemma_width}}"
        for index, width in enumerate(band_widths):
            line += f"  {row[3 + index]:>{width}}"
        print(line)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="List the unique lemmas of each EPUB/TXT file or folder.",
        epilog="Bare names are also looked up as books/<name>.epub and txtdir/<name>.txt.",
    )
    parser.add_argument("paths", nargs="+", help="EPUB/TXT files, bare book names, or folders to scan.")
    parser.add_argument("--out-dir", help="Write each file's full lemma list to <out-dir>/<stem>.txt (one lemma per line); the terminal only prints counts.")
    parser.add_argument("--no-exclude", action="store_true", help="Keep words from vocabulary_exclude/ instead of dropping them.")
    parser.add_argument("--anki-url", default=DEFAULT_ANKI_CONNECT_URL, help="AnkiConnect URL for the already-in-Anki check. Defaults to http://127.0.0.1:8765.")
    parser.add_argument("--no-anki-words", action="store_true", help="Do not remove words already in Anki; keep them in the frequency band counts.")
    parser.add_argument("--freq", default="10,5", help="Comma-separated frequency thresholds, e.g. '10,5,3'; each x adds Freq>=x and Freq<x columns. Defaults to 10,5.")
    parser.add_argument("--sort", choices=("alpha", "count"), default="alpha", help="Lemma order. Defaults to alpha.")
    parser.add_argument("--start", type=float, default=0, help="Start percent of EPUB content. Defaults to 0.")
    parser.add_argument("--end", type=float, default=100, help="End percent of EPUB content. Defaults to 100.")
    parser.add_argument("--workers", type=int, default=None, help="spaCy n_process worker count. Defaults to CPU count.")
    parser.add_argument("--model", default="en_core_web_sm", help="spaCy model name. Defaults to en_core_web_sm.")
    parser.add_argument("--vocabulary-dir", default=str(VOCABULARY_EXCLUDE_DIR), help="Directory of vocabulary files to exclude.")
    args = parser.parse_args(argv)

    if not (0 <= args.start < args.end <= 100):
        parser.error("Percent range must satisfy 0 <= start < end <= 100")
    try:
        freq_thresholds = parse_freq_thresholds(args.freq)
    except ValueError as exc:
        parser.error(str(exc))

    paths = resolve_input_paths(args.paths)
    excluded_words = set() if args.no_exclude else load_excluded_words(args.vocabulary_dir)
    out_dir = Path(args.out_dir).expanduser() if args.out_dir else None
    if out_dir is not None:
        out_dir.mkdir(parents=True, exist_ok=True)
    used_names: set[str] = set()

    nlp = load_spacy_model(args.model)

    # All files stream through one shared spaCy pass; the generator tags each
    # chunk with its file index and records per-file character totals.
    file_progress: list[tuple[Path, list[int]]] = []

    def _keyed_chunks() -> Iterable[tuple[int, str]]:
        for index, path in enumerate(paths):
            print(f"Processing {path.name} ...", file=sys.stderr)
            char_total = [0]
            file_progress.append((path, char_total))
            yield from (
                (index, chunk)
                for chunk in _counted_chunks(_file_text_chunks(path, args.start, args.end), char_total)
            )

    grouped = collect_grouped_lemma_frequencies(_keyed_chunks(), nlp, args.workers)

    # One AnkiConnect pass over every book's lemmas; a lemma is "known" if the
    # lemma or any of its inflected forms exists in Anki (same rule as the
    # Excel workflow). check_lemmas fails open: if Anki/AnkiConnect is not
    # reachable it prints a warning and removes nothing.
    anki_known: dict[str, bool] = {}
    if not args.no_anki_words:
        candidates = {
            lemma: {lemma, *forms.keys()}
            for book_lemmas in grouped.values()
            for lemma, forms in book_lemmas.items()
            if lemma not in excluded_words
        }
        anki_known = check_lemmas(candidates, url=args.anki_url)

    results: list[tuple] = []
    all_lemmas: set[str] = set()
    band_unions: list[tuple[set[str], set[str]]] = [(set(), set()) for _ in freq_thresholds]
    for index, (path, char_total) in enumerate(file_progress):
        grouped_book = grouped.get(index, {})
        lemmas = sorted_lemmas(grouped_book, excluded_words, args.sort)
        all_lemmas.update(lemmas)
        if out_dir is not None:
            out_path = _lemma_output_path(out_dir, path, used_names)
            out_path.write_text("".join(f"{lemma}\n" for lemma in lemmas), encoding="utf-8")
        members = frequency_band_members(grouped_book, excluded_words, anki_known, freq_thresholds)
        row: list = [path, char_total[0], len(lemmas)]
        for band_index, (ge, lt) in enumerate(members):
            band_unions[band_index][0].update(ge)
            band_unions[band_index][1].update(lt)
            row.extend((len(ge), len(lt)))
        results.append(tuple(row))

    # Table order: most unique lemmas first, book name as tie-breaker.
    results.sort(key=lambda row: (-row[2], row[0].name.lower()))

    # Footer rows: 合计 sums each column across books; 去重合计 counts distinct
    # lemmas across books (Characters has no deduped value).
    band_count = len(freq_thresholds) * 2
    totals = ["-", sum(row[1] for row in results), sum(row[2] for row in results)]
    totals += [sum(row[3 + i] for row in results) for i in range(band_count)]
    distincts = ["-", "-", len(all_lemmas)]
    for ge, lt in band_unions:
        distincts.extend((len(ge), len(lt)))
    _print_table(results, freq_thresholds, totals, distincts)
    if out_dir is not None:
        print(f"\nLemma lists written under: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
