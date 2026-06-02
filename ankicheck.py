#!/usr/bin/env python3
"""Check whether words already exist in local Anki through AnkiConnect."""

from __future__ import annotations

import argparse
import html
import json
import re
import urllib.error
import urllib.request
from collections.abc import Collection
from dataclasses import dataclass, field
from typing import Any


DEFAULT_ANKI_CONNECT_URL = "http://127.0.0.1:8765"
DEFAULT_MODEL_NAME = "OnlineDictHelper"
DEFAULT_FIELD_NAME = "expression"
WORD_RE = re.compile(r"^[A-Za-z]+$")
HTML_TAG_RE = re.compile(r"<[^>]+>")
NOTES_INFO_CHUNK_SIZE = 500


class AnkiConnectError(RuntimeError):
    """Raised when AnkiConnect is unavailable or returns an error."""


def _clean_field_text(value: str) -> str:
    return html.unescape(HTML_TAG_RE.sub(" ", value)).lower()


def _escape_query_value(value: str) -> str:
    escaped = value.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"' if any(char.isspace() for char in escaped) else escaped


@dataclass
class AnkiChecker:
    url: str = DEFAULT_ANKI_CONNECT_URL
    model_name: str = DEFAULT_MODEL_NAME
    field_name: str = DEFAULT_FIELD_NAME
    timeout: float = 10
    _cache: dict[str, bool] = field(default_factory=dict)
    _known_words: set[str] | None = None

    def request(self, action: str, params: dict[str, Any] | None = None) -> Any:
        payload = json.dumps({"action": action, "version": 6, "params": params or {}}).encode()
        request = urllib.request.Request(
            self.url,
            data=payload,
            headers={"Content-Type": "application/json"},
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                body = json.loads(response.read().decode("utf-8"))
        except (OSError, urllib.error.URLError, TimeoutError, json.JSONDecodeError) as exc:
            raise AnkiConnectError(f"AnkiConnect request failed: {exc}") from exc

        if body.get("error") is not None:
            raise AnkiConnectError(str(body["error"]))
        return body.get("result")

    def available(self) -> bool:
        try:
            return bool(self.request("version"))
        except AnkiConnectError:
            return False

    def known_words(self) -> set[str]:
        if self._known_words is not None:
            return self._known_words

        note_ids = self.request("findNotes", {"query": f"note:{_escape_query_value(self.model_name)}"})
        known: set[str] = set()
        for index in range(0, len(note_ids or []), NOTES_INFO_CHUNK_SIZE):
            notes = self.request("notesInfo", {"notes": note_ids[index:index + NOTES_INFO_CHUNK_SIZE]})
            for note in notes or []:
                fields = note.get("fields", {})
                field_value = fields.get(self.field_name, {}).get("value", "")
                for match in re.finditer(r"[A-Za-z]+", _clean_field_text(field_value)):
                    known.add(match.group(0).lower())

        self._known_words = known
        return known

    def word_exists(self, word: str) -> bool:
        word = word.lower().strip()
        if not WORD_RE.fullmatch(word):
            return False
        if word in self._cache:
            return self._cache[word]

        result = word in self.known_words()
        self._cache[word] = result
        return result

    def lemma_exists(self, candidates: Collection[str]) -> bool:
        for candidate in sorted({item.lower().strip() for item in candidates}):
            if self.word_exists(candidate):
                return True
        return False


def check_lemmas(
    lemma_candidates: dict[str, Collection[str]],
    url: str = DEFAULT_ANKI_CONNECT_URL,
    model_name: str = DEFAULT_MODEL_NAME,
    field_name: str = DEFAULT_FIELD_NAME,
    fail_open: bool = True,
) -> dict[str, bool]:
    checker = AnkiChecker(url=url, model_name=model_name, field_name=field_name)
    try:
        return {
            lemma: checker.lemma_exists({lemma, *candidates})
            for lemma, candidates in lemma_candidates.items()
        }
    except AnkiConnectError as exc:
        if not fail_open:
            raise
        print(f"Warning: Anki check skipped: {exc}")
        return {lemma: False for lemma in lemma_candidates}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Check words in local Anki through AnkiConnect.")
    parser.add_argument("words", nargs="+", help="Words to check.")
    parser.add_argument("--anki-url", default=DEFAULT_ANKI_CONNECT_URL, help="AnkiConnect URL.")
    parser.add_argument("--model", default=DEFAULT_MODEL_NAME, help="Anki note model name.")
    parser.add_argument("--field", default=DEFAULT_FIELD_NAME, help="Anki field name to match.")
    args = parser.parse_args(argv)

    checker = AnkiChecker(args.anki_url, args.model, args.field)
    for word in args.words:
        print(f"{word}\t{'yes' if checker.word_exists(word) else ''}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
