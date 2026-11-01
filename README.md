# English_words_frequency

## Prerequisites

Install the Python dependencies, then download the spaCy English model used for lemmatization:

```bash
python -m pip install -r requirements.txt
python -m spacy download en_core_web_sm
```

## List unique lemmas per file

`list_lemmas.py` accepts EPUB files, TXT files, bare book names (resolved against `books/` and `txtdir/`), or folders. When done it prints a table per file — Book, Characters, Unique lemmas, and a Freq>=x / Freq<x column pair for every threshold x in `--freq` (default `10,5`); two footer rows follow every column — 合计 sums each column across books, 去重合计 counts the distinct lemmas across books (Characters has no deduped value). Pass `--out-dir` to also write each file's full lemma list to `<out-dir>/<stem>.txt` (one lemma per line):

```bash
# Single TXT (bare names are looked up in txtdir/)
python list_lemmas.py "The Magic Finger"

# Single EPUB (bare names are looked up in books/)
python list_lemmas.py "Death on the Nile"

# Folders scan for EPUBs by default
python list_lemmas.py books/

# Write full lemma lists to lemmasdir/<stem>.txt
python list_lemmas.py --out-dir lemmasdir books/

# Order by frequency instead of alphabet, and keep vocabulary_exclude/ words
python list_lemmas.py --out-dir lemmasdir --sort count --no-exclude "Wonder (R. J. Palacio)"
```

Folder scanning only picks up the suffixes listed in `FOLDER_SCAN_SUFFIXES` at the top of `list_lemmas.py` (default: `.epub` only); edit that list, e.g. add `".txt"`, to scan more formats.

All files are lemmatized in a single spaCy pass (parser and NER disabled — lemma results are identical), so a whole folder costs the multi-process startup only once. For tiny inputs, `--workers 1` skips worker spawning entirely and is often faster.

Words listed in `vocabulary_exclude/` are dropped by default (same rule as the Excel workflow); pass `--no-exclude` to keep them. The frequency band columns count the lemmas in each Freq>=x / Freq<x band, after dropping vocabulary_exclude/ words and the words already in Anki; pass `--freq "20,10,5"` (any comma-separated list of positive integers, duplicates removed) to choose your own thresholds. The Anki check queries the local Anki instance through AnkiConnect (same as the Excel workflow: a lemma counts as known if the lemma itself or any of its inflected forms appears in an `OnlineDictHelper` note's `expression` field); keep Anki running with the AnkiConnect plugin, or point `--anki-url` elsewhere. If AnkiConnect is unreachable the script prints a warning and removes nothing; pass `--no-anki-words` to skip the Anki check entirely. `--start/--end` restrict EPUB inputs to a reading-progress percentage range, like the other scripts.
