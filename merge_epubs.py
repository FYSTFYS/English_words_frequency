#!/usr/bin/env python3
"""Merge multiple EPUB files into one EPUB.

This script is intentionally separate from the word-frequency workflow. It
creates a new EPUB whose spine is the concatenation of the input EPUB spines.
"""

from __future__ import annotations

import argparse
import html
import posixpath
import re
import uuid
import zipfile
from dataclasses import dataclass
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
BOOKS_DIR = BASE_DIR / "books"
EPUB_MIMETYPE = b"application/epub+zip"
CONTAINER_NS = "urn:oasis:names:tc:opendocument:xmlns:container"
OPF_NS = "http://www.idpf.org/2007/opf"
DC_NS = "http://purl.org/dc/elements/1.1/"
XHTML_NS = "http://www.w3.org/1999/xhtml"
TAG_RE = re.compile(r"<\s*([A-Za-z_][\w:.-]*)([^>]*)>", re.DOTALL)
ATTR_RE = re.compile(r"([A-Za-z_][\w:.-]*)\s*=\s*(['\"])(.*?)\2", re.DOTALL)


@dataclass(frozen=True)
class ManifestItem:
    id: str
    href: str
    media_type: str
    properties: str | None = None


@dataclass(frozen=True)
class EpubPackage:
    path: Path
    title: str
    opf_path: str
    manifest: list[ManifestItem]
    spine_ids: list[str]


def _local_name(name: str) -> str:
    return name.rsplit(":", 1)[-1]


def _decode_xml(data: bytes) -> str:
    return data.decode("utf-8", errors="replace")


def _iter_xml_tags(xml_text: str, local_name: str):
    for match in TAG_RE.finditer(xml_text):
        name, attrs_text = match.groups()
        if _local_name(name) != local_name:
            continue
        attrs = {
            _local_name(attr_name): html.unescape(value)
            for attr_name, _quote, value in ATTR_RE.findall(attrs_text)
        }
        yield attrs


def _first_text(xml_text: str, local_name: str) -> str | None:
    pattern = re.compile(
        rf"<\s*(?:[A-Za-z_][\w.-]*:)?{re.escape(local_name)}(?:\s[^>]*)?>(.*?)</\s*(?:[A-Za-z_][\w.-]*:)?{re.escape(local_name)}\s*>",
        re.DOTALL,
    )
    match = pattern.search(xml_text)
    if not match:
        return None
    text = re.sub(r"<[^>]+>", " ", match.group(1))
    text = re.sub(r"\s+", " ", html.unescape(text)).strip()
    return text or None


def _safe_id(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9_.-]+", "_", value).strip("_")
    if not value or not re.match(r"[A-Za-z_]", value):
        value = f"id_{value}"
    return value


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


def _container_opf_path(epub: zipfile.ZipFile, epub_path: Path) -> str:
    try:
        xml_text = _decode_xml(epub.read("META-INF/container.xml"))
    except KeyError as exc:
        raise ValueError(f"EPUB is missing META-INF/container.xml: {epub_path}") from exc

    for rootfile in _iter_xml_tags(xml_text, "rootfile"):
        full_path = rootfile.get("full-path")
        if full_path:
            return full_path

    raise ValueError(f"EPUB container does not contain a rootfile entry: {epub_path}")


def _read_package(epub_path: Path) -> EpubPackage:
    with zipfile.ZipFile(epub_path) as epub:
        opf_path = _container_opf_path(epub, epub_path)
        try:
            opf_text = _decode_xml(epub.read(opf_path))
        except KeyError as exc:
            raise ValueError(f"EPUB rootfile is missing: {epub_path}:{opf_path}") from exc

    title = epub_path.stem
    if package_title := _first_text(opf_text, "title"):
        title = package_title

    manifest: list[ManifestItem] = []
    id_to_item: dict[str, ManifestItem] = {}
    for item in _iter_xml_tags(opf_text, "item"):
        item_id = item.get("id")
        href = item.get("href")
        media_type = item.get("media-type")
        if not item_id or not href or not media_type:
            continue
        manifest_item = ManifestItem(
            id=item_id,
            href=href,
            media_type=media_type,
            properties=item.get("properties"),
        )
        manifest.append(manifest_item)
        id_to_item[item_id] = manifest_item

    if not manifest:
        raise ValueError(f"EPUB package has no manifest items: {epub_path}")

    spine_ids = [
        idref
        for itemref in _iter_xml_tags(opf_text, "itemref")
        if (idref := itemref.get("idref")) and idref in id_to_item
    ]
    if not spine_ids:
        raise ValueError(f"EPUB spine does not reference readable manifest items: {epub_path}")

    return EpubPackage(
        path=epub_path,
        title=title,
        opf_path=opf_path,
        manifest=manifest,
        spine_ids=spine_ids,
    )


def _join_zip_path(*parts: str) -> str:
    return posixpath.normpath(posixpath.join(*parts)).lstrip("/")


def _package_zip_path(package: EpubPackage, item: ManifestItem) -> str:
    opf_dir = posixpath.dirname(package.opf_path)
    return _join_zip_path(opf_dir, item.href)


def _unique_output_path(path: str, used_paths: set[str]) -> str:
    candidate = path
    index = 2
    while candidate in used_paths:
        stem, suffix = posixpath.splitext(path)
        candidate = f"{stem}_{index}{suffix}"
        index += 1
    used_paths.add(candidate)
    return candidate


def _merged_nav_html(title: str, packages: list[EpubPackage], first_hrefs: list[str]) -> bytes:
    links = "\n".join(
        f'      <li><a href="{html.escape(href, quote=True)}">{html.escape(package.title)}</a></li>'
        for package, href in zip(packages, first_hrefs)
    )
    return f"""<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html>
<html xmlns="{XHTML_NS}" lang="en">
  <head>
    <title>{html.escape(title)}</title>
  </head>
  <body>
    <nav epub:type="toc" id="toc">
      <h1>{html.escape(title)}</h1>
      <ol>
{links}
      </ol>
    </nav>
  </body>
</html>
""".encode("utf-8")


def _merged_opf(title: str, author: str | None, manifest: list[ManifestItem], spine_ids: list[str]) -> bytes:
    identifier = f"urn:uuid:{uuid.uuid4()}"
    author_xml = f"    <dc:creator>{html.escape(author)}</dc:creator>\n" if author else ""
    manifest_xml = "\n".join(
        (
            f'    <item id="{html.escape(item.id, quote=True)}" '
            f'href="{html.escape(item.href, quote=True)}" '
            f'media-type="{html.escape(item.media_type, quote=True)}"'
            f'{f" properties=\"{html.escape(item.properties, quote=True)}\"" if item.properties else ""}/>'
        )
        for item in manifest
    )
    spine_xml = "\n".join(
        f'    <itemref idref="{html.escape(item_id, quote=True)}"/>'
        for item_id in spine_ids
    )
    return f"""<?xml version="1.0" encoding="utf-8"?>
<package xmlns="{OPF_NS}" unique-identifier="bookid" version="3.0">
  <metadata xmlns:dc="{DC_NS}">
    <dc:identifier id="bookid">{identifier}</dc:identifier>
    <dc:title>{html.escape(title)}</dc:title>
{author_xml}    <dc:language>en</dc:language>
  </metadata>
  <manifest>
{manifest_xml}
  </manifest>
  <spine>
{spine_xml}
  </spine>
</package>
""".encode("utf-8")


def _container_xml() -> bytes:
    return f"""<?xml version="1.0" encoding="utf-8"?>
<container version="1.0" xmlns="{CONTAINER_NS}">
  <rootfiles>
    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>
""".encode("utf-8")


def merge_epubs(
    epub_paths: list[str | Path],
    output_path: str | Path,
    title: str = "Merged EPUB",
    author: str | None = None,
) -> Path:
    if len(epub_paths) < 2:
        raise ValueError("At least two EPUB inputs are required")

    resolved_paths = [resolve_epub_path(path) for path in epub_paths]
    packages = [_read_package(path) for path in resolved_paths]
    output_path = Path(output_path).expanduser()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    merged_manifest: list[ManifestItem] = [
        ManifestItem(id="nav", href="nav.xhtml", media_type="application/xhtml+xml", properties="nav")
    ]
    merged_spine: list[str] = []
    first_hrefs: list[str] = []
    used_output_paths = {"OEBPS/nav.xhtml", "OEBPS/content.opf"}
    copied_files: list[tuple[Path, str, str]] = []

    for book_index, package in enumerate(packages, start=1):
        prefix = f"book_{book_index}"
        item_id_map: dict[str, str] = {}
        item_href_map: dict[str, str] = {}

        for item in package.manifest:
            source_zip_path = _package_zip_path(package, item)
            target_zip_path = _unique_output_path(f"OEBPS/{prefix}/{source_zip_path}", used_output_paths)
            merged_href = posixpath.relpath(target_zip_path, "OEBPS")
            merged_id = _safe_id(f"b{book_index}_{item.id}")
            item_id_map[item.id] = merged_id
            item_href_map[item.id] = merged_href
            copied_files.append((package.path, source_zip_path, target_zip_path))
            merged_manifest.append(
                ManifestItem(
                    id=merged_id,
                    href=merged_href,
                    media_type=item.media_type,
                    properties=None if item.properties == "nav" else item.properties,
                )
            )

        first_hrefs.append(item_href_map[package.spine_ids[0]])
        for idref in package.spine_ids:
            merged_id = item_id_map[idref]
            merged_spine.append(merged_id)

    with zipfile.ZipFile(output_path, "w") as output:
        output.writestr(zipfile.ZipInfo("mimetype"), EPUB_MIMETYPE, compress_type=zipfile.ZIP_STORED)
        output.writestr("META-INF/container.xml", _container_xml(), compress_type=zipfile.ZIP_DEFLATED)
        output.writestr("OEBPS/nav.xhtml", _merged_nav_html(title, packages, first_hrefs), compress_type=zipfile.ZIP_DEFLATED)
        output.writestr("OEBPS/content.opf", _merged_opf(title, author, merged_manifest, merged_spine), compress_type=zipfile.ZIP_DEFLATED)

        for source_epub_path, source_zip_path, target_zip_path in copied_files:
            with zipfile.ZipFile(source_epub_path) as source_epub:
                try:
                    data = source_epub.read(source_zip_path)
                except KeyError as exc:
                    raise ValueError(f"EPUB manifest references missing file: {source_epub_path}:{source_zip_path}") from exc
            output.writestr(target_zip_path, data, compress_type=zipfile.ZIP_DEFLATED)

    return output_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Merge multiple EPUB files into one EPUB.")
    parser.add_argument("epubs", nargs="+", help="EPUB paths, filenames, or stems inside books/.")
    parser.add_argument("-o", "--output", required=True, help="Output EPUB path.")
    parser.add_argument("--title", default="Merged EPUB", help="Title for the merged EPUB.")
    parser.add_argument("--author", help="Author/creator metadata for the merged EPUB.")
    args = parser.parse_args(argv)

    try:
        output_path = merge_epubs(args.epubs, args.output, args.title, args.author)
    except (FileNotFoundError, ValueError, zipfile.BadZipFile) as exc:
        parser.error(str(exc))

    print(f"Merged {len(args.epubs)} EPUBs into {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
