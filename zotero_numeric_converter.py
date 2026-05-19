#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
zotero_numeric_converter.py

Convert plain numeric citations in a Word .docx document, such as [1], [1,2],
or [3-6], into Word ADDIN fields carrying Zotero-style CSL_CITATION JSON.

Typical use:
    python zotero_numeric_converter.py --docx paper.docx --bib references.bib --zotero-sqlite C:\\Users\\YOU\\Zotero\\zotero.sqlite --out paper_zotero.docx --dry-run --require-zotero-uris
    python zotero_numeric_converter.py --docx paper.docx --bib references.bib --zotero-sqlite C:\\Users\\YOU\\Zotero\\zotero.sqlite --out paper_zotero.docx

Important notes:
- The script keeps a detailed report and never edits the original file in place.
- It relies mainly on DOI matching between the Reference section and the BibTeX file.
- For true library-linked Zotero Word fields, use --zotero-sqlite. Standard
  Zotero RDF often does not include library item URIs/keys; it may contain only
  DOI URLs and is therefore insufficient by itself. After conversion, open the
  output .docx in Word with Zotero installed and use Zotero -> Refresh.
"""

from __future__ import annotations

import argparse
import csv
import datetime as _dt
import json
import os
import random
import re
import shutil
import sqlite3
import string
import sys
import tempfile
import uuid
import zipfile
from copy import deepcopy
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

try:
    from lxml import etree
except Exception as exc:  # pragma: no cover
    raise SystemExit("Missing dependency: lxml. Install with: pip install lxml") from exc

try:
    import bibtexparser  # type: ignore
except Exception:  # pragma: no cover
    bibtexparser = None


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
RDF_ABOUT = "{http://www.w3.org/1999/02/22-rdf-syntax-ns#}about"
RDF_RESOURCE = "{http://www.w3.org/1999/02/22-rdf-syntax-ns#}resource"
XML_NS = "http://www.w3.org/XML/1998/namespace"
W = NS["w"]

for prefix, uri in NS.items():
    etree.register_namespace(prefix, uri)


def qn(tag: str) -> str:
    """Qualified OOXML tag, e.g. qn('w:p')."""
    prefix, local = tag.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


DOI_RE = re.compile(r"10\.\d{4,9}/[^\s\"<>]+", re.IGNORECASE)
CITATION_RE = re.compile(
    r"\[(\s*\d+\s*(?:(?:[,，;；]\s*\d+\s*)|(?:[-–—]\s*\d+\s*))*?)\]"
)
REFERENCE_HEADING_RE = re.compile(
    r"""^\s*
    (?:[第\s]*\d+(?:\.\d+)*[章节部篇\s]*[\.．、:：)]*\s*)?
    (?:
        references?|
        bibliography|bibliographies|
        works\s+cited|literature\s+cited|
        reference\s+list|cited\s+references|
        参考文献|参考资料|参考书目|引用文献|文献
    )
    \s*(?:[:：。.]|$)\s*$""",
    re.IGNORECASE | re.VERBOSE,
)
REFERENCE_START_RE = re.compile(
    r"^\s*(?:\[(\d{1,5})\]|(\d{1,5})\s*[\.)、]|(\d{1,5})\s{2,})(.*)$"
)


@dataclass
class BibRecord:
    key: str
    entry_type: str
    raw: dict
    doi: str = ""
    title: str = ""
    csl_item: dict = field(default_factory=dict)
    uris: List[str] = field(default_factory=list)
    zotero_item_id: Optional[int] = None
    zotero_item_key: str = ""
    zotero_library_id: Optional[int] = None


@dataclass
class RdfRecord:
    uri: str
    item_key: str = ""
    doi: str = ""
    title: str = ""
    csl_item: dict = field(default_factory=dict)
    source: str = "rdf"


@dataclass
class RefMatch:
    number: int
    reference_text: str
    doi: str = ""
    bib_key: str = ""
    status: str = "unmatched"
    message: str = ""
    title_similarity: float = 0.0
    record: Optional[BibRecord] = None


@dataclass
class ConversionLog:
    part: str
    paragraph_index: int
    citation_text: str
    numbers: str
    action: str
    message: str


# ---------------------------------------------------------------------------
# Basic text normalization
# ---------------------------------------------------------------------------


def strip_latex_braces(value: str) -> str:
    if not value:
        return ""
    value = re.sub(r"[{}]", "", value)
    value = value.replace("~", " ")
    value = re.sub(r"\\&", "&", value)
    value = re.sub(r"\\%", "%", value)
    value = re.sub(r"\\_", "_", value)
    value = re.sub(r"\\[a-zA-Z]+\s*", "", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def norm_space(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def normalize_title(value: str) -> str:
    value = strip_latex_braces(value).lower()
    value = re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", value)
    return norm_space(value)


def normalize_doi(value: str) -> str:
    """Return a normalized DOI from a DOI field, URL, or free text."""
    if not value:
        return ""
    s = str(value).strip()
    s = s.replace("\u200b", "").replace("\ufeff", "")
    s = re.sub(r"https?://(?:dx\.)?doi\.org/", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^doi\s*:\s*", "", s, flags=re.IGNORECASE)
    m = DOI_RE.search(s)
    doi = m.group(0) if m else s
    doi = doi.strip().strip(" <>\t\r\n")
    # Common trailing punctuation introduced by reference-list formatting.
    while doi and doi[-1] in ".,;，；。]}":
        doi = doi[:-1]
    # A closing parenthesis is often punctuation, but can be part of a DOI.
    # Remove only clearly unbalanced trailing parentheses.
    while doi.endswith(")") and doi.count("(") < doi.count(")"):
        doi = doi[:-1]
    return doi.lower()


# ---------------------------------------------------------------------------
# BibTeX -> CSL JSON helpers
# ---------------------------------------------------------------------------


CSL_TYPE_MAP = {
    "article": "article-journal",
    "journalarticle": "article-journal",
    "inproceedings": "paper-conference",
    "conference": "paper-conference",
    "proceedings": "paper-conference",
    "book": "book",
    "inbook": "chapter",
    "incollection": "chapter",
    "chapter": "chapter",
    "phdthesis": "thesis",
    "mastersthesis": "thesis",
    "thesis": "thesis",
    "techreport": "report",
    "report": "report",
    "misc": "article",
    "online": "webpage",
    "webpage": "webpage",
}


def split_bibtex_names(author_field: str) -> List[str]:
    """Split a BibTeX author/editor field conservatively on ' and '."""
    if not author_field:
        return []
    # This simple splitter is sufficient for standard Zotero BibTeX exports.
    parts = re.split(r"\s+and\s+", author_field.strip(), flags=re.IGNORECASE)
    return [strip_latex_braces(p).strip() for p in parts if p.strip() and p.strip().lower() != "others"]


def bibtex_name_to_csl(name: str) -> dict:
    name = norm_space(name)
    if not name:
        return {}
    if "," in name:
        chunks = [c.strip() for c in name.split(",")]
        family = chunks[0]
        given = " ".join(c for c in chunks[1:] if c)
    else:
        chunks = name.split()
        if len(chunks) == 1:
            family, given = chunks[0], ""
        else:
            family, given = chunks[-1], " ".join(chunks[:-1])
    out = {"family": family}
    if given:
        out["given"] = given
    return out


def parse_year(entry: dict) -> Optional[List[List[int]]]:
    for key in ("year", "date"):
        value = strip_latex_braces(entry.get(key, ""))
        if value:
            m = re.search(r"(\d{4})", value)
            if m:
                return [[int(m.group(1))]]
    return None


def get_field(entry: dict, *names: str) -> str:
    lower = {str(k).lower(): v for k, v in entry.items()}
    for name in names:
        if name.lower() in lower and lower[name.lower()] is not None:
            return strip_latex_braces(str(lower[name.lower()]))
    return ""


def extract_zotero_uris(entry: dict) -> List[str]:
    """Try to recover Zotero item URIs from richer BibTeX exports.

    Standard Zotero BibTeX exports often do not include internal item URIs. Better
    BibTeX or customized exports sometimes include fields such as uri, zotero-uri,
    or values containing zotero.org/.../items/KEY.
    """
    uris: List[str] = []
    for key, value in entry.items():
        text = str(value or "")
        if "zotero.org/" in text and "/items/" in text:
            found = re.findall(r"https?://zotero\.org/[^\s,;{}]+/items/[A-Za-z0-9]+", text)
            uris.extend(found)
        if "zotero://select/" in text:
            uris.append(text.strip())
    # Remove duplicates while preserving order.
    seen = set()
    out = []
    for uri in uris:
        if uri not in seen:
            out.append(uri)
            seen.add(uri)
    return out


def bibtex_entry_to_csl(entry: dict) -> dict:
    key = entry.get("ID") or entry.get("id") or entry.get("key") or str(uuid.uuid4())
    entry_type = (entry.get("ENTRYTYPE") or entry.get("entrytype") or "misc").lower()
    csl_type = CSL_TYPE_MAP.get(entry_type, "article")

    title = get_field(entry, "title")
    doi = normalize_doi(get_field(entry, "doi"))
    url = get_field(entry, "url")
    container = get_field(entry, "journal", "journaltitle", "booktitle", "container-title")

    item = {
        "id": key,
        "type": csl_type,
    }
    if title:
        item["title"] = title
    if container:
        item["container-title"] = container
    if doi:
        item["DOI"] = doi
    if url:
        item["URL"] = url

    issued = parse_year(entry)
    if issued:
        item["issued"] = {"date-parts": issued}

    for bib_name, csl_name in [
        ("volume", "volume"),
        ("number", "issue"),
        ("issue", "issue"),
        ("pages", "page"),
        ("page", "page"),
        ("publisher", "publisher"),
        ("address", "publisher-place"),
        ("edition", "edition"),
        ("isbn", "ISBN"),
        ("issn", "ISSN"),
    ]:
        val = get_field(entry, bib_name)
        if val and csl_name not in item:
            item[csl_name] = val

    authors = split_bibtex_names(get_field(entry, "author"))
    author_csl = [bibtex_name_to_csl(a) for a in authors]
    author_csl = [a for a in author_csl if a]
    if author_csl:
        item["author"] = author_csl

    editors = split_bibtex_names(get_field(entry, "editor"))
    editor_csl = [bibtex_name_to_csl(e) for e in editors]
    editor_csl = [e for e in editor_csl if e]
    if editor_csl:
        item["editor"] = editor_csl

    return item


def _find_top_level_comma(text: str) -> int:
    brace = 0
    in_quote = False
    esc = False
    for i, ch in enumerate(text):
        if in_quote:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_quote = False
            continue
        if ch == '"':
            in_quote = True
        elif ch == "{":
            brace += 1
        elif ch == "}":
            brace = max(0, brace - 1)
        elif ch == "," and brace == 0:
            return i
    return -1


def _parse_braced_or_quoted_value(text: str, pos: int) -> Tuple[str, int]:
    while pos < len(text) and text[pos].isspace():
        pos += 1
    if pos >= len(text):
        return "", pos
    if text[pos] == "{":
        depth = 1
        i = pos + 1
        out = []
        while i < len(text) and depth > 0:
            ch = text[i]
            if ch == "{":
                depth += 1
                out.append(ch)
            elif ch == "}":
                depth -= 1
                if depth > 0:
                    out.append(ch)
            else:
                out.append(ch)
            i += 1
        return "".join(out).strip(), i
    if text[pos] == '"':
        i = pos + 1
        out = []
        esc = False
        while i < len(text):
            ch = text[i]
            if esc:
                out.append(ch)
                esc = False
            elif ch == "\\":
                out.append(ch)
                esc = True
            elif ch == '"':
                i += 1
                break
            else:
                out.append(ch)
            i += 1
        return "".join(out).strip(), i
    # Bare value: read to next comma.
    j = text.find(",", pos)
    if j == -1:
        return text[pos:].strip(), len(text)
    return text[pos:j].strip(), j


def simple_parse_bibtex(text: str) -> List[dict]:
    """Small fallback BibTeX parser for standard Zotero exports.

    It supports entries like @article{key, title={...}, doi={...}} and nested
    braces in values. For heavily macro-based BibTeX, install bibtexparser.
    """
    entries: List[dict] = []
    i = 0
    while True:
        m = re.search(r"@\s*([A-Za-z]+)\s*([\{\(])", text[i:])
        if not m:
            break
        entry_type = m.group(1)
        open_ch = m.group(2)
        close_ch = "}" if open_ch == "{" else ")"
        start = i + m.end()
        depth = 1
        j = start
        in_quote = False
        esc = False
        while j < len(text) and depth > 0:
            ch = text[j]
            if in_quote:
                if esc:
                    esc = False
                elif ch == "\\":
                    esc = True
                elif ch == '"':
                    in_quote = False
            else:
                if ch == '"':
                    in_quote = True
                elif ch == open_ch:
                    depth += 1
                elif ch == close_ch:
                    depth -= 1
            j += 1
        body = text[start : j - 1]
        i = j
        comma = _find_top_level_comma(body)
        if comma == -1:
            continue
        key = body[:comma].strip()
        field_text = body[comma + 1 :]
        entry = {"ENTRYTYPE": entry_type.lower(), "ID": key}
        pos = 0
        while pos < len(field_text):
            while pos < len(field_text) and (field_text[pos].isspace() or field_text[pos] == ","):
                pos += 1
            name_match = re.match(r"([A-Za-z][A-Za-z0-9_\-]*)\s*=", field_text[pos:])
            if not name_match:
                break
            name = name_match.group(1).lower()
            pos += name_match.end()
            value, pos = _parse_braced_or_quoted_value(field_text, pos)
            entry[name] = value
            # move to comma after value
            while pos < len(field_text) and field_text[pos] not in ",":
                if not field_text[pos].isspace():
                    break
                pos += 1
        entries.append(entry)
    return entries


def load_bibtex_entries(bib_path: Path) -> List[dict]:
    """Load BibTeX entries with a forgiving fallback.

    Some Zotero/Better BibTeX exports, or manually edited BibTeX files, contain
    bare month/string macros such as ``month = sept``. The upstream
    ``bibtexparser`` package may raise ``UndefinedString: 'sept'`` for these
    entries before we have a chance to use the DOI/title fields. In that case we
    fall back to the built-in lightweight parser, which treats bare values as
    plain text and is sufficient for DOI/title matching.
    """
    text = bib_path.read_text(encoding="utf-8-sig", errors="replace")

    if bibtexparser is not None:
        try:
            # common_strings=True handles standard BibTeX month macros such as
            # jan, feb, ..., sep. Non-standard forms such as ``sept`` are handled
            # by the fallback below.
            try:
                from bibtexparser.bparser import BibTexParser  # type: ignore

                parser = BibTexParser(common_strings=True)
                db = parser.parse(text)
            except TypeError:
                # Older bibtexparser versions may not accept common_strings.
                from io import StringIO

                db = bibtexparser.load(StringIO(text))
            return list(db.entries)
        except Exception as exc:
            print(
                "Warning: bibtexparser could not parse the BibTeX file; "
                "falling back to the built-in tolerant parser. Reason: "
                f"{exc}",
                file=sys.stderr,
            )

    return simple_parse_bibtex(text)


def load_bibtex(bib_path: Path) -> Tuple[Dict[str, BibRecord], Dict[str, BibRecord], List[BibRecord]]:
    entries = load_bibtex_entries(bib_path)

    by_doi: Dict[str, BibRecord] = {}
    by_title: Dict[str, BibRecord] = {}
    records: List[BibRecord] = []

    for entry in entries:
        key = str(entry.get("ID") or entry.get("id") or uuid.uuid4())
        entry_type = str(entry.get("ENTRYTYPE") or entry.get("entrytype") or "misc")
        doi = normalize_doi(get_field(entry, "doi"))
        if not doi:
            # Some exports put DOI in URL or note.
            doi = normalize_doi(get_field(entry, "url", "note", "annote"))
        title = get_field(entry, "title")
        csl = bibtex_entry_to_csl(entry)
        uris = extract_zotero_uris(entry)
        rec = BibRecord(key=key, entry_type=entry_type, raw=entry, doi=doi, title=title, csl_item=csl, uris=uris)
        records.append(rec)
        if doi and doi not in by_doi:
            by_doi[doi] = rec
        ntitle = normalize_title(title)
        if ntitle and ntitle not in by_title:
            by_title[ntitle] = rec
    return by_doi, by_title, records

# ---------------------------------------------------------------------------
# Zotero RDF -> item URI helpers
# ---------------------------------------------------------------------------


def _local_name(tag: str) -> str:
    if not tag:
        return ""
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    if ":" in tag:
        return tag.rsplit(":", 1)[1]
    return tag


def _iter_text_values(elem: etree._Element) -> List[str]:
    values: List[str] = []
    for child in elem.iter():
        if child.text and child.text.strip():
            values.append(child.text.strip())
        if child.tail and child.tail.strip():
            values.append(child.tail.strip())
        for val in child.attrib.values():
            if val and str(val).strip():
                values.append(str(val).strip())
    return values


def extract_item_key_from_uri(uri: str) -> str:
    if not uri:
        return ""
    m = re.search(r"/items/([A-Za-z0-9]+)", uri)
    if m:
        return m.group(1)
    m = re.search(r"items/([A-Za-z0-9]+)", uri)
    if m:
        return m.group(1)
    return ""


def make_zotero_uri_from_key(item_key: str, library_prefix: str = "", user_id: str = "", group_id: str = "") -> str:
    item_key = (item_key or "").strip()
    if not item_key:
        return ""
    if library_prefix:
        prefix = library_prefix.strip().rstrip("/")
        return f"{prefix}/items/{item_key}"
    if group_id:
        return f"http://zotero.org/groups/{group_id.strip()}/items/{item_key}"
    if user_id:
        return f"http://zotero.org/users/{user_id.strip()}/items/{item_key}"
    return ""


def extract_rdf_item_key(elem: etree._Element) -> str:
    # Common sources: rdf:about URI, z:itemKey child, RDF resource URI, or free text.
    candidates: List[str] = []
    about = elem.get(RDF_ABOUT, "")
    if about:
        candidates.append(about)
    for child in elem.iter():
        local = _local_name(child.tag).lower()
        if local in {"itemkey", "key"} and child.text:
            candidates.append(child.text.strip())
        for attr_val in child.attrib.values():
            if attr_val:
                candidates.append(str(attr_val))
    for candidate in candidates:
        key = extract_item_key_from_uri(candidate)
        if key:
            return key
        # Zotero keys are usually eight uppercase alphanumeric characters.
        cand = candidate.strip()
        if re.fullmatch(r"[A-Za-z0-9]{8,12}", cand):
            return cand
    return ""


def extract_rdf_uri(elem: etree._Element, library_prefix: str = "", user_id: str = "", group_id: str = "") -> str:
    candidates: List[str] = []
    about = elem.get(RDF_ABOUT, "")
    if about:
        candidates.append(about)
    for child in elem.iter():
        for attr_name in (RDF_RESOURCE, RDF_ABOUT):
            val = child.get(attr_name, "")
            if val:
                candidates.append(val)
        if child.text and ("zotero.org/" in child.text or "zotero://select/" in child.text):
            candidates.append(child.text.strip())
    for candidate in candidates:
        candidate = candidate.strip().strip("{}")
        if candidate.startswith("http") and "zotero.org/" in candidate and "/items/" in candidate:
            return candidate
        if candidate.startswith("zotero://select/") and "/items/" in candidate:
            # Less ideal than http://zotero.org/... but still keeps a selectable item pointer.
            return candidate
    key = extract_rdf_item_key(elem)
    return make_zotero_uri_from_key(key, library_prefix=library_prefix, user_id=user_id, group_id=group_id)


def extract_rdf_title(elem: etree._Element) -> str:
    # Prefer dc:title / title-like fields, but avoid publicationTitle as the item title.
    fallback = ""
    for child in elem.iter():
        local = _local_name(child.tag).lower()
        text = norm_space(" ".join((child.text or "").split()))
        if not text:
            continue
        if local == "title":
            return text
        if local in {"shorttitle", "booktitle"} and not fallback:
            fallback = text
    return fallback


def extract_rdf_doi(elem: etree._Element) -> str:
    # DOI may occur in bibo:doi, dc:identifier, prism:doi, rdf:value, URL, or notes.
    for child in elem.iter():
        local = _local_name(child.tag).lower()
        text_blob = " ".join([child.text or ""] + [str(v) for v in child.attrib.values()])
        if local in {"doi", "identifier", "value", "url", "relation"}:
            doi = normalize_doi(text_blob)
            if doi and doi.startswith("10."):
                return doi
    all_text = " ".join(_iter_text_values(elem))
    doi = normalize_doi(all_text)
    return doi if doi.startswith("10.") else ""


def rdf_element_to_csl(elem: etree._Element, uri: str, key: str, doi: str, title: str) -> dict:
    item = {"id": uri or key or str(uuid.uuid4()), "type": "article-journal"}
    if title:
        item["title"] = title
    if doi:
        item["DOI"] = doi
    for child in elem.iter():
        local = _local_name(child.tag).lower()
        text = norm_space(child.text or "")
        if not text:
            continue
        if local in {"publicationtitle", "journaltitle"} and "container-title" not in item:
            item["container-title"] = text
        elif local == "volume" and "volume" not in item:
            item["volume"] = text
        elif local in {"issue", "number"} and "issue" not in item:
            item["issue"] = text
        elif local in {"pages", "page"} and "page" not in item:
            item["page"] = text
        elif local in {"date", "issued"} and "issued" not in item:
            m = re.search(r"(\d{4})", text)
            if m:
                item["issued"] = {"date-parts": [[int(m.group(1))]]}
    return item


def load_zotero_rdf(
    rdf_path: Path,
    library_prefix: str = "",
    user_id: str = "",
    group_id: str = "",
) -> Tuple[Dict[str, RdfRecord], Dict[str, RdfRecord], List[RdfRecord]]:
    """Load Zotero RDF and recover Zotero item URIs.

    Zotero RDF exports typically store the library-linked item URI in rdf:about,
    for example http://zotero.org/users/123456/items/ABCDEFGH or
    http://zotero.org/groups/123456/items/ABCDEFGH. The DOI is often stored in
    dc:identifier/rdf:value or bibo:doi. This loader is deliberately permissive
    because RDF namespace details vary across Zotero versions and export options.
    """
    parser = etree.XMLParser(remove_blank_text=False, recover=True, huge_tree=True)
    root = etree.parse(str(rdf_path), parser).getroot()
    by_doi: Dict[str, RdfRecord] = {}
    by_title: Dict[str, RdfRecord] = {}
    records: List[RdfRecord] = []
    seen_uris = set()

    for elem in root.iter():
        # Ignore tiny leaf nodes; bibliography items normally contain multiple child nodes.
        if len(elem) == 0:
            continue
        uri = extract_rdf_uri(elem, library_prefix=library_prefix, user_id=user_id, group_id=group_id)
        doi = extract_rdf_doi(elem)
        title = extract_rdf_title(elem)
        key = extract_rdf_item_key(elem) or extract_item_key_from_uri(uri)
        # Keep only likely Zotero bibliographic items. Attachments/notes often have no DOI/title.
        if not uri and not key:
            continue
        if not doi and not title:
            continue
        uri_key = uri or key
        if uri_key in seen_uris:
            continue
        seen_uris.add(uri_key)
        csl = rdf_element_to_csl(elem, uri=uri, key=key, doi=doi, title=title)
        rec = RdfRecord(uri=uri, item_key=key, doi=doi, title=title, csl_item=csl)
        records.append(rec)
        if doi and doi not in by_doi:
            by_doi[doi] = rec
        ntitle = normalize_title(title)
        if ntitle and ntitle not in by_title:
            by_title[ntitle] = rec
    return by_doi, by_title, records


def attach_rdf_uris_to_bib_records(
    bib_by_doi: Dict[str, BibRecord],
    bib_by_title: Dict[str, BibRecord],
    all_records: List[BibRecord],
    rdf_by_doi: Dict[str, RdfRecord],
    rdf_by_title: Dict[str, RdfRecord],
    min_title_similarity: float = 0.90,
) -> Dict[str, int]:
    """Attach Zotero RDF URIs to BibRecord objects by DOI/title matching."""
    stats = {
        "rdf_records": len(set([r.uri or r.item_key for r in list(rdf_by_doi.values()) + list(rdf_by_title.values())])),
        "rdf_records_with_doi": len(rdf_by_doi),
        "attached_by_doi": 0,
        "attached_by_title": 0,
        "already_had_uri": 0,
        "unresolved_bib_records": 0,
    }
    for rec in all_records:
        if rec.uris:
            stats["already_had_uri"] += 1
            continue
        rdf_rec: Optional[RdfRecord] = None
        if rec.doi:
            rdf_rec = rdf_by_doi.get(rec.doi)
            if rdf_rec and rdf_rec.uri:
                rec.uris = [rdf_rec.uri]
                stats["attached_by_doi"] += 1
                continue
        title_norm = normalize_title(rec.title)
        if title_norm:
            rdf_rec = rdf_by_title.get(title_norm)
            if rdf_rec and rdf_rec.uri:
                rec.uris = [rdf_rec.uri]
                stats["attached_by_title"] += 1
                continue
            best_score = 0.0
            best_rdf: Optional[RdfRecord] = None
            for rdf_title_norm, candidate in rdf_by_title.items():
                if not candidate.uri:
                    continue
                score = SequenceMatcher(None, title_norm, rdf_title_norm).ratio()
                if score > best_score:
                    best_score = score
                    best_rdf = candidate
            if best_rdf and best_score >= min_title_similarity:
                rec.uris = [best_rdf.uri]
                stats["attached_by_title"] += 1
                continue
        stats["unresolved_bib_records"] += 1
    return stats


# ---------------------------------------------------------------------------
# Zotero local database -> library-linked URI helpers
# ---------------------------------------------------------------------------


def default_zotero_sqlite_path() -> Path:
    """Return Zotero's common default data-directory SQLite path."""
    return Path.home() / "Zotero" / "zotero.sqlite"


def sqlite_tables(conn: sqlite3.Connection) -> set:
    return {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}


def sqlite_columns(conn: sqlite3.Connection, table: str) -> List[str]:
    try:
        return [row[1] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()]
    except Exception:
        return []


def copy_sqlite_for_reading(sqlite_path: Path) -> Path:
    """Copy Zotero DB to a temp file so it can be read while Zotero is open."""
    tmpdir = Path(tempfile.mkdtemp(prefix="zotero_sqlite_read_"))
    dst = tmpdir / "zotero.sqlite"
    shutil.copy2(sqlite_path, dst)
    for suffix in ("-wal", "-shm"):
        src = Path(str(sqlite_path) + suffix)
        if src.exists():
            shutil.copy2(src, Path(str(dst) + suffix))
    return dst


def find_zotero_local_user_key(conn: sqlite3.Connection) -> str:
    """Try to recover Zotero's local user key from settings-like tables.

    Zotero Word fields can use URIs like
    http://zotero.org/users/local/<localUserKey>/items/<itemKey>.
    Zotero database layouts vary by version, so this scans settings-like rows.
    """
    tables = sqlite_tables(conn)
    blobs: List[str] = []
    for table in ("settings", "syncedSettings"):
        if table not in tables:
            continue
        try:
            rows = conn.execute(f"SELECT * FROM {table} LIMIT 5000").fetchall()
        except Exception:
            continue
        for row in rows:
            blobs.append(" ".join("" if v is None else str(v) for v in row))
    joined = "\n".join(blobs)
    for pat in (
        r'localUserKey["\']?\s*[:=]\s*["\']?([A-Za-z0-9]{6,20})',
        r'localUserID["\']?\s*[:=]\s*["\']?([A-Za-z0-9]{6,20})',
    ):
        m = re.search(pat, joined, flags=re.IGNORECASE)
        if m:
            return m.group(1)
    for line in blobs:
        if re.search(r"account|user|local", line, flags=re.IGNORECASE):
            m = re.search(r"\b([A-Za-z0-9]{8})\b", line)
            if m and not m.group(1).isdigit():
                return m.group(1)
    return ""


def find_zotero_user_id(conn: sqlite3.Connection) -> str:
    """Try to recover synced numeric Zotero userID from settings-like tables."""
    tables = sqlite_tables(conn)
    blobs: List[str] = []
    for table in ("settings", "syncedSettings"):
        if table not in tables:
            continue
        try:
            rows = conn.execute(f"SELECT * FROM {table} LIMIT 5000").fetchall()
        except Exception:
            continue
        for row in rows:
            blobs.append(" ".join("" if v is None else str(v) for v in row))
    joined = "\n".join(blobs)
    for pat in (
        r'userID["\']?\s*[:=]\s*["\']?(\d{2,12})',
        r'userid["\']?\s*[:=]\s*["\']?(\d{2,12})',
    ):
        m = re.search(pat, joined, flags=re.IGNORECASE)
        if m:
            return m.group(1)
    return ""


def load_zotero_group_library_map(conn: sqlite3.Connection) -> Dict[int, str]:
    """Map Zotero libraryID to web groupID where possible."""
    if "groups" not in sqlite_tables(conn):
        return {}
    cols = sqlite_columns(conn, "groups")
    if "libraryID" not in cols or "groupID" not in cols:
        return {}
    out: Dict[int, str] = {}
    try:
        for library_id, group_id in conn.execute("SELECT libraryID, groupID FROM groups"):
            if library_id is not None and group_id is not None:
                out[int(library_id)] = str(group_id)
    except Exception:
        pass
    return out


def make_zotero_uri_for_sqlite_item(
    item_key: str,
    library_id: int,
    group_map: Dict[int, str],
    user_id: str = "",
    local_user_key: str = "",
    library_prefix: str = "",
) -> str:
    item_key = (item_key or "").strip()
    if not item_key:
        return ""
    if library_prefix:
        return f"{library_prefix.strip().rstrip('/')}/items/{item_key}"
    if library_id in group_map:
        return f"http://zotero.org/groups/{group_map[library_id]}/items/{item_key}"
    if user_id:
        return f"http://zotero.org/users/{user_id}/items/{item_key}"
    if local_user_key:
        return f"http://zotero.org/users/local/{local_user_key}/items/{item_key}"
    return ""


def load_zotero_sqlite_records(
    sqlite_path: Path,
    user_id: str = "",
    local_user_key: str = "",
    library_prefix: str = "",
) -> Tuple[Dict[str, dict], Dict[str, dict], Dict[str, int]]:
    """Load Zotero itemID/key/libraryID/URI records from zotero.sqlite."""
    sqlite_copy = copy_sqlite_for_reading(sqlite_path)
    conn = sqlite3.connect(str(sqlite_copy))
    conn.row_factory = sqlite3.Row
    tables = sqlite_tables(conn)
    required = {"items", "itemData", "itemDataValues", "fields"}
    missing = required - tables
    if missing:
        raise RuntimeError(f"zotero.sqlite missing required tables: {sorted(missing)}")

    inferred_user_id = user_id or find_zotero_user_id(conn)
    inferred_local_key = local_user_key or find_zotero_local_user_key(conn)
    group_map = load_zotero_group_library_map(conn)

    deleted_clause = ""
    if "deletedItems" in tables:
        deleted_clause = " AND i.itemID NOT IN (SELECT itemID FROM deletedItems)"

    query = f"""
        SELECT i.itemID, i.key, i.libraryID, f.fieldName, v.value
        FROM items i
        JOIN itemData d ON i.itemID = d.itemID
        JOIN fields f ON d.fieldID = f.fieldID
        JOIN itemDataValues v ON d.valueID = v.valueID
        WHERE f.fieldName IN ('DOI', 'title') {deleted_clause}
    """
    per_item: Dict[int, dict] = {}
    for row in conn.execute(query):
        item_id = int(row["itemID"])
        rec = per_item.setdefault(item_id, {
            "itemID": item_id,
            "key": str(row["key"]),
            "libraryID": int(row["libraryID"] or 1),
            "doi": "",
            "title": "",
            "uri": "",
        })
        field = row["fieldName"]
        value = str(row["value"] or "")
        if field == "DOI":
            doi = normalize_doi(value)
            if doi.startswith("10."):
                rec["doi"] = doi
        elif field == "title":
            rec["title"] = norm_space(value)

    by_doi: Dict[str, dict] = {}
    by_title: Dict[str, dict] = {}
    for rec in per_item.values():
        rec["uri"] = make_zotero_uri_for_sqlite_item(
            rec["key"], rec["libraryID"], group_map,
            user_id=inferred_user_id,
            local_user_key=inferred_local_key,
            library_prefix=library_prefix,
        )
        if rec["doi"] and rec["uri"]:
            by_doi.setdefault(rec["doi"], rec)
        nt = normalize_title(rec["title"])
        if nt and rec["uri"]:
            by_title.setdefault(nt, rec)

    stats = {
        "sqlite_items_scanned": len(per_item),
        "sqlite_items_with_doi": sum(1 for r in per_item.values() if r.get("doi")),
        "sqlite_items_with_uri": sum(1 for r in per_item.values() if r.get("uri")),
        "sqlite_records_by_doi": len(by_doi),
        "sqlite_records_by_title": len(by_title),
        "group_libraries_detected": len(group_map),
        "used_numeric_user_id": 1 if inferred_user_id else 0,
        "used_local_user_key": 1 if inferred_local_key and not inferred_user_id else 0,
    }
    conn.close()
    return by_doi, by_title, stats


def attach_sqlite_uris_to_bib_records(
    all_records: List[BibRecord],
    sqlite_by_doi: Dict[str, dict],
    sqlite_by_title: Dict[str, dict],
    min_title_similarity: float = 0.90,
) -> Dict[str, int]:
    """Attach true Zotero item URIs and numeric itemIDs from local zotero.sqlite."""
    stats = {
        "attached_by_doi": 0,
        "attached_by_title": 0,
        "already_had_uri": 0,
        "unresolved_bib_records": 0,
    }
    for rec in all_records:
        if rec.uris and rec.zotero_item_id is not None:
            stats["already_had_uri"] += 1
            continue
        zrec = None
        if rec.doi:
            zrec = sqlite_by_doi.get(rec.doi)
            if zrec:
                stats["attached_by_doi"] += 1
        if zrec is None:
            nt = normalize_title(rec.title)
            if nt:
                zrec = sqlite_by_title.get(nt)
                if zrec:
                    stats["attached_by_title"] += 1
                else:
                    best_score = 0.0
                    best = None
                    for tnorm, cand in sqlite_by_title.items():
                        score = SequenceMatcher(None, nt, tnorm).ratio()
                        if score > best_score:
                            best_score = score
                            best = cand
                    if best and best_score >= min_title_similarity:
                        zrec = best
                        stats["attached_by_title"] += 1
        if zrec and zrec.get("uri"):
            rec.uris = [zrec["uri"]]
            rec.zotero_item_id = int(zrec["itemID"])
            rec.zotero_item_key = str(zrec["key"])
            rec.zotero_library_id = int(zrec["libraryID"])
            rec.csl_item["id"] = rec.zotero_item_id
        else:
            stats["unresolved_bib_records"] += 1
    return stats


# ---------------------------------------------------------------------------
# DOCX reading and Reference parsing
# ---------------------------------------------------------------------------


def paragraph_text(p: etree._Element) -> str:
    texts = p.xpath(".//w:t", namespaces=NS)
    return "".join(t.text or "" for t in texts)


def get_body_paragraphs(root: etree._Element) -> List[etree._Element]:
    return root.xpath(".//w:body/w:p", namespaces=NS)


def normalize_reference_heading_candidate(value: str) -> str:
    """Normalize a possible Reference heading from Word paragraphs.

    Word often stores headings with section numbers, full-width punctuation,
    non-breaking spaces, or trailing colons. This helper makes detection less
    brittle without changing document content.
    """
    value = norm_space(value).replace("\u00a0", " ").replace("\u3000", " ")
    value = re.sub(r"^[\s\u200b\ufeff]+|[\s\u200b\ufeff]+$", "", value)
    value = re.sub(r"^[第\s]*\d+(?:\.\d+)*[章节部篇\s]*[\.．、:：)]*\s*", "", value)
    value = re.sub(r"^[A-Z]\s*[\.．、:：)]\s*", "", value, flags=re.IGNORECASE)
    value = value.strip(" \t:：。.;；")
    return value.lower()


def is_reference_heading_text(value: str) -> bool:
    raw = norm_space(value)
    if not raw:
        return False
    if REFERENCE_HEADING_RE.match(raw):
        return True
    normalized = normalize_reference_heading_candidate(raw)
    return bool(REFERENCE_HEADING_RE.match(normalized))


def find_reference_heading_index(paragraphs: List[etree._Element], custom_heading: str = "") -> Optional[int]:
    if custom_heading:
        custom = normalize_reference_heading_candidate(custom_heading)
        for idx, p in enumerate(paragraphs):
            txt = normalize_reference_heading_candidate(paragraph_text(p))
            if txt == custom or custom in txt:
                return idx
    for idx, p in enumerate(paragraphs):
        txt = paragraph_text(p)
        if is_reference_heading_text(txt):
            return idx
    return None


def find_reference_entries_start_index(paragraphs: List[etree._Element]) -> Optional[int]:
    """Fallback: locate the first numbered reference when no heading is found.

    The function looks for a block near the end of the document that starts
    with [1], 1., 1), or 1、 and is followed by several increasing reference
    numbers. It returns the index of the first reference paragraph.
    """
    texts = [norm_space(paragraph_text(p)) for p in paragraphs]
    starts: List[Tuple[int, int]] = []
    for idx, txt in enumerate(texts):
        m = REFERENCE_START_RE.match(txt)
        if not m:
            continue
        try:
            num = int(next(g for g in m.groups()[:3] if g))
        except Exception:
            continue
        # Avoid treating very early numbered lists as references unless they start close to document end.
        if idx < max(0, int(len(paragraphs) * 0.45)) and num != 1:
            continue
        starts.append((idx, num))

    for pos, (idx, num) in enumerate(starts):
        if num != 1:
            continue
        following = [n for _, n in starts[pos:pos + 8]]
        if len(following) >= 3 and following[:3] == [1, 2, 3]:
            return idx
        if len(following) >= 2 and following[:2] == [1, 2]:
            return idx
    return None


def describe_reference_heading_candidates(paragraphs: List[etree._Element], limit: int = 25) -> str:
    nonempty = []
    for idx, p in enumerate(paragraphs):
        txt = norm_space(paragraph_text(p))
        if txt:
            nonempty.append((idx, txt))
    tail = nonempty[-limit:]
    lines = []
    for idx, txt in tail:
        preview = txt[:160] + ("..." if len(txt) > 160 else "")
        lines.append(f"  paragraph {idx}: {preview}")
    return "\n".join(lines)


def extract_references_from_paragraphs(paragraphs: List[etree._Element], start_idx: int) -> Dict[int, str]:
    refs: Dict[int, str] = {}
    current_num: Optional[int] = None
    current_text: List[str] = []

    def flush() -> None:
        nonlocal current_num, current_text
        if current_num is not None:
            refs[current_num] = norm_space(" ".join(current_text))
        current_num = None
        current_text = []

    for p in paragraphs[start_idx + 1 :]:
        txt = norm_space(paragraph_text(p))
        if not txt:
            continue
        m = REFERENCE_START_RE.match(txt)
        if m:
            flush()
            num = int(next(g for g in m.groups()[:3] if g))
            rest = m.group(4) or ""
            current_num = num
            current_text = [rest.strip()]
        else:
            if current_num is not None:
                current_text.append(txt)
    flush()

    # Fallback for references pasted as one long block.
    if not refs:
        blob = "\n".join(norm_space(paragraph_text(p)) for p in paragraphs[start_idx + 1 :])
        matches = list(re.finditer(r"(?:^|\n)\s*(?:\[(\d{1,5})\]|(\d{1,5})[\.)、])\s+", blob))
        for i, m in enumerate(matches):
            num = int(m.group(1) or m.group(2))
            a = m.end()
            b = matches[i + 1].start() if i + 1 < len(matches) else len(blob)
            refs[num] = norm_space(blob[a:b])
    return refs


def read_document_xml(docx_path: Path) -> etree._Element:
    with zipfile.ZipFile(docx_path, "r") as zf:
        with zf.open("word/document.xml") as f:
            parser = etree.XMLParser(remove_blank_text=False, recover=True)
            return etree.parse(f, parser).getroot()


# ---------------------------------------------------------------------------
# Citation parsing and field generation
# ---------------------------------------------------------------------------


def expand_citation_numbers(text_inside_brackets: str) -> List[int]:
    nums: List[int] = []
    parts = re.split(r"\s*[,，;；]\s*", text_inside_brackets.strip())
    for part in parts:
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(\d+)\s*[-–—]\s*(\d+)$", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            step = 1 if b >= a else -1
            nums.extend(range(a, b + step, step))
        else:
            if re.match(r"^\d+$", part):
                nums.append(int(part))
            else:
                raise ValueError(f"Unrecognized citation fragment: {part!r}")
    # De-duplicate while preserving order.
    out: List[int] = []
    seen = set()
    for n in nums:
        if n not in seen:
            out.append(n)
            seen.add(n)
    return out


def random_citation_id() -> str:
    alphabet = string.ascii_letters + string.digits
    return "".join(random.choice(alphabet) for _ in range(8))


def make_citation_json(matches: List[RefMatch], display_text: str) -> str:
    citation_items = []
    for match in matches:
        if not match.record:
            raise ValueError(f"Reference number {match.number} is not matched")
        # Zotero-generated Word fields normally use a numeric local itemID as
        # `id`, plus both `uris` and `uri` arrays pointing to the library item.
        # If no SQLite/RDF link is available, fall back to the BibTeX key, but
        # Zotero may then treat the citation as embedded/orphaned.
        if match.record.zotero_item_id is not None:
            item_id = match.record.zotero_item_id
        else:
            item_id = match.record.key
        item_data = deepcopy(match.record.csl_item)
        item_data["id"] = item_id
        item = {
            "id": item_id,
            "itemData": item_data,
        }
        if match.record.uris:
            item["uris"] = match.record.uris
            item["uri"] = match.record.uris
        citation_items.append(item)

    plain = display_text.strip()
    if plain.startswith("[") and plain.endswith("]"):
        plain = plain[1:-1]

    csl_citation = {
        "citationID": random_citation_id(),
        "properties": {
            "formattedCitation": display_text,
            "plainCitation": plain,
            "noteIndex": 0,
        },
        "citationItems": citation_items,
        "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json",
    }
    return json.dumps(csl_citation, ensure_ascii=False, separators=(",", ":"))


def make_run_with_text(text: str, rpr_template: Optional[etree._Element] = None) -> etree._Element:
    r = etree.Element(qn("w:r"))
    if rpr_template is not None:
        r.append(deepcopy(rpr_template))
    t = etree.SubElement(r, qn("w:t"))
    if text.startswith(" ") or text.endswith(" "):
        t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text
    return r


def make_fldchar_run(kind: str) -> etree._Element:
    r = etree.Element(qn("w:r"))
    fld = etree.SubElement(r, qn("w:fldChar"))
    fld.set(qn("w:fldCharType"), kind)
    if kind == "begin":
        fld.set(qn("w:dirty"), "true")
    return r


def make_instr_run(text: str) -> etree._Element:
    r = etree.Element(qn("w:r"))
    instr = etree.SubElement(r, qn("w:instrText"))
    instr.set(f"{{{XML_NS}}}space", "preserve")
    instr.text = text
    return r


def make_zotero_field_runs(citation_json: str, display_text: str, rpr_template: Optional[etree._Element]) -> List[etree._Element]:
    instr = " ADDIN ZOTERO_ITEM CSL_CITATION " + citation_json
    runs: List[etree._Element] = [make_fldchar_run("begin")]
    # Split long instruction text; Word tolerates multiple instrText runs in one field.
    chunk_size = 1500
    for i in range(0, len(instr), chunk_size):
        runs.append(make_instr_run(instr[i : i + chunk_size]))
    runs.append(make_fldchar_run("separate"))
    runs.append(make_run_with_text(display_text, rpr_template))
    runs.append(make_fldchar_run("end"))
    return runs


def paragraph_has_complex_objects(p: etree._Element) -> bool:
    complex_tags = [
        ".//w:drawing",
        ".//w:pict",
        ".//w:object",
        ".//m:oMath",
        ".//m:oMathPara",
        ".//w:fldChar",
        ".//w:instrText",
    ]
    for xp in complex_tags:
        if p.xpath(xp, namespaces=NS):
            return True
    return False


def first_run_rpr(p: etree._Element) -> Optional[etree._Element]:
    rprs = p.xpath(".//w:r[w:t][1]/w:rPr", namespaces=NS)
    return deepcopy(rprs[0]) if rprs else None


def rewrite_paragraph_plain_text_to_fields(
    p: etree._Element,
    matches_by_number: Dict[int, RefMatch],
    part_name: str,
    paragraph_index: int,
    logs: List[ConversionLog],
    force_complex: bool = False,
) -> bool:
    txt = paragraph_text(p)
    if not txt or not CITATION_RE.search(txt):
        return False

    if paragraph_has_complex_objects(p) and not force_complex:
        logs.append(
            ConversionLog(
                part=part_name,
                paragraph_index=paragraph_index,
                citation_text="",
                numbers="",
                action="skipped",
                message="Paragraph contains existing fields/drawings/math objects; use --force-complex to rewrite it.",
            )
        )
        return False

    rpr_template = first_run_rpr(p)
    new_children: List[etree._Element] = []
    ppr = p.find(qn("w:pPr"))
    if ppr is not None:
        new_children.append(deepcopy(ppr))

    pos = 0
    changed = False
    for m in CITATION_RE.finditer(txt):
        raw = m.group(0)
        inside = m.group(1)
        try:
            nums = expand_citation_numbers(inside)
        except Exception as exc:
            nums = []
            logs.append(
                ConversionLog(part_name, paragraph_index, raw, inside, "left_plain", f"Cannot parse citation: {exc}")
            )

        missing = [n for n in nums if n not in matches_by_number or matches_by_number[n].status != "matched"]
        if not nums or missing:
            if pos < m.end():
                # Preserve original text up to and including the unsupported citation.
                pass
            logs.append(
                ConversionLog(
                    part=part_name,
                    paragraph_index=paragraph_index,
                    citation_text=raw,
                    numbers=",".join(map(str, nums)) if nums else inside,
                    action="left_plain",
                    message=f"No matched BibTeX/DOI record for reference number(s): {missing}",
                )
            )
            continue

        # Add text before citation.
        if m.start() > pos:
            new_children.append(make_run_with_text(txt[pos : m.start()], rpr_template))
        ref_matches = [matches_by_number[n] for n in nums]
        citation_json = make_citation_json(ref_matches, raw)
        new_children.extend(make_zotero_field_runs(citation_json, raw, rpr_template))
        logs.append(
            ConversionLog(
                part=part_name,
                paragraph_index=paragraph_index,
                citation_text=raw,
                numbers=",".join(map(str, nums)),
                action="converted",
                message="Converted to ADDIN ZOTERO_ITEM CSL_CITATION field.",
            )
        )
        changed = True
        pos = m.end()

    if not changed:
        return False

    if pos < len(txt):
        new_children.append(make_run_with_text(txt[pos:], rpr_template))

    # Replace paragraph children. This preserves paragraph-level formatting, but not
    # mixed run-level formatting inside the paragraph.
    for child in list(p):
        p.remove(child)
    for child in new_children:
        p.append(child)
    return True


# ---------------------------------------------------------------------------
# Reference matching
# ---------------------------------------------------------------------------


def match_references(
    ref_texts: Dict[int, str],
    bib_by_doi: Dict[str, BibRecord],
    bib_by_title: Dict[str, BibRecord],
    all_records: List[BibRecord],
    allow_title_fallback: bool = True,
    min_title_similarity: float = 0.88,
) -> Dict[int, RefMatch]:
    out: Dict[int, RefMatch] = {}
    for num, ref in sorted(ref_texts.items()):
        doi = normalize_doi(ref)
        match = RefMatch(number=num, reference_text=ref, doi=doi)
        if doi and doi in bib_by_doi:
            rec = bib_by_doi[doi]
            match.record = rec
            match.bib_key = rec.key
            match.status = "matched"
            match.message = "Matched by DOI"
        elif doi:
            match.status = "doi_not_in_bibtex"
            match.message = "DOI found in Reference section but not found in BibTeX"
        else:
            match.status = "doi_missing"
            match.message = "No DOI found in this reference"

        if match.status != "matched" and allow_title_fallback:
            # Conservative title fallback: compare the whole reference text with BibTeX titles.
            ref_norm = normalize_title(ref)
            best_rec = None
            best_score = 0.0
            for rec in all_records:
                title_norm = normalize_title(rec.title)
                if not title_norm:
                    continue
                if title_norm in ref_norm:
                    score = 1.0
                else:
                    score = SequenceMatcher(None, title_norm, ref_norm).ratio()
                if score > best_score:
                    best_score = score
                    best_rec = rec
            match.title_similarity = round(best_score, 4)
            if best_rec and best_score >= min_title_similarity:
                match.record = best_rec
                match.bib_key = best_rec.key
                match.status = "matched"
                match.message = f"Matched by title fallback, similarity={best_score:.3f}"

        out[num] = match
    return out


# ---------------------------------------------------------------------------
# DOCX conversion
# ---------------------------------------------------------------------------


def docx_xml_parts(zf: zipfile.ZipFile) -> List[str]:
    candidates = []
    for name in zf.namelist():
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        base = os.path.basename(name)
        if base in {"document.xml", "footnotes.xml", "endnotes.xml", "comments.xml"}:
            candidates.append(name)
        elif re.match(r"header\d+\.xml$", base) or re.match(r"footer\d+\.xml$", base):
            candidates.append(name)
    return candidates


def remove_reference_section(root: etree._Element, heading_idx: Optional[int]) -> bool:
    if heading_idx is None:
        return False
    body = root.find(".//w:body", namespaces=NS)
    if body is None:
        return False
    children = list(body)
    paras = [c for c in children if c.tag == qn("w:p")]
    if heading_idx >= len(paras):
        return False
    heading_p = paras[heading_idx]
    remove_started = False
    removed = False
    for child in list(body):
        if child is heading_p:
            remove_started = True
        if remove_started and child.tag != qn("w:sectPr"):
            body.remove(child)
            removed = True
    return removed


def convert_docx(
    docx_path: Path,
    out_path: Path,
    matches_by_number: Dict[int, RefMatch],
    dry_run: bool,
    remove_refs: bool,
    custom_reference_heading: str,
    force_complex: bool,
) -> Tuple[List[ConversionLog], Dict[str, int]]:
    logs: List[ConversionLog] = []
    stats = {"converted": 0, "left_plain": 0, "skipped": 0, "parts_changed": 0}

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        with zipfile.ZipFile(docx_path, "r") as zf:
            zf.extractall(tmpdir_path)
            parts = docx_xml_parts(zf)

        parser = etree.XMLParser(remove_blank_text=False, recover=True)
        changed_parts: Dict[str, bytes] = {}

        for part in parts:
            part_file = tmpdir_path / part
            if not part_file.exists():
                continue
            try:
                root = etree.parse(str(part_file), parser).getroot()
            except Exception as exc:
                logs.append(ConversionLog(part, -1, "", "", "skipped", f"Cannot parse XML part: {exc}"))
                continue

            part_changed = False
            heading_idx: Optional[int] = None
            skip_after_ref = False
            body_paras: List[etree._Element] = []
            if part == "word/document.xml":
                body_paras = get_body_paragraphs(root)
                heading_idx = find_reference_heading_index(body_paras, custom_reference_heading)
                if heading_idx is None:
                    first_ref_idx = find_reference_entries_start_index(body_paras)
                    if first_ref_idx is not None:
                        heading_idx = max(-1, first_ref_idx - 1)
                skip_after_ref = heading_idx is not None

            all_paras = root.xpath(".//w:p", namespaces=NS)
            body_para_id_to_idx = {id(p): i for i, p in enumerate(body_paras)}

            for idx, p in enumerate(all_paras):
                if part == "word/document.xml" and skip_after_ref and id(p) in body_para_id_to_idx:
                    if body_para_id_to_idx[id(p)] >= heading_idx:  # type: ignore[arg-type]
                        continue
                before_log_count = len(logs)
                changed = rewrite_paragraph_plain_text_to_fields(
                    p,
                    matches_by_number,
                    part,
                    idx,
                    logs,
                    force_complex=force_complex,
                )
                if changed:
                    part_changed = True
                    stats["converted"] += sum(1 for lg in logs[before_log_count:] if lg.action == "converted")
                else:
                    for lg in logs[before_log_count:]:
                        if lg.action == "left_plain":
                            stats["left_plain"] += 1
                        elif lg.action == "skipped":
                            stats["skipped"] += 1

            if part == "word/document.xml" and remove_refs and heading_idx is not None:
                removed = remove_reference_section(root, heading_idx)
                part_changed = part_changed or removed
                if removed:
                    logs.append(ConversionLog(part, heading_idx, "", "", "removed_references", "Removed old plain-text Reference section."))

            if part_changed:
                stats["parts_changed"] += 1
                changed_parts[part] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

        if dry_run:
            return logs, stats

        # Write new .docx.
        if out_path.exists():
            out_path.unlink()
        with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = changed_parts.get(item.filename)
                if data is None:
                    data = zin.read(item.filename)
                zout.writestr(item, data)

    return logs, stats


# ---------------------------------------------------------------------------
# Reports
# ---------------------------------------------------------------------------


def write_match_report(path: Path, matches: Dict[int, RefMatch]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["number", "status", "doi", "bib_key", "zotero_uri_count", "first_zotero_uri", "zotero_item_id", "zotero_item_key", "zotero_library_id", "title_similarity", "message", "reference_text"])
        for num in sorted(matches):
            m = matches[num]
            uri_count = len(m.record.uris) if m.record and m.record.uris else 0
            first_uri = m.record.uris[0] if m.record and m.record.uris else ""
            zid = m.record.zotero_item_id if m.record and m.record.zotero_item_id is not None else ""
            zkey = m.record.zotero_item_key if m.record else ""
            zlib = m.record.zotero_library_id if m.record and m.record.zotero_library_id is not None else ""
            writer.writerow([m.number, m.status, m.doi, m.bib_key, uri_count, first_uri, zid, zkey, zlib, m.title_similarity, m.message, m.reference_text])


def write_conversion_log(path: Path, logs: List[ConversionLog]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["part", "paragraph_index", "citation_text", "numbers", "action", "message"])
        for lg in logs:
            writer.writerow([lg.part, lg.paragraph_index, lg.citation_text, lg.numbers, lg.action, lg.message])


def write_json_summary(path: Path, summary: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")


# ---------------------------------------------------------------------------
# Main CLI
# ---------------------------------------------------------------------------


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert plain numeric Word citations [1] to Zotero-style ADDIN fields using Reference DOI + BibTeX matching."
    )
    parser.add_argument("--docx", required=True, help="Input .docx file with plain numeric citations.")
    parser.add_argument("--bib", required=True, help="BibTeX file exported from Zotero. Keep using this for CSL metadata and DOI matching.")
    parser.add_argument("--rdf", default="", help="Optional Zotero RDF export. Useful for DOI metadata, but many RDF exports do NOT contain library item URIs.")
    parser.add_argument("--zotero-sqlite", default="", help="Path to Zotero local database zotero.sqlite. Best option for true library-linked Word fields. Default if omitted: ~/Zotero/zotero.sqlite when it exists.")
    parser.add_argument("--zotero-library-uri-prefix", default="", help="Optional forced prefix, e.g. http://zotero.org/users/123456, http://zotero.org/users/local/LOCALKEY, or http://zotero.org/groups/123456.")
    parser.add_argument("--zotero-user-id", default="", help="Optional numeric Zotero user ID for constructing http://zotero.org/users/<id>/items/<key>.")
    parser.add_argument("--zotero-local-user-key", default="", help="Optional local Zotero user key for constructing http://zotero.org/users/local/<key>/items/<itemKey> if no numeric user ID is available.")
    parser.add_argument("--zotero-group-id", default="", help="Optional group ID fallback for constructing http://zotero.org/groups/<id>/items/<key> if RDF has item keys but no full URIs.")
    parser.add_argument("--require-zotero-uris", action="store_true", help="Abort if not every matched reference has a Zotero URI/itemID. Recommended before final conversion.")
    parser.add_argument("--out", required=True, help="Output .docx path. Ignored in --dry-run except for naming reports.")
    parser.add_argument("--report-dir", default="", help="Directory for CSV/JSON reports. Default: <out>_reports")
    parser.add_argument("--dry-run", action="store_true", help="Only parse and report; do not write converted .docx.")
    parser.add_argument("--remove-reference-section", action="store_true", help="Remove old plain-text Reference section after converting citations.")
    parser.add_argument("--reference-heading", default="", help="Custom exact heading text for the Reference section, if automatic detection fails.")
    parser.add_argument("--no-title-fallback", action="store_true", help="Disable conservative title-based fallback matching when DOI is missing.")
    parser.add_argument("--min-title-similarity", type=float, default=0.88, help="Minimum title fallback similarity, default 0.88.")
    parser.add_argument("--force-complex", action="store_true", help="Rewrite paragraphs even if they contain existing fields/drawings/math. Use carefully.")
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)

    docx_path = Path(args.docx).expanduser().resolve()
    bib_path = Path(args.bib).expanduser().resolve()
    rdf_path = Path(args.rdf).expanduser().resolve() if args.rdf else None
    zotero_sqlite_path = Path(args.zotero_sqlite).expanduser().resolve() if args.zotero_sqlite else None
    if zotero_sqlite_path is None:
        default_sqlite = default_zotero_sqlite_path()
        if default_sqlite.exists():
            zotero_sqlite_path = default_sqlite.resolve()
    out_path = Path(args.out).expanduser().resolve()

    if not docx_path.exists():
        print(f"ERROR: Input DOCX not found: {docx_path}", file=sys.stderr)
        return 2
    if not bib_path.exists():
        print(f"ERROR: BibTeX file not found: {bib_path}", file=sys.stderr)
        return 2
    if rdf_path is not None and not rdf_path.exists():
        print(f"ERROR: Zotero RDF file not found: {rdf_path}", file=sys.stderr)
        return 2
    if zotero_sqlite_path is not None and not zotero_sqlite_path.exists():
        print(f"ERROR: Zotero SQLite database not found: {zotero_sqlite_path}", file=sys.stderr)
        return 2
    if docx_path.suffix.lower() != ".docx":
        print("ERROR: Input must be a .docx file. Save .doc/.odt as .docx first.", file=sys.stderr)
        return 2

    report_dir = Path(args.report_dir).expanduser().resolve() if args.report_dir else out_path.with_suffix("").parent / f"{out_path.stem}_reports"
    report_dir.mkdir(parents=True, exist_ok=True)

    print("Loading BibTeX...")
    bib_by_doi, bib_by_title, all_records = load_bibtex(bib_path)
    print(f"  BibTeX entries: {len(all_records)}; entries with DOI: {len(bib_by_doi)}")

    rdf_attach_stats: Optional[Dict[str, int]] = None
    if rdf_path is not None:
        print("Loading Zotero RDF and attaching item URIs...")
        rdf_by_doi, rdf_by_title, rdf_records = load_zotero_rdf(
            rdf_path,
            library_prefix=args.zotero_library_uri_prefix,
            user_id=args.zotero_user_id,
            group_id=args.zotero_group_id,
        )
        rdf_attach_stats = attach_rdf_uris_to_bib_records(
            bib_by_doi,
            bib_by_title,
            all_records,
            rdf_by_doi,
            rdf_by_title,
            min_title_similarity=max(args.min_title_similarity, 0.90),
        )
        linked_count = sum(1 for rec in all_records if rec.uris)
        print(f"  RDF candidate records: {len(rdf_records)}; records with DOI: {len(rdf_by_doi)}")
        print(f"  BibTeX records with Zotero URI after RDF merge: {linked_count}/{len(all_records)}")
        if linked_count == 0:
            print("  Warning: no Zotero item URIs were attached from RDF. This is common: standard Zotero RDF often contains DOI URLs but not library item keys.", file=sys.stderr)

    sqlite_attach_stats: Optional[Dict[str, int]] = None
    sqlite_load_stats: Optional[Dict[str, int]] = None
    if zotero_sqlite_path is not None:
        print("Loading Zotero local database and attaching true item URIs...")
        try:
            sqlite_by_doi, sqlite_by_title, sqlite_load_stats = load_zotero_sqlite_records(
                zotero_sqlite_path,
                user_id=args.zotero_user_id,
                local_user_key=args.zotero_local_user_key,
                library_prefix=args.zotero_library_uri_prefix,
            )
            sqlite_attach_stats = attach_sqlite_uris_to_bib_records(
                all_records,
                sqlite_by_doi,
                sqlite_by_title,
                min_title_similarity=max(args.min_title_similarity, 0.90),
            )
            linked_count = sum(1 for rec in all_records if rec.uris)
            linked_with_ids = sum(1 for rec in all_records if rec.uris and rec.zotero_item_id is not None)
            print(f"  SQLite records by DOI: {len(sqlite_by_doi)}; by title: {len(sqlite_by_title)}")
            print(f"  BibTeX records with Zotero URI after SQLite merge: {linked_count}/{len(all_records)}")
            print(f"  BibTeX records with Zotero numeric itemID after SQLite merge: {linked_with_ids}/{len(all_records)}")
            if sqlite_load_stats and not (sqlite_load_stats.get('used_numeric_user_id') or sqlite_load_stats.get('used_local_user_key') or args.zotero_library_uri_prefix):
                print("  Warning: could not infer Zotero user ID/local user key. Add --zotero-user-id or --zotero-local-user-key if URIs are missing.", file=sys.stderr)
        except Exception as exc:
            print(f"  Warning: failed to read Zotero SQLite database: {exc}", file=sys.stderr)
    else:
        print(r"  Note: no Zotero SQLite database was provided/found. If Zotero still reports field-code errors, rerun with --zotero-sqlite C:\Users\<you>\Zotero\zotero.sqlite")

    print("Reading DOCX Reference section...")
    root = read_document_xml(docx_path)
    body_paras = get_body_paragraphs(root)
    heading_idx = find_reference_heading_index(body_paras, args.reference_heading)
    reference_start_idx: Optional[int] = None
    if heading_idx is None:
        reference_start_idx = find_reference_entries_start_index(body_paras)
        if reference_start_idx is None:
            print("ERROR: Could not find the Reference section heading or a numbered reference block.", file=sys.stderr)
            print('Tip 1: If your heading is Chinese, try: --reference-heading "参考文献"', file=sys.stderr)
            print("Tip 2: If your heading is custom, copy it exactly after --reference-heading.", file=sys.stderr)
            print("Last non-empty paragraphs detected:", file=sys.stderr)
            print(describe_reference_heading_candidates(body_paras), file=sys.stderr)
            return 3
        # extraction starts at start_idx + 1, so pass the paragraph before the first reference
        heading_idx = reference_start_idx - 1
        print(f"  Warning: Reference heading not found; using numbered reference block starting at paragraph {reference_start_idx}.")
    ref_texts = extract_references_from_paragraphs(body_paras, heading_idx)
    if not ref_texts:
        print("ERROR: Found the Reference section heading/block, but could not parse numbered references.", file=sys.stderr)
        print("Last non-empty paragraphs detected:", file=sys.stderr)
        print(describe_reference_heading_candidates(body_paras), file=sys.stderr)
        return 3
    print(f"  Parsed Reference entries: {len(ref_texts)}")

    matches = match_references(
        ref_texts,
        bib_by_doi,
        bib_by_title,
        all_records,
        allow_title_fallback=not args.no_title_fallback,
        min_title_similarity=args.min_title_similarity,
    )
    matched_count = sum(1 for m in matches.values() if m.status == "matched")
    uri_matched_count = sum(1 for m in matches.values() if m.status == "matched" and m.record and m.record.uris)
    sqlite_id_matched_count = sum(1 for m in matches.values() if m.status == "matched" and m.record and m.record.zotero_item_id is not None)
    print(f"  Matched references: {matched_count}/{len(matches)}")
    print(f"  Matched references with Zotero URI: {uri_matched_count}/{matched_count}")
    print(f"  Matched references with Zotero numeric itemID: {sqlite_id_matched_count}/{matched_count}")
    if args.require_zotero_uris and (uri_matched_count < matched_count or sqlite_id_matched_count < matched_count):
        print("ERROR: --require-zotero-uris was set, but not all matched references have true Zotero URI + numeric itemID.", file=sys.stderr)
        print("Use --zotero-sqlite pointing to Zotero's zotero.sqlite, or provide --zotero-user-id/--zotero-local-user-key if automatic URI construction failed.", file=sys.stderr)
        return 4

    match_report = report_dir / "reference_match_report.csv"
    write_match_report(match_report, matches)

    print("Scanning/converting citations...")
    logs, stats = convert_docx(
        docx_path=docx_path,
        out_path=out_path,
        matches_by_number=matches,
        dry_run=args.dry_run,
        remove_refs=args.remove_reference_section,
        custom_reference_heading=args.reference_heading,
        force_complex=args.force_complex,
    )
    conversion_log = report_dir / "conversion_log.csv"
    write_conversion_log(conversion_log, logs)

    summary = {
        "time": _dt.datetime.now().isoformat(timespec="seconds"),
        "input_docx": str(docx_path),
        "bibtex": str(bib_path),
        "zotero_rdf": str(rdf_path) if rdf_path is not None else None,
        "zotero_sqlite": str(zotero_sqlite_path) if zotero_sqlite_path is not None else None,
        "rdf_attach_stats": rdf_attach_stats,
        "sqlite_load_stats": sqlite_load_stats,
        "sqlite_attach_stats": sqlite_attach_stats,
        "bib_records_with_zotero_uri": sum(1 for rec in all_records if rec.uris),
        "bib_records_with_zotero_numeric_item_id": sum(1 for rec in all_records if rec.zotero_item_id is not None),
        "output_docx": None if args.dry_run else str(out_path),
        "dry_run": bool(args.dry_run),
        "reference_entries": len(matches),
        "matched_references": matched_count,
        "unmatched_references": len(matches) - matched_count,
        "conversion_stats": stats,
        "reports": {
            "reference_match_report": str(match_report),
            "conversion_log": str(conversion_log),
        },
    }
    write_json_summary(report_dir / "summary.json", summary)

    print("Done.")
    print(f"  Reference match report: {match_report}")
    print(f"  Conversion log:         {conversion_log}")
    print(f"  Summary JSON:           {report_dir / 'summary.json'}")
    if args.dry_run:
        print("  Dry run only: no DOCX was written.")
    else:
        print(f"  Converted DOCX:         {out_path}")
        print("Next step: open the converted .docx in Microsoft Word with Zotero installed, click Zotero -> Refresh, and inspect citations.")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
