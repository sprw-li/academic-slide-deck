#!/usr/bin/env python3
"""ingest.py — extract ``material.json`` (``material/v1``) from the case's real sources.

Contract: ``work/CONTRACT.md`` §3.  Usage::

    python ingest.py --case <case.json> --work <workdir> [--input <dir>] [--out <path>]

Exit codes: ``0`` material written (warnings allowed), ``2`` usage/IO error.

Sources handled: PDF (``pdfplumber``, falling back to ``pypdf``), DOCX
(``python-docx``), MD / TXT / HTML.  **No PyMuPDF used for text extraction.**

Extracted per source
--------------------
``pages`` (text + number), ``sections`` (id/title/page/char offsets), numbered
``figures`` (label, page, bbox-or-null, caption, locator), numbered ``tables``
(label, page, caption, locator, cells/n_rows/n_cols), ``numbers`` (value/raw/
unit/context/kind + locator) and candidate ``claims`` (text + locator).

The HARD RULE (contract §3)
---------------------------
Every figure, table, number and claim MUST carry a ``locator`` whose ``quote`` is
**verbatim source text** (checked with ``common.is_verbatim``).  Anything that
fails that check is **dropped** and recorded in ``warnings``.  Nothing is ever
fabricated, rounded or inferred from a filename.

Figure images are best-effort.  The raster backends are tried in order
(``fitz`` → ``pypdfium2`` → ``pdfplumber.to_image`` → portable ImageMagick /
Ghostscript → a ``$DSH_PDF_RASTER_CMD`` hook).  When none is available the figure
keeps ``kind="unknown"`` and **no image**, plus an explicit warning; downstream
stages must then emit nothing rather than a placeholder.

All paths written into ``material.json`` are **relative to the work directory**
(where ``material.json`` itself lives) and POSIX-separated; no absolute
author-machine path is ever emitted.
"""
from __future__ import annotations

import sys as _sys

_sys.dont_write_bytecode = True

import argparse
import html
import json
import os
import re
import shutil
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in _sys.path:
    _sys.path.insert(0, HERE)

import common as C  # noqa: E402  (path set above)

MAX_PDF_PAGES = 80
MAX_PAGE_CHARS = 250_000
#: the ledger must attest EVERY digit sequence a slide could honestly repeat, so the
#: cap is generous; the gate asks that any number printed on a slide exist verbatim in
#: the source, and a partial ledger would make honest pages fail for lack of a record.
MAX_NUMBERS = 2500
MAX_CLAIMS = 120
MAX_FIGURES = 60
MAX_TABLES = 40
MAX_CAPTION_CONT = 200

CAPTION_FIG_RE = re.compile(
    r"^\s*(?:Figure|FIGURE|Fig\.?|FIG\.?)\s*(\d{1,3}[A-Za-z]?)\s*[.:\u3002\uff1a\-\u2013\u2014]?\s*(.*)$")
CAPTION_TAB_RE = re.compile(
    r"^\s*(?:Table|TABLE|Tab\.?|TAB\.?)\s*(\d{1,3}[A-Za-z]?)\s*[.:\u3002\uff1a\-\u2013\u2014]?\s*(.*)$")

HEADING_RE = re.compile(
    r"^\s*(?:\d{1,2}(?:\.\d{1,2})*\.?\s+)?"
    r"(Abstract|摘要|Introduction|引言|Background|Related Work|相关工作|"
    r"Materials and Methods|Methods|Method|Methodology|Experimental|"
    r"材料与方法|方法|Results|Results and Discussion|结果|Discussions?|讨论|"
    r"Conclusion|Conclusions|结论|Limitations?|局限性|"
    r"Future Work|Acknowledg(?:e)?ments?|致谢|References|Bibliography|参考文献)\b[.:\s]*$",
    re.IGNORECASE)

NUMBER_RE = re.compile(
    r"(?<![\w.])(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)"      # magnitude
    r"(?:\s*(%|\u2030|\u00b0C|\u00b5g|ug|mg|kg|ng|pg|mL|ml|\u00b5L|uL|\u00b5m|um|nm|mm|cm|km|"
    r"kDa|Da|bp|kb|Mb|GB|MB|KB|kHz|MHz|GHz|Hz|kPa|MPa|GPa|Pa|mM|\u00b5M|uM|nM|pM|"
    r"rpm|min|hours?|hrs?|h|days?|weeks?|months?|years?|fold|times|ms|ns|s|d)\b)?",
    re.IGNORECASE)

#: every digit sequence, exactly as a reader (or the deck gate) would see it — this
#: catches numbers embedded in identifiers (``TiO2``, ``C60``) and unitless values
ALL_DIGITS_RE = re.compile(r"\d+(?:[.,]\d+)?")

#: a "bullet" beginning with one of these is a continuation fragment, not a statement
CONTINUATION_START = {
    "and", "or", "but", "which", "while", "whereas", "then", "that", "also", "thus",
    "hence", "with", "whose", "where", "when", "so", "because", "if", "although",
    "however", "therefore", "namely", "not", "both", "either", "each", "respectively",
}

RESULT_CUES = re.compile(
    r"\b(increas|decreas|reduc|higher|lower|greater|less|improv|enhanc|"
    r"significan|compared|versus|vs\.?|observed|achieved|reached|outperform|"
    r"correlat|difference|effect|yield|accuracy|sensitivity|speedup|提升|降低|"
    r"提高|显著|相比)\w*", re.IGNORECASE)
PARAM_CUES = re.compile(
    r"(\bn\s*=|\bp\s*[<=>]|95\s*%\s*CI|confidence interval|threshold|cut-?off|"
    r"dose|concentration|temperature|\bpH\b|baseline|parameter|assumed|default)",
    re.IGNORECASE)

CLAIM_SECTION_RE = re.compile(
    r"(abstract|摘要|result|结果|conclusion|结论|discussion|讨论|summary|findings)",
    re.IGNORECASE)

#: a bibliography entry is a citation, not a finding: it must never become a claim
REFERENCE_ENTRY_RE = re.compile(
    r"^\s*\[\d{1,3}\]"                       # [12] Author …
    r"|^\s*\d{1,3}\.\s+[A-Z][a-z]+,"         # 12. Author, …
    r"|\(\d{4}[a-z]?\)\s*\d{1,4}\s*[-–]\s*\d{1,4}"   # (2023) 44-58
    r"|\bet al\.,?\s*(?:19|20)\d{2}"
    r"|\bdoi:\s*10\."
    r"|\b(?:Journal|Proceedings|Nature|Science|Physical Review|arXiv)\b[^.]*\b(?:19|20)\d{2}\b",
    re.IGNORECASE)


# --------------------------------------------------------------------------- #
# small text utilities
# --------------------------------------------------------------------------- #
def _norm_ws(text):
    return re.sub(r"\s+", " ", text or "").strip()


def _strip_leading_number(text):
    """``"3. Results"`` -> ``"Results"`` (heading numbering is not content)."""
    return re.sub(r"^\s*\d+(?:[./]\d+)*[.)]?\s+", "", text or "").strip()


def _iter_lines(text):
    """Yield ``(start, end, line)`` for every line (offsets into ``text``)."""
    pos = 0
    for line in text.split("\n"):
        yield pos, pos + len(line), line
        pos += len(line) + 1


def _to_float(raw):
    try:
        return float(raw.replace(",", ""))
    except (AttributeError, ValueError):
        return None


def _posix_rel(path, base):
    rel = C.relpath(path, base)
    return rel.replace("\\", "/")


# --------------------------------------------------------------------------- #
# generic page-level extraction (works for every source kind)
# --------------------------------------------------------------------------- #
def find_sections(pages, source_id, warnings):
    """Sections from heading-like lines; falls back to a single full-text section."""
    full = "\n".join(p["text"] for p in pages)
    hits = []
    for page in pages:
        for start, _end, line in _iter_lines(page["text"]):
            bare = line.strip()
            if not bare or len(bare) > 120:
                continue
            m = HEADING_RE.match(bare)
            if m:
                hits.append((page["n"], bare))
    sections = []
    if not hits:
        if full.strip():
            sections.append({"id": "%s-S1" % source_id, "title": "(no headings detected)",
                             "page": pages[0]["n"] if pages else 0,
                             "char_start": 0, "char_end": len(full)})
            warnings.append("%s: no numbered/section headings detected; treated the "
                            "whole document as one section" % source_id)
        return sections
    for page_no, title in hits:
        # find the position of this heading line in the full text, from the end of
        # the previous section onwards (so repeated heading text cannot alias)
        from_pos = sections[-1]["char_end"] if sections else 0
        idx = full.find(title, from_pos)
        if idx < 0:
            idx = full.find(title)
        if idx < 0:
            continue
        if sections:
            sections[-1]["char_end"] = idx
        sections.append({"id": "%s-S%d" % (source_id, len(sections) + 1),
                         "title": title, "page": page_no,
                         "char_start": idx, "char_end": len(full)})
    if sections:
        sections[-1]["char_end"] = len(full)
    return sections


def find_captions(pages, source_id, warnings):
    """Numbered figure/table captions.  Each caption ``quote`` is a verbatim line."""
    figures, tables = [], []
    for page in pages:
        lines = list(_iter_lines(page["text"]))
        for i, (start, _end, line) in enumerate(lines):
            bare = line.strip()
            if not bare:
                continue
            m_fig = CAPTION_FIG_RE.match(bare)
            m_tab = CAPTION_TAB_RE.match(bare) if not m_fig else None
            if not m_fig and not m_tab:
                continue
            label_kind, num, rest = ("Figure", m_fig.group(1), m_fig.group(2)) if m_fig \
                else ("Table", m_tab.group(1), m_tab.group(2))
            # optional single continuation line (caption wrapped by the PDF extractor)
            caption_lines = [bare]
            if rest and not rest.rstrip().endswith((".", "!", "?", "\u3002")):
                for j in (i + 1, i + 2):
                    if j >= len(lines):
                        break
                    nxt = lines[j][2].strip()
                    if (not nxt or len(nxt) > MAX_CAPTION_CONT
                            or CAPTION_FIG_RE.match(nxt) or CAPTION_TAB_RE.match(nxt)
                            or HEADING_RE.match(nxt)):
                        break
                    caption_lines.append(nxt)
                    if nxt.endswith((".", "!", "?", "\u3002")):
                        break
            caption = " ".join(caption_lines)
            if not C.is_verbatim(caption_lines[0], page["text"]):
                warnings.append("%s: dropped %s %s on page %s — caption quote not "
                                "verbatim in the page text" % (source_id, label_kind,
                                                               num, page["n"]))
                continue
            record = {"label": "%s %s" % (label_kind, num), "number": num,
                      "page": page["n"], "caption": caption,
                      "locator": C.locator(source_id, page["n"], caption_lines[0])}
            (figures if label_kind == "Figure" else tables).append(record)
    return figures[:MAX_FIGURES], tables[:MAX_TABLES]


def find_numbers(pages, source_id, warnings):
    """Numbers with a verbatim locator.  ``raw`` is printed exactly as in the source.

    Two passes per line: the first recognises a value with an optional unit (rich
    context), the second recognises **every** digit sequence, so numbers embedded in
    identifiers (``TiO2``, ``C60``, ``PBDB-T``) are attested too.  A deck gate that
    asks "is this number in the source?" must not fail merely because the ledger was
    conservative — the answer has to be truthful in both directions.
    """
    seen, out = set(), []

    def add(raw, unit, page, line, match_start, match_end):
        value = _to_float(raw)
        if value is None:
            return
        quote = C.sentence_around(line, match_start, match_end) or line.strip()
        if not quote or not C.is_verbatim(quote, page["text"]):
            warnings.append("%s: dropped number %r on page %s — no verbatim locator"
                            % (source_id, raw, page["n"]))
            return
        key = (raw, unit, page["n"], quote)
        if key in seen:
            return
        seen.add(key)
        if PARAM_CUES.search(quote):
            kind = "parameter"
        elif RESULT_CUES.search(quote):
            kind = "result"
        else:
            kind = "other"
        out.append({"value": value, "raw": raw, "unit": unit,
                    "context": C.shorten(line), "kind": kind,
                    "locator": C.locator(source_id, page["n"], quote)})

    for page in pages:
        for _start, _end, line in _iter_lines(page["text"]):
            bare = line.strip()
            if not bare or len(bare) > 400:
                continue
            for match in NUMBER_RE.finditer(bare):
                raw, unit = match.group(1), (match.group(2) or "")
                if not unit and "." not in raw and len(raw) < 2:
                    continue                       # bare single digits (list markers)
                add(raw, unit, page, bare, match.start(), match.end())
                if len(out) >= MAX_NUMBERS:
                    warnings.append("%s: number extraction capped at %d items"
                                    % (source_id, MAX_NUMBERS))
                    return out
            for match in ALL_DIGITS_RE.finditer(bare):
                add(match.group(0), "", page, bare, match.start(), match.end())
                if len(out) >= MAX_NUMBERS:
                    warnings.append("%s: number extraction capped at %d items"
                                    % (source_id, MAX_NUMBERS))
                    return out
    return out


HEADING_WORD_HINTS = {
    "scope", "selection", "method", "methods", "methodology", "results", "result",
    "discussion", "conclusion", "conclusions", "introduction", "abstract",
    "background", "limitations", "limitation", "analysis", "experimental", "study",
    "data", "references", "table", "figure", "supplementary", "appendix",
    "keywords", "materials", "overview", "summary",
}


def _strip_leading_heading_phrase(text):
    """Strip an unpunctuated heading the text extractor glued to a sentence.

    ``"Scope and Literature Selection A synthetic search returned 412 records."``
    starts with a heading that the source put on its own line; without heading
    styles there is nothing else to recognise it by.  Only a short (≤6 word)
    Title-Case run that *contains a known heading word* is removed; small connector
    words (``and``, ``of``, ``the`` …) are allowed inside the run, and an ordinary
    sentence is left alone.
    """
    words = text.split()
    limit = min(6, len(words) - 3)
    connectors = {"and", "of", "the", "for", "in", "on", "to", "with", "a", "an",
                  "or", "at", "by", "as"}
    for size in range(limit, 1, -1):
        head = words[:size]
        tail = words[size] if size < len(words) else ""
        if not tail or not tail[0].islower():
            continue
        shaped = all(word[:1].isupper() or word.isupper() or word.lower() in connectors
                     for word in head if word[:1].isalpha())
        if not shaped:
            continue
        if not any(word[:1].isupper() for word in head):
            continue
        lowered = {re.sub(r"[^a-z]", "", word.lower()) for word in head}
        if not (lowered & HEADING_WORD_HINTS):
            continue
        return " ".join(words[size:]).strip()
    return text


def find_claims(pages, sections, source_id, warnings):
    """Candidate page claims: complete verbatim sentences from the paper's sections.

    Only **whole statements** qualify: a fragment that starts with a continuation
    word, starts mid-clause, is too short, or was cut mid-value is not a claim and is
    never offered to the planner (the planner may only drop, never repair).
    """
    prod_pages = {s["page"] for s in sections
                  if CLAIM_SECTION_RE.search(s.get("title") or "")}
    titles_by_page = {}
    for section in sections:
        title = _norm_ws(_strip_leading_number(section.get("title") or ""))
        if title:
            titles_by_page.setdefault(section.get("page"), []).append(title)
    out, seen = [], set()
    for page in pages:
        if prod_pages and page["n"] not in prod_pages:
            continue
        reflow = re.sub(r"\s+", " ", page["text"])
        for sentence in re.split(r"(?<=[.!?\u3002\uff01\uff1f])\s+", reflow):
            text = _norm_ws(sentence)
            # a sentence that begins where a heading ended drags the heading in with
            # it ("Result Phase analysis detected ..."); the heading is not content
            for title in titles_by_page.get(page["n"], []):
                if text.lower().startswith(title.lower() + " "):
                    text = text[len(title):].strip()
                    break
            text = _norm_ws(re.sub(r"^\d+(?:\.\d+)*[.)]?\s+", "", text))
            text = _strip_leading_heading_phrase(text)
            if not (25 <= len(text) <= 320) or text in seen:
                continue
            if CAPTION_FIG_RE.match(text) or CAPTION_TAB_RE.match(text):
                continue
            if REFERENCE_ENTRY_RE.search(text):
                continue                       # a citation is not a finding
            if sum(ch.isalpha() for ch in text) < 25:
                continue
            if text[0].islower() or text[0] in ",;:":
                continue
            first_word = re.match(r"[A-Za-z'\u2019]+", text)
            if first_word and first_word.group(0).lower() in CONTINUATION_START:
                continue
            if re.search(r"\b(?:and|or|but|which|then|with|of|at|for|to|in|on|"
                         r"the|a|an|by|from|as|is|are|was|were)$", text, re.I):
                continue                       # dangling tail: cut mid-clause
            if not text.endswith((".", "!", "?", "\u3002")):  # cut mid-value
                warnings.append("%s: dropped unterminated candidate claim on page %s "
                                "(%r)" % (source_id, page["n"], text[-30:]))
                continue
            if not C.is_verbatim(text, page["text"]):
                warnings.append("%s: dropped candidate claim on page %s — not "
                                "verbatim" % (source_id, page["n"]))
                continue
            seen.add(text)
            out.append({"text": text, "locator": C.locator(source_id, page["n"], text)})
            if len(out) >= MAX_CLAIMS:
                return out
    return out


def common_extraction(pages, source_id, warnings):
    """Generic per-page extraction shared by every source kind.

    Returns ``sections``, ``figure_captions`` (a ``(figures, tables)`` tuple),
    ``numbers`` and ``claims``; every item carries a verbatim locator.
    """
    sections = find_sections(pages, source_id, warnings)
    figure_captions = find_captions(pages, source_id, warnings)
    return {
        "sections": sections,
        "figure_captions": figure_captions,
        "numbers": find_numbers(pages, source_id, warnings),
        "claims": find_claims(pages, sections, source_id, warnings),
    }


# --------------------------------------------------------------------------- #
# PDF
# --------------------------------------------------------------------------- #
def _dedupe(records, key_fn, prefer, warnings, what):
    """Drop duplicate records (same paper supplied twice), keeping the richer one.

    ``prefer`` marks the *better* record; when both or neither qualify the first
    one wins, so the result is deterministic.
    """
    best = {}
    order = []
    for record in records:
        key = key_fn(record)
        if key not in best:
            best[key] = record
            order.append(key)
            continue
        incumbent = best[key]
        if prefer(record) and not prefer(incumbent):
            best[key] = record
            warnings.append("kept the richer duplicate for %s %r (%s) and dropped the "
                            "copy from %s"
                            % (what, str(key[0]) if isinstance(key, tuple) else key,
                               (record.get("locator") or {}).get("source"),
                               (incumbent.get("locator") or {}).get("source")))
    return [best[key] for key in order]


def _is_task_card(path):
    """True for the platform's own task card, which is an instruction file.

    ``task.md`` (and ``task.json``) sit beside ``case.json`` and describe what the
    case wants; their sentences are requirements, not research content, so they
    are ingested (for constraint reading) but never become claims, numbers,
    figures or tables.
    """
    name = os.path.basename(path or "").lower()
    return name in ("task.md", "task.txt", "task.json", "instructions.md",
                    "readme.md", "case.md")


def extract_pdf(path, source_id, work_dir, warnings):
    """PDF via pdfplumber (text/tables/figures), pypdf as a text-only fallback."""
    pdfplumber = C.try_import("pdfplumber")
    pages, figures, tables = [], [], []
    if pdfplumber is not None:
        try:
            with pdfplumber.open(path) as pdf:
                page_cap = min(len(pdf.pages), MAX_PDF_PAGES)
                if len(pdf.pages) > MAX_PDF_PAGES:
                    warnings.append("%s: %d pages, only the first %d were ingested"
                                    % (source_id, len(pdf.pages), MAX_PDF_PAGES))
                for index in range(page_cap):
                    page = pdf.pages[index]
                    text = page.extract_text(x_tolerance=1.5, y_tolerance=2.0) or ""
                    text = text[:MAX_PAGE_CHARS]
                    pages.append({"n": index + 1, "text": text})
                    figures.extend(_pdf_figures_on_page(page, index + 1, source_id,
                                                       work_dir, warnings))
                    tables.extend(_pdf_tables_on_page(page, index + 1, source_id,
                                                      warnings))
        except Exception as exc:
            warnings.append("%s: pdfplumber failed (%s: %s); falling back to pypdf"
                            % (source_id, type(exc).__name__, exc))
            pages = []
    if not pages:
        pages = _pdf_text_pypdf(path, source_id, warnings)
    return {"pages": pages, "figures": figures, "tables": tables}


def _pdf_text_pypdf(path, source_id, warnings):
    pypdf = C.try_import("pypdf")
    if pypdf is None:
        warnings.append("%s: neither pdfplumber nor pypdf is importable — no PDF text "
                        "extracted" % source_id)
        return []
    try:
        reader = pypdf.PdfReader(path)
        pages = []
        for index, page in enumerate(reader.pages[:MAX_PDF_PAGES]):
            try:
                text = page.extract_text() or ""
            except Exception as exc:
                warnings.append("%s: pypdf could not read page %d (%s)"
                                % (source_id, index + 1, exc))
                text = ""
            pages.append({"n": index + 1, "text": text[:MAX_PAGE_CHARS]})
        if len(reader.pages) > MAX_PDF_PAGES:
            warnings.append("%s: %d pages, only the first %d were ingested"
                            % (source_id, len(reader.pages), MAX_PDF_PAGES))
        warnings.append("%s: PDF text came from pypdf — figure bbox and table cells "
                        "are unavailable on this path" % source_id)
        return pages
    except Exception as exc:
        warnings.append("%s: pypdf failed too (%s: %s) — no text extracted"
                        % (source_id, type(exc).__name__, exc))
        return []


def _pdf_caption_bbox(page, caption_line):
    """Bbox of a caption line: ``page.search`` first, then a word-sequence match.

    The word-sequence fallback matters because a caption whose line was wrapped or
    whose glyphs were re-encoded will not match ``search`` verbatim, and without a
    caption bbox the figure region cannot be located at all.
    """
    probe = _norm_ws(caption_line)[:90]
    for query in (probe, probe[:40]):
        if not query:
            continue
        try:
            hits = page.search(query)
        except Exception:
            hits = []
        if hits:
            h = hits[0]
            return [round(h["x0"], 2), round(h["top"], 2),
                    round(h["x1"], 2), round(h["bottom"], 2)]
    words = re.findall(r"[\w%./°-]+", probe.lower())[:8]
    if len(words) < 3:
        return None
    try:
        page_words = page.extract_words(x_tolerance=1.5, y_tolerance=2.0) or []
    except Exception:
        return None
    strip = lambda s: re.sub(r"[^\w%./°-]", "", s.lower())
    tokens = [strip(w.get("text", "")) for w in page_words]
    for start in range(len(tokens) - len(words) + 1):
        if tokens[start] == words[0] and \
                tokens[start:start + len(words)] == words:
            run = page_words[start:start + len(words)]
            return [round(min(w["x0"] for w in run), 2),
                    round(min(w["top"] for w in run), 2),
                    round(max(w["x1"] for w in run), 2),
                    round(max(w["bottom"] for w in run), 2)]
    return None


def _pdf_figures_on_page(page, page_no, source_id, work_dir, warnings):
    """Figures = numbered captions on this page + a best-effort image region hint."""
    text = page.extract_text(x_tolerance=1.5, y_tolerance=2.0) or ""
    out = []
    images = []
    try:
        images = list(page.images or [])
    except Exception:
        images = []
    drawings = 0
    try:
        drawings = len(page.curves or []) + len(page.rects or []) + len(page.lines or [])
    except Exception:
        drawings = 0
    for start, _end, line in _iter_lines(text):
        bare = line.strip()
        m = CAPTION_FIG_RE.match(bare) if bare else None
        if not m:
            continue
        caption = bare
        if not C.is_verbatim(caption, text):
            warnings.append("%s: dropped figure %s on page %d — caption not verbatim"
                            % (source_id, m.group(1), page_no))
            continue
        cap_bbox = _pdf_caption_bbox(page, bare)
        region = _nearest_image_region(images, cap_bbox)
        if region is None and drawings and cap_bbox is not None:
            region = _drawing_region(page, cap_bbox)
        kind = "raster" if _region_has_image(images, region) else (
            "vector" if region is not None else "unknown")
        out.append({"label": "Figure %s" % m.group(1), "number": m.group(1),
                    "page": page_no, "caption": caption,
                    "caption_bbox": cap_bbox,
                    "locator": C.locator(source_id, page_no, caption),
                    "bbox": region, "kind": kind})
    return out


def _nearest_image_region(images, cap_bbox, max_gap=600.0, below_gap=260.0):
    """Image region nearest the caption: above it first, then just below it."""
    if not images or not cap_bbox:
        return None
    cap_top, cap_bottom = cap_bbox[1], cap_bbox[3]

    def box(img):
        try:
            return [round(float(img["x0"]), 2), round(float(img["top"]), 2),
                    round(float(img["x1"]), 2), round(float(img["bottom"]), 2)]
        except (KeyError, TypeError, ValueError):
            return None

    best, best_gap = None, None
    for img in images:
        try:
            bottom = float(img.get("bottom"))
        except (TypeError, ValueError):
            continue
        gap = cap_top - bottom
        if -5.0 <= gap <= max_gap and (best_gap is None or gap < best_gap):
            best, best_gap = img, gap
    if best is not None:
        return box(best)
    for img in images:                     # caption above the figure
        try:
            top = float(img.get("top"))
        except (TypeError, ValueError):
            continue
        gap = top - cap_bottom
        if -5.0 <= gap <= below_gap and (best_gap is None or gap < best_gap):
            best, best_gap = img, gap
    return box(best) if best is not None else None


def _region_has_image(images, region):
    return bool(images) and region is not None


def _drawing_region(page, cap_bbox):
    """Union bbox of vector drawings sitting directly above the caption."""
    boxes = []
    for attr in ("curves", "rects", "lines"):
        try:
            items = getattr(page, attr) or []
        except Exception:
            items = []
        for item in items:
            try:
                top, bottom = float(item["top"]), float(item["bottom"])
            except (KeyError, TypeError, ValueError):
                continue
            gap = cap_bbox[1] - bottom
            if -5.0 <= gap <= 600.0:
                boxes.append((float(item["x0"]), top, float(item["x1"]), bottom))
    if not boxes:
        return None
    return [round(min(b[0] for b in boxes), 2), round(min(b[1] for b in boxes), 2),
            round(max(b[2] for b in boxes), 2), round(max(b[3] for b in boxes), 2)]


def _pdf_tables_on_page(page, page_no, source_id, warnings):
    """Numbered table captions on this page + cells from the nearest detected table."""
    text = page.extract_text(x_tolerance=1.5, y_tolerance=2.0) or ""
    found = []
    try:
        found = list(page.find_tables() or [])
    except Exception:
        found = []
    extracted = []
    if not found:
        try:
            extracted = [t for t in (page.extract_tables() or []) if t]
        except Exception:
            extracted = []
    out, used = [], set()
    for _start, _end, line in _iter_lines(text):
        bare = line.strip()
        m = CAPTION_TAB_RE.match(bare) if bare else None
        if not m:
            continue
        if not C.is_verbatim(bare, text):
            warnings.append("%s: dropped table %s on page %d — caption not verbatim"
                            % (source_id, m.group(1), page_no))
            continue
        cap_bbox = _pdf_caption_bbox(page, bare)
        cells, n_rows, n_cols = None, 0, 0
        cells, n_rows, n_cols, used = _match_table_cells(found, extracted, cap_bbox,
                                                        used)
        out.append({"label": "Table %s" % m.group(1), "number": m.group(1),
                    "page": page_no, "caption": bare,
                    "locator": C.locator(source_id, page_no, bare),
                    "cells": cells, "n_rows": n_rows, "n_cols": n_cols})
        if cells is None:
            warnings.append("%s: table %s on page %d kept with no cells (no matching "
                            "detected grid)" % (source_id, m.group(1), page_no))
    return out


def _match_table_cells(found, extracted, cap_bbox, used):
    """Pick the detected grid directly above the caption; return (cells, rows, cols)."""
    pick = None
    if found and cap_bbox:
        best_gap = None
        for index, table in enumerate(found):
            if index in used:
                continue
            bbox = getattr(table, "bbox", None)
            if not bbox:
                continue
            gap = cap_bbox[1] - float(bbox[3])
            if -10.0 <= gap <= 400.0 and (best_gap is None or gap < best_gap):
                pick, best_gap = index, gap
    elif extracted:
        for index in range(len(extracted)):
            if index not in used:
                pick = index
                break
    if pick is None:
        return None, 0, 0, used
    used = set(used) | {pick}
    try:
        rows = found[pick].extract() if found else extracted[pick]
    except Exception:
        rows = None
    if not rows:
        return None, 0, 0, used
    clean = [[("" if c is None else str(c)) for c in row] for row in rows]
    return clean, len(clean), max((len(r) for r in clean), default=0), used


# --------------------------------------------------------------------------- #
# DOCX
# --------------------------------------------------------------------------- #
def extract_docx(path, source_id, work_dir, warnings):
    docx = C.try_import("docx")
    if docx is None:
        warnings.append("%s: python-docx is not importable — DOCX skipped" % source_id)
        return {"pages": [], "figures": [], "tables": []}
    try:
        document = docx.Document(path)
    except Exception as exc:
        warnings.append("%s: python-docx could not open the file (%s: %s)"
                        % (source_id, type(exc).__name__, exc))
        return {"pages": [], "figures": [], "tables": []}

    chunks, table_records = [], []
    last_paragraph = ""
    for block in document.element.body.iterchildren():
        tag = block.tag.split("}")[-1]
        if tag == "p":
            text = "".join(node.text or "" for node in block.iter() if node.tag.endswith("}t"))
            if text.strip():
                chunks.append(text)
                last_paragraph = text.strip()
        elif tag == "tbl":
            rows = []
            for tr in block.iter():
                if not tr.tag.endswith("}tr"):
                    continue
                cells = []
                for tc in tr:
                    if not tc.tag.endswith("}tc"):
                        continue
                    cells.append("".join(n.text or "" for n in tc.iter()
                                         if n.tag.endswith("}t")).strip())
                if cells:
                    rows.append(cells)
            if rows:
                # the paragraph immediately before the table is usually its caption
                caption = last_paragraph if CAPTION_TAB_RE.match(last_paragraph) else ""
                table_records.append((rows, caption))
                last_paragraph = ""
    text = "\n".join(chunks)
    # Table cell text is real source text, but python-docx does not expose it through
    # the body's top-level paragraphs.  Appending it (row by row) lets cell values
    # carry a verbatim locator and be attested in material.numbers — otherwise a
    # truthful native table on a slide could not be traced back to the source.
    for rows, _caption in table_records:
        for row in rows:
            text += "\n" + " | ".join(row)
    pages = [{"n": 1, "text": text[:MAX_PAGE_CHARS]}]

    tables = []
    for index, (rows, docx_caption) in enumerate(table_records[:MAX_TABLES]):
        quote = docx_caption or (" | ".join(rows[0]) if rows else "")
        cells = [[c for c in row] for row in rows[:2000]]
        if not quote or not C.is_verbatim(quote, text):
            # fall back to any verbatim cell text so the table still has a locator
            quote = ""
            for row in rows:
                for cell in row:
                    if cell and C.is_verbatim(cell, text):
                        quote = cell
                        break
                if quote:
                    break
        if not quote:
            warnings.append("%s: dropped table #%d — no verbatim locator available"
                            % (source_id, index + 1))
            continue
        number_match = CAPTION_TAB_RE.match(quote)
        number = number_match.group(1) if number_match else str(index + 1)
        tables.append({"label": "Table %s" % number, "number": number,
                       "page": 1, "caption": quote,
                       "locator": C.locator(source_id, 1, quote),
                       "cells": cells, "n_rows": len(cells),
                       "n_cols": max((len(r) for r in cells), default=0)})
        if not number_match:
            warnings.append("%s: table #%d has no 'Table N' caption; locator uses a "
                            "verbatim cell instead" % (source_id, index + 1))

    figures = []
    images = _docx_image_parts(document, warnings, source_id)
    captions = []
    for _start, _end, line in _iter_lines(text):
        bare = line.strip()
        if bare and CAPTION_FIG_RE.match(bare) and C.is_verbatim(bare, text):
            captions.append(bare)
    for index, caption in enumerate(captions[:MAX_FIGURES]):
        figure = {"label": "Figure %s" % CAPTION_FIG_RE.match(caption).group(1),
                  "number": CAPTION_FIG_RE.match(caption).group(1),
                  "page": 1, "caption": caption,
                  "locator": C.locator(source_id, 1, caption),
                  "bbox": None, "kind": "unknown"}
        if index < len(images):
            figure["_image_bytes"] = images[index]
            figure["kind"] = "raster"
        figures.append(figure)
    if captions and not images:
        warnings.append("%s: %d figure caption(s) found but no embedded images could be "
                        "extracted" % (source_id, len(captions)))
    return {"pages": pages, "figures": figures, "tables": tables}


def _docx_image_parts(document, warnings, source_id):
    """Embedded image blobs in document order (best effort)."""
    out = []
    try:
        rels = document.part.rels
    except Exception:
        return out
    for rel in rels.values():
        try:
            if "image" not in str(getattr(rel, "reltype", "")):
                continue
            part = rel.target_part
            blob = getattr(part, "blob", None)
            name = str(getattr(part, "partname", "image"))
            if blob:
                out.append((blob, os.path.splitext(name)[1].lstrip(".") or "png"))
        except Exception as exc:
            warnings.append("%s: could not read an embedded image (%s)" % (source_id, exc))
    return out


# --------------------------------------------------------------------------- #
# MD / TXT / HTML
# --------------------------------------------------------------------------- #
def _html_to_text(raw):
    try:
        parser_cls = C.try_import("html.parser").HTMLParser
    except AttributeError:  # pragma: no cover
        return re.sub(r"<[^>]+>", " ", raw)

    class _Strip(parser_cls):
        def __init__(self):
            super().__init__(convert_charrefs=True)
            self.parts = []
            self.skip = 0

        def handle_starttag(self, tag, attrs):
            if tag in ("script", "style"):
                self.skip += 1
            elif tag in ("p", "div", "br", "li", "tr", "h1", "h2", "h3", "h4"):
                self.parts.append("\n")

        def handle_endtag(self, tag):
            if tag in ("script", "style") and self.skip:
                self.skip -= 1
            elif tag in ("p", "div", "li", "tr", "h1", "h2", "h3", "h4"):
                self.parts.append("\n")

        def handle_data(self, data):
            if not self.skip:
                self.parts.append(data)

    stripper = _Strip()
    stripper.feed(raw)
    return html.unescape("".join(stripper.parts))


def extract_text_file(path, source_id, kind, warnings):
    try:
        with open(path, "r", encoding="utf-8-sig", errors="replace") as fh:
            raw = fh.read()
    except OSError as exc:
        warnings.append("%s: cannot read (%s)" % (source_id, exc))
        return {"pages": [], "figures": [], "tables": []}
    if kind == "html":
        text = _html_to_text(raw)
    else:
        text = raw
    return {"pages": [{"n": 1, "text": text[:MAX_PAGE_CHARS]}],
            "figures": [], "tables": []}


# --------------------------------------------------------------------------- #
# figure image capture (best effort, honest failure)
# --------------------------------------------------------------------------- #
def _candidate_tool_dirs():
    """Portable tool locations derived from env + ancestors.  No absolute author path."""
    dirs = []
    for var in ("SKILL_TOOLCHAINS", "DSH_SKILL_REPOS", "SKILL_DOWNLOADED_REPOS"):
        value = os.environ.get(var)
        if value:
            dirs.append(value)
    bases = [os.getcwd(), os.path.dirname(os.path.abspath(__file__))]
    for base in bases:
        node = base
        for _ in range(4):
            node = os.path.dirname(node) or node
            dirs.append(os.path.join(node, "SKILL_Downloaded", "repos", "08-toolchains"))
            dirs.append(os.path.join(node, "08-toolchains"))
            if node == os.path.dirname(node):
                break
    seen, out = set(), []
    for d in dirs:
        if d and d not in seen:
            seen.add(d)
            out.append(d)
    return out


def _find_imagemagick():
    for name in ("magick", "magick.exe", "convert", "convert.exe"):
        found = shutil.which(name)
        if found:
            return [found]
    for base in _candidate_tool_dirs():
        if not os.path.isdir(base):
            continue
        for name in ("magick.exe", "magick", "convert.exe", "convert"):
            hit = os.path.join(base, name)
            if os.path.isfile(hit):
                return [hit]
        try:
            for entry in sorted(os.listdir(base)):
                for name in ("magick.exe", "magick"):
                    hit = os.path.join(base, entry, name)
                    if os.path.isfile(hit):
                        return [hit]
        except OSError:
            continue
    return None


def _find_ghostscript():
    for name in ("gs", "gs.exe", "gswin64c.exe", "gswin32c.exe"):
        found = shutil.which(name)
        if found:
            return found
    for base in _candidate_tool_dirs():
        if not os.path.isdir(base):
            continue
        for entry in sorted(os.listdir(base) or []):
            for name in ("gswin64c.exe", "gs.exe"):
                hit = os.path.join(base, entry, "bin", name)
                if os.path.isfile(hit):
                    return hit
    return None


def _pdf_page_bbox_px(page, bbox, resolution=200):
    """Convert a PDF-point bbox to an integer pixel box (y flipped)."""
    width, height = float(page.width), float(page.height)
    x0, top, x1, bottom = [float(v) for v in bbox]
    x0, x1 = max(0.0, min(x0, x1)), min(width, max(x0, x1))
    top, bottom = max(0.0, min(top, bottom)), min(height, max(top, bottom))
    if x1 - x0 < 2 or bottom - top < 2:
        return None
    scale = resolution / 72.0
    return (int(x0 * scale), int(top * scale), int(x1 * scale), int(bottom * scale))


def _capture_with_pdf_backends(pdf_path, page_no, bbox, out_png, warnings, key):
    """fitz → pypdfium2 → pdfplumber.to_image.  Returns backend name or ``None``."""
    fitz = C.try_import("fitz")
    if fitz is not None:
        try:
            doc = fitz.open(pdf_path)
            page = doc.load_page(page_no - 1)
            clip = fitz.Rect(*[float(v) for v in bbox]) if bbox else None
            pix = page.get_pixmap(clip=clip, dpi=200) if clip else page.get_pixmap(dpi=200)
            pix.save(out_png)
            return "fitz"
        except Exception as exc:
            warnings.append("%s: fitz rasterisation failed (%s: %s)"
                            % (key, type(exc).__name__, exc))
    pdfium = C.try_import("pypdfium2")
    if pdfium is not None:
        try:
            pdf = pdfium.PdfDocument(pdf_path)
            page = pdf[page_no - 1]
            bitmap = page.render(scale=200 / 72.0)
            image = bitmap.to_pil()
            if bbox:
                width, height = page.get_size()
                x0, top, x1, bottom = [float(v) for v in bbox]
                sx, sy = image.width / float(width), image.height / float(height)
                box = (int(x0 * sx), int(top * sy), int(x1 * sx), int(bottom * sy))
                if box[2] - box[0] > 2 and box[3] - box[1] > 2:
                    image = image.crop(box)
            image.save(out_png)
            return "pypdfium2"
        except Exception as exc:
            warnings.append("%s: pypdfium2 rasterisation failed (%s: %s)"
                            % (key, type(exc).__name__, exc))
    pdfplumber = C.try_import("pdfplumber")
    if pdfplumber is not None:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[page_no - 1]
                target = page.crop(tuple(float(v) for v in bbox)) if bbox else page
                image = target.to_image(resolution=200).original
                image.save(out_png)
            return "pdfplumber.to_image"
        except Exception as exc:
            warnings.append("%s: pdfplumber.to_image failed (%s: %s)"
                            % (key, type(exc).__name__, exc))
    return None


def _capture_with_tools(pdf_path, page_no, bbox, out_png, warnings, key):
    """Portable ImageMagick / Ghostscript / ``$DSH_PDF_RASTER_CMD`` hook."""
    magick = _find_imagemagick()
    if magick and bbox:
        try:
            res = subprocess.run(
                magick + ["-density", "200", "%s[%d]" % (pdf_path, page_no - 1),
                          "-crop", "%dx%d+%d+%d" % (
                              int(float(bbox[2]) - float(bbox[0])) * 200 // 72 or 1,
                              int(float(bbox[3]) - float(bbox[1])) * 200 // 72 or 1,
                              int(float(bbox[0])) * 200 // 72,
                              int(float(bbox[1])) * 200 // 72),
                          "+repage", out_png],
                capture_output=True, timeout=120)
            if res.returncode == 0 and os.path.isfile(out_png):
                return "imagemagick"
            warnings.append("%s: imagemagick exited %s" % (key, res.returncode))
        except Exception as exc:
            warnings.append("%s: imagemagick call failed (%s: %s)"
                            % (key, type(exc).__name__, exc))
    gs = _find_ghostscript()
    if gs:
        try:
            res = subprocess.run(
                [gs, "-q", "-dNOPAUSE", "-dBATCH", "-sDEVICE=png16m", "-r200",
                 "-dFirstPage=%d" % page_no, "-dLastPage=%d" % page_no,
                 "-sOutputFile=" + out_png, pdf_path],
                capture_output=True, timeout=180)
            if res.returncode == 0 and os.path.isfile(out_png):
                return "ghostscript"
            warnings.append("%s: ghostscript exited %s" % (key, res.returncode))
        except Exception as exc:
            warnings.append("%s: ghostscript call failed (%s: %s)"
                            % (key, type(exc).__name__, exc))
    hook = os.environ.get("DSH_PDF_RASTER_CMD")
    if hook:
        try:
            cmd = (hook.replace("{input}", pdf_path).replace("{page}", str(page_no))
                   .replace("{out}", out_png))
            res = subprocess.run(cmd, shell=True, capture_output=True, timeout=300)
            if res.returncode == 0 and os.path.isfile(out_png):
                return "raster-hook"
            warnings.append("%s: $DSH_PDF_RASTER_CMD exited %s" % (key, res.returncode))
        except Exception as exc:
            warnings.append("%s: $DSH_PDF_RASTER_CMD failed (%s: %s)"
                            % (key, type(exc).__name__, exc))
    return None


def capture_figure_image(pdf_path, page_no, bbox, out_png, warnings, key):
    """Crop a figure region from the source PDF; returns ``(backend, error)``."""
    backend = _capture_with_pdf_backends(pdf_path, page_no, bbox, out_png, warnings, key)
    if backend is None:
        backend = _capture_with_tools(pdf_path, page_no, bbox, out_png, warnings, key)
    if backend is None:
        if os.path.isfile(out_png):
            try:
                os.remove(out_png)
            except OSError:
                pass
        return None, "no raster backend available"
    if not os.path.isfile(out_png) or os.path.getsize(out_png) == 0:
        return None, "backend %s produced no image" % backend
    return backend, None


# --------------------------------------------------------------------------- #
# orchestration
# --------------------------------------------------------------------------- #
def _rewrite_image_bytes(figure, work_dir, figures_dir, warnings, key):
    """Save a DOCX-embedded image blob; returns the relative path or ``None``."""
    blob, ext = figure.pop("_image_bytes")
    out_png = os.path.join(figures_dir, "%s.png" % figure["id"])
    try:
        if ext.lower() in ("png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "webp"):
            from PIL import Image  # noqa: WPS433 (guarded optional import)
            import io
            with Image.open(io.BytesIO(blob)) as img:
                img.convert("RGB").save(out_png, "PNG")
        else:
            return None
    except Exception as exc:
        warnings.append("%s: could not convert an embedded DOCX image (%s: %s)"
                        % (key, type(exc).__name__, exc))
        try:
            import io
            with open(out_png, "wb") as fh:
                fh.write(blob)
        except OSError:
            return None
    return _posix_rel(out_png, work_dir)


def ingest(case_path, work_dir, input_dir=None, out_path=None, log=None):
    """Extract ``material.json`` for one case.  Returns the material dict."""
    warnings = []
    case = C.load_case(case_path)
    case_id = C.extract_case_id(case, case_path)
    case_dir = os.path.dirname(os.path.abspath(case_path))
    C.ensure_dir(work_dir)
    figures_dir = C.ensure_dir(os.path.join(work_dir, "figures"))

    sources_out, figures, tables, numbers, claims = [], [], [], [], []
    resolved = C.discover_source_files(case, case_path, input_dir=input_dir)
    if not resolved:
        warnings.append("no source file (pdf/docx/md/txt/html) could be located for this "
                        "case; material.json contains no sources")
    if log:
        log.event("sources_found", count=len(resolved),
                  files=[os.path.basename(s["path"]) for s in resolved])

    for index, item in enumerate(resolved):
        source_id = "src%d" % (index + 1)
        path, kind = item["path"], item["kind"]
        key = "%s(%s)" % (source_id, os.path.basename(path))
        task_card = _is_task_card(path)
        if log:
            log.info("ingesting %s" % key, kind=kind, role="task_card" if task_card
                     else "material")
        if kind == "pdf":
            data = extract_pdf(path, source_id, work_dir, warnings)
        elif kind == "docx":
            data = extract_docx(path, source_id, work_dir, warnings)
        else:
            data = extract_text_file(path, source_id, kind, warnings)

        pages = data.get("pages") or []
        if not pages or not any(p["text"].strip() for p in pages):
            warnings.append("%s: no text extracted" % key)

        common_data = common_extraction(pages, source_id, warnings)
        source_sections = common_data["sections"]
        if task_card:
            # the task card is read for its constraints (see common.extract_constraints)
            # but its sentences are requirements, not research findings
            source_sections = []
            cleaned = []
            for note in warnings:
                if note.startswith(key) and "no numbered/section headings" in note:
                    continue
                cleaned.append(note)
            warnings[:] = cleaned
            for rec in data.get("figures") or []:
                warnings.append("%s is the case task card; its figure caption %r is not "
                                "treated as research material"
                                % (key, rec.get("label")))
        else:
            # --- figures: PDF-specific (bbox/image) plus generic text-source captions --
            # A source may yield the same numbered caption twice (e.g. DOCX embedded
            # images are picked up by the DOCX path and their captions by the generic
            # path); keep the richer record.
            figure_records = list(data.get("figures") or [])
            seen_labels = {rec.get("label") for rec in figure_records}
            for cap in common_data["figure_captions"][0]:
                if cap.get("label") in seen_labels:
                    continue
                seen_labels.add(cap.get("label"))
                figure_records.append(dict(cap, bbox=None, kind="unknown"))
            for rec in figure_records:
                rec = dict(rec)
                rec["source"] = source_id
                figures.append(rec)
            table_records = list(data.get("tables") or [])
            for cap in common_data["figure_captions"][1]:
                if not any(t.get("label") == cap["label"] for t in table_records):
                    table_records.append(dict(cap, cells=None, n_rows=0, n_cols=0))
            for rec in table_records:
                rec = dict(rec)
                rec["source"] = source_id
                tables.append(rec)
            for rec in common_data["numbers"]:
                rec = dict(rec)
                rec["source"] = source_id
                numbers.append(rec)
            for rec in common_data["claims"]:
                rec = dict(rec)
                rec["source"] = source_id
                claims.append(rec)

        sources_out.append({
            "id": source_id, "kind": kind,
            "role": "task_card" if task_card else "material",
            "path": _posix_rel(path, work_dir),
            "sha256": C.sha256_file(path),
            "bytes": os.path.getsize(path) if os.path.isfile(path) else 0,
            "pages": pages, "sections": source_sections,
            "text": "\n".join(p["text"] for p in pages),
            "warnings": [w for w in warnings if w.startswith(key)],
        })

    # ---- cross-format de-duplication -------------------------------------------
    # The same paper is often supplied as PDF *and* DOCX.  Without this, every
    # caption, number and sentence would appear twice (once per format) and the
    # deck would repeat itself.  Records with a captured image / with cells win.
    stem_by_source = {s["id"]: os.path.splitext(os.path.basename(s["path"]))[0]
                      for s in sources_out}
    figures = _dedupe(figures,
                      lambda r: (r.get("label"),
                                 stem_by_source.get(r.get("source"), "")),
                      prefer=lambda r: ("_image_bytes" in r or "image" in r),
                      warnings=warnings, what="figure")
    tables = _dedupe(tables, lambda r: r.get("label"),
                     prefer=lambda r: bool(r.get("cells")), warnings=warnings,
                     what="table")
    numbers = _dedupe(numbers, lambda r: (r.get("raw"), r.get("unit"),
                                          _norm_ws((r.get("locator") or {})
                                                   .get("quote") or "")),
                      prefer=lambda r: True, warnings=warnings, what="number")
    claims = _dedupe(claims, lambda r: _norm_ws(r.get("text") or ""),
                     prefer=lambda r: True, warnings=warnings, what="claim")

    # A figure/table caption is the caption of that exhibit, not a page claim.  Its
    # sentence otherwise reaches the deck twice: once as a bullet, once as the figure
    # page's title — an overlap the deck gate (correctly) reports as reused content.
    caption_keys = set()
    for record in figures + tables:
        body = re.sub(r"^\s*(?:Figure|Fig\.?|Table|Tab\.?)\s*\d+[A-Za-z]?\s*[.:]?\s*",
                      "", record.get("caption") or "", flags=re.IGNORECASE)
        body = _norm_ws(body).lower()
        if len(body) >= 25:
            caption_keys.add(body[:60])
    if caption_keys:
        kept_claims = []
        for claim in claims:
            key = _norm_ws(claim.get("text") or "").lower()[:60]
            if key and key in caption_keys:
                warnings.append("claim %r dropped: it is the wording of a figure/table "
                                "caption, which already appears on the exhibit's own "
                                "page" % claim.get("text", "")[:60])
                continue
            kept_claims.append(claim)
        claims = kept_claims

    # ---- assign stable ids, capture images, and re-verify every locator ----------
    for i, rec in enumerate(figures, start=1):
        rec["id"] = "fig%d" % i
    for i, rec in enumerate(tables, start=1):
        rec["id"] = "tab%d" % i
    for i, rec in enumerate(numbers, start=1):
        rec["id"] = "num%d" % i
    for i, rec in enumerate(claims, start=1):
        rec["id"] = "cl%d" % i

    by_id = {s["id"]: s for s in sources_out}
    kept_figures = []
    for rec in figures:
        src = by_id.get(rec.pop("source", None))
        text = src["text"] if src else ""
        loc = rec.get("locator") or {}
        if not loc.get("quote") or not C.is_verbatim(loc.get("quote", ""), text):
            warnings.append("dropped figure %s — locator quote not verbatim"
                            % rec.get("label", "?"))
            continue
        if "_image_bytes" in rec and src is not None:
            rel = _rewrite_image_bytes(rec, work_dir, figures_dir, warnings, rec["id"])
            if rel:
                rec["image"] = {"path": rel, "sha256": C.sha256_file(
                    os.path.join(work_dir, rel)), "backend": "docx-embedded"}
        if "image" not in rec and src is not None and src["kind"] == "pdf" \
                and rec.get("bbox") and rec.get("page"):
            src_path = os.path.join(work_dir, src["path"])
            out_png = os.path.join(figures_dir, "%s.png" % rec["id"])
            backend, error = capture_figure_image(src_path, rec["page"], rec["bbox"],
                                                  out_png, warnings, rec["id"])
            if backend:
                rec["image"] = {"path": _posix_rel(out_png, work_dir),
                                "sha256": C.sha256_file(out_png), "backend": backend}
            else:
                rec["kind"] = "unknown"
                warnings.append("figure %s has no image: %s (deck builder must emit "
                                "nothing for it)" % (rec["id"], error))
        elif "image" not in rec and rec.get("kind") == "raster":
            rec["kind"] = "unknown"
        rec.pop("caption_bbox", None)
        rec.pop("_image_bytes", None)
        rec.pop("source", None)
        kept_figures.append(rec)

    kept_tables = []
    for rec in tables:
        src = by_id.get(rec.pop("source", None))
        text = src["text"] if src else ""
        loc = rec.get("locator") or {}
        if not loc.get("quote") or not C.is_verbatim(loc.get("quote", ""), text):
            warnings.append("dropped table %s — locator quote not verbatim"
                            % rec.get("label", "?"))
            continue
        kept_tables.append(rec)

    kept_numbers = []
    for rec in numbers:
        src = by_id.get(rec.pop("source", None))
        text = src["text"] if src else ""
        loc = rec.get("locator") or {}
        if not loc.get("quote") or not C.is_verbatim(loc.get("quote", ""), text):
            warnings.append("dropped number %s — locator quote not verbatim"
                            % rec.get("raw", "?"))
            continue
        kept_numbers.append(rec)

    kept_claims = []
    for rec in claims:
        src = by_id.get(rec.pop("source", None))
        text = src["text"] if src else ""
        loc = rec.get("locator") or {}
        if not loc.get("quote") or not C.is_verbatim(loc.get("quote", ""), text):
            warnings.append("dropped claim — locator quote not verbatim")
            continue
        kept_claims.append(rec)

    material = {
        "schema": C.SCHEMA_MATERIAL,
        "case_id": case_id,
        "generated": C.now_iso(),
        "sources": sources_out,
        "figures": kept_figures,
        "tables": kept_tables,
        "numbers": kept_numbers,
        "claims": kept_claims,
        "warnings": warnings,
        "stats": {"sources": len(sources_out), "pages": sum(len(s["pages"])
                                                            for s in sources_out),
                  "figures": len(kept_figures), "tables": len(kept_tables),
                  "numbers": len(kept_numbers), "claims": len(kept_claims),
                  "warnings": len(warnings)},
        "uncertainty": [],
    }
    return material


def build_arg_parser():
    parser = argparse.ArgumentParser(
        prog="ingest.py",
        description="Extract material.json (material/v1) from a case's real sources.")
    parser.add_argument("--case", required=True, help="path to case.json")
    parser.add_argument("--in", dest="in_path", default=None,
                        help="alias for --case (stage-chaining override)")
    parser.add_argument("--work", default=None, help="work directory (default: "
                        "<case dir>/work)")
    parser.add_argument("--input", default=None,
                        help="directory holding the case's source materials")
    parser.add_argument("--out", default=None,
                        help="material.json path (default <work>/material.json)")
    return parser


def main(argv=None):
    args = build_arg_parser().parse_args(argv)
    case_path = args.case or args.in_path
    work_dir = C.derive_work_dir(case_path, args.work)
    out_path = args.out or os.path.join(work_dir, "material.json")
    log = C.RunLog(os.path.join(work_dir, "ingest.log"))
    try:
        with log.stage("ingest", case=os.path.basename(case_path)):
            material = ingest(case_path, work_dir, input_dir=args.input,
                              out_path=out_path, log=log)
        C.safe_write_json(out_path, material)
    except SystemExit:
        raise
    except Exception as exc:
        log.error("ingest failed: %s: %s" % (type(exc).__name__, exc))
        C.safe_write_json(out_path, {
            "schema": C.SCHEMA_MATERIAL, "case_id": "unknown-case",
            "sources": [], "figures": [], "tables": [], "numbers": [], "claims": [],
            "warnings": ["ingest failed: %s: %s" % (type(exc).__name__, exc)],
            "stats": {}, "uncertainty": ["material extraction did not complete"],
        })
        print("error: ingest failed; a minimal material.json was written to %s"
              % out_path, file=_sys.stderr)
        return 2
    print(json.dumps({"material": out_path, "stats": material["stats"],
                      "warnings": len(material["warnings"])}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
