#!/usr/bin/env python3
"""common.py — shared, dependency-light helpers for the ARMI:2611.000074v1 PPT Skill.

This module is the ONLY cross-module import allowed by the pipeline contract
(``work/CONTRACT.md`` §3).  Its public API is deliberately small and stable; other
stages (``plan.py``, ``build_deck.py``, ``render.py``, ``evidence.py``,
``verify_deck.py``, ``result_contract.py``, ``selftest.py``) import it.  Do not
rename or re-sign existing names; add new ones instead.

Public API
----------
Schema identifiers
    ``SCHEMA_MATERIAL``, ``SCHEMA_PLAN``, ``SCHEMA_BUILD_REPORT``,
    ``SCHEMA_RENDER_REPORT``, ``SCHEMA_EVIDENCE_MAP``, ``SCHEMA_VERIFY_REPORT``,
    ``SCHEMA_RESULT`` — the exact ``"<name>/v1"`` strings from the contract.

Hashing / IO
    ``sha256_bytes(b)``, ``sha256_file(path)``, ``read_json(path, default=None)``,
    ``safe_write_json(path, obj)``, ``write_text(path, text)``,
    ``ensure_dir(path)``, ``relpath(path, start)``.

Case loading (the contract-driven part)
    ``load_case(case_path)`` -> dict
    ``extract_case_id(case, case_path)``
    ``get_output_contract(case)`` -> dict
    ``extract_constraints(case, case_path)`` -> dict with keys
        ``audience, duration_min, slide_count, slide_count_range, language,
        template_path, starting_deck_path, notes``
    ``discover_source_files(case, case_path, input_dir=None)`` -> ordered list of
        ``{"role", "path", "kind"}`` for every real file the case points at.
    ``find_task_card(case_path, input_dir=None)``, ``parse_task_card(text)``,
    ``extract_constraints_merged(case, case_path, input_dir=None)`` -> the
        ``case.json`` constraints merged with the ``task.md`` instruction card
        (which may be the ONLY place the fixed audience / duration / slide count
        are declared).  Returns ``(constraints, sources, task_path)`` where
        ``sources[field]`` records which document the value came from.

Key-synonym tolerant ``output_contract`` mapping
    ``normalize_key(key)``, ``canonical_role(key)``, ``CONTRACT_SYNONYMS``,
    ``map_contract_keys(contract, resolver)`` -> ``(filled, unmapped, roles)``.
    ``resolver`` is a mapping ``{role: value}`` or a callable ``role -> value``.

Locator / provenance helpers
    ``locator(source, page, quote, **extra)``, ``is_verbatim(quote, text)``,
    ``shorten(text, n)``, ``sentence_around(text, start, end)``.

Logging / timing
    ``RunLog(path)`` with ``.event(name, **fields)``, ``.info(msg)``,
    ``.warn(msg)``, ``.error(msg)``, ``.stage(name)`` (contextmanager, records
    ``duration_ms``), ``.records`` (list of dicts).

Environment
    ``try_import(name)`` -> module or ``None`` (never raises).
    Importing this module sets ``sys.dont_write_bytecode = True``.
"""
from __future__ import annotations

import sys as _sys

_sys.dont_write_bytecode = True

import contextlib
import hashlib
import importlib
import json
import os
import re
import time
import unicodedata
from datetime import datetime, timezone

# --------------------------------------------------------------------------- #
# Schema identifiers (CONTRACT.md §3 — frozen)
# --------------------------------------------------------------------------- #
SCHEMA_MATERIAL = "material/v1"
SCHEMA_PLAN = "plan/v1"
SCHEMA_BUILD_REPORT = "build_report/v1"
SCHEMA_RENDER_REPORT = "render_report/v1"
SCHEMA_EVIDENCE_MAP = "evidence_map/v1"
SCHEMA_VERIFY_REPORT = "verify_report/v1"
SCHEMA_RESULT = "result/v1"

#: Extensions ingest.py is expected to handle natively.
SOURCE_KINDS = {".pdf": "pdf", ".docx": "docx", ".md": "md", ".markdown": "md",
                ".txt": "txt", ".text": "txt", ".html": "html", ".htm": "html"}


# --------------------------------------------------------------------------- #
# Environment
# --------------------------------------------------------------------------- #
def try_import(name):
    """Import ``name``; return the module, or ``None`` if it is unavailable.

    Optional dependencies (``pptx``, ``pdfplumber``, ``pypdf``, ``docx``,
    ``matplotlib`` …) are always guarded through this helper so a missing package
    degrades honestly instead of crashing a stage.
    """
    try:
        return importlib.import_module(name)
    except Exception:
        return None


def now_iso():
    """UTC timestamp, second resolution, ISO-8601 with ``Z``."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def monotonic():
    return time.monotonic()


# --------------------------------------------------------------------------- #
# Hashing / paths / IO
# --------------------------------------------------------------------------- #
def ensure_dir(path):
    """Create ``path`` (and parents).  Never uses tempfile (broken on this host)."""
    if path:
        os.makedirs(path, exist_ok=True)
    return path


def sha256_bytes(data):
    return hashlib.sha256(data).hexdigest()


def sha256_file(path):
    """SHA-256 of a file, or ``""`` when it cannot be read."""
    h = hashlib.sha256()
    try:
        with open(path, "rb") as fh:
            for chunk in iter(lambda: fh.read(1024 * 1024), b""):
                h.update(chunk)
    except OSError:
        return ""
    return h.hexdigest()


def read_json(path, default=None):
    """Read JSON with UTF-8 (BOM tolerated).  Returns ``default`` on failure."""
    try:
        with open(path, "r", encoding="utf-8-sig") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return default


def safe_write_json(path, obj):
    """Atomically write ``obj`` as pretty, key-sorted, non-ASCII JSON.

    Returns ``path``.  No absolute author-machine path is ever injected here.
    """
    ensure_dir(os.path.dirname(os.path.abspath(path)) or ".")
    text = json.dumps(obj, ensure_ascii=False, indent=2, sort_keys=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(text)
    os.replace(tmp, path)
    return path


def write_text(path, text):
    ensure_dir(os.path.dirname(os.path.abspath(path)) or ".")
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(text)
    return path


def relpath(path, start):
    """Best-effort relative path; falls back to the original string."""
    if not path:
        return path
    try:
        return os.path.relpath(path, start)
    except ValueError:
        return path


def derive_work_dir(case_path, work_dir=None):
    """Choose the work directory for a case when ``--work`` was not given."""
    if work_dir:
        return os.path.abspath(work_dir)
    if case_path:
        return os.path.join(os.path.dirname(os.path.abspath(case_path)), "work")
    return os.path.abspath("work")


# --------------------------------------------------------------------------- #
# Case loading
# --------------------------------------------------------------------------- #
def load_case(case_path):
    """Load ``case.json``; raise ``SystemExit(2)`` when it is missing/unparseable."""
    if not case_path or not os.path.isfile(case_path):
        print("error: case file not found: %s" % (case_path,), file=_sys.stderr)
        raise SystemExit(2)
    case = read_json(case_path, None)
    if not isinstance(case, dict):
        print("error: case file is not a JSON object: %s" % (case_path,),
              file=_sys.stderr)
        raise SystemExit(2)
    return case


def extract_case_id(case, case_path=None):
    """``case_id`` under any of the plausible spellings; never invented."""
    for key in ("case_id", "caseId", "case", "id", "task_id", "name"):
        val = case.get(key)
        if isinstance(val, str) and val.strip():
            return val.strip()
    if case_path:
        return os.path.splitext(os.path.basename(case_path))[0]
    return "unknown-case"


def _lookup(obj, names, depth=0):
    """Recursive, case-insensitive lookup of the first present synonym."""
    if depth > 4 or not isinstance(obj, dict):
        return None
    lowered = {normalize_key(k): v for k, v in obj.items()}
    for name in names:
        key = normalize_key(name)
        if key in lowered and lowered[key] not in (None, "", [], {}):
            return lowered[key]
    for value in obj.values():
        if isinstance(value, dict):
            found = _lookup(value, names, depth + 1)
            if found is not None:
                return found
    return None


def get_output_contract(case):
    """The declared ``output_contract`` dict (empty dict when absent)."""
    value = _lookup(case, ["output_contract", "outputContract", "contract"])
    if isinstance(value, dict):
        return value
    return {}


_CONSTRAINT_NAMES = {
    "audience": ["audience", "target_audience", "listener", "受众", "object"],
    "duration_min": ["duration_min", "duration_minutes", "minutes", "talk_length",
                     "duration", "时长"],
    "slide_count": ["slide_count", "slides", "num_slides", "n_slides",
                    "slide_number", "page_count", "页数"],
    "slide_count_range": ["slide_count_range", "slide_range", "pages_range",
                          "slide_count_min_max"],
    "language": ["language", "lang", "locale", "语言"],
    "template_path": ["template_path", "template", "pptx_template", "master_template",
                      "模板"],
    "starting_deck_path": ["starting_deck_path", "starting_deck", "start_deck",
                           "seed_deck", "base_deck", "initial_deck", "起稿"],
    "sources": ["sources", "materials", "material", "inputs", "input_files",
                "files", "documents", "paper", "paper_path", "pdf", "docx"],
    "title": ["title", "deck_title", "presentation_title", "topic"],
    "tone": ["tone", "style", "register"],
}


def extract_constraints(case, case_path=None):
    """Tolerant extraction of the case constraints used by plan/build.

    Unknown constraints stay ``None`` — they are never guessed.  Paths are
    resolved relative to the case file (and to the case directory's ``input/``).
    """
    out = {}
    for field, names in _CONSTRAINT_NAMES.items():
        out[field] = _lookup(case, names)

    if out["slide_count_range"] is not None:
        rng = out["slide_count_range"]
        if isinstance(rng, dict):
            lo = rng.get("min", rng.get("lo", rng.get("from")))
            hi = rng.get("max", rng.get("hi", rng.get("to")))
            rng = [lo, hi]
        if isinstance(rng, (list, tuple)) and len(rng) == 2:
            try:
                out["slide_count_range"] = [int(rng[0]), int(rng[1])]
            except (TypeError, ValueError):
                out["slide_count_range"] = None
        else:
            out["slide_count_range"] = None

    if out["slide_count"] is not None:
        try:
            out["slide_count"] = int(out["slide_count"])
        except (TypeError, ValueError):
            out["slide_count"] = None

    if out["duration_min"] is not None:
        try:
            out["duration_min"] = float(out["duration_min"])
        except (TypeError, ValueError):
            out["duration_min"] = None

    for field in ("template_path", "starting_deck_path"):
        out[field] = _resolve_case_path(out[field], case_path)
    return out


def _resolve_case_path(value, case_path):
    """Resolve a path-ish value found in the case against the case's directories."""
    if not isinstance(value, str) or not value.strip():
        return None
    candidate = value.strip()
    if os.path.isabs(candidate):
        return candidate if os.path.exists(candidate) else candidate
    bases = []
    if case_path:
        cdir = os.path.dirname(os.path.abspath(case_path))
        bases += [cdir, os.path.dirname(cdir), os.path.join(cdir, "input"),
                  os.path.join(os.path.dirname(cdir), "input")]
    bases.append(os.getcwd())
    for base in bases:
        full = os.path.join(base, candidate)
        if os.path.exists(full):
            return os.path.abspath(full)
    return candidate


def _classify(path):
    return SOURCE_KINDS.get(os.path.splitext(path)[1].lower(), "other")


def discover_source_files(case, case_path, input_dir=None):
    """Return the real source files a case refers to, in a stable order.

    Sources come from, in priority order: explicit ``input_dir`` contents, path
    strings anywhere in the case JSON, and the case directory / its ``input/``
    sibling (the platform layout is ``/input/case.json`` + ``/input/**``).
    """
    found = {}  # abspath -> role

    def add(path, role):
        if not path or not isinstance(path, str):
            return
        full = _resolve_case_path(path, case_path)
        if not full or not os.path.isfile(full):
            return
        ext = os.path.splitext(full)[1].lower()
        if ext not in SOURCE_KINDS:
            return
        found.setdefault(os.path.abspath(full), role)

    roots = []
    if input_dir:
        roots.append(input_dir)
    if case_path:
        cdir = os.path.dirname(os.path.abspath(case_path))
        roots += [os.path.join(cdir, "input"), cdir]
    for root in roots:
        if os.path.isdir(root):
            for base, _dirs, files in os.walk(root):
                for name in sorted(files):
                    if os.path.splitext(name)[1].lower() in SOURCE_KINDS:
                        add(os.path.join(base, name), "input_dir")

    for key, value in _walk_scalars(case):
        if isinstance(value, str):
            add(value, "case." + key)

    result = []
    for path in sorted(found):
        result.append({"role": found[path], "path": path, "kind": _classify(path),
                       "name": os.path.basename(path)})
    return result


def _walk_scalars(obj, prefix="", depth=0):
    if depth > 5:
        return
    if isinstance(obj, dict):
        for key, value in obj.items():
            if isinstance(value, (dict, list)):
                yield from _walk_scalars(value, str(key), depth + 1)
            else:
                yield str(key), value
    elif isinstance(obj, list):
        for item in obj:
            if not isinstance(item, (dict, list)):
                yield prefix, item
            else:
                yield from _walk_scalars(item, prefix, depth + 1)


# --------------------------------------------------------------------------- #
# Task-card (task.md) constraints — the case may declare them ONLY here
# --------------------------------------------------------------------------- #
TASK_CARD_NAMES = ("task.md", "task.txt", "task.json", "instructions.md",
                   "instruction.md", "case.md", "readme.md")

_SLIDE_EXACT_RES = (
    re.compile(r"exactly\s+(\d{1,2})\s+(?:slides?|pages?)", re.I),
    re.compile(r"(\d{1,2})\s+(?:slides?|pages?)\s+(?:exactly|in\s+total)", re.I),
    re.compile(r"slide\s*count\s*[:：=]?\s*(\d{1,2})", re.I),
    re.compile(r"(\d{1,2})\s*(?:页|张|页幻灯片|页PPT)"),
)
_SLIDE_RANGE_RES = (
    re.compile(r"min\s*(\d{1,2})\s*,?\s*max\s*(\d{1,2})", re.I),
    re.compile(r"(\d{1,2})\s*(?:-|–|~|to)\s*(\d{1,2})\s+(?:slides?|pages?)", re.I),
    re.compile(r"(\d{1,2})\s*[到至]\s*(\d{1,2})\s*页"),
)
_DURATION_RES = (
    re.compile(r"duration\s*[:：=]?\s*(\d{1,3})\s*(?:minutes?|mins?|min)\b", re.I),
    re.compile(r"(\d{1,3})\s*(?:minutes?|mins?)\s+(?:of\s+)?speaking", re.I),
    re.compile(r"(\d{1,3})\s*分钟"),
)
_AUDIENCE_RES = (
    re.compile(r"^\s*[-*]?\s*audience\s*[:：]\s*(.+)$", re.I | re.M),
    re.compile(r"^\s*[-*]?\s*受众\s*[:：]\s*(.+)$", re.M),
    re.compile(r"^\s*[-*]?\s*target\s+audience\s*[:：]\s*(.+)$", re.I | re.M),
)
_LANGUAGE_RES = (
    re.compile(r"^\s*[-*]?\s*language\s*[:：]\s*([A-Za-z\u4e00-\u9fff ]{2,40})$",
               re.I | re.M),
    re.compile(r"^\s*[-*]?\s*语言\s*[:：]\s*(.+)$", re.M),
)
_TEMPLATE_RES = re.compile(r"template[^\n]*?[`\"']([^`\"'\n]+\.pptx)[`\"']", re.I)
_STARTING_DECK_RES = re.compile(
    r"starting\s+deck[^\n]*?[`\"']([^`\"'\n]+\.pptx)[`\"']", re.I)


def find_task_card(case_path, input_dir=None):
    """Locate the case's ``task.md`` instruction card, if there is one."""
    candidates = []
    if input_dir:
        candidates.append(input_dir)
    if case_path:
        case_dir = os.path.dirname(os.path.abspath(case_path))
        candidates += [case_dir, os.path.join(case_dir, "input"),
                       os.path.dirname(case_dir)]
    for base in candidates:
        if not base or not os.path.isdir(base):
            continue
        for name in TASK_CARD_NAMES:
            hit = os.path.join(base, name)
            if os.path.isfile(hit):
                return os.path.abspath(hit)
    return None


def parse_task_card(text, task_path=None):
    """Tolerant constraint extraction from a task card.  Never guesses.

    Returns ``{field: value}`` for the fields it can actually read; anything
    ambiguous is left out rather than defaulted.
    """
    out = {}
    if not text:
        return out
    for pattern in _SLIDE_RANGE_RES:
        match = pattern.search(text)
        if match:
            low, high = int(match.group(1)), int(match.group(2))
            if 1 <= low <= high <= 100:
                out["slide_count_range"] = [low, high]
                break
    if "slide_count_range" not in out:
        for pattern in _SLIDE_EXACT_RES:
            match = pattern.search(text)
            if match:
                value = int(match.group(1))
                if 1 <= value <= 100:
                    out["slide_count"] = value
                    break
    else:
        for pattern in _SLIDE_EXACT_RES:
            exact = pattern.search(text)
            if exact:
                out["slide_count"] = int(exact.group(1))
                break
    for pattern in _DURATION_RES:
        match = pattern.search(text)
        if match:
            out["duration_min"] = int(match.group(1))
            break
    for pattern in _AUDIENCE_RES:
        match = pattern.search(text)
        if match:
            out["audience"] = match.group(1).strip().rstrip(".")
            break
    for pattern in _LANGUAGE_RES:
        match = pattern.search(text)
        if match:
            value = match.group(1).strip()
            folded = value.lower()
            if folded.startswith("english"):
                out["language"] = "en"
            elif "中文" in value or folded.startswith("chinese"):
                out["language"] = "zh"
            break
    match = _TEMPLATE_RES.search(text)
    if match:
        out["template_path"] = match.group(1).strip()
    match = _STARTING_DECK_RES.search(text)
    if match:
        out["starting_deck_path"] = match.group(1).strip()
    if task_path:
        out["_task_card"] = task_path
    return out


def _values_agree(left, right):
    """True when two constraint values mean the same thing (10 == 10.0 == '10')."""
    if left == right:
        return True
    try:
        return float(left) == float(right)
    except (TypeError, ValueError):
        pass
    try:
        return [float(x) for x in left] == [float(x) for x in right]
    except (TypeError, ValueError):
        return str(left).strip() == str(right).strip()


def extract_constraints_merged(case, case_path, input_dir=None):
    """Merge ``case.json`` constraints with the ``task.md`` card.

    ``case.json`` is the structured source and wins where both speak; the task card
    fills the gaps (the platform may declare the fixed audience / duration / slide
    count only in ``task.md``).  Returns ``(constraints, sources, task_path)`` where
    ``sources[field]`` says which document supplied the value that is in use.
    """
    merged = extract_constraints(case, case_path)
    sources = {field: ("case.json" if merged.get(field) not in (None, "", [], {})
                       else None) for field in merged}
    task_path = find_task_card(case_path, input_dir=input_dir)
    task_values = {}
    if task_path:
        try:
            with open(task_path, "r", encoding="utf-8-sig", errors="replace") as fh:
                task_values = parse_task_card(fh.read(), task_path)
        except OSError:
            task_values = {}
    for field in ("slide_count", "slide_count_range", "duration_min", "audience",
                  "language", "template_path", "starting_deck_path"):
        if field not in task_values:
            continue
        value = task_values[field]
        if value in (None, "", [], {}):
            continue
        if field in ("template_path", "starting_deck_path") and case_path:
            value = _resolve_case_path(value, case_path)
        current = merged.get(field)
        if current in (None, "", [], {}):
            merged[field] = value
            sources[field] = "task.md"
        elif not _values_agree(current, value):
            sources[field] = "case.json (task.md says %r)" % (value,)
    merged["_sources"] = sources
    merged["_task_card"] = task_path
    return merged, sources, task_path


# --------------------------------------------------------------------------- #
# output_contract key mapping (tolerant; CONTRACT.md §3 "Key-name mapping")
# --------------------------------------------------------------------------- #
CONTRACT_SYNONYMS = {
    "deck_pptx": ["deck_pptx", "pptx", "presentation", "editable_deck",
                  "slide_file", "deck", "deck_path", "slides_path", "pptx_path"],
    "deck_pdf": ["deck_pdf", "pdf", "export", "exported_pdf", "pdf_path",
                 "presentation_pdf"],
    "slide_previews": ["slide_previews", "previews", "images", "png", "pngs",
                       "renders", "thumbnails", "preview_images", "slide_images"],
    "evidence_map": ["page_evidence_map", "evidence", "evidence_map", "source_map",
                     "evidence_mapping", "evidence_ledger"],
    "speaker_notes": ["speaker_notes", "notes", "script", "speaker_script",
                      "notes_artifact", "presenter_notes"],
    "plan": ["plan", "outline", "page_plan", "slide_plan", "deck_plan",
             "slide_outline"],
    "verify_report": ["verify_report", "verification", "self_check", "checks",
                      "validation_report", "verify"],
    "build_report": ["build_report", "build", "build_log"],
    "render_report": ["render_report", "render", "render_log"],
    "material": ["material", "materials", "material_json", "extraction",
                 "ingested_material"],
    "slide_count": ["slide_count", "num_slides", "n_slides", "page_count",
                    "slides_count"],
    "deck_title": ["deck_title", "title", "presentation_title", "topic",
                   "subject"],
    "summary": ["summary", "abstract", "overview", "tl_dr"],
    "uncertainty": ["uncertainty", "limitations", "caveats", "open_questions",
                    "risks"],
    "sources": ["sources", "source_files", "inputs", "provenance"],
    "status": ["status", "run_status", "state"],
    "change_records": ["change_records", "changes", "history", "revision_log",
                       "version_records", "changelog"],
    "case_id": ["case_id", "caseId", "case", "id"],
    "run_log": ["run_log", "log", "logfile", "execution_log"],
    "artifact_dir": ["artifact_dir", "output_dir", "artifacts_dir", "work_dir"],
}


def normalize_key(key):
    """``"Deck-PPTX "`` -> ``"deck_pptx"``; CJK is preserved, case-folded."""
    text = unicodedata.normalize("NFKC", str(key)).strip()
    text = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "_", text)
    return text.strip("_").lower()


_ROLE_BY_KEY = None


def _role_index():
    global _ROLE_BY_KEY
    if _ROLE_BY_KEY is None:
        index = {}
        for role, names in CONTRACT_SYNONYMS.items():
            index[normalize_key(role)] = role
            for name in names:
                index.setdefault(normalize_key(name), role)
        _ROLE_BY_KEY = index
    return _ROLE_BY_KEY


def canonical_role(key):
    """Map one contract key to a canonical role, or ``None`` when unknown."""
    return _role_index().get(normalize_key(key))


def map_contract_keys(contract, resolver):
    """Fill every declared ``output_contract`` key from canonical roles.

    ``resolver`` is either a ``{role: value}`` mapping or a callable
    ``role -> value``.  Returns ``(filled, unmapped, roles)`` where

    * ``filled``   — ``{declared_key: value}``; unmapped/undeliverable keys get
      ``None`` (never dropped: dropping a declared key is a contract defect);
    * ``unmapped`` — declared keys with no exact/synonym role (also used by
      ``result_contract.py`` for ``_unmapped_contract_keys``);
    * ``roles``    — ``{declared_key: role|None}``.
    """
    filled, unmapped, roles = {}, [], {}
    if not isinstance(contract, dict):
        return filled, unmapped, roles
    for key in contract:
        role = canonical_role(key)
        roles[key] = role
        if role is None:
            filled[key] = None
            unmapped.append(key)
            continue
        try:
            value = resolver(role) if callable(resolver) else resolver.get(role)
        except Exception:
            value = None
        filled[key] = value
    return filled, unmapped, roles


# --------------------------------------------------------------------------- #
# Numeric-token guard (shared by plan/build/verify — V04)
# --------------------------------------------------------------------------- #
#: A numeric literal as it appears in prose.  Never rounded, never rewritten.
NUMBER_TOKEN_RE = re.compile(
    r"(?<![\w.])(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)(?![\w])")


def numeric_tokens(text):
    """Every numeric literal in ``text``, exactly as printed (strings).

    ``"19.2 % to 22.7 % (n=240)"`` -> ``["19.2", "22.7", "240"]``.  A token that
    is part of an identifier (``C60``, ``p53``) is not a numeric token.
    """
    if not text:
        return []
    return [m.group(1) for m in NUMBER_TOKEN_RE.finditer(str(text))]


def normalize_number_raw(raw):
    """Comparison key for a numeric literal: ``"1,200"`` -> ``"1200.0"``."""
    if raw is None:
        return None
    try:
        return repr(float(str(raw).replace(",", "").strip()))
    except ValueError:
        return None


def numbers_covered(text, allowed_raws):
    """Return ``(ok, unmatched)`` for every numeric literal in ``text``.

    ``allowed_raws`` is an iterable of ``raw`` strings from
    ``material.json["numbers"]``.  A literal is covered when its numeric value is
    present (so ``"19.20 %"`` is covered by raw ``"19.2"``, but ``"20 %"`` is not
    covered by ``"19.2"``).  This is the mechanical basis of V04.
    """
    allowed = set()
    for raw in allowed_raws or ():
        key = normalize_number_raw(raw)
        if key is not None:
            allowed.add(key)
    unmatched = []
    for token in numeric_tokens(text):
        if normalize_number_raw(token) not in allowed:
            unmatched.append(token)
    return (not unmatched), unmatched


# --------------------------------------------------------------------------- #
# Provenance helpers — the "verbatim locator" rule
# --------------------------------------------------------------------------- #
def locator(source, page, quote, **extra):
    """Build a locator dict.  ``quote`` MUST be verbatim source text."""
    out = {"source": source, "page": page, "quote": quote}
    out.update(extra)
    return out


def is_verbatim(quote, text):
    """True iff ``quote`` occurs literally in ``text`` (after newline folding).

    Whitespace-only differences are tolerated *only* as runs of whitespace being
    collapsed on both sides — the character sequence itself must be present.
    """
    if not quote or not text:
        return False
    if quote in text:
        return True
    fold = lambda s: re.sub(r"\s+", " ", s)
    return fold(quote) in fold(text)


def shorten(text, n=160):
    """One-line trimming for ``context``/``detail`` fields."""
    if text is None:
        return ""
    flat = re.sub(r"\s+", " ", str(text)).strip()
    return flat if len(flat) <= n else flat[: n - 1].rstrip() + "\u2026"


_SENT_SPLIT = re.compile(r"(?<=[.!?\u3002\uff01\uff1f;])\s+")


def sentence_around(text, start, end, pad=0):
    """Verbatim sentence enclosing ``text[start:end]`` (``""`` when absent)."""
    if not text or start is None:
        return ""
    start = max(0, start)
    end = min(len(text), end if end is not None else start)
    left = 0
    for m in _SENT_SPLIT.finditer(text):
        if m.end() <= start:
            left = m.end()
        else:
            break
    right = len(text)
    for m in _SENT_SPLIT.finditer(text):
        if m.start() >= end:
            right = m.start()
            break
    quote = text[left:right].strip()
    # collapse internal newlines only; never rewrite words
    return re.sub(r"\s+", " ", quote)


# --------------------------------------------------------------------------- #
# Logging / timing
# --------------------------------------------------------------------------- #
class RunLog:
    """Append-only JSONL + human-readable run log.

    ``run.py`` points it at ``<out>/run.log``.  Every stage records its own
    ``duration_ms`` so the reviewer can attribute time per stage.
    """

    def __init__(self, path, echo=True):
        self.path = path
        self.echo = echo
        self.records = []
        self._stage = {}
        if path:
            ensure_dir(os.path.dirname(os.path.abspath(path)) or ".")
            with open(path, "a", encoding="utf-8", newline="\n") as fh:
                fh.write("# run.log %s\n" % now_iso())

    def _write(self, record):
        self.records.append(record)
        if self.path:
            with open(self.path, "a", encoding="utf-8", newline="\n") as fh:
                fh.write(json.dumps(record, ensure_ascii=False, sort_keys=True) + "\n")
        if self.echo:
            level = record.get("level", "info").upper()
            print("[%s] %s %s" % (level, record.get("event", "-"),
                                  record.get("message", "")), file=_sys.stderr)

    def event(self, name, level="info", **fields):
        record = {"ts": now_iso(), "event": name, "level": level}
        record.update(fields)
        self._write(record)
        return record

    def info(self, message, **fields):
        return self.event("info", level="info", message=message, **fields)

    def warn(self, message, **fields):
        return self.event("warn", level="warn", message=message, **fields)

    def error(self, message, **fields):
        return self.event("error", level="error", message=message, **fields)

    @contextlib.contextmanager
    def stage(self, name, **fields):
        """Context manager recording ``<name>`` start/end with ``duration_ms``."""
        t0 = monotonic()
        self.event("stage_start", stage=name, **fields)
        record = {"ok": True, "error": None}
        try:
            yield record
        except Exception as exc:  # re-raised after logging
            self.event("stage_end", level="error", stage=name,
                       duration_ms=int((monotonic() - t0) * 1000),
                       ok=False, error="%s: %s" % (type(exc).__name__, exc))
            raise
        self.event("stage_end", stage=name, duration_ms=int((monotonic() - t0) * 1000),
                   ok=bool(record.get("ok", True)), error=record.get("error"))
        return


def touch_file(path):
    """Create/refresh an empty marker file (used for honest partial writes)."""
    ensure_dir(os.path.dirname(os.path.abspath(path)) or ".")
    with open(path, "a", encoding="utf-8"):
        pass
    return path


if __name__ == "__main__":  # tiny self-check, no side effects
    print("common.py ok — schemas:", SCHEMA_MATERIAL, SCHEMA_PLAN)
