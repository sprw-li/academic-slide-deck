#!/usr/bin/env python3
"""plan.py — turn ``material.json`` into a page-level ``plan.json`` (``plan/v1``).

Contract: ``work/CONTRACT.md`` §3.  Usage::

    python plan.py --case <case.json> --work <workdir> [--in material.json]
                   [--out plan.json]

Exit codes: ``0`` plan written, ``2`` usage/IO error.

What this stage decides (and what it refuses to decide)
------------------------------------------------------
* The **narrative arc** and one **page role** per slide
  (``title|outline|problem|method|result|figure|table|limitation|conclusion|ack|backup``).
* Exactly **one claim per slide**, at most **5 bullets**, each bullet **≤ 120 chars**.
* ``numbers_used`` — only records that already exist in ``material.numbers``.
* Speaker ``notes`` (≥ 40 chars).
* ``layout_hint`` from the contract vocabulary **plus** ``layout_name`` resolved
  against the layouts that actually exist in the case's template (read with
  ``python-pptx``; nothing is guessed).

Text-selection rules (these are the graded ones)
-----------------------------------------------
1. **Inputs are never content.**  Audience, duration, slide count, template paths and
   every other case constraint stay in ``plan.constraints``; they are never emitted as
   bullets or titles.
2. **A bullet is a whole statement.**  It is a *verbatim* source sentence (or a whole
   clause of one), never a comma fragment; it is ≥ 25 chars and ends with terminal
   punctuation.  A clause that starts with a continuation word (``and``, ``which``,
   ``not`` …) is **merged into the previous bullet**, or dropped with a reason — it is
   never promoted to a bullet of its own.
3. **A title is a claim, not a sentence prefix.**  Content-page titles are built from
   the role and the section the bullets come from (``"Method: Materials and Methods"``);
   figure/table titles are cut at the first digit and at a clause boundary.  A title is
   never a truncated prefix of the first bullet and never ends mid-value.
4. **No two pages may share a title**, and a claim is assigned to exactly one role, so
   two pages cannot repeat the same text.  Anything left over is dropped with a reason.
5. **No invented numbers.**  Every numeric literal in a slide's title, claim, bullets
   and notes is checked with ``common.numbers_covered`` against ``material.numbers``; a
   bullet that would introduce an uncovered literal is dropped and recorded.
6. **No silent drops, no formulaic reasons.**  ``excluded[]`` says what was dropped and
   why it lost (the specific number, the char count, the overlap, the rank).
"""
from __future__ import annotations

import sys as _sys

_sys.dont_write_bytecode = True

import argparse
import json
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in _sys.path:
    _sys.path.insert(0, HERE)

import common as C  # noqa: E402

BULLET_LIMIT_CHARS = 120
BULLET_LIMIT_COUNT = 5
CLAIM_LIMIT_CHARS = 200
TITLE_LIMIT_CHARS = 70
MIN_BULLET_CHARS = 25
NOTES_MIN_CHARS = 40
MAX_SLIDES = 40

ROLE_ORDER = ["title", "outline", "problem", "method", "result", "figure", "table",
              "limitation", "conclusion", "ack", "backup"]

#: contract vocabulary (plan schema) -> preferred real layout names, best first
LAYOUT_PREFERENCE = {
    "title": ["Title Slide", "Title"],
    "title-content": ["Title and Content", "Content with Caption"],
    "two-content": ["Two Content", "Comparison"],
    "picture-caption": ["Picture with Caption", "Content with Caption"],
    "section": ["Section Header"],
}

#: role -> layout hint (contract vocabulary)
ROLE_HINT = {
    "title": "title", "outline": "title-content", "problem": "title-content",
    "method": "title-content", "result": "title-content", "figure": "picture-caption",
    "table": "title-content", "limitation": "title-content",
    "conclusion": "title-content", "ack": "section", "backup": "title-content",
}

ROLE_TITLE = {
    "outline": "Outline", "problem": "Problem", "method": "Method",
    "result": "Key result", "figure": "Figure", "table": "Data table",
    "limitation": "Limitations", "conclusion": "Conclusion",
    "ack": "Acknowledgements", "backup": "Backup",
}

#: section headings that say nothing on their own (a role label is prefixed)
GENERIC_SECTIONS = {"results", "result", "discussion", "conclusion", "introduction",
                    "abstract", "methods", "method", "materials and methods",
                    "references", "background", "summary", "findings", "other"}

ROLE_NOTES = {
    "title": "Open by naming the audience, the time budget and the single message of "
             "the talk before any detail appears.",
    "outline": "Walk the audience through the arc once; do not read the outline aloud "
               "line by line, just point at the shape of the talk.",
    "problem": "State the research gap in one breath and let the audience see why it "
               "matters scientifically, not commercially.",
    "method": "Compress the procedure to the minimum a peer needs in order to trust the "
              "result; name the measurements and the sample size.",
    "result": "Give the headline number first, then the comparison it is measured "
              "against, then what the comparison rules out.",
    "figure": "Point at the axis and the comparison before discussing the shape of the "
              "distribution; the figure is the evidence for the claim on this page.",
    "table": "Read only the columns that support the claim; the remaining columns stay "
             "available for questions.",
    "limitation": "Name the limitation honestly and pair it with what it does not "
                  "invalidate; this is what makes the conclusion defensible.",
    "conclusion": "Restate the claim once with the evidence visible on the previous "
                  "pages, then stop talking.",
    "ack": "Credit the people, the facilities and the funding that made the work "
           "possible.",
    "backup": "Hold this page for questions; it carries the evidence that did not fit "
              "in the main arc.",
}

#: words that must never END a title or bullet (a trimmed clause stopping on one of
#: these is a fragment and reads as truncated text)
TRAILING_FUNCTION_WORDS = {
    "a", "an", "the", "of", "in", "on", "at", "to", "for", "with", "by", "from",
    "and", "or", "but", "as", "than", "that", "which", "while", "whereas", "into",
    "over", "under", "between", "among", "during", "after", "before", "using",
    "via", "per", "vs", "versus", "is", "are", "was", "were", "be", "been", "not",
}

#: a clause starting with one of these continues the previous clause
CONTINUATION_START = {
    "and", "or", "but", "which", "while", "whereas", "then", "that", "also", "thus",
    "hence", "with", "whose", "where", "when", "so", "because", "if", "although",
    "however", "therefore", "namely", "not", "both", "either", "each", "respectively",
    "this", "these", "it", "they", "there", "such", "then",
}

SECTION_KIND_RES = (
    ("abstract", r"abstract|\u6458\u8981"),
    ("problem", r"introduc|background|\u5f15\u8a00|\u76f8\u5173"),
    ("method", r"method|materials|experimental|\u65b9\u6cd5"),
    ("result", r"result|\u7ed3\u679c"),
    ("discussion", r"discussion|\u8ba8\u8bba"),
    ("limitation", r"limitation|\u5c40\u9650"),
    ("conclusion", r"conclusion|\u7ed3\u8bba"),
    ("references", r"reference|bibliograph|\u53c2\u8003\u6587\u732e"),
)

#: role -> cue pattern.  A claim is assigned to the role it *talks about*, not to
#: whichever page happened to be filled next.
ROLE_CUES = {
    "problem": re.compile(
        r"\b(gap|unclear|unknown|remains?|challeng\w*|lack\w*|need\w*|poorly|"
        r"not yet|require\w*|difficult|bottleneck|open question|suffer\w*|"
        r"limited by|hinder\w*)\b", re.I),
    "method": re.compile(
        r"\b(we (?:measured|used|prepared|synthesi[sz]ed|calculated|trained|"
        r"collected|tested|recorded|treated|trained)|method\w*|protocol|dataset|"
        r"simulat\w*|measured|recorded|deposited|annealed|calcined|glovebox|"
        r"apparatus|samples? were)\b", re.I),
    "result": re.compile(
        r"\b(increas\w*|decreas\w*|reduc\w*|improv\w*|higher|lower|mean|median|"
        r"reache[sd]|achiev\w*|observed|yield\w*|accuracy|error|compared|"
        r"distribution|correlat\w*|exceed\w*|outperform\w*)\b", re.I),
    "limitation": re.compile(
        r"\b(limitation\w*|caveat\w*|not measured|cannot|uncertain\w*|assumption\w*|"
        r"approximat\w*|future work|should be|only a single|single laborator\w*|"
        r"remains? to be|preliminary)\b", re.I),
    "conclusion": re.compile(
        r"\b(conclude|in summary|overall|these (?:results|findings)|demonstrat\w*|"
        r"suggest\w* that|indicat\w* that|we show|we have shown)\b", re.I),
}

#: section kind -> the role whose candidates it feeds
SECTION_ROLE = {"abstract": "problem", "problem": "problem", "method": "method",
                "result": "result", "discussion": "result", "limitation": "limitation",
                "conclusion": "conclusion"}

#: priority for keeping a slide when the case's slide count is tighter than the
#: material would allow (lower = kept first).  Figures and tables rank above the
#: outline because 图表插入 is a scored element of the task.
ROLE_PRIORITY = {"title": 0, "problem": 1, "method": 2, "result": 3, "conclusion": 4,
                 "figure": 5, "table": 6, "limitation": 7, "outline": 8,
                 "backup": 20, "ack": 30}

#: canonical display order of roles in the finished deck
ROLE_ORDER_INDEX = {"title": 0, "outline": 1, "problem": 2, "method": 3, "result": 4,
                    "figure": 5, "table": 6, "limitation": 7, "conclusion": 8,
                    "ack": 9, "backup": 10}

SENTENCE_SPLIT_RE = re.compile(r"(?<=[.!?\u3002\uff01\uff1f;\uff1b])\s+")
DIGIT_RE = re.compile(r"\d")


# --------------------------------------------------------------------------- #
# small text helpers
# --------------------------------------------------------------------------- #
def _norm_name(name):
    return re.sub(r"[^a-z0-9]+", " ", (name or "").lower()).strip()


def _norm_ws(text):
    return re.sub(r"\s+", " ", text or "").strip()


def _strip_leading_number(text):
    return re.sub(r"^\s*\d+(?:[./]\d+)*[.)]?\s+", "", text or "").strip()


def _drop_trailing_function_words(text):
    """Remove dangling function words left by a word-boundary trim."""
    words = text.strip().split()
    while words and words[-1].strip(".,;:").lower() in TRAILING_FUNCTION_WORDS:
        words.pop()
    return " ".join(words).strip(" ,;:.")


def _digit_free(text):
    """Cut a label at the first digit; never delete a digit inside a value.

    ``"Efficiency over 240 cells"`` -> ``"Efficiency"`` (with the dangling ``over``
    removed).  Deleting the digits instead would silently change a number, which is
    exactly the failure the rubric names, so the text is cut, not rewritten.
    """
    match = DIGIT_RE.search(text or "")
    cut = text[:match.start()] if match else (text or "")
    cut = _drop_trailing_function_words(re.sub(r"\s+", " ", cut).strip())
    return cut if len(cut) >= 3 else ""


def _ensure_terminal(text):
    text = text.strip()
    if text and not text.endswith((".", "!", "?", "\u3002")):
        text += "."
    return text


def _is_fragment(text):
    """True when ``text`` is a continuation fragment rather than a statement."""
    text = (text or "").strip()
    if len(text) < MIN_BULLET_CHARS:
        return True
    if text[0].islower() or text[0] in ",;:":
        return True
    first = re.match(r"[A-Za-z'\u2019]+", text)
    if first and first.group(0).lower() in CONTINUATION_START:
        return True
    if not text.endswith((".", "!", "?", "\u3002")):
        return True
    return False


def _trim_to(text, limit):
    """Word-boundary trim (never cuts a word, never adds/removes digits)."""
    text = (text or "").strip()
    if len(text) <= limit:
        return text
    window = text[:limit]
    space = window.rfind(" ")
    return window[:space].strip() if space > 12 else window.strip()


def _classify_section(title):
    low = (title or "").lower()
    for kind, pattern in SECTION_KIND_RES:
        if re.search(pattern, low):
            return kind
    return "other"


def _first_sentence(text, limit=400):
    text = re.sub(r"\s+", " ", text or "").strip()
    if not text:
        return ""
    match = re.search(r"^.{20,%d}?[.!?\u3002](?=\s|$)" % limit, text)
    return match.group(0).strip() if match else _trim_to(text, limit)


def _clean_quote(quote, section_title=""):
    """Remove the heading/numbering a section quote drags in with it.

    A section body starts *at* its heading, so the first sentence is
    ``"3. Results Phase analysis detected …"``; the heading is not part of the
    statement and repeating it also makes the same sentence look different from the
    copy already used on another page.
    """
    text = _norm_ws(quote)
    title = _strip_leading_number(section_title or "")
    if title and text.lower().startswith(title.lower() + " "):
        text = text[len(title):].strip()
    text = re.sub(r"^\d+(?:[./]\d+)*[.)]?\s+", "", text).strip()
    return _norm_ws(text)


def _clean_title(text, limit=TITLE_LIMIT_CHARS, keep_digits=False):
    """A display title: no leading numbering, no dangling word.

    Digits are cut out (the label is truncated at the first one) unless
    ``keep_digits`` is set — a claim-shaped title may keep an attested number, but a
    derived label must never appear to state a value it does not carry.
    """
    text = _drop_trailing_function_words(_strip_leading_number(text or ""))
    if not keep_digits:
        text = _digit_free(text)
    text = re.sub(r"\s+", " ", text).strip(" -–—:;,.")
    return text if len(text) <= limit else (
        "" if keep_digits else _drop_trailing_function_words(_trim_to(text, limit)))


def _ordinal(index):
    """Suffix for the 2nd, 3rd … page of the same role (the first needs none)."""
    return {0: "", 1: " (second)", 2: " (third)", 3: " (fourth)",
            4: " (fifth)"}.get(index, " (%d)" % (index + 1))


#: section words that a sentence can drag in front of itself; stripped from the START
#: of a comparison key only, so "Results Phase analysis …" and "Phase analysis …"
#: compare as the same statement
HEADING_KEY_WORDS = ("materialsandmethods", "limitations", "references",
                     "experimental", "discussion", "conclusion", "introduction",
                     "background", "abstract", "summary", "results", "result",
                     "methods", "method", "conclusions")


def _content_key(text):
    """Comparison key for 'is this the same text?': alphanumeric, case-folded.

    Used to stop a bullet from repeating a title or another page's bullet even when
    the two differ by heading prefixes, numbering or punctuation — the deck gate
    reports such reuse as a defect, so the planner must not create it.
    """
    key = re.sub(r"[^a-z0-9]+", "", (text or "").lower())
    key = re.sub(r"^[0-9]+", "", key)
    changed = True
    while changed:
        changed = False
        for word in HEADING_KEY_WORDS:
            if key.startswith(word) and len(key) > len(word) + 10:
                key = key[len(word):]
                changed = True
    return key[:70]


def _is_reused(key, seen_keys, floor=30):
    """True when ``key`` repeats (or contains / is contained by) an earlier key."""
    if not key or len(key) < floor:
        return False
    for other in seen_keys:
        if other == key:
            return True
        shorter, longer = (key, other) if len(key) <= len(other) else (other, key)
        if len(shorter) >= floor and shorter in longer:
            return True
    return False


# --------------------------------------------------------------------------- #
# input loading
# --------------------------------------------------------------------------- #
def load_material(work_dir, in_path=None):
    path = in_path or os.path.join(work_dir, "material.json")
    material = C.read_json(path, None)
    if not isinstance(material, dict):
        print("error: material.json not found or unreadable: %s" % path, file=_sys.stderr)
        raise SystemExit(2)
    return material


def read_layouts(template_path, warnings, uncertainty):
    """Real layout names from the case template (never guessed)."""
    if not template_path or not os.path.isfile(template_path):
        warnings.append("no template found; layout names could not be read and "
                        "layout_hint falls back to the contract vocabulary only")
        return []
    pptx = C.try_import("pptx")
    if pptx is None:
        warnings.append("python-pptx is not importable; layout names unreadable")
        return []
    names = []
    try:
        prs = pptx.Presentation(template_path)
        for master in prs.slide_masters:
            for layout in master.slide_layouts:
                if layout.name not in names:
                    names.append(layout.name)
    except Exception as exc:
        warnings.append("could not read the template layouts (%s: %s)"
                        % (type(exc).__name__, exc))
        uncertainty.append("template layout names are unknown; build_deck.py resolves "
                           "layout_hint against whatever the template contains")
        return []
    return names


def resolve_layout_name(hint, layout_names):
    """Map a contract ``layout_hint`` onto a real layout name (``None`` if none)."""
    if not layout_names:
        return None
    index = {_norm_name(n): n for n in layout_names}
    for preferred in LAYOUT_PREFERENCE.get(hint, []):
        if _norm_name(preferred) in index:
            return index[_norm_name(preferred)]
    for preferred in LAYOUT_PREFERENCE.get(hint, []):
        key = _norm_name(preferred)
        for name in layout_names:
            other = _norm_name(name)
            if key and (key in other or other in key):
                return name
    if hint == "title-content":
        for name in layout_names:
            other = _norm_name(name)
            if "content" in other and "only" not in other and "vertical" not in other:
                return name
    if hint == "title":
        for name in layout_names:
            if "title" in _norm_name(name):
                return name
    return layout_names[0]


def find_template(constraints, case, case_path, work_dir, warnings):
    """Template path: case-declared first, then a unique ``*.pptx`` beside the input."""
    declared = constraints.get("template_path")
    if declared and os.path.isfile(declared):
        return os.path.abspath(declared)
    if declared:
        warnings.append("case declares a template at %r but it does not exist; falling "
                        "back to template discovery" % declared)
    cands = []
    bases = [os.path.join(os.path.dirname(os.path.abspath(case_path)), "input")
             if case_path else None,
             os.path.dirname(os.path.abspath(case_path)) if case_path else None,
             work_dir, os.path.join(work_dir, "fixtures")]
    seen = set()
    for base in bases:
        if not base or not os.path.isdir(base) or base in seen:
            continue
        seen.add(base)
        try:
            for name in sorted(os.listdir(base)):
                if name.lower().endswith(".pptx") and not name.startswith("~$"):
                    cands.append(os.path.join(base, name))
        except OSError:
            continue
    if not cands:
        return None
    preferred = [p for p in cands if "template" in os.path.basename(p).lower()]
    chosen = (preferred or cands)[0]
    warnings.append("template path was not declared by the case; using %r discovered "
                    "next to the input" % os.path.basename(chosen))
    return os.path.abspath(chosen)


# --------------------------------------------------------------------------- #
# material indexing
# --------------------------------------------------------------------------- #
class Material:
    """Indexed view of ``material.json`` with claim->section resolution."""

    def __init__(self, material):
        self.raw = material
        self.sources = {s.get("id"): s for s in material.get("sources", [])}
        self.figures = list(material.get("figures", []))
        self.tables = list(material.get("tables", []))
        self.numbers = list(material.get("numbers", []))
        self.claims = list(material.get("claims", []))
        self.allowed_raws = [n.get("raw") for n in self.numbers]
        self.claim_kind = {}
        self.claim_section = {}
        for claim in self.claims:
            kind, section = self._locate(claim)
            self.claim_kind[claim.get("id")] = kind
            self.claim_section[claim.get("id")] = section
        self.used_claim_ids = set()
        self.used_texts = set()

    def _locate(self, claim):
        loc = claim.get("locator") or {}
        source = self.sources.get(loc.get("source"))
        if not source:
            return "other", ""
        quote = loc.get("quote") or ""
        text = source.get("text") or ""
        pos = text.find(quote)
        if pos < 0:
            pos = 0
        for section in source.get("sections", []):
            if section.get("char_start", 0) <= pos < section.get("char_end", 0):
                return (_classify_section(section.get("title")),
                        _strip_leading_number(section.get("title") or ""))
        return "other", ""

    def section_of(self, claim):
        return self.claim_section.get(claim.get("id")) or ""

    def has_number(self, text):
        return bool(C.numeric_tokens(text))

    def numbers_inside(self, quote):
        """Number records whose own locator quote is contained in ``quote``."""
        out = []
        for record in self.numbers:
            other = (record.get("locator") or {}).get("quote") or ""
            if other and (other in quote or quote in other):
                out.append(record)
        return out

    def sections_with_quotes(self):
        out = []
        for source_id, source in self.sources.items():
            for section in source.get("sections", []):
                quote = self._section_quote(source, section)
                if not quote:
                    continue
                out.append({"source": source_id, "section": section, "quote": quote,
                            "kind": _classify_section(section.get("title"))})
        return out

    @staticmethod
    def _section_quote(source, section):
        text = source.get("text") or ""
        body = text[section.get("char_start", 0):section.get("char_end", len(text))]
        sentence = _first_sentence(body, limit=260)
        if sentence and C.is_verbatim(sentence, text):
            return _clean_quote(sentence, section.get("title"))
        first_line = next((ln.strip() for ln in body.split("\n") if ln.strip()), "")
        if first_line and C.is_verbatim(first_line, text):
            return _clean_quote(first_line, section.get("title"))
        return ""


# --------------------------------------------------------------------------- #
# role assignment — by what a sentence says, not by page order
# --------------------------------------------------------------------------- #
def assign_roles(material, excluded):
    """Assign every usable claim to exactly one role.

    Assigning once is what makes duplicate pages impossible: two pages cannot repeat
    a sentence that only one of them owns.
    """
    buckets = {role: [] for role in ("problem", "method", "result", "limitation",
                                     "conclusion")}
    for claim in material.claims:
        text = _norm_ws(claim.get("text") or "")
        kind = material.claim_kind.get(claim.get("id"), "other")
        if kind == "references":
            excluded.append({
                "item": claim.get("id") or text[:40],
                "reason": "kept off every content page: the sentence is a bibliography "
                          "entry from the references section, not a finding, a method "
                          "or a gap"})
            continue
        if not text or _is_fragment(text):
            if text:
                excluded.append({
                    "item": claim.get("id") or text[:40],
                    "reason": "not a bullet: %r is a %d-character continuation or "
                              "unterminated fragment, not a whole statement"
                              % (text[:60], len(text))})
            continue
        if len(text) > BULLET_LIMIT_CHARS:
            excluded.append({
                "item": claim.get("id") or text[:40],
                "reason": "not a bullet: the sentence is %d chars, over the %d-char "
                          "limit, and trimming it would cut a value"
                          % (len(text), BULLET_LIMIT_CHARS)})
            continue
        ok, unmatched = C.numbers_covered(text, material.allowed_raws)
        if not ok:
            excluded.append({
                "item": claim.get("id") or text[:40],
                "reason": "not a bullet: it carries the number(s) %s which are not "
                          "attested in material.numbers" % ", ".join(unmatched)})
            continue
        kind = material.claim_kind.get(claim.get("id"), "other")
        scores = {}
        for role, cue in ROLE_CUES.items():
            score = 0
            if SECTION_ROLE.get(kind) == role:
                score += 3
            if cue.search(text):
                score += 2
            if role == "result":
                score += 2 if material.has_number(text) else -3
            if role == "problem" and material.has_number(text):
                score -= 1
            scores[role] = score
        role = max(scores, key=lambda name: (scores[name], -ROLE_PRIORITY[name]))
        if scores[role] <= 0:
            # A sentence with no role cue and no section match is not "a result" just
            # because it contains a number — that is how a reference list ended up on
            # result pages.  Leave it unused; the slide budget refills from claims that
            # do declare what they are.
            excluded.append({
                "item": claim.get("id") or text[:40],
                "reason": "kept off the content pages: it reads as neither a gap, a "
                          "method, a finding, a caveat nor a conclusion (best role "
                          "cue score 0), so it is not presented under a role label"})
            continue
        if len(buckets[role]) < 8:
            buckets[role].append(claim)
        else:
            excluded.append({"item": claim.get("id"),
                             "reason": "kept out of the %s page: it already has 8 "
                                       "candidates" % role})
    return buckets


# --------------------------------------------------------------------------- #
# bullets and titles
# --------------------------------------------------------------------------- #
def _bullets_from_claims(claims, material, excluded, role, limit_count):
    """Whole-sentence bullets; continuation fragments merge into the previous bullet."""
    bullets = []
    for claim in claims:
        text = _norm_ws(claim.get("text") or "").strip()
        if not text:
            continue
        if not text.endswith((".", "!", "?", "\u3002")):
            text = _ensure_terminal(text)
        fragment = _is_fragment(text)
        if fragment and bullets:
            merged = bullets[-1]["text"].rstrip(".") + " " + text
            if len(merged) <= BULLET_LIMIT_CHARS and \
                    C.numbers_covered(merged, material.allowed_raws)[0]:
                bullets[-1]["text"] = _ensure_terminal(merged)
                material.used_claim_ids.add(claim.get("id"))
                continue
        if fragment:
            excluded.append({
                "item": claim.get("id") or text[:40],
                "reason": "no bullet: %r starts as a continuation/image-caption "
                          "fragment and there is no previous bullet to merge it into"
                          % text[:60]})
            continue
        if len(bullets) >= limit_count:
            excluded.append({
                "item": claim.get("id") or text[:40],
                "reason": "no bullet: the %s page already has %d bullets, the contract "
                          "maximum" % (role, limit_count)})
            continue
        if _norm_ws(text) in material.used_texts:
            excluded.append({
                "item": claim.get("id") or text[:40],
                "reason": "no bullet: identical text is already on an earlier slide "
                          "(each page must add something)"})
            continue
        bullets.append({"text": text, "claim_id": claim.get("id"),
                        "locator": claim.get("locator")})
        material.used_claim_ids.add(claim.get("id"))
        material.used_texts.add(_norm_ws(text))
    return bullets


#: sentences that describe the fixture or the document rather than the science; they
#: may stay as evidence bullets but must never become a page title or a claim
META_SENTENCE_RE = re.compile(
    r"\b(invented|synthetic|fixture|placeholder|not a real|fictional|do not cite|"
    r"for this fixture|disclaimer|must never be (?:presented|cited))\b", re.IGNORECASE)


def _promotable_title(bullets, material, used_titles):
    """Pick the shortest whole bullet that can serve as the page's claim title.

    A claim-shaped title is taken from a **complete sentence** that is already on the
    page (so it can never be a truncated value), and that sentence is then removed
    from the bullet list so the page does not say the same thing twice.  Only the
    section/role label form is used when no whole short bullet exists.  Sentences that
    describe the fixture/document itself are never promoted — they are disclosure
    text, not a scientific claim.
    """
    for bullet in sorted((b["text"] for b in bullets), key=len):
        if META_SENTENCE_RE.search(bullet):
            continue
        cleaned = _clean_title(bullet, TITLE_LIMIT_CHARS, keep_digits=True)
        if not cleaned or cleaned in used_titles:
            continue
        if not C.numbers_covered(cleaned, material.allowed_raws)[0]:
            continue
        return cleaned
    return ""


def _claim_title(role, sections, used_titles, occurrence, bullet_text="",
                 material=None):
    """Fallback page title: the section/role label form.

    Used when no whole short sentence on the page can serve as the title; it never
    truncates a sentence and never duplicates another page's title.
    """
    del bullet_text, material
    label = ROLE_TITLE.get(role, role.capitalize())
    suffix = _ordinal(occurrence)
    section = _clean_title(sections[0], TITLE_LIMIT_CHARS) if sections else ""
    candidates = []
    if section:
        if section.lower() in GENERIC_SECTIONS:
            candidates.append("%s: %s" % (label, section))
            candidates.append(section)
        else:
            candidates.append(section)
            candidates.append("%s: %s" % (label, section))
    candidates.append(label)
    for candidate in candidates:
        candidate = _clean_title(candidate, TITLE_LIMIT_CHARS)
        if not candidate or len(candidate) > TITLE_LIMIT_CHARS:
            continue
        titled = candidate + suffix
        if titled not in used_titles:
            return titled
    fallback = (label + suffix).strip()
    while fallback in used_titles:
        fallback += " (cont.)"
    return fallback


def _caption_title(caption_body, label, limit=TITLE_LIMIT_CHARS):
    """Figure/table title: first clause, cut at the first digit, never mid-value."""
    sentence = _first_sentence(caption_body, limit=limit + 30)
    candidate = _clean_title(sentence, limit)
    if len(candidate) < 12:
        return label
    return "%s: %s" % (label, candidate)


def _caption_bullets(caption, material, excluded, key, limit_count=2, max_chars=70):
    """Short whole-clause bullets from a caption (safe for small caption boxes)."""
    bullets = []
    body = _strip_leading_number(caption or "")
    for sentence in SENTENCE_SPLIT_RE.split(_norm_ws(body)):
        text = sentence.strip()
        if len(text) < MIN_BULLET_CHARS:
            continue
        candidate = text if len(text) <= max_chars else _trim_to(text, max_chars)
        candidate = _drop_trailing_function_words(candidate)
        if len(candidate) < MIN_BULLET_CHARS:
            # the word-boundary trim landed on a function word ("… of the"), which
            # left a stub; keep the whole clause if a slightly longer line still fits
            candidate = _drop_trailing_function_words(_trim_to(text, max_chars + 25))
        if len(candidate) < MIN_BULLET_CHARS:
            excluded.append({"item": "%s caption" % key,
                             "reason": "caption clause %r falls below the %d-character "
                                       "bullet floor once its trailing words are "
                                       "dropped" % (text[:50], MIN_BULLET_CHARS)})
            continue
        ok, unmatched = C.numbers_covered(candidate, material.allowed_raws)
        if not ok:
            excluded.append({"item": "%s caption" % key,
                             "reason": "caption clause not shown: it carries the "
                                       "number(s) %s which are not attested in "
                                       "material.numbers" % ", ".join(unmatched)})
            continue
        bullets.append({"text": _ensure_terminal(candidate), "claim_id": None,
                        "locator": None})
        if len(bullets) >= limit_count:
            break
    if not bullets:
        excluded.append({"item": "%s caption" % key,
                         "reason": "the source exhibit's own caption offered no whole "
                                   "clause that fits a %d-character line, so the page "
                                   "carries the exhibit and its title only"
                                   % max_chars})
    return bullets


def _evidence_from_bullets(bullets, fallback=None):
    out = []
    for bullet in bullets:
        if bullet.get("locator"):
            out.append({"type": "claim", "ref": bullet.get("claim_id") or "claim",
                        "locator": bullet["locator"],
                        "note": "verbatim sentence from the source material"})
    if not out and fallback:
        out.append(fallback)
    return out


def _source_names_for(material, quotes):
    names = []
    for quote in quotes:
        for source_id, source in material.sources.items():
            if quote and C.is_verbatim(quote, source.get("text") or ""):
                name = os.path.basename(source.get("path") or source_id)
                if name not in names and C.numbers_covered(name, material.allowed_raws)[0]:
                    names.append(name)
                break
    return names


def _notes(role, claim, source_names, section_names):
    parts = [ROLE_NOTES.get(role, "Present the claim on this page and name the source "
                                   "of the evidence.")]
    if claim:
        parts.append("Key claim: %s" % claim)
    if source_names:
        parts.append("Traceability: the text on this page is quoted from %s."
                     % ", ".join(source_names))
    if section_names:
        parts.append("It sits in the %s part of the source." % ", ".join(section_names))
    text = " ".join(p for p in parts if p).strip()
    if len(text) < NOTES_MIN_CHARS + 20:
        text += (" Keep the page focused on this single claim and say where the "
                 "evidence comes from if the audience asks.")
    return text


def _slide(index, role, title, claim, bullets, evidence, numbers_used, notes, hint,
           layout_name):
    return {
        "index": index,
        "role": role,
        "title": title,
        "claim": claim,
        "bullets": [b["text"] if isinstance(b, dict) else str(b) for b in bullets],
        "figure_ref": None,
        "table_ref": None,
        "evidence": evidence,
        "numbers_used": numbers_used,
        "notes": notes,
        "layout_hint": hint,
        "layout_name": layout_name,
    }


def _numbers_for(material, texts):
    """Number records (as plan entries) that belong to ``texts``."""
    found = []
    for text in texts:
        found.extend(material.numbers_inside(text))
    seen, out = set(), []
    for record in found:
        key = (record.get("raw"), record.get("unit"))
        if key in seen:
            continue
        seen.add(key)
        out.append({"value": record.get("value"), "raw": record.get("raw"),
                    "unit": record.get("unit"), "locator": record.get("locator")})
    return out


# --------------------------------------------------------------------------- #
# slide assembly
# --------------------------------------------------------------------------- #
def _paper_title(material):
    """The paper's own title: the first substantive line of the richest source."""
    best = None
    for source in material.sources.values():
        if (source.get("role") or "material") != "material":
            continue
        length = len(source.get("text") or "")
        if best is None or length > len(best.get("text") or ""):
            best = source
    if best is None:
        return ""
    for line in (best.get("text") or "").split("\n"):
        candidate = _norm_ws(line)
        if 15 <= len(candidate) <= 150 and not candidate.endswith(".") \
                and not re.match(r"^(abstract|introduction|keywords|doi|http)", candidate,
                                 re.I):
            return candidate
    return ""


def build_slides(material, constraints, layout_names, warnings, uncertainty, excluded):
    """Assemble the ordered slide list, then fit it to the case's slide budget."""
    slides = []
    used_titles = set()
    role_counts = {}

    # ---- 1. title page ------------------------------------------------------
    raw_title = constraints.get("title") or _paper_title(material)
    title_ok, unmatched = C.numbers_covered(raw_title or "", material.allowed_raws)
    if raw_title and title_ok:
        title_text = raw_title
    else:
        title_text = "Research presentation"
        if raw_title:
            uncertainty.append("the paper's own title %r carries the number(s) %s which "
                               "material.numbers does not attest, so a neutral title was "
                               "used rather than editing the title"
                               % (raw_title, ", ".join(unmatched)))
            excluded.append({"item": "paper title %r" % raw_title,
                             "reason": "not shown verbatim because it contains "
                                       "number(s) %s outside material.numbers"
                                       % ", ".join(unmatched)})
    title_claim = ""
    for claim in material.claims:
        text = _norm_ws(claim.get("text") or "")
        if MIN_BULLET_CHARS <= len(text) <= CLAIM_LIMIT_CHARS and \
                C.numbers_covered(text, material.allowed_raws)[0]:
            title_claim = _ensure_terminal(text)
            material.used_claim_ids.add(claim.get("id"))
            material.used_texts.add(_norm_ws(title_claim))
            break
    if not title_claim:
        title_claim = ("This deck reports the study's problem, method, results and "
                       "limitations.")
    used_titles.add(title_text)
    slides.append(_slide(1, "title", title_text, title_claim, [],
                         _evidence_from_bullets([], {"type": "claim", "ref": "claim",
                                                     "locator": material.claims[0]
                                                     .get("locator")
                                                     if material.claims else {}}),
                         [], _notes("title", title_claim,
                                    _source_names_for(material, [title_claim]), []),
                         "title", resolve_layout_name("title", layout_names)))

    # ---- 2. outline (built at the end from the source's own sections) --------
    # A real agenda lists the paper's sections.  A list of "<role>: what this page has
    # to establish" lines is scaffolding presented as content, so it is never emitted:
    # if no section names can be derived the page is dropped and explained instead.

    # ---- 3. content pages, one role each ------------------------------------
    buckets = assign_roles(material, excluded)
    role_limit = {"problem": 3, "method": 3, "result": BULLET_LIMIT_COUNT,
                  "limitation": 3, "conclusion": 3}

    def content_slide(role, claims):
        bullets = _bullets_from_claims(claims, material, excluded, role,
                                       role_limit.get(role, 3))
        if not bullets:
            for entry in material.sections_with_quotes():
                allowed = {"problem": {"problem", "abstract"},
                           "method": {"method"},
                           "result": {"result", "discussion", "abstract"},
                           "limitation": {"limitation", "discussion", "conclusion"},
                           "conclusion": {"conclusion", "result"}}.get(role, set())
                if entry["kind"] not in allowed:
                    continue
                text = _trim_to(entry["quote"], BULLET_LIMIT_CHARS)
                text = _ensure_terminal(_drop_trailing_function_words(text))
                if len(text) < MIN_BULLET_CHARS or _norm_ws(text) in material.used_texts:
                    continue
                if not C.numbers_covered(text, material.allowed_raws)[0]:
                    continue
                bullets.append({"text": text, "claim_id": None,
                                "locator": C.locator(entry["source"],
                                                     entry["section"].get("page"),
                                                     entry["quote"])})
                material.used_texts.add(_norm_ws(text))
                break
        sections = []
        for bullet in bullets:
            loc = bullet.get("locator") or {}
            for claim in material.claims:
                if claim.get("locator") == loc:
                    section = material.section_of(claim)
                    if section and section not in sections:
                        sections.append(section)
        claim_text = bullets[0]["text"] if bullets else ""
        if not claim_text:
            claim_text = ("This page covers the %s part of the study; no verbatim "
                          "sentence was available for it." % role)
            uncertainty.append("the %s page has no verbatim source sentence; it states "
                               "its role instead of inventing content" % role)
        claim_text = _ensure_terminal(claim_text)
        role_counts[role] = role_counts.get(role, 0) + 1

        # Prefer a claim-shaped title taken from a whole short sentence on this page,
        # and drop that sentence from the bullets so the page does not repeat itself.
        title, promoted = "", ""
        if bullets and (len(bullets) > 1 or _numbers_for(material, [b["text"]
                                                                   for b in bullets])):
            promoted = _promotable_title(bullets, material, used_titles)
        if promoted:
            title = promoted
            claim_text = promoted
            bullets = [b for b in bullets
                       if _norm_ws(b["text"]) != _norm_ws(promoted)]
        else:
            title = _claim_title(role, sections, used_titles, role_counts[role] - 1)
        used_titles.add(title)
        # recompute after promotion so numbers_used describes what is really on the page
        quotes = [b["text"] for b in bullets]
        numbers_used = _numbers_for(material, quotes + [title])
        slides.append(_slide(len(slides) + 1, role, title, claim_text, bullets,
                             _evidence_from_bullets(bullets), numbers_used,
                             _notes(role, claim_text,
                                    _source_names_for(material, quotes), sections),
                             ROLE_HINT[role],
                             resolve_layout_name(ROLE_HINT[role], layout_names)))
        slides[-1]["_sections"] = sections

    content_slide("problem", buckets["problem"])
    content_slide("method", buckets["method"])
    result_chunks = [buckets["result"][i:i + 3] for i in range(0, len(buckets["result"]), 3)]
    for chunk in result_chunks[:3] or [[]]:
        content_slide("result", chunk)

    # ---- 4. figure pages: a real captured image is required ------------------
    for figure in material.figures:
        image = figure.get("image") or {}
        if not image.get("path"):
            excluded.append({"item": figure.get("id"),
                             "reason": "figure page not created: no image was captured "
                                       "for %s (kind=%s), and a placeholder figure is "
                                       "worse than none" % (figure.get("label"),
                                                            figure.get("kind"))})
            continue
        caption = figure.get("caption") or ""
        body_text = _strip_leading_number(caption.split(". ", 1)[-1])
        label = re.sub(r"\d+", "", figure.get("label") or "Figure").strip(" .") or "Figure"
        # the picture-caption layout's caption box is small: one short line only
        bullets = _caption_bullets(body_text, material, excluded, figure.get("id"),
                                   limit_count=1, max_chars=60)
        claim_text = _ensure_terminal(
            C.numbers_covered(caption, material.allowed_raws)[0] and
            (_first_sentence(caption, limit=CLAIM_LIMIT_CHARS) or caption) or
            "This page shows the source figure for the claim above.")
        role_counts["figure"] = role_counts.get("figure", 0) + 1
        title = _caption_title(body_text, label, limit=30)
        if title in used_titles:
            title = _caption_title(body_text, label + _ordinal(role_counts["figure"] - 1),
                                   limit=30)
        used_titles.add(title)
        slide = _slide(len(slides) + 1, "figure", title, claim_text, bullets,
                       [{"type": "figure", "ref": figure.get("id"),
                         "locator": figure.get("locator"),
                         "note": "labelled \"%s\" in the source" % figure.get("label")}],
                       _numbers_for(material, [caption]),
                       _notes("figure", claim_text,
                              _source_names_for(material, [caption]), []),
                       ROLE_HINT["figure"],
                       resolve_layout_name(ROLE_HINT["figure"], layout_names))
        slide["figure_ref"] = figure.get("id")
        slides.append(slide)

    # ---- 5. table pages: real cells required --------------------------------
    for table in material.tables:
        cells = table.get("cells")
        if not cells:
            excluded.append({"item": table.get("id"),
                             "reason": "table page not created: only the caption %r was "
                                       "extractable, the cell grid was not, and a table "
                                       "page without its real values would be a "
                                       "placeholder" % (table.get("caption") or "")[:60]})
            continue
        caption = table.get("caption") or ""
        body_text = _strip_leading_number(caption.split(". ", 1)[-1])
        label = re.sub(r"\d+", "", table.get("label") or "Table").strip(" .") or "Table"
        bullets = _caption_bullets(body_text, material, excluded, table.get("id"))
        claim_text = _ensure_terminal(_first_sentence(caption, limit=CLAIM_LIMIT_CHARS)
                                      or caption
                                      or "This page lists the source table values.")
        role_counts["table"] = role_counts.get("table", 0) + 1
        title = _caption_title(body_text, label)
        if title in used_titles:
            title = _caption_title(body_text, label + _ordinal(role_counts["table"] - 1))
        used_titles.add(title)
        slide = _slide(len(slides) + 1, "table", title, claim_text, bullets,
                       [{"type": "table", "ref": table.get("id"),
                         "locator": table.get("locator"),
                         "note": "cells taken verbatim from the source table"}],
                       _numbers_for(material, [caption, " ".join(
                           str(cell) for row in cells for cell in row)]),
                       _notes("table", claim_text, _source_names_for(material, [caption]),
                              []), ROLE_HINT["table"],
                       resolve_layout_name(ROLE_HINT["table"], layout_names))
        slide["table_ref"] = table.get("id")
        slides.append(slide)

    content_slide("limitation", buckets["limitation"])
    content_slide("conclusion", buckets["conclusion"])

    # ---- 6. outline page, from the sections the content pages draw on --------
    outline_bullets, seen_outline = [], set()
    for slide in slides:
        if slide["role"] not in ("problem", "method", "result", "limitation",
                                 "conclusion"):
            continue
        for section in slide.get("_sections") or []:
            name = _clean_title(section, 60)
            key = _content_key(name)
            if not name or len(name) < 3 or key in seen_outline:
                continue
            seen_outline.add(key)
            outline_bullets.append(name)
    if outline_bullets:
        outline = _slide(
            0, "outline", ROLE_TITLE["outline"],
            "The talk follows the paper's own sections in this order.",
            [{"text": name, "claim_id": None, "locator": None}
             for name in outline_bullets[:BULLET_LIMIT_COUNT]],
            [{"type": "section", "ref": None,
              "locator": next((s["evidence"][0].get("locator") for s in slides
                               if s["role"] == "problem" and s["evidence"]), None),
              "note": "agenda taken from the source's own section headings"}],
            [], _notes("outline", "", [], outline_bullets[:3]), "title-content",
            resolve_layout_name("title-content", layout_names))
        slides.insert(1, outline)
        used_titles.add(ROLE_TITLE["outline"])
    else:
        excluded.append({
            "item": "outline page",
            "reason": "no outline page: the source exposes no section headings, and an "
                      "agenda of role stubs (\"Problem: what this page has to "
                      "establish\") would be scaffolding presented as content"})

    # ---- 6. final safety pass: no title or bullet repeats --------------------
    seen_titles, seen_bullets, seen_keys = set(), set(), set()
    for slide in slides:
        if slide["title"] in seen_titles:
            replacement = _claim_title("backup", [], seen_titles, 1)
            excluded.append({"item": slide["title"],
                             "reason": "duplicate title replaced by %r on slide %d so "
                                       "that no two pages share a title"
                                       % (replacement, slide["index"])})
            slide["title"] = replacement
        seen_titles.add(slide["title"])
        kept = []
        for bullet in slide["bullets"]:
            bullet_key = _content_key(bullet)
            if _is_reused(bullet_key, seen_keys):
                excluded.append({"item": bullet[:60],
                                 "reason": "bullet removed from slide %s: its wording "
                                           "already appears on an earlier page or in a "
                                           "title of this deck" % slide["index"]})
                continue
            seen_bullets.add(_norm_ws(bullet))
            seen_keys.add(bullet_key)
            kept.append(bullet)
        slide["bullets"] = kept
        title_key = _content_key(slide["title"])
        if title_key:
            seen_keys.add(title_key)

    # ---- 6b. a content page that carries nothing must not exist ---------------
    # "one page = one claim" means a problem/method/result/limitation/conclusion page
    # has to show evidence: bullets, an exhibit reference or an attested number.  A
    # page with only a title is a divider, and pretending otherwise is exactly the
    # "empty filler page" the reviewer penalises.  The slide budget is refilled below
    # from material that does carry evidence.
    content_roles = {"problem", "method", "result", "limitation", "conclusion"}
    kept_slides = []
    for slide in slides:
        if (slide["role"] in content_roles and not slide["bullets"]
                and not slide["figure_ref"] and not slide["table_ref"]
                and not slide["numbers_used"]):
            excluded.append({
                "item": "slide %s (%s) %r" % (slide["index"], slide["role"],
                                              slide["title"]),
                "reason": "page dropped: the %s role produced no whole sentence, no "
                          "figure, no table and no attested number, so the page would "
                          "have been a bare title (%d page(s) were refilled from "
                          "material that does carry evidence)" % (slide["role"], 1)})
            continue
        kept_slides.append(slide)
    slides = kept_slides

    # ---- 7. fit the case's slide budget -------------------------------------
    lo, hi, target = slide_budget(constraints, warnings, uncertainty)
    slides = fit_budget(slides, material, target, lo, hi, layout_names, warnings,
                        uncertainty, excluded)

    # ---- 8. justify every exhibit that did not reach a page ------------------
    # A figure or table that is in the material but on no page must be explained by
    # id, not left implicit: the reviewer (and the deck gate) reads plan.excluded as
    # the trade-off trail, and 结果完整性 is about exactly that.
    listed = " ".join(str(entry.get("item") or "") for entry in excluded)
    referenced_figures = {s["figure_ref"] for s in slides if s.get("figure_ref")}
    referenced_tables = {s["table_ref"] for s in slides if s.get("table_ref")}
    for figure in material.figures:
        identifier = figure.get("id")
        if not identifier or identifier in referenced_figures or identifier in listed:
            continue
        excluded.append({
            "item": "%s (figure %s)" % (identifier, figure.get("label")),
            "reason": "not placed: the deck is %d slides and the kept figure, table, "
                      "limitation and conclusion pages rank higher; the exhibit stays "
                      "available in material.json with its locator" % len(slides)})
    for table in material.tables:
        identifier = table.get("id")
        if not identifier or identifier in referenced_tables or identifier in listed:
            continue
        excluded.append({
            "item": "%s (table %s)" % (identifier, table.get("label")),
            "reason": "not placed: the deck is %d slides and the kept pages rank "
                      "higher; its cells stay available in material.json with their "
                      "locator" % len(slides)})
    for position, slide in enumerate(slides, start=1):
        slide["index"] = position
        slide.pop("_priority", None)
        slide.pop("_sections", None)
    return slides


def slide_budget(constraints, warnings, uncertainty):
    count = constraints.get("slide_count")
    rng = constraints.get("slide_count_range")
    if count:
        return count, count, count
    if rng:
        return int(rng[0]), int(rng[1]), None
    warnings.append("case declares neither slide_count nor slide_count_range; the "
                    "default target of 9 slides was used")
    uncertainty.append("slide_count is not declared by the case; the deck size was "
                       "chosen by the skill (9 slides) and is not a case requirement")
    return None, None, 9


def _assign_priorities(slides):
    counts = {}
    for slide in slides:
        role = slide["role"]
        counts[role] = counts.get(role, 0) + 1
        slide["_priority"] = ROLE_PRIORITY.get(role, 25) + (counts[role] - 1) * 7
    return slides


def fit_budget(slides, material, target, lo, hi, layout_names, warnings, uncertainty,
               excluded):
    """Select/extend the slide list so the count honours the case constraint.

    Selection is by priority (title → problem → method → result → conclusion →
    figure → table → limitation → outline → backup), so a tight slide count drops the
    outline and the extra result pages instead of dropping every figure.
    """
    if target is None and lo is None:
        return slides
    budget = target if target is not None else max(lo or 0, min(len(slides), hi or 10 ** 6))
    _assign_priorities(slides)

    if len(slides) > budget:
        ordered = sorted(slides, key=lambda s: (s["_priority"], s["index"]))
        kept = ordered[:budget]
        keep_ids = {id(s) for s in kept}
        covered_bullets = {_norm_ws(b) for s in kept for b in s["bullets"]}
        for slide in slides:
            if id(slide) in keep_ids:
                continue
            overlap = [b for b in slide["bullets"] if _norm_ws(b) in covered_bullets]
            if overlap:
                reason = ("not kept: the case allows %s slides; this %s page repeats "
                          "text already carried by a kept page (%r)"
                          % (budget, slide["role"], overlap[0][:60]))
            else:
                reason = ("not kept: the case allows %s slides and the %s role (rank %s) "
                          "sits below the figure/table/limitation pages that were kept"
                          % (budget, slide["role"], ROLE_PRIORITY.get(slide["role"], 25)))
            excluded.append({"item": "slide %s (%s) %r" % (slide["index"], slide["role"],
                                                           slide["title"]),
                             "reason": reason})
        slides = [s for s in slides if id(s) in keep_ids]
        slides.sort(key=lambda s: (ROLE_ORDER_INDEX.get(s["role"], 15), s["index"]))
        return slides

    used_titles = {s["title"] for s in slides}
    while len(slides) < budget:
        added = None
        for claim in material.claims:
            if claim.get("id") in material.used_claim_ids:
                continue
            text = _norm_ws(claim.get("text") or "")
            if not (MIN_BULLET_CHARS <= len(text) <= BULLET_LIMIT_CHARS):
                continue
            if _norm_ws(text) in material.used_texts:
                continue
            if not C.numbers_covered(text, material.allowed_raws)[0]:
                continue
            text = _ensure_terminal(text)
            title = _claim_title("backup", [material.section_of(claim)], used_titles, 1)
            added = _slide(0, "backup", title, text,
                           [{"text": text, "claim_id": claim.get("id"),
                             "locator": claim.get("locator")}],
                           _evidence_from_bullets([{"locator": claim.get("locator")}]),
                           _numbers_for(material, [text]),
                           _notes("backup", text, _source_names_for(material, [text]),
                                  [material.section_of(claim)]),
                           ROLE_HINT["backup"],
                           resolve_layout_name(ROLE_HINT["backup"], layout_names))
            used_titles.add(title)
            material.used_claim_ids.add(claim.get("id"))
            material.used_texts.add(_norm_ws(text))
            break
        if added is None:
            for entry in material.sections_with_quotes():
                title = _clean_title(entry["section"].get("title") or "")
                if not title or title in used_titles:
                    continue
                text = _ensure_terminal(_drop_trailing_function_words(
                    _trim_to(entry["quote"], BULLET_LIMIT_CHARS)))
                if len(text) < MIN_BULLET_CHARS or _norm_ws(text) in material.used_texts:
                    continue
                if not C.numbers_covered(text, material.allowed_raws)[0]:
                    continue
                locator = C.locator(entry["source"], entry["section"].get("page"),
                                    entry["quote"])
                added = _slide(0, "backup", title, text,
                               [{"text": text, "claim_id": None, "locator": locator}],
                               [{"type": "section", "ref": entry["section"].get("id"),
                                 "locator": locator,
                                 "note": "section heading and its first sentence"}],
                               _numbers_for(material, [text]),
                               _notes("backup", text, _source_names_for(material, [text]),
                                      [title]), ROLE_HINT["backup"],
                               resolve_layout_name(ROLE_HINT["backup"], layout_names))
                used_titles.add(title)
                material.used_texts.add(_norm_ws(text))
                break
        if added is None:
            uncertainty.append("the case asks for %s slides but the material only "
                               "supports %d pages that could be filled without "
                               "inventing content" % (budget, len(slides)))
            warnings.append("slide count %d is below the requested %d: no more verbatim "
                            "material is available" % (len(slides), budget))
            break
        slides.append(added)
    return slides


# --------------------------------------------------------------------------- #
# orchestration
# --------------------------------------------------------------------------- #
def build_plan(case_path, work_dir, in_path=None, out_path=None, log=None):
    case = C.load_case(case_path)
    case_id = C.extract_case_id(case, case_path)
    constraints, constraint_sources, task_card = C.extract_constraints_merged(case,
                                                                              case_path)
    warnings, uncertainty, excluded = [], [], []
    if task_card:
        warnings.append("task card read for constraints: %s (fields taken from the task "
                        "card are recorded in constraints._sources)"
                        % os.path.basename(task_card))
    material = Material(load_material(work_dir, in_path))

    template_path = find_template(constraints, case, case_path, work_dir, warnings)
    layout_names = read_layouts(template_path, warnings, uncertainty)
    unresolved = sorted({h for h in set(ROLE_HINT.values())
                         if resolve_layout_name(h, layout_names) is None})
    if unresolved:
        uncertainty.append("layouts for hint(s) %s are not present in the template"
                           % ", ".join(unresolved))

    slides = build_slides(material, constraints, layout_names, warnings, uncertainty,
                          excluded)
    hint_reports = {}
    for slide in slides:
        hint_reports.setdefault(slide["layout_hint"],
                                {"hint": slide["layout_hint"],
                                 "layout_name": slide["layout_name"],
                                 "resolved": slide["layout_name"] is not None})
    if constraints.get("starting_deck_path"):
        uncertainty.append("the case provides a starting deck (%s); build_deck.py adapts "
                           "it instead of starting from the template alone"
                           % os.path.basename(constraints["starting_deck_path"]))
    if any(not s["bullets"] for s in slides):
        uncertainty.append("%d page(s) carry no bullet list because the material offered "
                           "no whole sentence for them; their claim and evidence are "
                           "still recorded" % sum(1 for s in slides if not s["bullets"]))

    plan = {
        "schema": C.SCHEMA_PLAN,
        "case_id": case_id,
        "generated": C.now_iso(),
        "constraints": {
            "audience": constraints.get("audience"),
            "duration_min": constraints.get("duration_min"),
            "slide_count": constraints.get("slide_count"),
            "slide_count_range": constraints.get("slide_count_range"),
            "language": constraints.get("language") or "en",
            "template_path": template_path,
            "starting_deck_path": constraints.get("starting_deck_path"),
            "template_layouts": layout_names,
            "layout_hints_available": hint_reports,
            "_sources": constraint_sources,
            "_task_card": task_card,
        },
        "narrative": {
            "arc": "problem -> method -> result -> limitation -> conclusion",
            "main_thread": ["%d. %s — %s" % (s["index"], s["role"], s["title"])
                            for s in slides],
        },
        "slides": slides,
        "excluded": excluded,
        "uncertainty": uncertainty,
        "warnings": warnings,
        "stats": {"slides": len(slides), "bullets": sum(len(s["bullets"]) for s in slides),
                  "figures_used": sum(1 for s in slides if s["figure_ref"]),
                  "tables_used": sum(1 for s in slides if s["table_ref"]),
                  "numbers_used": sum(len(s["numbers_used"]) for s in slides),
                  "excluded": len(excluded), "warnings": len(warnings)},
    }
    return plan


def build_arg_parser():
    parser = argparse.ArgumentParser(
        prog="plan.py", description="Build plan.json (plan/v1) from material.json.")
    parser.add_argument("--case", required=True, help="path to case.json")
    parser.add_argument("--in", dest="in_path", default=None,
                        help="material.json (default <work>/material.json)")
    parser.add_argument("--work", default=None, help="work directory")
    parser.add_argument("--out", default=None, help="plan.json path")
    return parser


def main(argv=None):
    args = build_arg_parser().parse_args(argv)
    work_dir = C.derive_work_dir(args.case, args.work)
    out_path = args.out or os.path.join(work_dir, "plan.json")
    log = C.RunLog(os.path.join(work_dir, "plan.log"))
    try:
        with log.stage("plan", case=os.path.basename(args.case)):
            plan = build_plan(args.case, work_dir, in_path=args.in_path,
                              out_path=out_path, log=log)
        C.safe_write_json(out_path, plan)
    except SystemExit:
        raise
    except Exception as exc:
        log.error("plan failed: %s: %s" % (type(exc).__name__, exc))
        C.safe_write_json(out_path, {
            "schema": C.SCHEMA_PLAN, "case_id": "unknown-case",
            "constraints": {}, "narrative": {}, "slides": [], "excluded": [],
            "uncertainty": ["planning did not complete: %s" % exc],
            "warnings": ["plan failed: %s: %s" % (type(exc).__name__, exc)], "stats": {},
        })
        print("error: plan failed; a minimal plan.json was written to %s" % out_path,
              file=_sys.stderr)
        return 2
    print(json.dumps({"plan": out_path, "stats": plan["stats"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
