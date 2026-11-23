#!/usr/bin/env python3
"""evidence.py -- build the page-level evidence map (``evidence_map.json``).

Stage 5 of the 8-CLI pipeline (CONTRACT.md section 3).  For every slide it records each
claim / figure / table / number together with the locator that justifies it, and marks an
item ``verified`` **only when its locator quote really occurs in the source text**.

The verification rule is deliberately mechanical:

* ``exact``       -- the locator quote appears verbatim (whitespace-normalised) in the text
                     of the source it names, so ``verified`` is true.
* ``paraphrase``  -- the quote is not a verbatim substring but most of its words occur in the
                     source, so the page is related to the source but the wording is the
                     skill's; ``verified`` stays **false**, because "related" is not "quoted".
* ``unsupported`` -- neither, so the item lands in the page's ``unverified`` list with a
                     reason.  This file never turns a missing source into a green tick.

CLI
---

    evidence.py --case case.json --work WORK [--out evidence_map.json]
                [--plan plan.json] [--material material.json]
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

SCHEMA = "evidence_map/v1"

try:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from common import dump_json as _dump_json  # type: ignore
except Exception:  # noqa: BLE001
    _dump_json = None


def load_json(path, default=None):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return default


def dump_json(path, obj):
    if _dump_json is not None:
        try:
            _dump_json(path, obj)
            return
        except Exception:  # noqa: BLE001
            pass
    path = Path(path)
    if str(path.parent) not in ("", "."):
        path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(obj, fh, ensure_ascii=False, indent=2)
        fh.write("\n")


def norm(text):
    return re.sub(r"\s+", " ", text or "").strip()


def sentences(text):
    return [s for s in re.split(r"(?<=[.!?。！？])\s+|\n+", text or "") if s.strip()]


def source_index(material):
    """``{source_id: normalised text}`` for every source in material.json."""
    index = {}
    for src in (material or {}).get("sources", []) or []:
        if not isinstance(src, dict):
            continue
        chunks = []
        for page in src.get("pages") or []:
            if isinstance(page, dict):
                chunks.append(page.get("text") or "")
            elif isinstance(page, str):
                chunks.append(page)
        for sec in src.get("sections") or []:
            if isinstance(sec, dict):
                chunks.append(sec.get("title") or "")
        index[str(src.get("id", ""))] = norm("\n".join(chunks))
    if not index:
        # tolerate a material.json whose sources have no id but do carry text
        joined = []
        for src in (material or {}).get("sources", []) or []:
            if isinstance(src, dict):
                for page in src.get("pages") or []:
                    if isinstance(page, dict):
                        joined.append(page.get("text") or "")
        if joined:
            index[""] = norm("\n".join(joined))
    return index


def classify(quote, source_id, index):
    """Return ``(match, verified, reason)`` for one locator quote."""
    quote = norm(quote)
    if not quote:
        return "unsupported", False, "locator carries no quote"
    candidates = []
    if source_id and str(source_id) in index:
        candidates.append((str(source_id), index[str(source_id)]))
    else:
        candidates += list(index.items())
    for sid, text in candidates:
        if not text:
            continue
        if quote in text:
            return "exact", True, ""
    # paraphrase: most of the quote's words appear inside one source sentence
    q_tokens = {t for t in re.findall(r"[\w%.\-]+", quote.lower()) if len(t) > 1}
    if q_tokens:
        for sid, text in candidates:
            for sentence in sentences(text):
                s_tokens = {t for t in re.findall(r"[\w%.\-]+", sentence.lower())
                            if len(t) > 1}
                if not s_tokens:
                    continue
                overlap = len(q_tokens & s_tokens) / float(len(q_tokens))
                if overlap >= 0.8:
                    return "paraphrase", False, (
                        f"{overlap:.0%} of the quote's words appear in one source sentence but "
                        f"the quote is not verbatim")
    return "unsupported", False, "quote not found in the named source text"


def locate(item, index, kind):
    """Classify one evidence-bearing item; returns ``(entry, unverified_or_None, verified_bool)``."""
    locator = item.get("locator") or {}
    if not isinstance(locator, dict):
        locator = {}
    quote = locator.get("quote", "")
    source_id = locator.get("source", "")
    match, verified, reason = classify(quote, source_id, index)
    entry = dict(item)
    entry["locator"] = locator
    entry["match"] = match
    entry["verified"] = verified
    if not verified:
        un = {"kind": kind, "reason": reason,
              "item": (item.get("claim") or item.get("raw") or item.get("ref")
                       or item.get("label") or "")}
        return entry, un, False
    return entry, None, True


def numbers_on(text):
    """Numeric tokens on a slide, with caption labels removed (mirrors verify_deck V04)."""
    stripped = re.sub(
        r"\b(?:fig(?:ure)?|tab(?:le)?|sec(?:tion)?|eq(?:uation)?|appendix|step|page|slide|ref|no)"
        r"\s*\.?\s*\d+", " ", text or "", flags=re.I)
    return re.findall(r"\d+(?:[.,]\d+)?", stripped)


def build_map(case, work, plan_path=None, material_path=None):
    work = Path(work)
    plan_path = Path(plan_path) if plan_path else work / "plan.json"
    material_path = Path(material_path) if material_path else work / "material.json"
    plan = load_json(plan_path, {}) or {}
    material = load_json(material_path, {}) or {}
    index = source_index(material)

    numbers_by_raw = {}
    for item in (material.get("numbers") or []):
        if not isinstance(item, dict):
            continue
        for rep in {str(item.get("raw", "")).strip(), str(item.get("value", "")).strip()}:
            if rep:
                numbers_by_raw.setdefault(rep, item)
    figures_by_id = {str(f.get("id")): f for f in (material.get("figures") or [])
                     if isinstance(f, dict)}
    tables_by_id = {str(t.get("id")): t for t in (material.get("tables") or [])
                    if isinstance(t, dict)}

    pages = []
    warnings = []
    for i, slide in enumerate(plan.get("slides") or [], 1):
        if not isinstance(slide, dict):
            continue
        page = {"slide": int(slide.get("index", i) or i),
                "role": slide.get("role", ""),
                "claims": [], "figures": [], "tables": [], "numbers": [], "unverified": []}

        # ---- claims: the slide's own claim plus any explicit claim/section evidence entry ----
        page_claim = norm(slide.get("claim") or "")
        if page_claim:
            ev = None
            for item in slide.get("evidence") or []:
                if isinstance(item, dict) and item.get("locator"):
                    ev = item
                    break
            entry = {"claim": page_claim,
                     "locator": (ev or {}).get("locator") or {},
                     "text": page_claim}
            e2, un, _ = locate(entry, index, "claim")
            page["claims"].append({k: e2[k] for k in ("claim", "locator", "match", "verified")})
            if un:
                page["unverified"].append(un)
        for item in slide.get("evidence") or []:
            if not isinstance(item, dict):
                continue
            etype = item.get("type")
            # numbers are resolved below against material.numbers, and figure/table locators
            # are carried by the figure/table entries themselves -- do not duplicate them here
            if etype in ("figure", "table", "number"):
                continue
            text = item.get("text") or item.get("ref") or ""
            if not text:
                continue
            entry = {"claim": text, "locator": item.get("locator") or {}}
            e2, un, _ = locate(entry, index, etype or "evidence")
            page["claims"].append({k: e2[k] for k in ("claim", "locator", "match", "verified")})
            if un:
                page["unverified"].append(un)

        # ---- figures / tables ----
        for ref, table in ((slide.get("figure_ref"), False), (slide.get("table_ref"), True)):
            if not ref:
                continue
            ref = str(ref)
            src = tables_by_id.get(ref) if table else figures_by_id.get(ref)
            if src is None:
                page["unverified"].append({
                    "kind": "table" if table else "figure",
                    "reason": f"{ref!r} is not present in material.json",
                    "item": ref})
                continue
            locator = src.get("locator") or {}
            entry = {"ref": ref, "label": src.get("label", ""),
                     "caption": src.get("caption", ""),
                     "source_ref": (src.get("locator") or {}).get("source", ""),
                     "locator": locator}
            e2, un, _ = locate(entry, index, "table" if table else "figure")
            bucket = "tables" if table else "figures"
            page[bucket].append({k: e2[k] for k in
                                 ("ref", "label", "caption", "locator", "match", "verified")})
            if un:
                page["unverified"].append(un)

        # ---- numbers: what is actually on the slide, resolved to material.locator ----
        slide_text = "\n".join(str(x) for x in
                              ([slide.get("title") or "", slide.get("claim") or ""]
                               + list(slide.get("bullets") or [])
                               + [str(n.get("raw", "")) for n in slide.get("numbers_used") or []
                                  if isinstance(n, dict)]))
        seen_raw = set()
        for raw in numbers_on(slide_text):
            if raw in seen_raw:
                continue
            seen_raw.add(raw)
            src = numbers_by_raw.get(raw)
            locator = (src or {}).get("locator") or {}
            entry = {"raw": raw, "locator": locator}
            if src is None:
                page["unverified"].append({
                    "kind": "number", "reason": f"{raw!r} has no counterpart in "
                                               f"material.numbers", "item": raw})
                continue
            e2, un, _ = locate(entry, index, "number")
            page["numbers"].append({"raw": raw, "locator": locator,
                                    "match": e2["match"], "verified": e2["verified"]})
            if un:
                page["unverified"].append(un)

        pages.append(page)

    verified_slides = sum(1 for p in pages
                          if any(c.get("verified") for c in p["claims"])
                          or any(f.get("verified") for f in p["figures"])
                          or any(t.get("verified") for t in p["tables"])
                          or any(n.get("verified") for n in p["numbers"]))
    with_evidence = sum(1 for p in pages
                        if p["claims"] or p["figures"] or p["tables"] or p["numbers"])
    if not plan.get("slides"):
        warnings.append("plan.json carried no slides, so the evidence map is empty; this is "
                        "not a pass and verify_deck V14 will fail on it")
    if not material.get("sources"):
        warnings.append("material.json carried no sources, so no quote could be verified")
    if len(pages) != len(plan.get("slides") or []):
        warnings.append("some plan slides were not dict-like and were skipped")

    out = {
        "schema": SCHEMA,
        "case_id": (case or {}).get("case_id") or plan.get("case_id"),
        "pages": pages,
        "coverage": {"slides_total": len(pages), "slides_with_evidence": with_evidence},
        "warnings": warnings,
    }
    out["coverage"]["slides_with_verified_evidence"] = verified_slides
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description="Build the page-level evidence map.")
    ap.add_argument("--case", default=None)
    ap.add_argument("--work", required=True)
    ap.add_argument("--out", default=None)
    ap.add_argument("--plan", default=None)
    ap.add_argument("--material", default=None)
    args = ap.parse_args(argv)

    work = Path(args.work)
    report_path = Path(args.out) if args.out else work / "evidence_map.json"
    case = load_json(args.case, {}) if args.case else {}
    emap = build_map(case, work, args.plan, args.material)
    dump_json(report_path, emap)
    print(json.dumps({"schema": emap["schema"], "pages": len(emap["pages"]),
                      "coverage": emap["coverage"], "out": str(report_path)},
                     ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
