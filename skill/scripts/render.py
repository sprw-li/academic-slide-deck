#!/usr/bin/env python3
"""render.py -- turn the built .pptx into inspectable previews.

Stage 4 of the 8-CLI pipeline (CONTRACT.md section 3).  Produces:

    <work>/render/deck.pdf
    <work>/render/slide-XX.png        (one per slide, zero-padded, slide numbering = 1-based)
    <work>/render_report.json         (schema render_report/v1)

Design rules (CONTRACT.md sections 3 and 5, "honesty rules"):

*   A preview is only reported when the file was actually produced and non-empty.
    No fabricated paths -- if nothing worked we emit ``status="skipped"``,
    ``renderer.available=false`` and a matching ``uncertainty`` entry.
*   No absolute author-machine path is hard-coded.  Renderers are discovered through
    environment overrides, then ``PATH``, then known installation directories, then a
    bounded search -- and the winning strategy is recorded in the report so the
    discovery is auditable rather than magic.
*   Every external tool is optional.  Missing tools degrade; they never crash the run.

Renderer strategies, in order
-----------------------------

1. ``libreoffice-kit`` CLI through Node (preferred -- this is the DSH-sidecar route and it
   both exports PDF and rasterises PNGs page by page).
2. ``soffice``/``libreoffice``/``impress`` binary on PATH or in a known install dir,
   ``--headless --convert-to pdf`` (Windows gets a private ``-env:UserInstallation`` profile).
3. ``--renderer-cli`` / ``--node`` command-line overrides, which pin strategy 1 exactly.

PDF -> PNG fallback when the chosen renderer cannot rasterise: ``pdftoppm`` (Poppler) is
tried; ``PIL``/``pypdf`` deliberately are not, because neither can rasterise a PDF vector
page.  ImageMagick's ``convert`` is explicitly *not* used: on Windows ``convert.exe`` is a
filesystem-conversion utility, not ImageMagick, and silently running it would be a trap.

CLI
---

    render.py --case case.json --work WORK [--in deck.pptx] [--out DIR_OR_JSON]
              [--renderer-cli PATH] [--node PATH] [--pages 1,2,3] [--dpi N]

``--out`` may name either the output directory (``.../render``) or the report json path;
the two are disambiguated by the ``.json`` suffix.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path

SCHEMA = "render_report/v1"

# --------------------------------------------------------------------------- #
# shared helpers (common.py is owned by deck-core; import it when present and
# fall back to local equivalents so this stage stays independently runnable --
# a hard requirement of CONTRACT.md section 3).
# --------------------------------------------------------------------------- #

try:  # pragma: no cover - depends on sibling stage being present
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from common import dump_json as _dump_json  # type: ignore
    from common import sha256_file as _sha256_file  # type: ignore
except Exception:  # noqa: BLE001
    _dump_json = None
    _sha256_file = None


def sha256_file(path):
    """Return the hex sha256 of *path*, or ``""`` when it cannot be read."""
    if _sha256_file is not None:
        try:
            return _sha256_file(path)
        except Exception:  # noqa: BLE001
            pass
    try:
        h = hashlib.sha256()
        with open(path, "rb") as fh:
            for chunk in iter(lambda: fh.read(1 << 20), b""):
                h.update(chunk)
        return h.hexdigest()
    except OSError:
        return ""


def dump_json(path, obj):
    """Write *obj* as UTF-8 JSON, creating parent directories."""
    if _dump_json is not None:
        try:
            _dump_json(path, obj)
            return
        except Exception:  # noqa: BLE001
            pass
    path = Path(path)
    if path.parent and not str(path.parent) in ("", "."):
        path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(obj, fh, ensure_ascii=False, indent=2)
        fh.write("\n")


def load_json(path, default=None):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return default


def is_windows():
    return os.name == "nt" or platform.system().lower().startswith("win")


# --------------------------------------------------------------------------- #
# renderer discovery
# --------------------------------------------------------------------------- #

KIT_CLI_REL = os.path.join("node_modules", "@deepseek-ai", "libreoffice-kit", "lib", "cli.js")

#: env vars that may point straight at an executable/binary.
BIN_ENV_VARS = ("SOFFICE_BIN", "LIBREOFFICE_KIT_BIN", "LO_KIT_BIN", "LIBREOFFICE_BIN")
#: env vars that may point at the kit's node CLI entry point.
KIT_CLI_ENV_VARS = ("LO_KIT_CLI", "LIBREOFFICE_KIT_CLI", "DSH_LO_KIT_CLI")
#: env vars that may point at the node executable to drive that CLI.
KIT_NODE_ENV_VARS = ("LO_KIT_NODE", "LIBREOFFICE_KIT_NODE")

SOFFICE_NAMES = ("soffice", "soffice.exe", "soffice.bin", "libreoffice", "impress", "localc")


def _which(names):
    for name in names:
        found = shutil.which(name)
        if found:
            return found
    return None


def _node_exe():
    for var in KIT_NODE_ENV_VARS:
        val = os.environ.get(var)
        if val and Path(val).exists():
            return val
    return shutil.which("node") or shutil.which("node.exe")


def _known_soffice_dirs():
    """Install directories that are named by the platform, not by this author."""
    dirs = []
    if is_windows():
        for var in ("PROGRAMFILES", "PROGRAMFILES(X86)", "ProgramW6432", "LOCALAPPDATA"):
            base = os.environ.get(var)
            if base:
                dirs.append(Path(base) / "LibreOffice" / "program")
                dirs.append(Path(base) / "Programs" / "LibreOffice" / "program")
    else:
        dirs += [Path("/usr/bin"), Path("/usr/local/bin"), Path("/opt/libreoffice/program"),
                 Path("/usr/lib/libreoffice/program"), Path("/snap/bin")]
        for pattern_root in ("/opt",):
            root = Path(pattern_root)
            if root.is_dir():
                try:
                    for child in sorted(root.glob("libreoffice*")):
                        dirs.append(child / "program")
                except OSError:
                    pass
        dirs.append(Path("/Applications/LibreOffice.app/Contents/MacOS"))
    return dirs


KIT_PKG_REL = os.path.join("@deepseek-ai", "libreoffice-kit", "lib", "cli.js")


def _iter_roots():
    """Yield plausible roots that may contain a ``node_modules`` with the kit.

    Every root is *derived* from the environment at run time -- a PATH entry, an env var
    that happens to name a module directory, the node executable, ``DSH_*`` variables, or
    the working directory and its ancestors.  No source-machine path is written here.
    """
    raw = []

    node = _node_exe()
    if node:
        raw += [str(p) for p in list(Path(node).resolve().parents)[:8]]

    # PATH entries and their ancestors: a runtime's bin dir sits inside the app tree.
    for entry in (os.environ.get("PATH") or "").split(os.pathsep)[:60]:
        entry = entry.strip()
        if not entry or not Path(entry).is_dir():
            continue
        try:
            raw += [str(p) for p in list(Path(entry).resolve().parents)[:5]]
        except OSError:
            pass

    # Any env var whose value names a module/asar location (the harness exports several),
    # plus its ancestors -- walking up from one reaches the ``node_modules`` parent.
    for key, val in os.environ.items():
        if not val or not isinstance(val, str):
            continue
        low = val.lower()
        if "node_modules" in low or "app.asar" in low or key.upper().startswith("DSH"):
            cleaned = val.lstrip("\\?")
            try:
                path = Path(cleaned)
            except (OSError, ValueError):
                continue
            raw.append(str(path))
            try:
                raw += [str(p) for p in list(path.resolve().parents)[:6]]
            except OSError:
                pass

    try:
        here = Path.cwd().resolve()
        raw += [str(p) for p in list(here.parents)[:6]]
    except OSError:
        pass

    seen = set()
    for item in raw:
        if item not in seen:
            seen.add(item)
            yield Path(item)


def _bounded_kit_search():
    """Return the first ``cli.js`` found under a bounded derivation of the environment."""
    probes = (KIT_CLI_REL, KIT_PKG_REL,
              os.path.join("resources", "app.asar.unpacked", "dsh", KIT_CLI_REL),
              os.path.join("app.asar.unpacked", "dsh", KIT_CLI_REL))
    for root in list(_iter_roots())[:400]:
        for rel in probes:
            cand = root / rel
            if cand.is_file():
                return cand
    return None


def _nearby_node_candidates(kit_cli):
    """Node executables plausibly belonging to the same installation as *kit_cli*."""
    found = []
    if kit_cli:
        try:
            parents = list(Path(kit_cli).resolve().parents)[:8]
        except OSError:
            parents = []
        for parent in parents:
            for pattern in ("resources/runtime/*/dependencies/node/bin/node",
                            "runtime/*/dependencies/node/bin/node",
                            "resources/runtime/*/dependencies/node/bin/node.exe",
                            "runtime/*/dependencies/node/bin/node.exe"):
                try:
                    found += [str(p) for p in parent.glob(pattern) if p.is_file()]
                except OSError:
                    pass
    return found


def discover_renderer(cli_override=None, node_override=None):
    """Return an ordered list of viable renderer descriptions.

    Each entry: ``{"name", "command", "kind", "strategy"}`` where ``kind`` is
    ``"kit-cli"`` (Node CLI) or ``"soffice"`` (LibreOffice binary).  Kit entries are
    emitted once per plausible node executable so a broken node does not mask a working one.
    """
    out = []

    def add_kit(cli_path, strategy, nodes):
        if not cli_path or not Path(cli_path).is_file():
            return
        tried = []
        for node in (nodes or []):
            if not node or not Path(node).is_file() or node in tried:
                continue
            tried.append(node)
            out.append({"name": "libreoffice-kit", "kind": "kit-cli", "strategy": strategy,
                        "command": [str(node), str(cli_path)]})

    def add_soffice(bin_path, strategy):
        if not bin_path or not Path(bin_path).is_file():
            return
        out.append({"name": "libreoffice-kit" if "libreoffice-kit" in str(bin_path).lower()
                    else "soffice",
                    "kind": "soffice", "strategy": strategy, "command": [str(bin_path)]})

    def nodes_for(cli_path):
        return ([node_override] if node_override else []) + \
               [n for n in (_node_exe(),) if n] + _nearby_node_candidates(cli_path)

    # 0. explicit command-line override (most deterministic)
    if cli_override:
        add_kit(cli_override, "--renderer-cli", nodes_for(cli_override))

    # 1. env: kit CLI entry point
    for var in KIT_CLI_ENV_VARS:
        val = os.environ.get(var)
        if val:
            add_kit(val, f"env:{var}", nodes_for(val))
            break

    # 2. env: direct LibreOffice / kit binary
    for var in BIN_ENV_VARS:
        val = os.environ.get(var)
        if val:
            add_soffice(val, f"env:{var}")
            break

    # 3. PATH
    add_soffice(_which(SOFFICE_NAMES), "PATH")
    kit_on_path = _which(("libreoffice-kit", "libreoffice-kit-cli"))
    if kit_on_path:
        add_kit(kit_on_path, "PATH", nodes_for(kit_on_path))

    # 4. known install directories
    for directory in _known_soffice_dirs():
        for name in SOFFICE_NAMES:
            cand = Path(directory) / name
            if cand.is_file():
                add_soffice(cand, "known-dir")
                break

    # 5. bounded search for the bundled kit CLI
    searched = _bounded_kit_search()
    if searched:
        add_kit(searched, "bounded-search", nodes_for(searched))

    # de-dup identical commands while keeping order
    seen, ordered = set(), []
    for entry in out:
        key = tuple(entry["command"])
        if key not in seen:
            seen.add(key)
            ordered.append(entry)
    return ordered


def _kill_tree(proc):
    """Kill *proc* and every descendant.

    The kit's Node CLI spawns a worker **grandchild**.  Killing only the Node process leaves
    that worker alive holding the inherited stdout pipe -- measured on this host as an
    indefinite hang, so the tree has to go.
    """
    try:
        pid = proc.pid
    except Exception:  # noqa: BLE001
        return
    if is_windows():
        try:
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)],  # noqa: S603,S607
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=30)
        except Exception:  # noqa: BLE001
            pass
    else:
        try:
            import signal
            os.killpg(os.getpgid(pid), signal.SIGKILL)
        except Exception:  # noqa: BLE001
            pass
    try:
        proc.kill()
    except Exception:  # noqa: BLE001
        pass


def _run(cmd, timeout, env=None, scratch=None):
    """Run *cmd*; return ``(returncode, stdout, stderr)`` and never raise.

    stdout/stderr go to **files**, not pipes.  With pipes, a surviving grandchild keeps the
    pipe handle open and ``communicate()`` blocks forever even after the timeout fires; with
    files, the timeout actually returns.  A timed-out command is killed as a whole tree.
    """
    import uuid
    scratch = Path(scratch) if scratch else Path.cwd()
    try:
        scratch.mkdir(parents=True, exist_ok=True)
    except OSError:
        scratch = Path.cwd()
    tag = uuid.uuid4().hex[:12]
    out_f = scratch / f".child-{tag}.out"
    err_f = scratch / f".child-{tag}.err"
    rc, note = None, ""
    kwargs = {}
    if is_windows():
        kwargs["creationflags"] = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    else:
        kwargs["start_new_session"] = True
    try:
        with open(out_f, "wb") as o, open(err_f, "wb") as e:
            try:
                proc = subprocess.Popen(  # noqa: S603
                    cmd, stdout=o, stderr=e, stdin=subprocess.DEVNULL, env=env, **kwargs)
                try:
                    rc = proc.wait(timeout=timeout)
                except subprocess.TimeoutExpired:
                    note = f"timeout after {timeout}s (process tree killed)"
                    _kill_tree(proc)
                    try:
                        proc.wait(timeout=15)
                    except Exception:  # noqa: BLE001
                        pass
                    rc = 124
            except FileNotFoundError as exc:
                rc, note = 127, f"not found: {exc}"
            except OSError as exc:
                rc, note = 126, f"os error: {exc}"
        def _read(path):
            try:
                return path.read_text("utf-8", "replace")
            except OSError:
                return ""
        out, err = _read(out_f), _read(err_f)
    finally:
        for path in (out_f, err_f):
            try:
                path.unlink()
            except OSError:
                pass
    if note:
        err = (err + "\n" + note).strip()
    return (rc if rc is not None else 126), out, err


# --------------------------------------------------------------------------- #
# per-strategy rendering
# --------------------------------------------------------------------------- #

def _kit_convert_pdf(renderer, pptx, pdf_path, timeout, log):
    cmd = renderer["command"] + ["convert", "--input", str(pptx), "--output", str(pdf_path)]
    log.append({"step": "kit.convert", "command": cmd})
    rc, out, err = _run(cmd, timeout, scratch=Path(pdf_path).parent)
    info = {}
    for line in out.splitlines():
        line = line.strip()
        if line.startswith("{"):
            try:
                info = json.loads(line)
            except ValueError:
                pass
    return rc, out, err, info


def _kit_render_png(renderer, pptx, outdir, timeout, log, pages=None):
    """Rasterise with the kit CLI.

    Measured on the bundled kit (2026-11-23): ``render`` creates ``--output-dir`` itself and
    fails with ``EEXIST: file already exists, mkdir`` when the directory is already there.
    A stale directory would therefore break every re-run, so we always render into a scratch
    directory that we clear first.
    """
    outdir = Path(outdir)
    if outdir.exists():
        try:
            shutil.rmtree(outdir)
        except OSError:
            pass
    cmd = renderer["command"] + ["render", "--input", str(pptx), "--output-dir", str(outdir)]
    if pages:
        cmd += ["--pages", ",".join(str(p) for p in pages)]
    log.append({"step": "kit.render", "command": cmd})
    rc, out, err = _run(cmd, timeout, scratch=outdir.parent)
    info = {}
    for line in out.splitlines():
        line = line.strip()
        if line.startswith("{"):
            try:
                info = json.loads(line)
            except ValueError:
                pass
    return rc, out, err, info


def _soffice_convert_pdf(renderer, pptx, pdf_path, timeout, log, workdir):
    """Convert with a LibreOffice binary, using a private user profile.

    On Windows a shared profile makes soffice fail when another instance is live, so a
    throwaway ``-env:UserInstallation`` directory is always supplied.
    """
    profile = Path(workdir) / ".lo-profile"
    try:
        profile.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass
    uri = profile.resolve().as_uri()
    outdir = Path(pdf_path).parent
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = renderer["command"] + [
        "--headless", "--norestore", "--nolockcheck", "--nodefault", "--nologo",
        f"-env:UserInstallation={uri}",
        "--convert-to", "pdf", "--outdir", str(outdir), str(pptx),
    ]
    log.append({"step": "soffice.convert", "command": cmd})
    rc, out, err = _run(cmd, timeout, scratch=outdir)
    produced = outdir / (Path(pptx).stem + ".pdf")
    if rc == 0 and produced.is_file() and produced != Path(pdf_path):
        try:
            shutil.move(str(produced), str(pdf_path))
        except OSError:
            return rc, out, err + "\ncould not move produced pdf into place"
    return rc, out, err


def _pdftoppm_pdf_to_png(pdf_path, outdir, timeout, log, dpi=144):
    """Rasterise a PDF with Poppler when the primary renderer cannot do PNG itself."""
    exe = shutil.which("pdftoppm") or shutil.which("pdftoppm.exe")
    if not exe:
        return None, "pdftoppm not found on PATH"
    prefix = Path(outdir) / "page"
    cmd = [exe, "-png", "-r", str(dpi), str(pdf_path), str(prefix)]
    log.append({"step": "pdftoppm", "command": cmd})
    rc, out, err = _run(cmd, timeout, scratch=Path(outdir))
    if rc != 0:
        return None, f"pdftoppm exit {rc}: {err.strip()[:200]}"
    produced = sorted(Path(outdir).glob("page*.png"))
    if not produced:
        return None, "pdftoppm wrote no png"
    return produced, None


# --------------------------------------------------------------------------- #
# image probing (pure stdlib PNG header parse + optional PIL blank check)
# --------------------------------------------------------------------------- #

def probe_png(path):
    """Return ``{width, height, bytes}`` from the PNG's own IHDR, not from a library claim."""
    info = {"path": str(path), "bytes": 0, "width": 0, "height": 0}
    try:
        info["bytes"] = os.path.getsize(path)
    except OSError:
        return info
    try:
        with open(path, "rb") as fh:
            head = fh.read(33)
        if head[:8] == b"\x89PNG\r\n\x1a\n" and head[12:16] == b"IHDR":
            info["width"] = int.from_bytes(head[16:20], "big")
            info["height"] = int.from_bytes(head[20:24], "big")
    except OSError:
        pass
    return info


def png_content_ratio(path):
    """Fraction of pixels that differ from the most common pixel value.

    A uniformly blank/transparent preview is a *renderer* failure, not a valid slide, so we
    measure it instead of trusting exit code 0.  Returns ``None`` when PIL is unavailable
    (the honest answer is "not measured", not 1.0).
    """
    try:
        from PIL import Image  # type: ignore
    except Exception:  # noqa: BLE001
        return None
    try:
        with Image.open(path) as img:
            small = img.convert("RGB")
            small.thumbnail((160, 160))
            colors = small.getcolors(maxcolors=1 << 20)
        if not colors:
            return None
        total = sum(count for count, _ in colors)
        if not total:
            return None
        modal = max(count for count, _ in colors)
        return round(1.0 - (modal / total), 6)
    except Exception:  # noqa: BLE001
        return None


def pdf_page_count(path):
    """Page count via pypdf when available, else a stdlib ``/Type /Page`` count."""
    try:
        from pypdf import PdfReader  # type: ignore
        try:
            return len(PdfReader(str(path)).pages)
        except Exception:  # noqa: BLE001
            pass
    except Exception:  # noqa: BLE001
        pass
    try:
        data = Path(path).read_bytes()
        return data.count(b"/Type /Page") - data.count(b"/Type /Pages")
    except OSError:
        return 0


def pdf_fonts_embedded(path):
    """Font names actually embedded in the produced PDF, read out of ``/Resources /Font``.

    The kit tells us which font families it could *not* find; only the PDF itself shows which
    face was substituted in.  Reading it back converts "the renderer says a font is missing"
    into measured evidence, which is what makes the font-fidelity limitation checkable rather
    than asserted.  Returns ``None`` when pypdf is unavailable (not measured != no fonts).
    """
    try:
        from pypdf import PdfReader  # type: ignore
    except Exception:  # noqa: BLE001
        return None
    names = set()
    try:
        reader = PdfReader(str(path))
        for page in reader.pages[:12]:
            res = page.get("/Resources")
            if res is None:
                continue
            res = res.get_object()
            fonts = res.get("/Font")
            if fonts is None:
                continue
            fonts = fonts.get_object()
            for key in fonts:
                try:
                    f = fonts[key].get_object()
                except Exception:  # noqa: BLE001
                    continue
                base = f.get("/BaseFont")
                if base:
                    names.add(str(base))
                desc = f.get("/FontDescriptor")
                if desc is not None:
                    try:
                        desc = desc.get_object()
                        nm = desc.get("/FontName")
                        if nm:
                            names.add(str(nm))
                    except Exception:  # noqa: BLE001
                        pass
    except Exception:  # noqa: BLE001
        return None
    return sorted(names)


def pptx_slide_count(pptx):
    try:
        from pptx import Presentation  # type: ignore
        return len(Presentation(str(pptx)).slides)
    except Exception:  # noqa: BLE001
        return 0


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #

def build_report(pptx, outdir=None, report_out=None, workdir=None,
                 renderer_cli=None, node=None, pages=None, dpi=144, timeout=150,
                 pdf_only=False, total_budget=300):
    pptx = Path(pptx)
    outdir = Path(outdir) if outdir else pptx.parent / "render"
    # A renderer that hangs must not consume the whole execution budget: every attempt is
    # capped by both `timeout` and what is left of `total_budget`.
    import time as _time
    deadline = _time.time() + max(30, int(total_budget))

    def remaining():
        return max(0, deadline - _time.time())
    report = {
        "schema": SCHEMA,
        "renderer": {"name": "none", "version": "", "command": "", "available": False},
        "pdf": {"path": "", "pages": 0, "bytes": 0, "sha256": ""},
        "pngs": [],
        "status": "skipped",
        "error": None,
        "uncertainty": [],
    }
    log = []

    if not pptx.is_file():
        report["status"] = "failed"
        report["error"] = f"deck not found: {pptx}"
        report["uncertainty"].append(
            "Deck file was not present, so no PDF/PNG preview could be produced.")
        if report_out:
            dump_json(report_out, report)
        return report, log

    try:
        outdir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        report["status"] = "failed"
        report["error"] = f"cannot create render dir {outdir}: {exc}"
        report["uncertainty"].append(f"Render directory is not writable: {exc}")
        if report_out:
            dump_json(report_out, report)
        return report, log

    # Stale previews from an earlier run must never be mistaken for this run's output:
    # the report is the only authority on what was produced, so old files are cleared.
    for stale in list(outdir.glob("slide-*.png")):
        try:
            stale.unlink()
        except OSError:
            pass

    candidates = discover_renderer(renderer_cli, node)
    if not candidates:
        report["status"] = "skipped"
        report["error"] = None
        report["uncertainty"].append(
            "No LibreOffice renderer was found (checked --renderer-cli, "
            "LO_KIT_CLI/LIBREOFFICE_KIT_CLI, SOFFICE_BIN/LIBREOFFICE_KIT_BIN/LO_KIT_BIN, "
            "PATH, known install directories and a bounded search). PDF/PNG previews were "
            "NOT produced; this is reported as skipped rather than faked.")
        report["_discovery"] = {
            "strategies_tried": ["override", "env:kit-cli", "env:bin", "PATH",
                                 "known-dir", "bounded-search"],
            "node": _node_exe() or None,
        }
        log.append({"step": "discovery", "result": "none"})
        if report_out:
            dump_json(report_out, report)
        return report, log

    pdf_path = outdir / "deck.pdf"
    errors = []
    png_source = None      # list of raw page png paths
    kit_manifest = {}
    chosen = None
    # ONE font measurement, recorded per call, reported once (see the note below the loop).
    font_report = {"kit_convert_missingFonts": None, "kit_render_missingFonts": None}

    def _record_renderer(cand, info):
        report["renderer"] = {
            "name": cand["name"], "version": str(info.get("backend") or info.get("version") or ""),
            "available": True, "strategy": cand["strategy"],
            "command": " ".join(str(c) for c in cand["command"]),
        }

    # --- PDF: try each discovered renderer until one produces a non-empty file ---
    for cand in candidates:
        log.append({"step": "try", "renderer": cand["name"], "strategy": cand["strategy"],
                    "command": " ".join(str(c) for c in cand["command"])})
        if remaining() <= 5:
            errors.append(f"render budget of {int(total_budget)}s exhausted before "
                          f"{cand['strategy']} could be tried")
            break
        # Measured on the bundled kit (2026-11-23): `convert` refuses to overwrite and fails
        # with EEXIST.  Clear the target so repeated runs into the same work dir behave.
        try:
            if pdf_path.exists():
                pdf_path.unlink()
        except OSError:
            pass
        attempt_timeout = min(timeout, max(5, int(remaining())))
        if cand["kind"] == "kit-cli":
            rc, out, err, info = _kit_convert_pdf(cand, pptx, pdf_path, attempt_timeout, log)
        else:
            rc, out, err, info = _soffice_convert_pdf(cand, pptx, pdf_path, attempt_timeout, log,
                                                     workdir or outdir)
        if rc == 124:
            errors.append(f"{cand['strategy']}/{cand['name']} timed out after "
                          f"{attempt_timeout}s (the renderer hung; process tree killed)")
        if rc == 0 and pdf_path.is_file() and pdf_path.stat().st_size > 0:
            report["pdf"] = {
                "path": str(pdf_path),
                "pages": pdf_page_count(pdf_path),
                "bytes": pdf_path.stat().st_size,
                "sha256": sha256_file(pdf_path),
            }
            font_report["kit_convert_missingFonts"] = info.get("missingFonts")
            _record_renderer(cand, info)
            chosen = cand
            break
        if rc != 124:
            errors.append(f"{cand['strategy']}/{cand['name']} pdf export failed "
                          f"(exit {rc}): {(err or out).strip()[:200]}")
        try:
            if pdf_path.is_file() and pdf_path.stat().st_size == 0:
                pdf_path.unlink()
        except OSError:
            pass

    # --- PNG: the kit can rasterise page by page; otherwise fall back to Poppler ---
    slide_count = pptx_slide_count(pptx)
    scratch_png = outdir / ".kit-render"
    if not pdf_only:
        for cand in candidates:
            if cand["kind"] != "kit-cli":
                continue
            if remaining() <= 5:
                errors.append(f"render budget of {int(total_budget)}s exhausted before "
                              f"{cand['strategy']} rasterisation could be tried")
                break
            attempt_timeout = min(timeout, max(5, int(remaining())))
            rc2, out2, err2, man = _kit_render_png(cand, pptx, scratch_png, attempt_timeout,
                                                  log, pages)
            if rc2 == 0 and man.get("images"):
                kit_manifest = man
                font_report["kit_render_missingFonts"] = man.get("missingFonts")
                png_source = [Path(img["path"]) for img in man["images"]]
                if chosen is None:
                    _record_renderer(cand, man)
                    chosen = cand
                break
            if rc2 == 124:
                errors.append(f"{cand['strategy']}/kit png render timed out after "
                              f"{attempt_timeout}s (process tree killed)")
            else:
                errors.append(f"{cand['strategy']}/kit png render failed "
                              f"(exit {rc2}): {(err2 or out2).strip()[:200]}")
        if not png_source and report["pdf"]["path"]:
            produced, why = _pdftoppm_pdf_to_png(pdf_path, outdir, min(timeout, max(5, int(remaining()) or 5)),
                                                 log, dpi)
            if produced:
                png_source = produced
            elif why:
                errors.append(why)

    # --- normalise names: page-NNNN.png -> slide-NN.png ---
    if png_source:
        for idx, src in enumerate(sorted(png_source, key=lambda p: str(p)), start=1):
            target = outdir / f"slide-{idx:02d}.png"
            try:
                shutil.copyfile(src, target)
            except OSError as exc:
                errors.append(f"could not place {target.name}: {exc}")
                continue
            probed = probe_png(target)
            entry = {"slide": idx, "path": str(target), "bytes": probed["bytes"],
                     "width": probed["width"], "height": probed["height"]}
            ratio = png_content_ratio(target)
            if ratio is not None:
                entry["nonbackground_ratio"] = ratio
            report["pngs"].append(entry)
        # the scratch raster directory is an implementation detail; do not ship it
        try:
            if scratch_png.is_dir():
                shutil.rmtree(scratch_png)
        except OSError:
            pass
        for stray in outdir.glob("page-*.png"):
            try:
                stray.unlink()
            except OSError:
                pass

    # --- status ---
    have_pdf = bool(report["pdf"]["path"])
    have_png = bool(report["pngs"])
    if have_pdf and have_png:
        report["status"] = "ok"
    elif have_pdf or have_png:
        report["status"] = "partial"
    else:
        report["status"] = "failed"

    if errors:
        report["error"] = " | ".join(errors)[:1500]
    if report["status"] != "ok":
        report["uncertainty"].append(
            f"Rendering did not fully succeed (status={report['status']}): "
            + (report["error"] or "no renderer output")[:400])

    # --- cross-checks that make the previews trustworthy or flag them ---
    if have_png:
        if slide_count and len(report["pngs"]) != slide_count:
            report["uncertainty"].append(
                f"Rendered page count ({len(report['pngs'])}) differs from the slide count "
                f"({slide_count}); the previews may not map 1:1 onto slides.")
        blank = [p["slide"] for p in report["pngs"]
                 if p.get("nonbackground_ratio") is not None and p["nonbackground_ratio"] < 5e-4]
        if blank:
            report["blank_slides_suspected"] = blank
            report["uncertainty"].append(
                "These previews are almost entirely a single colour and may be blank renders "
                f"(renderer failure, not an empty slide): slides {blank}")
        if all(p.get("nonbackground_ratio") is not None for p in report["pngs"]) and \
                all(p["nonbackground_ratio"] < 5e-4 for p in report["pngs"]):
            report["uncertainty"].append(
                "Every preview is essentially uniform; treat the whole render as failed and "
                "rely on the geometry checks in verify_deck.py instead of the images.")
    # --- fonts: ONE field, derived from ONE set of measurements ---------------------------
    # An earlier version emitted both `missing_fonts` (from the PNG manifest) and
    # `_missing_fonts` (from the PDF export), which read as a self-contradiction: one of them
    # looked like "no fonts missing" while the uncertainty text said otherwise.  Now the
    # per-call values are kept together under `_font_evidence`, and `missing_fonts` is the
    # union, always present as a (possibly empty) list.
    missing = set()
    for value in font_report.values():
        if isinstance(value, list):
            missing.update(str(v) for v in value)
    report["missing_fonts"] = sorted(missing)
    embedded = pdf_fonts_embedded(pdf_path) if report["pdf"]["path"] else None
    report["_font_evidence"] = dict(font_report)
    report["_font_evidence"]["pdf_fonts_embedded"] = embedded
    if embedded is None and report["pdf"]["path"]:
        report["_font_evidence"]["pdf_fonts_embedded_note"] = (
            "pypdf unavailable, so the fonts actually embedded in the PDF could not be read "
            "back; the substitution list is the renderer's own report only")
    if missing:
        detail = ("The renderer could not find these font families: "
                  + ", ".join(sorted(missing))
                  + " -- glyphs are therefore substituted, so the previews are a substitute "
                    "render and must not be read as a certificate of font fidelity.")
        if embedded:
            detail += (" Measured: the exported PDF embeds " + ", ".join(embedded)
                       + " instead.")
        report["uncertainty"].append(detail)
    elif embedded:
        # Positive case: say what was measured rather than staying silent.
        report["uncertainty"].append(
            "Fonts: the renderer reported no missing family and the exported PDF embeds "
            + ", ".join(embedded) + ", so no substitution was indicated for this render.")
    report["_log"] = log
    if report_out:
        dump_json(report_out, report)
    return report, log


def main(argv=None):
    ap = argparse.ArgumentParser(description="Render the built deck to PDF and PNG previews.")
    ap.add_argument("--case", default=None, help="case.json (used for case_id only)")
    ap.add_argument("--work", default=None, help="work dir holding deck.pptx / render/")
    ap.add_argument("--in", dest="inp", default=None, help="deck.pptx to render")
    ap.add_argument("--out", default=None,
                    help="output dir, or the render_report.json path")
    ap.add_argument("--renderer-cli", default=None, help="explicit path to kit lib/cli.js")
    ap.add_argument("--node", default=None, help="explicit path to the node executable")
    ap.add_argument("--pages", default=None, help="comma separated 1-based pages to rasterise")
    ap.add_argument("--dpi", type=int, default=144)
    ap.add_argument("--timeout", type=int, default=150,
                    help="per-attempt timeout; the renderer's process tree is killed on expiry")
    ap.add_argument("--budget", type=int, default=300,
                    help="total seconds allowed for all render attempts, so a hung renderer "
                         "cannot consume the whole execution budget")
    ap.add_argument("--pdf-only", action="store_true")
    args = ap.parse_args(argv)

    work = Path(args.work) if args.work else Path.cwd()
    pptx = Path(args.inp) if args.inp else work / "deck.pptx"
    if not args.inp:
        build_report_json = work / "build_report.json"
        data = load_json(build_report_json)
        if isinstance(data, dict) and data.get("pptx"):
            cand = Path(data["pptx"])
            if cand.is_file():
                pptx = cand

    outdir, report_out = work / "render", work / "render_report.json"
    if args.out:
        if str(args.out).lower().endswith(".json"):
            report_out = Path(args.out)
            outdir = report_out.parent / "render"
        else:
            outdir = Path(args.out)
            report_out = work / "render_report.json"

    pages = None
    if args.pages:
        try:
            pages = [int(p) for p in str(args.pages).replace(" ", "").split(",") if p]
        except ValueError:
            pages = None

    report, _ = build_report(pptx, outdir=outdir, report_out=report_out, workdir=work,
                             renderer_cli=args.renderer_cli, node=args.node, pages=pages,
                             dpi=args.dpi, timeout=args.timeout, pdf_only=args.pdf_only,
                             total_budget=args.budget)
    print(json.dumps({"status": report["status"],
                      "renderer": report["renderer"]["name"],
                      "strategy": report["renderer"].get("strategy", ""),
                      "pdf": report["pdf"]["path"],
                      "pngs": len(report["pngs"]),
                      "report": str(report_out)}, ensure_ascii=False))
    # exit 0 when previews exist; non-zero is reserved for "we produced nothing at all",
    # which run.py must surface as a partial result rather than a crash.
    return 0 if report["status"] in ("ok", "partial") else 1


if __name__ == "__main__":
    raise SystemExit(main())
