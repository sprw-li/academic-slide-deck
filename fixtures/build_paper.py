#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build the SYNTHETIC fixture papers for ARMI:2611.000074v1.

SYNTHETIC / NOT AN OFFICIAL CASE.  Every number in every paper built here is
invented for the fixture.  The numbers are internally consistent inside a case
(abstract, prose, tables and figure data agree) but they are NOT measurements
and must never be presented as real experimental or literature values.

What this script builds (per case):

    caseA/  experimental paper   -> materials/paper.docx, materials/paper.pdf,
                                    materials/figure{1..4}-*.png,
                                    case.json, task.md
    caseB/  review/theory paper  -> materials/paper.docx, materials/paper.pdf,
                                    materials/figure{1..3}-*.png,
                                    case.json, task.md

The .pdf is produced from the .docx with the DSH-bundled LibreOffice Kit Node
wrapper, so the PDF gets a real text layer AND real embedded raster images
(that is what makes figure "source-crop" possible downstream).

This file lives under ``work/fixtures/`` and is a build script, not part of the
shipped submission, so it may contain absolute host paths.  Nothing here is
copied into ``work/submission/``.

Usage::

    python build_paper.py                 # build both cases
    python build_paper.py --case caseA    # rebuild one case
    python build_paper.py --check         # build and re-verify the artifacts
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

os.environ.setdefault("PYTHONDONTWRITEBYTECODE", "1")

HERE = Path(__file__).resolve().parent                 # .../work/fixtures
REPO = HERE.parents[1]                                 # .../ARMI2611000074v1

CONTROOT = REPO.parents[0]                             # D:\Borium_Contest

FONT_TREE = CONTROOT / "SKILL_Downloaded" / "repos" / "06-fonts-icons" / "liberation"
SERIF_TTF = {
    "regular": FONT_TREE / "liberationserifregular" / "LiberationSerif-Regular.ttf",
    "bold": FONT_TREE / "liberationserifbold" / "LiberationSerif-Bold.ttf",
    "italic": FONT_TREE / "liberationserifitalic" / "LiberationSerif-Italic.ttf",
    "bolditalic": FONT_TREE / "liberationserifbolditalic" / "LiberationSerif-BoldItalic.ttf",
}

NODE = Path(r"D:\App\DeepSeek Harness\resources\runtime\primary-runtime\dependencies\node\bin\node.exe")
KIT_CLI = Path(r"D:\App\DeepSeek Harness\resources\app.asar.unpacked\dsh\node_modules"
               r"\@deepseek-ai\libreoffice-kit\lib\cli.js")

TEMPLATE = HERE / "template-organizer.pptx"
TEMPLATE_REL = "materials/template-organizer.pptx"

SYNTHETIC_BANNER = "SYNTHETIC / NOT AN OFFICIAL CASE"


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def log(msg: str) -> None:
    print("[build_paper] %s" % msg, flush=True)


def register_matplotlib_fonts():
    """Make Liberation Serif available to matplotlib for this process only."""
    import matplotlib
    matplotlib.use("Agg")
    from matplotlib import font_manager
    missing = [str(p) for p in SERIF_TTF.values() if not p.is_file()]
    if missing:
        raise SystemExit("missing Liberation Serif TTF(s): %s" % ", ".join(missing))
    for path in SERIF_TTF.values():
        font_manager.fontManager.addfont(str(path))
    matplotlib.rcParams["font.family"] = ["Liberation Serif"]
    matplotlib.rcParams["font.serif"] = ["Liberation Serif"]
    matplotlib.rcParams["savefig.dpi"] = 200
    matplotlib.rcParams["figure.dpi"] = 200
    return matplotlib


# --------------------------------------------------------------------------- #
# figures (matplotlib, Liberation Serif registered at runtime)
# --------------------------------------------------------------------------- #
ACCENT = "#1F6FEB"      # same accent as the organizer template (theme accent1)
ACCENT2 = "#B3541E"
INK = "#1A1A1A"


def _finish(fig, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.tight_layout()
    fig.savefig(str(path), facecolor="white")
    import matplotlib.pyplot as plt
    plt.close(fig)


def fig_morphology(path: Path) -> None:
    """Figure 1 (case A): synthesis route schematic. Synthetic illustration."""
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyArrow, FancyBboxPatch
    fig, ax = plt.subplots(figsize=(6.4, 2.5))
    ax.set_xlim(0, 10); ax.set_ylim(0, 3.2); ax.axis("off")
    boxes = [(0.15, "Ti(OiPr)4\n+ urea"), (2.35, "sol-gel\npH 3.2"),
             (4.55, "dry 80 C\n12 h"), (6.75, "calcine\n500 C, 4 h"),
             (8.75, "N-TiO2\n12.3 nm")]
    for x, label in boxes:
        ax.add_patch(FancyBboxPatch((x, 1.0), 1.5, 1.2, boxstyle="round,pad=0.06",
                                    linewidth=1.2, edgecolor=ACCENT, facecolor="#EAF1FB"))
        ax.text(x + 0.75, 1.6, label, ha="center", va="center", fontsize=8.5, color=INK)
    for x in (1.65, 3.85, 6.05, 8.25):
        ax.add_patch(FancyArrow(x, 1.6, 0.68, 0, width=0.06, head_width=0.18,
                                head_length=0.16, length_includes_head=True, color=ACCENT2))
    ax.text(0.15, 2.72, "Synthetic schematic - NOT a real micrograph",
            fontsize=7.5, style="italic", color="#666666")
    _finish(fig, path)


def fig_degradation(path: Path) -> None:
    """Figure 2 (case A): C/C0 vs time, synthetic first-order decay."""
    import matplotlib.pyplot as plt
    import numpy as np
    t = np.array([0, 10, 20, 30, 40, 50, 60], dtype=float)
    c_over_c0 = np.exp(-0.0312 * t)
    fig, ax = plt.subplots(figsize=(5.2, 3.4))
    ax.plot(t, c_over_c0, "o-", color=ACCENT, label="N-TiO2 (k = 0.0312 min$^{-1}$)")
    ax.plot(t, np.exp(-0.0061 * t), "s--", color=ACCENT2, label="P25 reference (k = 0.0061 min$^{-1}$)")
    ax.set_xlabel("Irradiation time (min)"); ax.set_ylabel("C / C$_0$")
    ax.set_xlim(-1, 61); ax.set_ylim(0, 1.05)
    ax.legend(fontsize=7.5, frameon=False); ax.grid(alpha=0.25)
    _finish(fig, path)


def fig_kinetics(path: Path) -> None:
    """Figure 3 (case A): pseudo-first-order fits, R2 = 0.994."""
    import matplotlib.pyplot as plt
    import numpy as np
    t = np.array([0, 10, 20, 30, 40, 50, 60], dtype=float)
    y = 0.0312 * t
    rng = np.random.default_rng(20261123)
    fig, ax = plt.subplots(figsize=(5.2, 3.4))
    ax.plot(t, y, "-", color=INK, linewidth=1.0, label="fit, k = 0.0312 min$^{-1}$")
    ax.plot(t, y + rng.normal(0, 0.012, t.size), "o", color=ACCENT, label="N-TiO2, R$^2$ = 0.994")
    ax.set_xlabel("Irradiation time (min)"); ax.set_ylabel("ln(C$_0$/C)")
    ax.legend(fontsize=7.5, frameon=False); ax.grid(alpha=0.25)
    _finish(fig, path)


def fig_cycles(path: Path) -> None:
    """Figure 4 (case A): cycling stability over five runs, 88.1 % retained."""
    import matplotlib.pyplot as plt
    runs = [1, 2, 3, 4, 5]
    removal = [92.4, 91.1, 90.0, 88.9, 81.4]
    fig, ax = plt.subplots(figsize=(5.2, 3.2))
    ax.bar(runs, removal, color=ACCENT, width=0.6)
    for x, y in zip(runs, removal):
        ax.text(x, y + 0.8, "%.1f" % y, ha="center", fontsize=7.5, color=INK)
    ax.set_xlabel("Reuse cycle"); ax.set_ylabel("MB removal after 60 min (%)")
    ax.set_ylim(0, 100); ax.grid(axis="y", alpha=0.25)
    _finish(fig, path)


def fig_selection(path: Path) -> None:
    """Figure 1 (case B): literature screening funnel, synthetic counts."""
    import matplotlib.pyplot as plt
    stages = ["records\nidentified\n(n = 412)", "title/abstract\nscreened\n(n = 412)",
              "full text\nassessed\n(n = 96)", "included in\nreview\n(n = 12)"]
    widths = [412, 412, 96, 12]
    fig, ax = plt.subplots(figsize=(5.6, 3.4))
    for i, (w, label) in enumerate(zip(widths, stages)):
        half = 0.5 * (w / 412.0) * 4.6
        ax.barh(-i, 2 * half, left=-half, height=0.62, color=ACCENT,
                alpha=0.85 - 0.12 * i, edgecolor="white")
        ax.text(0, -i, label, ha="center", va="center", fontsize=7.5, color="white")
    ax.set_ylim(-3.7, 0.7); ax.set_xlim(-3, 3); ax.axis("off")
    ax.text(-2.9, 0.45, "Synthetic counts - NOT a real systematic review",
            fontsize=7.5, style="italic", color="#666666")
    _finish(fig, path)


def fig_mae(path: Path) -> None:
    """Figure 2 (case B): reported MAE per method family, synthetic values."""
    import matplotlib.pyplot as plt
    families = ["GAP-MTP", "ACE-EAM", "MTP-SOAP", "ACE-NN", "EAM (classical)"]
    mae = [0.043, 0.061, 0.058, 0.079, 0.186]
    fig, ax = plt.subplots(figsize=(5.6, 3.0))
    ax.barh(families, mae, color=ACCENT)
    for i, v in enumerate(mae):
        ax.text(v + 0.004, i, "%.3f eV" % v, va="center", fontsize=7.5, color=INK)
    ax.set_xlabel("Reported MAE for vacancy formation energy (eV)")
    ax.set_xlim(0, 0.23); ax.grid(axis="x", alpha=0.25)
    _finish(fig, path)


def fig_parity(path: Path) -> None:
    """Figure 3 (case B): MLIP vs DFT vacancy energetics parity, synthetic."""
    import matplotlib.pyplot as plt
    import numpy as np
    rng = np.random.default_rng(74)
    dft = rng.uniform(0.9, 4.1, 64)
    mlip = dft + rng.normal(0, 0.045, dft.size)
    fig, ax = plt.subplots(figsize=(4.4, 4.0))
    ax.plot([0.8, 4.2], [0.8, 4.2], "--", color="#888888", linewidth=1.0)
    ax.plot(dft, mlip, "o", color=ACCENT, markersize=3.6)
    ax.set_xlabel("DFT formation energy (eV)"); ax.set_ylabel("MLIP formation energy (eV)")
    ax.set_xlim(0.8, 4.2); ax.set_ylim(0.8, 4.2); ax.grid(alpha=0.25)
    return _finish(fig, path)


# --------------------------------------------------------------------------- #
# paper content (all values invented for the fixture)
# --------------------------------------------------------------------------- #
T_A = {
    "template_text": "Organizer Academic 2026 - SYNTHETIC TEMPLATE",
    "title": "Nitrogen-Doped TiO2 Photocatalysts for Visible-Light Degradation of "
             "Methylene Blue: A Kinetic Study",
    "authors": "Synthetic Author Group \u2014 FIXTURE TEXT, NOT A REAL PUBLICATION",
    "banner": SYNTHETIC_BANNER,
}
T_B = {
    "template_text": "Organizer Academic 2026 - SYNTHETIC TEMPLATE",
    "title": "Machine-Learned Interatomic Potentials for Point-Defect Energetics in "
             "Metals: A Critical Review",
    "authors": "Synthetic Author Group \u2014 FIXTURE TEXT, NOT A REAL PUBLICATION",
    "banner": SYNTHETIC_BANNER,
}

DOC_A = [
    ("h1", "Abstract"),
    ("p", "Nitrogen-doped titanium dioxide (N-TiO2) powders were prepared by a urea-assisted "
          "sol-gel route and calcined at 500 \u00b0C for 4 h. The resulting material had a "
          "measured BET surface area of 78.6 m2/g, a crystallite size of 12.3 nm and an "
          "optical band gap of 2.84 eV, compared with 3.18 eV for the undoped reference. "
          "Under visible light (lambda > 420 nm, 300 W Xe lamp) the material removed "
          "92.4 % of methylene blue (MB, 20 mg/L, catalyst dose 0.75 g/L) within 60 min. "
          "The decay followed pseudo-first-order kinetics with a rate constant of "
          "0.0312 min-1 and R2 = 0.994. Five consecutive reuse cycles retained 88.1 % of "
          "the initial activity. Every value in this paragraph is invented for the fixture."),
    ("h1", "1. Introduction"),
    ("p", "Titanium dioxide is the most widely studied photocatalyst for aqueous pollutant "
          "degradation, but its 3.18 eV band gap restricts photoexcitation to the ultraviolet. "
          "Doping with nitrogen has been proposed as a route to visible-light activity, and "
          "reported band-gap reductions in the literature span 0.2-0.5 eV. Methylene blue is a "
          "convenient probe dye for such studies because its concentration is easy to follow "
          "spectrophotometrically at 664 nm."),
    ("p", "This work reports a single N-TiO2 series prepared at one calcination temperature "
          "(500 \u00b0C) and compares it with a commercial P25 reference under identical "
          "illumination conditions. The specific question is whether the visible-light rate "
          "constant gain is large enough to survive reuse."),
    ("h1", "2. Materials and Methods"),
    ("p", "Titanium tetraisopropoxide (Ti(OiPr)4, 97 %) was hydrolysed in ethanol at pH 3.2 "
          "with urea added at a urea/Ti molar ratio of 2.0. The gel was dried at 80 \u00b0C for "
          "12 h and calcined in static air at 500 \u00b0C for 4 h with a ramp of 5 \u00b0C/min. "
          "X-ray diffraction used Cu K-alpha radiation over 2theta = 20-80 degrees."),
    ("p", "Photocatalytic tests used 100 mL of 20 mg/L MB at a catalyst dose of 0.75 g/L in a "
          "jacketed reactor at 25 \u00b0C with a 420 nm cut-off filter and a 300 W Xe lamp. "
          "Aliquots of 3 mL were taken at 0, 10, 20, 30, 40, 50 and 60 min and centrifuged at "
          "9000 rpm for 5 min. Absorbance at 664 nm was converted to concentration with a "
          "seven-point calibration (R2 = 0.999). Each condition was run three times (n = 3); "
          "error bars in Figure 2 are +/- 1 standard deviation."),
    ("h1", "3. Results"),
    ("p", "Phase analysis detected anatase only; no rutile or TiN reflections appeared at the "
          "0.5 % detection limit. The nitrogen content from CHN analysis was 3.4 at.%. "
          "Diffuse-reflectance spectra gave a Tauc band gap of 2.84 eV for N-TiO2 versus "
          "3.18 eV for the undoped control, a shift of 0.34 eV."),
    ("figure", "figure1-morphology.png",
     "Figure 1. Synthetic schematic of the urea-assisted sol-gel route used for the N-TiO2 "
     "series (calcination 500 \u00b0C, 4 h). This panel is a drawing, not a micrograph."),
    ("p", "Surface area increased from 41.2 m2/g (undoped) to 78.6 m2/g (N-TiO2), and the "
          "mean crystallite size from Rietveld refinement fell from 18.9 nm to 12.3 nm. "
          "Pore volume was 0.21 cm3/g with a mean pore diameter of 9.4 nm."),
    ("table", "Table 1. Composition and textural properties of the synthetic series.",
     [["Sample", "N (at.%)", "BET (m2/g)", "d_XRD (nm)", "Eg (eV)"],
      ["Undoped TiO2", "0.0", "41.2", "18.9", "3.18"],
      ["N-TiO2-0.5", "0.9", "63.4", "14.6", "3.02"],
      ["N-TiO2-1.0", "1.8", "71.9", "13.1", "2.91"],
      ["N-TiO2-2.0", "3.4", "78.6", "12.3", "2.84"]]),
    ("p", "Under visible light the best sample removed 92.4 % of MB in 60 min, whereas the "
          "P25 reference removed 30.7 % over the same period. The corresponding "
          "pseudo-first-order rate constant was 0.0312 min-1 for N-TiO2-2.0 and 0.0061 min-1 "
          "for P25, a five-fold difference. The fit quality was R2 = 0.994 over seven points."),
    ("figure", "figure2-degradation.png",
     "Figure 2. Normalised methylene blue concentration versus irradiation time for N-TiO2-2.0 "
     "(k = 0.0312 min-1) and the P25 reference (k = 0.0061 min-1); n = 3, error bars +/- 1 SD."),
    ("figure", "figure3-kinetics.png",
     "Figure 3. Pseudo-first-order transformation of the N-TiO2-2.0 data; the slope gives "
     "k = 0.0312 min-1 with R2 = 0.994."),
    ("table", "Table 2. Kinetic parameters from the synthetic degradation runs.",
     [["Sample", "k (min-1)", "R2", "Removal at 60 min (%)", "t1/2 (min)"],
      ["Undoped TiO2", "0.0021", "0.961", "11.8", "330"],
      ["N-TiO2-2.0", "0.0312", "0.994", "92.4", "22.2"],
      ["P25 reference", "0.0061", "0.988", "30.7", "114"]]),
    ("p", "Reuse tests were run without regeneration between cycles. Removal decayed from "
          "92.4 % in cycle 1 to 81.4 % in cycle 5, so the material retained 88.1 % of its "
          "initial activity. The nitrogen content after five cycles was 2.6 at.%, a loss of "
          "0.8 at.% relative to the fresh sample."),
    ("figure", "figure4-cycles.png",
     "Figure 4. Removal of methylene blue after 60 min over five consecutive reuse cycles "
     "(92.4, 91.1, 90.0, 88.9 and 81.4 %)."),
    ("h1", "4. Discussion"),
    ("p", "The 0.34 eV band-gap narrowing is smaller than the 0.5 eV upper end of the range "
          "reported in the literature, but it is sufficient to place absorption inside the "
          "visible window used here (lambda > 420 nm, i.e. photon energy below 2.95 eV). "
          "The simultaneous increase in surface area from 41.2 to 78.6 m2/g makes it "
          "impossible to attribute the five-fold rate gain to the band gap alone; a control "
          "series at matched surface area would be required to separate the two effects."),
    ("p", "The rate constants in Table 2 differ by more than the pooled standard deviation of "
          "the triplicate runs (relative SD 3.1 % for N-TiO2-2.0). A two-sided t-test between "
          "N-TiO2-2.0 and P25 gave p = 0.008, which we report as an indication only, because "
          "n = 3 is small and the samples are not independent."),
    ("h1", "5. Limitations"),
    ("p", "The study uses a single dye, a single catalyst dose and one calcination "
          "temperature. Mineralisation was not measured; only decolourisation was followed at "
          "664 nm, so adsorption on the catalyst surface cannot be separated from true "
          "photocatalytic degradation. The nitrogen loss of 0.8 at.% over five cycles is "
          "within the reproducibility of the CHN measurement (+/- 0.3 at.%) and should not be "
          "over-interpreted."),
    ("h1", "References"),
    ("p", "[1] Synthetic reference A. Fictional Journal of Photocatalysis 12 (2024) 101-110. "
          "[2] Synthetic reference B. Fictional Materials Letters 88 (2023) 55-60. "
          "[3] Synthetic reference C. Fictional Applied Catalysis Bulletin 7 (2022) 1-19. "
          "These references are invented placeholders for the fixture."),
]

DOC_B = [
    ("h1", "Abstract"),
    ("p", "Machine-learned interatomic potentials (MLIPs) are increasingly used to compute "
          "point-defect energetics in metals at a fraction of the cost of density functional "
          "theory (DFT). This synthetic review compares twelve reported studies covering six "
          "face-centred-cubic metals. Across the included studies the mean absolute error "
          "(MAE) for monovacancy formation energies ranged from 0.043 eV (GAP-MTP) to "
          "0.186 eV (classical EAM), against a DFT self-consistency floor of 0.020 eV. "
          "Migration barriers, by contrast, were reproduced much less uniformly: reported MAE "
          "values span 0.058 to 0.310 eV over a barrier range of 0.8 to 4.2 eV. We conclude "
          "that vacancy energetics are close to being a solved problem for MLIPs, while "
          "migration barriers are not. All counts and values here are invented for the "
          "fixture."),
    ("h1", "1. Introduction"),
    ("p", "Point defects control diffusion, creep and radiation damage in metals, and their "
          "energetics are traditionally obtained from DFT supercell calculations. A single "
          "64-atom vacancy calculation costs hours of wall time, which makes the "
          "configuration space of realistic defect populations inaccessible. MLIPs trained on "
          "DFT data promise the same accuracy at 10^3 to 10^5 times lower cost."),
    ("p", "The literature, however, reports accuracy in inconsistent ways: some studies quote "
          "MAE over formation energies, others over migration barriers, and a few quote only "
          "the maximum deviation. This review asks a narrow question: for which quantity is "
          "the MLIP error small enough to be ignored relative to the DFT reference error?"),
    ("h1", "2. Scope and Literature Selection"),
    ("p", "A synthetic search returned 412 records. After title and abstract screening, 96 "
          "full texts were assessed and 12 studies were included; the flow is shown in "
          "Figure 1. Inclusion required a stated DFT reference level, at least 100 defect "
          "configurations and a reported error metric. Studies that reported only energies "
          "for surfaces, interfaces or grain boundaries were excluded, as were studies of "
          "systems with fewer than 30 atoms in the training set."),
    ("figure", "figure1-selection.png",
     "Figure 1. Synthetic screening flow for the review: 412 records identified, 96 full "
     "texts assessed and 12 studies included."),
    ("p", "The six metals covered are Cu, Ni, Al, Ag, Au and Pd. The median supercell size "
          "was 108 atoms (range 64 to 256 atoms) and the median training set contained 1,842 "
          "defect configurations. Eight of the twelve studies used a PBE reference, three used "
          "PBEsol and one used a hybrid functional."),
    ("h1", "3. Method Families"),
    ("p", "We group the included studies into five families: Gaussian approximation "
          "potentials with moment tensor potentials (GAP-MTP), atomic cluster expansion "
          "models (ACE-EAM and ACE-NN), moment tensor potentials with SOAP descriptors "
          "(MTP-SOAP) and classical embedded-atom models (EAM). The families differ mainly in "
          "descriptor richness and in the number of training configurations required: "
          "reported training set sizes grow from 300 configurations for classical EAM to "
          "12,400 for ACE-NN."),
    ("table", "Table 1. Synthetic summary of the included method families.",
     [["Family", "Studies", "Median training set", "Descriptor order", "Reference level"],
      ["GAP-MTP", "3", "2,400", "3-body + SOAP", "PBE"],
      ["ACE-EAM", "3", "5,100", "4-body", "PBE"],
      ["MTP-SOAP", "2", "3,800", "SOAP", "PBEsol"],
      ["ACE-NN", "2", "12,400", "learned", "PBE"],
      ["EAM (classical)", "2", "300", "pair only", "PBE"]]),
    ("h1", "4. Reported Accuracy"),
    ("p", "For monovacancy formation energies, the reported MAE values were 0.043 eV for "
          "GAP-MTP, 0.058 eV for MTP-SOAP, 0.061 eV for ACE-EAM, 0.079 eV for ACE-NN and "
          "0.186 eV for classical EAM. The DFT self-consistency floor, estimated from "
          "published convergence tests, is 0.020 eV. Only the three best families sit within "
          "a factor of four of that floor."),
    ("figure", "figure2-mae.png",
     "Figure 2. Reported mean absolute error for monovacancy formation energies by method "
     "family: 0.043, 0.061, 0.058, 0.079 and 0.186 eV."),
    ("table", "Table 2. Synthetic error metrics reported by the included studies.",
     [["Quantity", "Best MAE (eV)", "Worst MAE (eV)", "Typical DFT floor (eV)", "Studies"],
      ["Monovacancy formation", "0.043", "0.186", "0.020", "12"],
      ["Migration barrier", "0.058", "0.310", "0.035", "7"],
      ["Dilute solute binding", "0.081", "0.240", "0.030", "4"]]),
    ("p", "Migration barriers behave differently. Seven of the twelve studies reported "
          "barriers, with MAE between 0.058 and 0.310 eV over a barrier range of 0.8 to "
          "4.2 eV. In two studies the barrier MAE exceeded 15 % of the mean barrier, which is "
          "large enough to reorder the predicted diffusion pathway."),
    ("figure", "figure3-parity.png",
     "Figure 3. Synthetic parity plot of MLIP versus DFT formation energies for 64 "
     "configurations; the dashed line is y = x."),
    ("h1", "5. Limitations and Open Problems"),
    ("p", "Three limitations recur. First, only 12 studies met the inclusion criteria and "
          "seven of them come from four groups, so the sample is not independent. Second, "
          "training and test sets overlap in at least three studies, which inflates reported "
          "accuracy. Third, elevated-temperature behaviour above 1,200 K is almost never "
          "validated, even though 300 to 1,200 K is the range of practical interest. "
          "A shared benchmark with fixed splits, fixed DFT references and reported error "
          "bars would resolve the first two problems; the third requires new "
          "finite-temperature reference data."),
    ("p", "We deliberately do not rank the families: the reported metrics differ in "
          "definition, in reference functional and in system, so a single ranking would be "
          "unjustified. Table 2 should be read as a map of what is reported, not as a "
          "like-for-like comparison."),
    ("h1", "References"),
    ("p", "[1] Synthetic review reference A. Fictional Reviews in Computational Materials 9 "
          "(2025) 200-215. [2] Synthetic review reference B. Fictional Journal of Physics: "
          "Condensed Fictions 41 (2024) 335-349. [3] Synthetic review reference C. Fictional "
          "Data in Materials Science 3 (2023) 44-58. These references are invented "
          "placeholders for the fixture."),
]

DOCS = {"caseA": (DOC_A, T_A), "caseB": (DOC_B, T_B)}

FIGBUILD = {
    "caseA": [("figure1-morphology.png", fig_morphology),
              ("figure2-degradation.png", fig_degradation),
              ("figure3-kinetics.png", fig_kinetics),
              ("figure4-cycles.png", fig_cycles)],
    "caseB": [("figure1-selection.png", fig_selection),
              ("figure2-mae.png", fig_mae),
              ("figure3-parity.png", fig_parity)],
}

CASE_META = {
    "caseA": {
        "case_id": "SYNTHETIC-EXP-001",
        "paper_type": "experimental",
        "output_contract": {
            "case_id": "", "deck_pptx": "", "deck_pdf": "", "slide_previews": [],
            "speaker_notes": "", "page_evidence_map": "", "plan": "", "status": "",
            "uncertainty": [], "change_records": [],
        },
        "task_md": """# {case_id} — SYNTHETIC TASK CARD (NOT AN OFFICIAL CASE)

**{banner}.** The paper in `materials/` is invented for this fixture. Its numbers are
internally consistent but they are not measurements and must never be cited as such.

## Task
Turn the supplied research material into a presentation for a group meeting, using the
organizer template in `materials/template-organizer.pptx`.

## Fixed constraints
- Audience: 课题组组会 (supervisor + peer graduate students), mixed theoretical/experimental.
- Duration: 10 minutes speaking time.
- Slide count: exactly 8 slides (title through conclusion, references allowed as backup).
- Language: English slide text; speaker notes may be English.
- Template: `materials/template-organizer.pptx` (master, layouts, theme fonts and footer
  must be preserved — they are part of the deliverable).
- Sources: `materials/paper.pdf` and `materials/paper.docx` (same paper, two formats).

## Requirements
1. One page = one claim; at most 5 bullets per slide.
2. Every number on a slide must appear verbatim in the source material; no rounding and
   no invented values.
3. Every figure used must be taken from the source (crop) or redrawn from supplied data —
   placeholder or invented curves are not acceptable.
4. Speaker notes of at least 40 characters on every content slide.
5. Keep a page-level evidence map (slide -> figure/table/section/number) and a render of
   the deck (PDF + PNG previews).
6. Return every field named in `case.json` → `output_contract`.
""",
    },
    "caseB": {
        "case_id": "SYNTHETIC-REV-002",
        "paper_type": "review/theory",
        "output_contract": {
            "case_id": "", "pptx": "", "pdf": "", "previews": [],
            "notes": "", "evidence": "", "outline": "", "status": "", "uncertainty": [],
        },
        "task_md": """# {case_id} — SYNTHETIC TASK CARD (NOT AN OFFICIAL CASE)

**{banner}.** The paper in `materials/` is invented for this fixture, including all
literature counts and reported error metrics. Nothing here is a real meta-analysis.

## Task
Turn the supplied review material into a seminar presentation for a theory group, using
the organizer template in `materials/template-organizer.pptx`.

## Fixed constraints
- Audience: 理论组 seminar, including one external reviewer.
- Duration: 20 minutes speaking time.
- Slide count: 12 slides (min 11, max 12 — a range is acceptable).
- Language: English slide text.
- Template: `materials/template-organizer.pptx` (master, layouts, theme fonts and footer
  must be preserved).
- Sources: `materials/paper.pdf` and `materials/paper.docx`.

## Requirements
1. One page = one claim; at most 5 bullets per slide. Do not paste abstract paragraphs.
2. Every reported metric on a slide must be traceable to the source text or to Table 1/2;
   the review must not present synthetic literature values as real measurements.
3. Figures must come from the source or be redrawn from supplied numbers.
4. Speaker notes of at least 40 characters on every content slide.
5. Keep a page-level evidence map and a render of the deck (PDF + PNG previews).
6. Return every field named in `case.json` → `output_contract`.
""",
    },
}


# --------------------------------------------------------------------------- #
# docx builder
# --------------------------------------------------------------------------- #
def _set_normal_font(doc) -> None:
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    style = doc.styles["Normal"]
    style.font.name = "Liberation Serif"
    style.font.size = Pt(10.5)
    style.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), "Liberation Serif")


def build_docx(case: str, out_docx: Path, figdir: Path) -> dict:
    """Write the fixture .docx; return a small inventory for the self-check."""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt, RGBColor

    blocks, meta = DOCS[case]
    doc = Document()
    _set_normal_font(doc)

    title = doc.add_heading(meta["title"], level=0)
    for run in title.runs:
        run.font.name = "Liberation Serif"
        run.font.size = Pt(16)
    sub = doc.add_paragraph(meta["authors"])
    sub.runs[0].italic = True
    sub.runs[0].font.size = Pt(9)
    sub.runs[0].font.color.rgb = RGBColor(0x99, 0x33, 0x00)
    ban = doc.add_paragraph(meta["banner"] + " — every number in this file is invented.")
    ban.runs[0].bold = True
    ban.runs[0].font.size = Pt(9)
    ban.runs[0].font.color.rgb = RGBColor(0x99, 0x33, 0x00)

    inventory = {"figures": [], "tables": [], "numbers": [], "paragraphs": 0}
    body_text = []

    for block in blocks:
        kind = block[0]
        if kind == "h1":
            h = doc.add_heading(block[1], level=1)
            for run in h.runs:
                run.font.name = "Liberation Serif"
                run.font.size = Pt(12.5)
        elif kind == "p":
            p = doc.add_paragraph(block[1])
            p.paragraph_format.space_after = Pt(6)
            inventory["paragraphs"] += 1
            body_text.append(block[1])
        elif kind == "figure":
            name, caption = block[1], block[2]
            img = figdir / name
            if not img.is_file():
                raise SystemExit("missing figure image: %s" % img)
            doc.add_picture(str(img), width=Inches(5.2))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap = doc.add_paragraph(caption)
            for run in cap.runs:
                run.italic = True
                run.font.size = Pt(9)
            cap.paragraph_format.space_after = Pt(10)
            inventory["figures"].append({"file": name, "caption": caption,
                                         "number": len(inventory["figures"]) + 1})
            body_text.append(caption)
        elif kind == "table":
            caption, rows = block[1], block[2]
            cap = doc.add_paragraph(caption)
            for run in cap.runs:
                run.bold = True
                run.font.size = Pt(9.5)
            t = doc.add_table(rows=len(rows), cols=len(rows[0]))
            t.style = "Table Grid"
            for r, row in enumerate(rows):
                for c, value in enumerate(row):
                    t.cell(r, c).text = str(value)
                    for run in t.cell(r, c).paragraphs[0].runs:
                        run.font.size = Pt(9)
                        run.font.name = "Liberation Serif"
                        run.bold = (r == 0)
            doc.add_paragraph()
            inventory["tables"].append({"caption": caption, "rows": rows,
                                        "number": len(inventory["tables"]) + 1})
            body_text.append(caption)
            for row in rows:
                body_text.extend(str(c) for c in row)

    out_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_docx))

    joined = "\n".join(body_text)
    nums = sorted({m.group(0) for m in re.finditer(r"\d+(?:\.\d+)?", joined)})
    inventory["numbers"] = nums
    inventory["distinct_numbers"] = len(nums)
    inventory["text_chars"] = len(joined)
    return inventory


# --------------------------------------------------------------------------- #
# docx -> pdf via the DSH-bundled LibreOffice Kit
# --------------------------------------------------------------------------- #
def convert_pdf(docx: Path, pdf: Path) -> dict:
    for path in (NODE, KIT_CLI):
        if not path.is_file():
            raise SystemExit("LibreOffice Kit component missing: %s" % path)
    pdf.parent.mkdir(parents=True, exist_ok=True)
    if pdf.exists():
        pdf.unlink()
    cmd = [str(NODE), str(KIT_CLI), "convert",
           "--input", str(docx), "--output", str(pdf)]
    if FONT_TREE.is_dir():
        cmd += ["--font-directory", str(FONT_TREE)]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                          errors="replace")
    payload = {}
    try:
        payload = json.loads(proc.stdout.strip().splitlines()[-1]) if proc.stdout.strip() else {}
    except Exception:
        payload = {"raw": proc.stdout[-400:]}
    ok = proc.returncode == 0 and pdf.is_file() and pdf.stat().st_size > 0
    return {"exit_code": proc.returncode, "ok": ok, "kit": payload,
            "stderr": (proc.stderr or "")[-400:]}


def pdf_introspect(pdf: Path) -> dict:
    """Confirm the PDF has a text layer and embedded raster images."""
    out = {"pages": 0, "text_chars": 0, "image_xobjects": 0, "fonts": []}
    try:
        import pypdf
    except Exception as exc:  # pragma: no cover - pypdf is present on this host
        out["error"] = "pypdf unavailable: %s" % exc
        return out
    try:
        reader = pypdf.PdfReader(str(pdf))
        out["pages"] = len(reader.pages)
        chunks = []
        for page in reader.pages:
            chunks.append(page.extract_text() or "")
            res = page.get("/Resources")
            xo = res.get("/XObject") if res else None
            if xo:
                try:
                    for name in xo.keys():
                        obj = xo[name].get_object()
                        if obj.get("/Subtype") == "/Image":
                            out["image_xobjects"] += 1
                except Exception:
                    pass
        out["text_chars"] = sum(len(c) for c in chunks)
        fonts = set()
        for page in reader.pages:
            res = page.get("/Resources")
            fd = res.get("/Font") if res else None
            if not fd:
                continue
            for name in fd.keys():
                try:
                    fo = fd[name].get_object()
                    base = str(fo.get("/BaseFont", ""))
                    if base:
                        fonts.add(base.lstrip("/"))
                except Exception:
                    pass
        out["fonts"] = sorted(fonts)
        out["text_sample"] = "\n".join(chunks)[:1200]
    except Exception as exc:
        out["error"] = "%s: %s" % (type(exc).__name__, exc)
    return out


# --------------------------------------------------------------------------- #
# case.json / task.md
# --------------------------------------------------------------------------- #
def build_case(case: str) -> dict:
    meta = CASE_META[case]
    cdir = HERE / case
    matdir = cdir / "materials"
    figdir = matdir / "figures"
    matdir.mkdir(parents=True, exist_ok=True)
    figdir.mkdir(parents=True, exist_ok=True)

    if not TEMPLATE.is_file():
        raise SystemExit("organizer template missing: %s" % TEMPLATE)
    shutil.copy2(TEMPLATE, matdir / "template-organizer.pptx")

    matplotlib = register_matplotlib_fonts()
    figs = []
    for name, fn in FIGBUILD[case]:
        target = figdir / name
        fn(target)
        figs.append(target)

    docx_path = matdir / "paper.docx"
    inv = build_docx(case, docx_path, figdir)
    pdf_path = matdir / "paper.pdf"
    conv = convert_pdf(docx_path, pdf_path)
    if not conv["ok"]:
        raise SystemExit("docx->pdf conversion failed: %s" % json.dumps(conv)[:600])
    pdf_info = pdf_introspect(pdf_path)

    case_json = {
        "case_id": meta["case_id"],
        "note": ("SYNTHETIC / NOT AN OFFICIAL CASE. Fixture built by work/fixtures/"
                 "build_paper.py for self-testing only. The paper, the figures, the tables "
                 "and every number are invented for this fixture; they are internally "
                 "consistent inside the case but are NOT measurements and must never be "
                 "presented as real results. The template is an organizer-style synthetic "
                 "template, not a real organizer file."),
        "paper_type": meta["paper_type"],
        "output_contract": meta["output_contract"],
        "constraints": {
            "audience": ("课题组组会 (supervisor + peer graduate students)"
                         if case == "caseA" else
                         "理论组 seminar with one external reviewer"),
            "duration_min": 10 if case == "caseA" else 20,
            "slide_count": 8 if case == "caseA" else 12,
            "slide_count_range": None if case == "caseA" else [11, 12],
            "language": "en",
            "template_path": TEMPLATE_REL,
            "starting_deck_path": None,
        },
        "materials_dir": "materials",
        "template": TEMPLATE_REL,
        "sources": ["materials/paper.pdf", "materials/paper.docx"],
        "starting_deck": None,
        "synth": {
            "builder": "work/fixtures/build_paper.py",
            "figures": ["materials/figures/%s" % p.name for p in figs],
            "distinct_numbers_in_text": inv["distinct_numbers"],
            "pdf_pages": pdf_info.get("pages"),
            "pdf_embedded_images": pdf_info.get("image_xobjects"),
        },
    }
    (cdir / "case.json").write_text(
        json.dumps(case_json, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8", newline="\n")

    (cdir / "task.md").write_text(
        meta["task_md"].format(case_id=meta["case_id"], banner=SYNTHETIC_BANNER),
        encoding="utf-8", newline="\n")

    return {
        "case": case, "case_id": meta["case_id"], "dir": str(cdir),
        "docx": docx_path, "pdf": pdf_path, "figures": figs,
        "inventory": inv, "pdf_info": pdf_info, "conversion": conv,
    }


def report(results) -> dict:
    """Print and return the machine-readable build report."""
    summary = {"schema": "fixture_build/v1", "cases": []}
    for r in results:
        files = {}
        for p in [r["docx"], r["pdf"], HERE / r["case"] / "case.json",
                  HERE / r["case"] / "task.md"] + list(r["figures"]):
            files[str(p.relative_to(HERE)).replace("\\", "/")] = {
                "sha256": sha256(p), "bytes": p.stat().st_size}
        summary["cases"].append({
            "case": r["case"], "case_id": r["case_id"],
            "pdf_pages": r["pdf_info"].get("pages"),
            "pdf_text_chars": r["pdf_info"].get("text_chars"),
            "pdf_embedded_images": r["pdf_info"].get("image_xobjects"),
            "pdf_fonts": r["pdf_info"].get("fonts"),
            "figures": len(r["figures"]),
            "tables": len(r["inventory"]["tables"]),
            "distinct_numbers": r["inventory"]["distinct_numbers"],
            "text_chars": r["inventory"]["text_chars"],
            "files": files,
        })
    return summary


README_HEADER = """# `work/fixtures/` — SYNTHETIC fixture cases (NOT OFFICIAL CASES)

**Every case in this directory is SYNTHETIC.** The papers, the figures, the tables, the
literature counts and every number in them were invented by `build_paper.py` for
self-testing this Skill. They are internally consistent *inside a case* (abstract, prose,
tables and figure data agree) but they are **not measurements and not a meta-analysis**,
and must never be presented as real results. The `.pptx` template is an organizer-style
**synthetic** template (`build_template.py`), not a real organizer file.

The platform supplies its own 3 hidden cases; this directory exists so the pipeline, the
self-test and the gain harness have something real to run on.

## Contents

| path | what it is |
|---|---|
| `template-organizer.pptx` | master copy of the synthetic organizer template (11 layouts, theme fonts = Liberation Serif, accent1 = `1F6FEB`, master stamp text box `TemplateStamp` = `ORGANIZER-ACADEMIC-2026`, footer text `Organizer Academic 2026 - SYNTHETIC TEMPLATE`) |
| `build_template.py` | rebuilds the template: `python build_template.py template-organizer.pptx` |
| `build_paper.py` | rebuilds both cases end-to-end: `python build_paper.py` |
| `probe_mutations.py` | dev probe: applies all 12 defect injections of `scripts/selftest.py` to a throwaway deck, so an injection that silently does nothing is caught before a full pipeline run (`python probe_mutations.py` → `PROBE RESULT: OK`) |
| `make_example.py` | dev script: copies a **real** run into `work/submission/examples/dev-run-caseA/`, sanitizing absolute paths and building a small PNG contact sheet (`python make_example.py --case caseA/case.json --work _example_work --out _example_out`) |
| `known_good.txt` | which cases `scripts/selftest.py` asserts fully green; drop a case here to keep exercising it without hard-failing the harness while it is being fixed |
| `fixture_build_report.json` | machine-readable inventory + sha256 of every fixture artifact, written by `build_paper.py --json` |
| `caseA/` | experimental paper (photocatalysis), 8 slides, 10 min |
| `caseB/` | review / theory paper (ML interatomic potentials), 12 slides, 20 min |

Each case directory holds:

* `case.json` — platform shape `{case_id, note, output_contract:{...}}` **plus** optional
  `constraints` / `template` / `sources` keys so the case is self-contained;
* `task.md` — audience, duration, slide count (`task.md` and `case.json` agree);
* `materials/paper.docx` + `materials/paper.pdf` — the same synthetic paper in two
  formats; the PDF has a real text layer and **real embedded raster images**;
* `materials/figures/*.png` — the figure source images;
* `materials/template-organizer.pptx` — a copy of the template, so each case is portable.

## Why two cases with different `output_contract` spellings

The platform declares the result field names per case, and the names are not known in
advance. The two fixtures deliberately use **different spellings for the same roles**, so a
hard-coded field list fails loudly:

| role | caseA key | caseB key |
|---|---|---|
| editable deck | `deck_pptx` | `pptx` |
| exported PDF | `deck_pdf` | `pdf` |
| PNG previews | `slide_previews` | `previews` |
| page evidence map | `page_evidence_map` | `evidence` |
| speaker notes | `speaker_notes` | `notes` |
| page plan | `plan` | `outline` |

## Rebuild and re-verify

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
cd ARMI2611000074v1\\work\\fixtures
python build_template.py template-organizer.pptx      # optional: template
python build_paper.py --json fixture_build_report.json --readme
```

`build_paper.py` exits non-zero if any case loses its text layer, its embedded figure
images, its numbered figures/tables or its minimum number of distinct numbers.

## sha256 of fixture artifacts

Generated by `build_paper.py --json`; re-run it after any rebuild.
"""


def build_readme(summary: dict) -> str:
    lines = [README_HEADER, ""]
    tmpl = TEMPLATE
    if tmpl.is_file():
        lines += ["### template", "",
                  "| file | bytes | sha256 |", "|---|---|---|",
                  "| `template-organizer.pptx` | %d | `%s` |"
                  % (tmpl.stat().st_size, sha256(tmpl)), ""]
    for c in summary["cases"]:
        meta = CASE_META[c["case"]]
        lines += ["### %s — `%s` (%s paper)" % (c["case"], c["case_id"],
                                                meta["paper_type"]), "",
                  "- PDF pages: %s, text-layer chars: %s, embedded raster images: %s"
                  % (c["pdf_pages"], c["pdf_text_chars"], c["pdf_embedded_images"]),
                  "- numbered figures: %s, numbered tables: %s, distinct numbers in text: %s"
                  % (c["figures"], c["tables"], c["distinct_numbers"]),
                  "- `output_contract` keys: `%s`"
                  % "`, `".join(meta["output_contract"].keys()), "",
                  "| file | bytes | sha256 |", "|---|---|---|"]
        for rel, info in sorted(c["files"].items()):
            lines.append("| `%s` | %d | `%s` |" % (rel, info["bytes"], info["sha256"]))
        lines.append("")
    return "\n".join(lines) + "\n"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--case", choices=sorted(CASE_META), action="append",
                    help="build only this case (repeatable); default both")
    ap.add_argument("--json", help="also write the build report to this path")
    ap.add_argument("--readme", action="store_true",
                    help="rewrite README.md from the build report (keeps hashes fresh)")
    ap.add_argument("--readme-only", action="store_true",
                    help="rewrite README.md from an existing --json report without rebuilding")
    args = ap.parse_args(argv)

    if args.readme_only:
        if not args.json or not Path(args.json).is_file():
            print("--readme-only needs --json <existing report>", file=sys.stderr)
            return 2
        summary = json.loads(Path(args.json).read_text(encoding="utf-8"))
        (HERE / "README.md").write_text(build_readme(summary), encoding="utf-8", newline="\n")
        log("wrote README.md from %s" % args.json)
        print("RESULT: OK (README only)")
        return 0

    cases = args.case or sorted(CASE_META)
    results = []
    for case in cases:
        log("building %s ..." % case)
        r = build_case(case)
        results.append(r)
        log("  %s: pdf pages=%s text_chars=%s images=%s figures=%s tables=%s numbers=%s"
            % (case, r["pdf_info"].get("pages"), r["pdf_info"].get("text_chars"),
               r["pdf_info"].get("image_xobjects"), len(r["figures"]),
               len(r["inventory"]["tables"]), r["inventory"]["distinct_numbers"]))
        problems = []
        if (r["pdf_info"].get("pages") or 0) < 2:
            problems.append("pdf has fewer than 2 pages")
        if (r["pdf_info"].get("text_chars") or 0) < 2000:
            problems.append("pdf text layer too small")
        if (r["pdf_info"].get("image_xobjects") or 0) < 1:
            problems.append("pdf has no embedded raster image")
        if len(r["inventory"]["tables"]) < 2:
            problems.append("fewer than 2 tables")
        if len(r["figures"]) < 3:
            problems.append("fewer than 3 figures")
        if r["inventory"]["distinct_numbers"] < 12:
            problems.append("fewer than 12 distinct numbers")
        if problems:
            log("  FAIL %s: %s" % (case, "; ".join(problems)))
            return 1

    summary = report(results)
    if args.json:
        Path(args.json).write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
                                   encoding="utf-8", newline="\n")
    if args.readme:
        (HERE / "README.md").write_text(build_readme(summary),
                                        encoding="utf-8", newline="\n")
        log("wrote README.md")
    print("RESULT: OK (%d case(s))" % len(results))
    return 0


if __name__ == "__main__":
    sys.exit(main())
