#!/usr/bin/env python3
"""
Simply Supported Beam Analysis Report
Author: Soniya Kuchekar
PDF via PyLaTeX + pdflatex (2 runs for TOC + LastPage)
"""

from __future__ import annotations

import subprocess
from pathlib import Path
import pandas as pd
from pylatex import Document, Section, Subsection, Figure, Tabular, Package, NoEscape


# ====== USER DETAILS ======
AUTHOR_NAME = "Soniya Kuchekar"
REPORT_TITLE = "Simply Supported Beam Analysis Report"
REPORT_SUBTITLE = "Structural Beam Analysis using TikZ/pgfplots"
INSTITUTE_LINE = "VIT Bhopal University"
REPORT_ID = "BA-2026-SONIYA-KUCHEKAR-SSB"

# ====== FILES ======
INPUT_EXCEL = "beam_data.xlsx"
BEAM_IMAGE = "beam.png"
OUTPUT_BASENAME = "Simply_Supported_Beam_Analysis_Soniya_Kuchekar"
ERROR_LOG_FILE = "compile_error_tail.txt"


def pick_columns(df: pd.DataFrame) -> tuple[str, str, str]:
    """Detect the Position, Shear, and Moment columns."""
    df.columns = df.columns.str.strip()
    pos = shear = moment = None

    for c in df.columns:
        cl = c.lower()
        if pos is None and ("position" in cl or cl in ["x", "distance", "dist", "length"]):
            pos = c
        if shear is None and "shear" in cl:
            shear = c
        if moment is None and "moment" in cl:
            moment = c

    if not all([pos, shear, moment]):
        raise ValueError(
            f"Could not detect required columns. Found: {list(df.columns)}\n"
            "Need columns containing keywords like Position/X, Shear, Moment."
        )
    return pos, shear, moment


def run_pdflatex_twice(tex_path: Path) -> None:
    """
    Run pdflatex twice (TOC and LastPage need 2 runs).
    If it fails, saves log tail to compile_error_tail.txt.
    """
    workdir = tex_path.parent
    tex_file = tex_path.name
    cmd = ["pdflatex", "-interaction=nonstopmode", "-halt-on-error", tex_file]

    for run_no in (1, 2):
        p = subprocess.run(cmd, cwd=workdir, capture_output=True, text=True)
        if p.returncode != 0:
            tail = (p.stdout[-4500:] + "\n" + p.stderr[-4500:]).strip()
            (workdir / ERROR_LOG_FILE).write_text(
                f"PDLATEX FAILED ON RUN {run_no}\n\n{tail}\n",
                encoding="utf-8",
                errors="ignore",
            )
            raise RuntimeError(
                f"pdflatex failed (run {run_no}). "
                f"Open '{ERROR_LOG_FILE}' in the folder to see the exact LaTeX error."
            )


def sfd_plot(df: pd.DataFrame, x_col: str, v_col: str, L: float) -> NoEscape:
    """Shear Force Diagram using TikZ/pgfplots vector plot."""
    sdf = df.sort_values(x_col).reset_index(drop=True)
    vmin = float(sdf[v_col].min())
    vmax = float(sdf[v_col].max())
    pad = max(5.0, 0.12 * (abs(vmin) + abs(vmax)))
    ymin, ymax = vmin - pad, vmax + pad

    coords = "\n".join(
        f"({float(r[x_col])}, {float(r[v_col])})" for _, r in sdf.iterrows()
    )

    return NoEscape(rf"""
\begin{{figure}}[H]
\centering
\vspace{{2mm}}
\begin{{tikzpicture}}
\begin{{axis}}[
    width=15.6cm,
    height=6.6cm,
    title={{\textbf{{Shear Force Diagram (SFD)}}}},
    title style={{font=\large}},
    xlabel={{Position, $x$ (m)}},
    ylabel={{Shear Force, $V$ (kN)}},
    xmin=0, xmax={L:.2f},
    ymin={ymin:.2f}, ymax={ymax:.2f},
    grid=major,
    grid style={{dotted, gray!35}},
    axis line style={{black, thick}},
    tick label style={{font=\small}},
    label style={{font=\small}},
]
\addplot[blue!75!black, very thick, mark=*, mark size=1.7pt]
coordinates {{
{coords}
}};
\addplot[black, dashed, thick] coordinates {{(0,0) ({L:.2f},0)}};
\end{{axis}}
\end{{tikzpicture}}
\vspace{{-2mm}}
\end{{figure}}
""")


def bmd_plot(df: pd.DataFrame, x_col: str, m_col: str, L: float) -> NoEscape:
    """Bending Moment Diagram using TikZ/pgfplots vector plot."""
    sdf = df.sort_values(x_col).reset_index(drop=True)
    mmin = float(sdf[m_col].min())
    mmax = float(sdf[m_col].max())
    pad = max(5.0, 0.12 * (abs(mmin) + abs(mmax)))
    ymin, ymax = min(0.0, mmin - pad), mmax + pad

    coords = "\n".join(
        f"({float(r[x_col])}, {float(r[m_col])})" for _, r in sdf.iterrows()
    )

    return NoEscape(rf"""
\begin{{figure}}[H]
\centering
\vspace{{2mm}}
\begin{{tikzpicture}}
\begin{{axis}}[
    width=15.6cm,
    height=6.6cm,
    title={{\textbf{{Bending Moment Diagram (BMD)}}}},
    title style={{font=\large}},
    xlabel={{Position, $x$ (m)}},
    ylabel={{Bending Moment, $M$ (kNm)}},
    xmin=0, xmax={L:.2f},
    ymin={ymin:.2f}, ymax={ymax:.2f},
    grid=major,
    grid style={{dotted, gray!35}},
    axis line style={{black, thick}},
    tick label style={{font=\small}},
    label style={{font=\small}},
]
\addplot[red!80!black, very thick, mark=*, mark size=1.7pt]
coordinates {{
{coords}
}};
\end{{axis}}
\end{{tikzpicture}}
\vspace{{-2mm}}
\end{{figure}}
""")


def build_report() -> Path:
    df = pd.read_excel(INPUT_EXCEL)
    x_col, v_col, m_col = pick_columns(df)
    L = float(df[x_col].max())

    doc = Document(documentclass="report", document_options=["12pt", "a4paper"], lmodern=True)

    # ===== Packages =====
    doc.packages.append(Package("geometry", options=["margin=1in"]))
    doc.packages.append(Package("graphicx"))
    doc.packages.append(Package("float"))
    doc.packages.append(Package("booktabs"))
    doc.packages.append(Package("array"))
    doc.packages.append(Package("setspace"))
    doc.packages.append(Package("parskip"))
    doc.packages.append(Package("titlesec"))
    doc.packages.append(Package("fancyhdr"))
    doc.packages.append(Package("lastpage"))
    doc.packages.append(Package("needspace"))

    doc.packages.append(Package("xcolor", options=["table"]))
    doc.packages.append(Package("tikz"))
    doc.packages.append(Package("pgfplots"))
    doc.packages.append(Package("hyperref"))
    doc.packages.append(Package("microtype"))
    doc.preamble.append(NoEscape(r"\pgfplotsset{compat=1.18}"))

    # Hyperlinks
    doc.preamble.append(NoEscape(r"""
\hypersetup{
  colorlinks=true,
  linkcolor=black,
  urlcolor=black,
  citecolor=black
}
"""))

    # Spacing
    doc.preamble.append(NoEscape(r"\onehalfspacing"))
    doc.preamble.append(NoEscape(r"\setlength{\parindent}{0pt}"))
    doc.preamble.append(NoEscape(r"\setlength{\parskip}{9pt}"))
    doc.preamble.append(NoEscape(r"\setlength{\textfloatsep}{16pt}"))
    doc.preamble.append(NoEscape(r"\setlength{\intextsep}{16pt}"))
    doc.preamble.append(NoEscape(r"\setlength{\abovecaptionskip}{6pt}"))
    doc.preamble.append(NoEscape(r"\setlength{\belowcaptionskip}{6pt}"))
    doc.preamble.append(NoEscape(r"\renewcommand{\contentsname}{Table of Contents}"))

    # Headings
    doc.preamble.append(NoEscape(r"""
\titleformat{\section}{\Large\bfseries}{\thesection.}{0.9em}{}
\titleformat{\subsection}{\large\bfseries}{\thesubsection}{0.9em}{}
\titlespacing*{\section}{0pt}{22pt}{10pt}
\titlespacing*{\subsection}{0pt}{14pt}{8pt}
"""))

    # Header/footer
    doc.preamble.append(NoEscape(r"\setlength{\headheight}{15pt}"))
    doc.preamble.append(NoEscape(rf"""
\pagestyle{{fancy}}
\fancyhf{{}}
\lhead{{\textbf{{{REPORT_TITLE}}}}}
\rhead{{\textbf{{{AUTHOR_NAME}}}}}
\cfoot{{Page \thepage\ of \pageref{{LastPage}}}}
\renewcommand{{\headrulewidth}}{{0.4pt}}
\renewcommand{{\footrulewidth}}{{0.4pt}}
"""))

    def lead(text: str) -> None:
        doc.append(NoEscape(r"\vspace{-2pt}"))
        doc.append(NoEscape(rf"\textit{{{text}}}"))
        doc.append(NoEscape(r"\vspace{6pt}"))

    # =========================================================
    # 1) TITLE PAGE
    # =========================================================
    doc.append(NoEscape(r"\thispagestyle{empty}"))
    doc.append(NoEscape(r"\begin{tikzpicture}[remember picture,overlay]"))
    doc.append(NoEscape(r"\fill[black!6] (current page.north west) rectangle ([yshift=-3.8cm]current page.north east);"))
    doc.append(NoEscape(r"\fill[black!22] (current page.north west) rectangle ([xshift=0.35cm,yshift=-3.8cm]current page.north west);"))
    doc.append(NoEscape(r"\end{tikzpicture}"))

    doc.append(NoEscape(r"\vspace*{2.2cm}"))
    doc.append(NoEscape(r"\begin{center}"))
    doc.append(NoEscape(rf"{{\Huge\bfseries {REPORT_TITLE}}}\\[0.35cm]"))
    doc.append(NoEscape(rf"{{\Large {REPORT_SUBTITLE}}}\\[0.35cm]"))
    doc.append(NoEscape(rf"{{\large {INSTITUTE_LINE}}}\\[0.9cm]"))
    doc.append(NoEscape(r"\rule{0.74\textwidth}{0.7pt}\\[0.7cm]"))
    doc.append(NoEscape(rf"{{\Large \textbf{{Author:}} {AUTHOR_NAME}}}\\[0.25cm]"))
    doc.append(NoEscape(rf"{{\large \textbf{{Report ID:}} {REPORT_ID}}}\\[0.25cm]"))
    doc.append(NoEscape(r"{\large \textbf{Date:} \today}"))
    doc.append(NoEscape(r"\end{center}"))
    doc.append(NoEscape(r"\vfill"))
    doc.append(NoEscape(r"\newpage"))

    # =========================================================
    # 2) TABLE OF CONTENTS (ONLY ONCE ✅)
    # =========================================================
    doc.append(NoEscape(r"\tableofcontents"))
    doc.append(NoEscape(r"\newpage"))

    # =========================================================
    # 3) INTRODUCTION
    # =========================================================
    with doc.create(Section("Introduction")):
        lead("This section describes the beam system and the origin of the input dataset used for analysis.")

        with doc.create(Subsection("Beam Description")):
            doc.append(
                "The structure considered is a simply supported beam with a pinned support at the left end "
                "and a roller support at the right end. This arrangement allows rotation at both supports "
                "while preventing vertical displacement."
            )
            with doc.create(Figure(position="H")) as fig:
                fig.add_image(BEAM_IMAGE, width=NoEscape(r"0.82\textwidth"))
                fig.add_caption("Simply Supported Beam Configuration")

        with doc.create(Subsection("Data Source")):
            doc.append(
                f"The force and moment values used in this report are read directly from the provided Excel file "
                f"({INPUT_EXCEL}). The next section recreates the Excel table using LaTeX Tabular."
            )

    # =========================================================
    # 4) INPUT DATA
    # =========================================================
    with doc.create(Section("Input Data")):
        lead("This section recreates the Excel dataset using a LaTeX table (not inserted as an image).")

        sdf = df.sort_values(x_col).reset_index(drop=True)

        doc.append(NoEscape(r"\renewcommand{\arraystretch}{1.25}"))
        doc.append(NoEscape(r"\rowcolors{2}{black!4}{white}"))
        doc.append(NoEscape(r"\begin{center}"))

        with doc.create(Tabular(NoEscape(
            r"@{}>{\centering\arraybackslash}p{3.7cm}"
            r">{\centering\arraybackslash}p{5.3cm}"
            r">{\centering\arraybackslash}p{5.3cm}@{}"
        ))) as table:
            doc.append(NoEscape(r"\toprule"))
            table.add_row(
                NoEscape(r"\textbf{Position (m)}"),
                NoEscape(r"\textbf{Shear Force (kN)}"),
                NoEscape(r"\textbf{Bending Moment (kNm)}")
            )
            doc.append(NoEscape(r"\midrule"))
            for _, r in sdf.iterrows():
                table.add_row(
                    f"{float(r[x_col]):.2f}",
                    f"{float(r[v_col]):.2f}",
                    f"{float(r[m_col]):.2f}"
                )
            doc.append(NoEscape(r"\bottomrule"))

        doc.append(NoEscape(r"\end{center}"))
        doc.append(NoEscape(r"\rowcolors{2}{}{}"))

    # =========================================================
    # 5) ANALYSIS (alignment fixed)
    # =========================================================
    with doc.create(Section("Analysis")):
        lead("This section presents engineering diagrams generated as TikZ/pgfplots vector plots for high-quality output.")

        doc.append(NoEscape(r"\Needspace{16\baselineskip}"))
        with doc.create(Subsection("Shear Force Diagram")):
            doc.append(NoEscape(r"\begin{samepage}"))
            doc.append("The Shear Force Diagram (SFD) illustrates the variation of shear force along the beam span.")
            doc.append(sfd_plot(df, x_col, v_col, L))
            doc.append(NoEscape(r"\end{samepage}"))

        doc.append(NoEscape(r"\Needspace{16\baselineskip}"))
        with doc.create(Subsection("Bending Moment Diagram")):
            doc.append(NoEscape(r"\begin{samepage}"))
            doc.append("The Bending Moment Diagram (BMD) represents the bending moment distribution along the span.")
            doc.append(bmd_plot(df, x_col, m_col, L))
            doc.append(NoEscape(r"\end{samepage}"))

    # Write .tex
    base = Path(OUTPUT_BASENAME)
    doc.generate_tex(str(base))
    return base.with_suffix(".tex")


def main():
    tex_path = build_report()
    run_pdflatex_twice(tex_path)
    print(f"✅ PDF created: {OUTPUT_BASENAME}.pdf")


if __name__ == "__main__":
    main()
