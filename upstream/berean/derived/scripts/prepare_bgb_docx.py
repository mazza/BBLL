#!/usr/bin/env python3
"""
prepare_bgb_docx.py

Prepare the original BGB DOCX by applying the same structural edits
previously done in Google Docs:

1. Convert book titles to Heading 1.
2. Convert chapter titles to Heading 2.
3. Convert paragraphs with the native 'hdg' style to Heading 3
   (leaving the front matter untouched).
4. Explicitly set the red color (FF0000) on runs that use the 'red' style.

Usage:
    python3 prepare_bgb_docx.py
    python3 prepare_bgb_docx.py -i path/to/bgb.docx -o path/to/bgb-edited.docx
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = SCRIPT_DIR.parent / ".." / "raw" / "greek-nt" / "bgb.docx"
DEFAULT_OUTPUT = SCRIPT_DIR.parent / "bgb-edited.docx"

NT_BOOKS_SET = {
    "Matthew", "Mark", "Luke", "John", "Acts",
    "Romans", "1 Corinthians", "2 Corinthians", "Galatians",
    "Ephesians", "Philippians", "Colossians",
    "1 Thessalonians", "2 Thessalonians",
    "1 Timothy", "2 Timothy", "Titus", "Philemon",
    "Hebrews", "James", "1 Peter", "2 Peter",
    "1 John", "2 John", "3 John", "Jude", "Revelation",
}


def make_run_explicit_red(run) -> None:
    """Inject the XML tag <w:color w:val="FF0000"/> directly into the run element."""
    rPr = run._element.get_or_add_rPr()
    color = rPr.find(qn('w:color'))
    if color is None:
        color = OxmlElement('w:color')
        rPr.append(color)
    color.set(qn('w:val'), 'FF0000')


def prepare(input_path: Path, output_path: Path) -> None:
    print(f"Abrindo arquivo original: {input_path}")
    doc = Document(str(input_path))

    stats = {"books": 0, "chapters": 0, "sections": 0, "red_runs": 0}

    for para in doc.paragraphs:
        txt = para.text.strip()
        p_style = para.style.name if para.style else ''

        # 1. Book title (Heading 1)
        if txt in NT_BOOKS_SET:
            para.style = doc.styles['Heading 1']
            stats["books"] += 1
            continue

        # 2. Chapter title (Heading 2)
        match = re.match(r"^(.*?)\s+(\d+)$", txt)
        if match and match.group(1) in NT_BOOKS_SET:
            para.style = doc.styles['Heading 2']
            stats["chapters"] += 1
            continue

        # 3. Section title (uses the native 'hdg' style label from bgb.docx)
        if p_style == 'hdg':
            para.style = doc.styles['Heading 3']
            stats["sections"] += 1
            continue

        # 4. Bake Jesus's red colors into the XML (as Google Docs does)
        is_para_red = 'red' in p_style.lower()

        for run in para.runs:
            is_run_red = False
            rPr = run._element.rPr
            if rPr is not None:
                rstyle = rPr.find(qn('w:rStyle'))
                if rstyle is not None:
                    val = rstyle.attrib.get(qn('w:val'))
                    if val and 'red' in val.lower():
                        is_run_red = True

            if is_para_red or is_run_red:
                make_run_explicit_red(run)
                stats["red_runs"] += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))

    print(f"\n✓ Sucesso! Salvo em: {output_path}")
    print(f"  Livros marcados (Heading 1)   : {stats['books']}")
    print(f"  Capítulos marcados (Heading 2): {stats['chapters']}")
    print(f"  Seções marcadas (Heading 3)   : {stats['sections']}")
    print(f"  Trechos vermelhos processados : {stats['red_runs']}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Prepara o DOCX BGB aplicando Heading 1/2/3 e baking de cores.")
    parser.add_argument("-i", "--input", type=Path, default=DEFAULT_INPUT, help=f"DOCX de entrada (padrão: {DEFAULT_INPUT})")
    parser.add_argument("-o", "--output", type=Path, default=DEFAULT_OUTPUT, help=f"DOCX de saída (padrão: {DEFAULT_OUTPUT})")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if not args.input.is_file():
        print(f"Erro: arquivo não encontrado: {args.input}", file=sys.stderr)
        return 1
    prepare(args.input, args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
