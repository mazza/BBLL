#!/usr/bin/env python3
"""
bgb_to_markdown.py

Convert the prepared BGB DOCX into structured Markdown for the
New Testament Greek text.

Features:
- HTML provenance comment at the top
- Skip front matter (start at the first book heading)
- Bold verse numbers (at the beginning and in the middle of paragraphs)
- Balance red HTML color tags line by line
- Insert line breaks before mid-paragraph verse numbers
- Map footnotes from the DOCX hyperlinks
- Render words of Jesus in red and cross-references in blue
- Handle poetic indentation with block quotes and non-breaking spaces

Usage:
    python3 bgb_to_markdown.py
    python3 bgb_to_markdown.py -i bgb-edited.docx -o BGB_Greek_NT.md
"""

from __future__ import annotations

import argparse
import html
import re
import sys
from pathlib import Path

import docx
from docx import Document
from docx.oxml.ns import qn

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = SCRIPT_DIR.parent / "bgb-edited.docx"
DEFAULT_OUTPUT = SCRIPT_DIR.parent / "BGB_Greek_NT.md"

PROVENANCE_COMMENT = """\
<!--
  Generated from the Berean Greek Bible (BGB) New Testament DOCX
  published by Bible Hub / Berean Bible (https://berean.bible).

  Source text © Bible Hub. This Markdown extraction is a derived
  working file for the BBLL project and does not alter the original
  Greek text.
-->
"""

# Verse number at the start of the line (may be inside a <span...>)
START_VERSE_RE = re.compile(
    r"^(<span[^>]*>)?(\d{1,3})([\s\u00a0]+)"
)

# Verse number in the middle of the text
MID_VERSE_RE = re.compile(
    r"(?<!\*)(?<!\*\*)"                 # not already bold
    r"([\.\;\:\?\!\,’\”»\)\s]|</span>)" # separator before (including </span>)
    r"(<span[^>]*>)?"                   # optional <span...> tag
    r"(\d{1,3})"                        # verse number
    r"([\s\u00a0]+)"                    # space
    r"(?:<span[^>]*>)?"                 # optional post-space tag
    r"(?=[Α-Ωα-ωἀ-ὡἈ-Ὡ‹“‘\"\(])"        # lookahead: Greek, quotes, ‹, (
)


def get_left_indent_pt(para) -> float | None:
    try:
        if para.paragraph_format.left_indent is not None:
            return round(para.paragraph_format.left_indent.pt, 1)
    except Exception:
        pass
    return None


def get_paragraph_indentation(para) -> tuple[bool, str]:
    style_name = (para.style.name if para.style else "").lower()
    indent_pt = get_left_indent_pt(para)

    if style_name in ("indent2", "indentred2"):
        return True, "&nbsp;" * 8
    if style_name in ("indent1", "indentred1"):
        return True, "&nbsp;" * 4
    if style_name in ("indent1stline", "indent1stlinered"):
        return True, ""

    if indent_pt is not None:
        if indent_pt >= 35:
            return True, "&nbsp;" * 8
        if indent_pt >= 12:
            return True, "&nbsp;" * 4

    return False, ""


def is_run_red(run, para) -> bool:
    try:
        if run.font and run.font.color and run.font.color.rgb:
            c = str(run.font.color.rgb).upper()
            if c in {
                "FF0000", "RED", "C00000", "990000",
                "E00000", "CC0000", "D80000", "FF3333",
            }:
                return True
    except Exception:
        pass

    rPr = run._element.rPr
    if rPr is not None:
        color = rPr.find(qn("w:color"))
        if color is not None:
            val = color.attrib.get(qn("w:val"))
            if val and val.upper() in {
                "FF0000", "RED", "C00000", "990000",
                "E00000", "CC0000", "D80000", "FF3333",
            }:
                return True
        rstyle = rPr.find(qn("w:rStyle"))
        if rstyle is not None:
            val = rstyle.attrib.get(qn("w:val"))
            if val and "red" in val.lower():
                return True

    if para.style and "red" in (para.style.name or "").lower():
        return True

    return False


def is_run_blue(run) -> bool:
    rPr = run._element.rPr
    if rPr is not None:
        rstyle = rPr.find(qn("w:rStyle"))
        if rstyle is not None:
            val = rstyle.attrib.get(qn("w:val"))
            if val and "cross" in val.lower():
                return True
        color = rPr.find(qn("w:color"))
        if color is not None:
            val = color.attrib.get(qn("w:val"))
            if val and val.upper() in {
                "0092F2", "BLUE", "0000FF", "1B75BC",
                "0070C0", "00A2E8", "0066CC",
            }:
                return True
    return False


def parse_footnote_block(text: str) -> list[tuple[str, str, str]]:
    notes = []
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for line in lines:
        m = re.match(r"^([a-z])\s+(\d+)(?:-\d+)?\s+(.+)$", line, re.IGNORECASE | re.DOTALL)
        if m:
            notes.append((m.group(1).lower(), m.group(2), m.group(3).strip()))
        else:
            m2 = re.match(r"^([a-z])\s+(.+)$", line, re.IGNORECASE | re.DOTALL)
            if m2:
                notes.append((m2.group(1).lower(), "", m2.group(2).strip()))
    return notes


def bold_and_split_verse_numbers(content: str) -> list[str]:
    # Clean adjacent span fragmentations caused by footnotes
    content = re.sub(r'</span>\s*<span style="color:#FF0000">', '', content)

    # Bold at the start of the string
    content = START_VERSE_RE.sub(r"\1**\2**\3", content)

    # Bold and insert a line break in the middle
    content = MID_VERSE_RE.sub(r"\1\n\2**\3**\4", content)

    parts = [p.strip() for p in content.split("\n") if p.strip()]
    return fix_html_spans_across_lines(parts) if parts else [content.strip()]


def fix_html_spans_across_lines(lines: list[str]) -> list[str]:
    """
    Ensure each standalone line has perfectly matched open and close
    <span> / </span> tags when red text spans multiple verses.
    """
    fixed_lines = []
    in_red = False
    red_open_tag = '<span style="color:#FF0000">'
    close_tag = '</span>'

    for line in lines:
        if not line:
            fixed_lines.append(line)
            continue

        prefix = red_open_tag if in_red else ""
        current_line = prefix + line

        # Remove redundancies within the same line
        current_line = re.sub(r'</span>\s*<span style="color:#FF0000">', '', current_line)

        opens = len(re.findall(r'<span style="color:#FF0000">', current_line))
        closes = len(re.findall(r'</span>', current_line))

        diff = opens - closes
        if diff > 0:
            current_line += close_tag * diff
            in_red = True
        elif diff < 0:
            for _ in range(abs(diff)):
                current_line = red_open_tag + current_line
            in_red = False
        else:
            in_red = False

        fixed_lines.append(current_line)

    return fixed_lines


class VerseLine:
    def __init__(self, text: str, verse_num: str = "", is_heading: bool = False):
        self.text = text
        self.verse_num = verse_num
        self.is_heading = is_heading


class BGBConverter:
    def __init__(self, docx_path: Path):
        self.doc = Document(str(docx_path))
        self.lines: list[str] = []
        self.current_book = ""
        self.current_chapter = ""
        self.current_verse = ""

        self.chapter_verses: list[VerseLine] = []
        self.chapter_footnotes: list[tuple[str, str, str]] = []
        self.text_calls_map: dict[str, str] = {}
        self.last_was_indented = False
        self.started = False

    def flush_chapter(self) -> None:
        if not self.chapter_verses and not self.chapter_footnotes:
            return

        def_lines = []
        for letter, v_num_fn, content in self.chapter_footnotes:
            target_v = self.text_calls_map.get(
                letter, v_num_fn or self.current_verse or "0"
            )
            key = f"{self.current_book}_{self.current_chapter}_{target_v}_{letter}"
            def_lines.append(f"[^{key}]: {content}")

        for obj in self.chapter_verses:
            self.lines.append(obj.text)

        if def_lines:
            if self.lines and self.lines[-1] != "":
                self.lines.append("")
            for d in def_lines:
                self.lines.append(d)
            self.lines.append("")

        self.chapter_verses.clear()
        self.chapter_footnotes.clear()
        self.text_calls_map.clear()
        self.last_was_indented = False

    def convert(self) -> str:
        self.lines.append(PROVENANCE_COMMENT.rstrip())
        self.lines.append("")

        for para in self.doc.paragraphs:
            txt = para.text.strip()
            if not txt:
                continue

            style_name = para.style.name if para.style else "Normal"

            # ── 1. Books ───────────────────────────────────────────────
            if style_name == "Heading 1":
                self.flush_chapter()
                self.current_book = txt.replace(" ", "_")
                self.current_chapter = ""
                self.current_verse = ""
                self.started = True
                if self.lines and self.lines[-1] != "":
                    self.lines.append("")
                self.lines.append(f"# **{txt}**")
                self.lines.append("")
                self.last_was_indented = False
                continue

            if not self.started:
                continue

            # ── 2. Chapters ────────────────────────────────────────────
            if style_name == "Heading 2":
                self.flush_chapter()
                m = re.search(r"(\d+)$", txt)
                self.current_chapter = m.group(1) if m else txt
                self.current_verse = ""
                if self.lines and self.lines[-1] != "":
                    self.lines.append("")
                self.lines.append(f"## **{txt}**")
                self.lines.append("")
                self.last_was_indented = False
                continue

            # ── 3. Sections ────────────────────────────────────────────
            if style_name in ("Heading 3", "hdg"):
                title_part, ref_part = "", ""
                for r in para.runs:
                    r_txt = r.text.replace("⇔", "").replace("  ", " ")
                    if not r_txt:
                        continue
                    if is_run_blue(r):
                        ref_part += r_txt
                    else:
                        title_part += r_txt

                title_part = title_part.strip()
                ref_part = ref_part.strip()

                if not ref_part:
                    m_ref = re.match(
                        r"^(.*?)\s*(\([A-Za-z0-9\s:;,\-–—]+\))\s*$", title_part
                    )
                    if m_ref:
                        title_part = m_ref.group(1).strip()
                        ref_part = m_ref.group(2).strip()

                if ref_part:
                    blue_ref = (
                        f'<span style="color:#0092F2">'
                        f"{html.escape(ref_part, quote=False)}</span>"
                    )
                    sec_line = f"### **{title_part}**<br>*{blue_ref}*"
                else:
                    sec_line = f"### **{title_part}**"

                if self.current_chapter:
                    self.chapter_verses.append(VerseLine("", is_heading=True))
                    self.chapter_verses.append(VerseLine(sec_line, is_heading=True))
                    self.chapter_verses.append(VerseLine("", is_heading=True))
                else:
                    if self.lines and self.lines[-1] != "":
                        self.lines.append("")
                    self.lines.append(sec_line)
                    self.lines.append("")
                self.last_was_indented = False
                continue

            # ── 4. Footnotes ───────────────────────────────────────────
            if re.match(r"^[a-z]\s+\d+", txt, re.IGNORECASE) or style_name == "foot":
                notes = parse_footnote_block(txt)
                for ltr, v_num, cnt in notes:
                    self.chapter_footnotes.append((ltr, v_num, cnt))
                continue

            # ── 5. Content ─────────────────────────────────────────────
            parts: list[str] = []
            curr_red = None
            curr_txt: list[str] = []

            vm_head = re.match(r"^(\d+)[\s\xa0]", txt)
            if vm_head:
                self.current_verse = vm_head.group(1)

            for child in para._element:
                tag = child.tag.split("}")[-1]

                if tag == "r":
                    run = docx.text.run.Run(child, para)
                    r_text = run.text
                    if not r_text:
                        continue

                    rPr = child.find(qn("w:rPr"))
                    if rPr is not None:
                        rstyle_ele = rPr.find(qn("w:rStyle"))
                        if (
                            rstyle_ele is not None
                            and rstyle_ele.attrib.get(qn("w:val")) == "reftext1"
                        ):
                            m_v = re.search(r"(\d+)", r_text)
                            if m_v:
                                self.current_verse = m_v.group(1)

                    r_text_clean = r_text.replace("⇔", "").replace("  ", " ")
                    r_text_clean = r_text_clean.replace("*", "\\*")

                    run_red = is_run_red(run, para)

                    if run_red == curr_red:
                        curr_txt.append(r_text_clean)
                    else:
                        if curr_txt:
                            merged = "".join(curr_txt)
                            if curr_red:
                                parts.append(
                                    f'<span style="color:#FF0000">'
                                    f"{html.escape(merged, quote=False)}</span>"
                                )
                            else:
                                parts.append(merged)
                        curr_red = run_red
                        curr_txt = [r_text_clean]

                elif tag == "hyperlink":
                    if curr_txt:
                        merged = "".join(curr_txt)
                        if curr_red:
                            parts.append(
                                f'<span style="color:#FF0000">'
                                f"{html.escape(merged, quote=False)}</span>"
                            )
                        else:
                            parts.append(merged)
                        curr_txt = []
                        curr_red = None

                    hl_text = "".join(
                        t.text for t in child.findall(".//" + qn("w:t")) if t.text
                    ).strip().lower()
                    if len(hl_text) == 1 and hl_text.isalpha():
                        self.text_calls_map[hl_text] = self.current_verse
                        key = (
                            f"{self.current_book}_{self.current_chapter}_"
                            f"{self.current_verse}_{hl_text}"
                        )
                        parts.append(f"[^{key}]")

            if curr_txt:
                merged = "".join(curr_txt)
                if curr_red:
                    parts.append(
                        f'<span style="color:#FF0000">'
                        f"{html.escape(merged, quote=False)}</span>"
                    )
                else:
                    parts.append(merged)

            content = "".join(parts).replace("\xa0", " ")

            verse_lines = bold_and_split_verse_numbers(content)

            for vl in verse_lines:
                m_v = re.search(r"\*\*(\d+)\*\*", vl)
                if m_v:
                    self.current_verse = m_v.group(1)

            is_indented, nbsp_prefix = get_paragraph_indentation(para)

            for i, vl in enumerate(verse_lines):
                if is_indented:
                    line_str = f"> {nbsp_prefix}{vl}<br>"
                else:
                    line_str = vl

                if self.current_chapter:
                    if not (is_indented and self.last_was_indented) and i == 0:
                        if self.chapter_verses and self.chapter_verses[-1].text != "":
                            self.chapter_verses.append(
                                VerseLine("", verse_num=self.current_verse)
                            )
                    self.chapter_verses.append(
                        VerseLine(line_str, verse_num=self.current_verse)
                    )
                else:
                    if not (is_indented and self.last_was_indented) and i == 0:
                        if self.lines and self.lines[-1] != "":
                            self.lines.append("")
                    self.lines.append(line_str)

            self.last_was_indented = is_indented

        self.flush_chapter()

        result = "\n".join(self.lines)
        result = re.sub(r"\n{4,}", "\n\n\n", result)
        return result.strip() + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Converte BGB DOCX para Markdown estruturado."
    )
    parser.add_argument("-i", "--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("-o", "--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    if not args.input.is_file():
        print(f"Erro: arquivo não encontrado: {args.input}", file=sys.stderr)
        return 1

    print(f"Convertendo: {args.input}")
    print(f"Saída      : {args.output}\n")

    converter = BGBConverter(args.input)
    markdown = converter.convert()

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(markdown, encoding="utf-8")

    print("✓ Concluído")
    print(f"  Linhas geradas : {markdown.count(chr(10)) + 1:,}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
