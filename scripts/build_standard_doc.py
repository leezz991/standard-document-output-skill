#!/usr/bin/env python3
"""Build a standard DOCX in official or general mode from simple Markdown."""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from docx import Document


STYLE_BY_LEVEL = {
    1: "Title",
    2: "Heading 1",
    3: "Heading 2",
    4: "Heading 3",
    5: "Heading 4",
}


def parse_markdown(path: Path, mode: str) -> list[tuple[str, str]]:
    items: list[tuple[str, str]] = []
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        text = raw.strip()
        if not text:
            continue
        if text.startswith("> "):
            if mode != "general":
                raise ValueError("Subtitle marker '> ' is only supported in general mode.")
            items.append(("图表内容", text[2:].strip()))
            continue
        level = 0
        while level < len(text) and text[level] == "#":
            level += 1
        if 1 <= level <= 5 and len(text) > level and text[level].isspace():
            items.append((STYLE_BY_LEVEL[level], text[level:].strip()))
        else:
            items.append(("Normal", text))
    if not items:
        raise ValueError("Input contains no non-empty paragraphs.")
    return items


def clear_body(document: Document) -> None:
    body = document._element.body
    for child in list(body):
        if child.tag.endswith("}sectPr"):
            continue
        body.remove(child)


def build(input_path: Path, output_path: Path, template_path: Path, mode: str) -> None:
    if input_path.resolve() == output_path.resolve():
        raise ValueError("Input and output paths must differ.")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)
    document = Document(output_path)
    clear_body(document)
    for style_name, text in parse_markdown(input_path, mode):
        document.add_paragraph(text, style=style_name)
    document.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="UTF-8 Markdown source")
    parser.add_argument("output", type=Path, help="Output .docx path")
    parser.add_argument("--mode", required=True, choices=("official", "general"))
    parser.add_argument("--template", type=Path, help="Optional template override")
    args = parser.parse_args()
    if args.output.suffix.lower() != ".docx":
        parser.error("output must use the .docx extension")
    template = args.template or (
        Path(__file__).resolve().parents[1]
        / "assets"
        / ("gongwen.docx" if args.mode == "official" else "reference.docx")
    )
    build(args.input, args.output, template, args.mode)


if __name__ == "__main__":
    main()
