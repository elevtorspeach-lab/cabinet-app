#!/usr/bin/env python3
import re
import sys
from pathlib import Path


PAGE_WIDTH = 595
PAGE_HEIGHT = 842
MARGIN_X = 56
TOP_MARGIN = 60
BOTTOM_MARGIN = 56
FONT_REGULAR = "F1"
FONT_BOLD = "F2"
SIZE_BODY = 11
SIZE_H1 = 20
SIZE_H2 = 14
LINE_HEIGHT_BODY = 16
LINE_HEIGHT_H1 = 26
LINE_HEIGHT_H2 = 20


def escape_pdf_text(value: str) -> str:
    return value.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def wrap_text(text: str, max_chars: int) -> list[str]:
    words = text.split()
    if not words:
        return [""]
    lines = []
    current = words[0]
    for word in words[1:]:
        candidate = f"{current} {word}"
        if len(candidate) <= max_chars:
            current = candidate
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def parse_markdown(md: str) -> list[tuple[str, str]]:
    blocks: list[tuple[str, str]] = []
    for raw_line in md.splitlines():
        line = raw_line.rstrip()
        if not line.strip():
            blocks.append(("blank", ""))
            continue
        if line.startswith("# "):
            blocks.append(("h1", line[2:].strip()))
            continue
        if line.startswith("## "):
            blocks.append(("h2", line[3:].strip()))
            continue
        if line.startswith("- "):
            blocks.append(("bullet", line[2:].strip()))
            continue
        if re.match(r"^\d+\.\s+", line):
            blocks.append(("number", re.sub(r"^\d+\.\s+", "", line).strip()))
            continue
        blocks.append(("p", line.strip()))
    return blocks


def render_pages(blocks: list[tuple[str, str]]) -> list[list[tuple[str, str, int]]]:
    pages: list[list[tuple[str, str, int]]] = [[]]
    y = PAGE_HEIGHT - TOP_MARGIN

    def ensure_space(height: int):
        nonlocal y
        if y - height < BOTTOM_MARGIN:
            pages.append([])
            y = PAGE_HEIGHT - TOP_MARGIN

    for kind, text in blocks:
        if kind == "blank":
            y -= 8
            continue

        if kind == "h1":
            ensure_space(LINE_HEIGHT_H1 + 8)
            pages[-1].append((FONT_BOLD, text, y))
            pages[-1].append(("__SIZE__", str(SIZE_H1), y))
            y -= LINE_HEIGHT_H1 + 4
            continue

        if kind == "h2":
            ensure_space(LINE_HEIGHT_H2 + 6)
            pages[-1].append((FONT_BOLD, text, y))
            pages[-1].append(("__SIZE__", str(SIZE_H2), y))
            y -= LINE_HEIGHT_H2
            continue

        prefix = ""
        max_chars = 78
        if kind == "bullet":
            prefix = "- "
            max_chars = 74
        elif kind == "number":
            prefix = "1. "
            max_chars = 73

        wrapped = wrap_text(text, max_chars)
        for index, line in enumerate(wrapped):
            ensure_space(LINE_HEIGHT_BODY)
            content = f"{prefix}{line}" if index == 0 else f"{' ' * len(prefix)}{line}"
            pages[-1].append((FONT_REGULAR, content, y))
            pages[-1].append(("__SIZE__", str(SIZE_BODY), y))
            y -= LINE_HEIGHT_BODY
        y -= 2

    return pages


def build_stream(page_items: list[tuple[str, str, int]]) -> bytes:
    parts = ["BT", f"/{FONT_REGULAR} {SIZE_BODY} Tf", "0 g"]
    current_font = FONT_REGULAR
    current_size = SIZE_BODY
    for font, text, y in page_items:
        if font == "__SIZE__":
            continue
        for size_font, size_text, size_y in page_items:
            if size_font == "__SIZE__" and size_y == y:
                size = int(size_text)
                if current_size != size or current_font != font:
                    parts.append(f"/{font} {size} Tf")
                    current_font = font
                    current_size = size
                break
        parts.append(f"1 0 0 1 {MARGIN_X} {y} Tm ({escape_pdf_text(text)}) Tj")
    parts.append("ET")
    return "\n".join(parts).encode("latin-1", errors="replace")


def generate_pdf(markdown_path: Path, output_path: Path) -> None:
    text = markdown_path.read_text(encoding="utf-8")
    pages = render_pages(parse_markdown(text))

    objects: list[bytes] = []

    def add_object(data: bytes) -> int:
        objects.append(data)
        return len(objects)

    font_regular_id = add_object(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    font_bold_id = add_object(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>")

    page_ids = []
    content_ids = []
    placeholder_pages_id = add_object(b"<<>>")

    for page in pages:
        stream = build_stream(page)
        content_id = add_object(
            f"<< /Length {len(stream)} >>\nstream\n".encode("latin-1") + stream + b"\nendstream"
        )
        content_ids.append(content_id)
        page_id = add_object(b"<<>>")
        page_ids.append(page_id)

    kids = " ".join(f"{page_id} 0 R" for page_id in page_ids)
    objects[placeholder_pages_id - 1] = (
        f"<< /Type /Pages /Count {len(page_ids)} /Kids [{kids}] >>".encode("latin-1")
    )

    for idx, page_id in enumerate(page_ids):
        content_id = content_ids[idx]
        objects[page_id - 1] = (
            f"<< /Type /Page /Parent {placeholder_pages_id} 0 R "
            f"/MediaBox [0 0 {PAGE_WIDTH} {PAGE_HEIGHT}] "
            f"/Resources << /Font << /{FONT_REGULAR} {font_regular_id} 0 R /{FONT_BOLD} {font_bold_id} 0 R >> >> "
            f"/Contents {content_id} 0 R >>"
        ).encode("latin-1")

    catalog_id = add_object(f"<< /Type /Catalog /Pages {placeholder_pages_id} 0 R >>".encode("latin-1"))

    pdf = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for index, obj in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf.extend(f"{index} 0 obj\n".encode("latin-1"))
        pdf.extend(obj)
        pdf.extend(b"\nendobj\n")

    xref_start = len(pdf)
    pdf.extend(f"xref\n0 {len(objects) + 1}\n".encode("latin-1"))
    pdf.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        pdf.extend(f"{offset:010d} 00000 n \n".encode("latin-1"))
    pdf.extend(
        (
            f"trailer\n<< /Size {len(objects) + 1} /Root {catalog_id} 0 R >>\n"
            f"startxref\n{xref_start}\n%%EOF\n"
        ).encode("latin-1")
    )

    output_path.write_bytes(pdf)


def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: markdown_to_pdf.py input.md output.pdf")
        return 1
    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    generate_pdf(input_path, output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
