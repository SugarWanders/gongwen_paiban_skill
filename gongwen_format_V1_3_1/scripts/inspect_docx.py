from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python inspect_docx.py <document.docx>", file=sys.stderr)
        return 2

    path = Path(sys.argv[1])
    if not path.exists():
        print(f"File not found: {path}", file=sys.stderr)
        return 1

    try:
        from docx import Document
    except ImportError:
        print("Missing dependency python-docx. Install with: python -m pip install python-docx", file=sys.stderr)
        return 1

    document = Document(str(path))
    paragraphs = document.paragraphs
    data: dict[str, Any] = {
        "path": str(path),
        "paragraph_count": len(paragraphs),
        "title_text": paragraphs[0].text if paragraphs else "",
        "title_alignment": alignment_name(paragraphs[0].alignment) if paragraphs else None,
        "has_blank_after_title": len(paragraphs) > 1 and paragraphs[1].text == "",
        "paragraphs": [],
    }

    for index, paragraph in enumerate(paragraphs[:12]):
        first_run = paragraph.runs[0] if paragraph.runs else None
        data["paragraphs"].append(
            {
                "index": index,
                "text": paragraph.text,
                "alignment": alignment_name(paragraph.alignment),
                "first_line_indent_pt": paragraph.paragraph_format.first_line_indent.pt
                if paragraph.paragraph_format.first_line_indent
                else None,
                "line_spacing_pt": paragraph.paragraph_format.line_spacing.pt
                if hasattr(paragraph.paragraph_format.line_spacing, "pt")
                else paragraph.paragraph_format.line_spacing,
                "first_run_font": first_run.font.name if first_run else None,
                "first_run_size_pt": first_run.font.size.pt if first_run and first_run.font.size else None,
                "first_run_bold": first_run.bold if first_run else None,
            }
        )

    print(json.dumps(data, ensure_ascii=False, indent=2))
    return 0


def alignment_name(value: Any) -> str | None:
    if value is None:
        return None
    return str(value).split(" ")[0]


if __name__ == "__main__":
    raise SystemExit(main())
