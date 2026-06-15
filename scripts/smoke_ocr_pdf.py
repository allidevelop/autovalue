from __future__ import annotations

import argparse
from pathlib import Path

from realtify.ocr import ocr_image
from realtify.pdf_tools import extract_text_layer, render_pdf_pages


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--out", type=Path, default=Path("outputs/smoke_ocr"))
    args = parser.parse_args()

    text_layer = extract_text_layer(args.pdf)
    print(f"text_layer_chars={len(text_layer)}")
    pages = render_pdf_pages(args.pdf, args.out, first_page=1, last_page=1, dpi=200)
    print(f"rendered={pages[0]}")
    text = ocr_image(pages[0])
    print(f"ocr_chars={len(text)}")
    print("ocr_preview:")
    for line in [line.strip() for line in text.splitlines() if line.strip()][:30]:
        print(line)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

