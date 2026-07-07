#!/usr/bin/env python3
"""
Step 1: OCR all 162 pages, detect chapter structure.
"""

import fitz
import os
import json
import re
from rapidocr_onnxruntime import RapidOCR

PDF_PATH = "/home/HY/pdf_to_doc/液压元件的寿命试验(段长宝等编)(2).pdf"
OUTPUT_DIR = "/home/HY/pdf_to_doc/temp_pages"
OCR_RESULT = "/home/HY/pdf_to_doc/ocr_results.json"
CHAPTERS_FILE = "/home/HY/pdf_to_doc/chapters.json"

os.makedirs(OUTPUT_DIR, exist_ok=True)


def main():
    print("Loading RapidOCR...")
    ocr = RapidOCR()
    doc = fitz.open(PDF_PATH)
    total = doc.page_count
    print(f"Total pages: {total}")

    all_results = {}
    chapters = []

    for i in range(total):
        page_num = i + 1
        print(f"Page {page_num}/{total}...", end=" ", flush=True)

        pix = doc[i].get_pixmap(dpi=200)
        img_path = os.path.join(OUTPUT_DIR, f"page_{page_num:03d}.png")
        pix.save(img_path)

        result, _ = ocr(img_path)

        lines = []
        if result:
            for bbox, text, conf in result:
                text = text.strip()
                if text:
                    y_center = (bbox[0][1] + bbox[2][1]) / 2
                    x_left = bbox[0][0]
                    lines.append({
                        "text": text,
                        "confidence": round(float(conf), 4),
                        "y": round(float(y_center), 1),
                        "x": round(float(x_left), 1),
                        "bbox": [[round(float(p[0]), 1), round(float(p[1]), 1)] for p in bbox]
                    })

        lines.sort(key=lambda l: (l["y"], l["x"]))
        all_results[str(page_num)] = lines
        print(f"{len(lines)} lines")

        # Detect chapter headings
        for ln in lines[:10]:  # usually in top portion
            t = ln["text"]
            m = re.match(r"第([一二三四五六七八九十百\d]+)章", t)
            if m:
                chapters.append({
                    "page": page_num,
                    "type": "chapter",
                    "title": t,
                    "chapter_num": m.group(1)
                })
                print(f"  >>> CHAPTER: {t}")
                break

    doc.close()

    # Save
    with open(OCR_RESULT, "w", encoding="utf-8") as f:
        json.dump(all_results, f, ensure_ascii=False, indent=1)
    with open(CHAPTERS_FILE, "w", encoding="utf-8") as f:
        json.dump(chapters, f, ensure_ascii=False, indent=1)

    print(f"\n=== Chapter Structure ===")
    for ch in chapters:
        print(f"  Page {ch['page']:3d} | {ch['title']}")
    print(f"Total chapters: {len(chapters)}")


if __name__ == "__main__":
    main()
