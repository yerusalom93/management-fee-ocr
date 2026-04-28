from __future__ import annotations

import argparse
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import fitz
from PIL import Image, ImageFilter, ImageOps


def render_one(args: tuple[str, str, int, float, bool, bool, bool]) -> str:
    pdf_name, out_name, page_index, scale, preprocess, force, alpha = args
    out = Path(out_name)
    if out.exists() and not force:
        return f"{page_index + 1}\tcached\t{out}"

    doc = fitz.open(pdf_name)
    page = doc[page_index]
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=alpha)
    pix.save(out)
    doc.close()

    if preprocess:
        with Image.open(out) as img:
            gray = ImageOps.grayscale(img)
            gray = ImageOps.autocontrast(gray, cutoff=1)
            gray = gray.filter(ImageFilter.UnsharpMask(radius=1.2, percent=135, threshold=3))
            gray.save(out, optimize=True)

    return f"{page_index + 1}\t{pix.width}x{pix.height}\t{out}"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf")
    parser.add_argument("out_dir")
    parser.add_argument("--scale", type=float, default=3.0)
    parser.add_argument("--workers", type=int, default=0)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--preprocess", action="store_true")
    parser.add_argument("--alpha", action="store_true")
    ns = parser.parse_args()

    pdf_path = Path(ns.pdf)
    out_dir = Path(ns.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    page_count = doc.page_count
    doc.close()

    jobs = [
        (
            str(pdf_path),
            str(out_dir / f"page{idx + 1:03d}.png"),
            idx,
            ns.scale,
            ns.preprocess,
            ns.force,
            ns.alpha,
        )
        for idx in range(page_count)
    ]

    workers = ns.workers or min(4, page_count)
    if workers <= 1 or page_count <= 1:
        for job in jobs:
            print(render_one(job))
        return

    with ProcessPoolExecutor(max_workers=workers) as executor:
        for line in executor.map(render_one, jobs):
            print(line)


if __name__ == "__main__":
    main()
