from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image


REGIONS = {
    # Coordinates are in the historical 2x render coordinate space.
    "lower_table": (70, 1040, 900, 2380),
}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("image_dir")
    parser.add_argument("out_dir")
    parser.add_argument("--region", choices=sorted(REGIONS), default="lower_table")
    parser.add_argument("--scale", type=float, default=3.0)
    parser.add_argument("--base-scale", type=float, default=2.0)
    parser.add_argument("--force", action="store_true")
    ns = parser.parse_args()

    image_dir = Path(ns.image_dir)
    out_dir = Path(ns.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    ratio = ns.scale / ns.base_scale
    box = tuple(int(round(value * ratio)) for value in REGIONS[ns.region])
    for src in sorted(image_dir.glob("page*.png")):
        out = out_dir / src.name
        if out.exists() and not ns.force:
            print(f"{src.stem}\tcached\t{out}")
            continue
        with Image.open(src) as img:
            img.crop(box).save(out, optimize=True)
        print(f"{src.stem}\t{ns.region}\t{out}")


if __name__ == "__main__":
    main()
