"""Record what Pillow reports (format, size, info["dpi"]) for the images used by
test/image/headers.test.ts, so the header reader can be checked against Pillow without python.

Usage: python scripts/image_golden.py test/fixtures/images_pillow.json test/fixtures/images test/test_files
"""
import glob
import json
import os
import sys

from PIL import Image

EXTS = ("png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "wmf", "emf", "eps")
out = {}
for d in sys.argv[2:]:
    for f in sorted(glob.glob(d + "/**/*", recursive=True)):
        if f.lower().rsplit(".", 1)[-1] not in EXTS:
            continue
        key = os.path.relpath(f).replace("\\", "/")
        try:
            im = Image.open(f)
            dpi = im.info.get("dpi")
            dpi = [float(x) for x in dpi] if isinstance(dpi, tuple) else (float(dpi) if dpi is not None else None)
            out[key] = {"format": im.format, "size": list(im.size), "dpi": dpi}
        except Exception as e:
            out[key] = {"error": type(e).__name__}
json.dump(out, open(sys.argv[1], "w"), indent=1, sort_keys=True)
print(len(out), "images")
