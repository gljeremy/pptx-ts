"""Dump what python-pptx reports for every upstream test deck (slides, layouts, masters, shapes,
placeholders, notes, core properties) plus the SHA-1 of each part after re-saving.

Usage: python scripts/crosscheck/py_side.py <python-pptx checkout> py.json
Then:  npx tsx scripts/crosscheck/ts_side.ts py.json <python-pptx checkout>
"""
import glob, io, json, sys, zipfile
from pptx import Presentation
from pptx.util import Emu
decks = sorted(glob.glob(sys.argv[1] + "/tests/test_files/*.pptx") + glob.glob(sys.argv[1] + "/features/steps/test_files/*.pptx"))
out = {}
def shape_info(sh):
    d = {"name": sh.name, "id": sh.shape_id, "ph": sh.is_placeholder}
    for a in ("left", "top", "width", "height"):
        try: d[a] = getattr(sh, a)
        except Exception as e: d[a] = "ERR"
    if sh.is_placeholder:
        pf = sh.placeholder_format
        d["phIdx"] = pf.idx; d["phType"] = pf.type.name if pf.type is not None else None
    return d
for f in decks:
    key = f.replace("\\", "/").split("python-pptx/")[1]
    try:
        prs = Presentation(f)
        info = {
            "slides": [{"id": s.slide_id, "name": s.name, "layout": s.slide_layout.name, "shapes": [shape_info(x) for x in s.shapes],
                        "placeholders": [p.placeholder_format.idx for p in s.placeholders], "hasNotes": s.has_notes_slide,
                        "followBg": s.follow_master_background} for s in prs.slides],
            "masters": [{"layouts": [l.name for l in m.slide_layouts], "shapes": [x.name for x in m.shapes]} for m in prs.slide_masters],
            "size": [prs.slide_width, prs.slide_height],
            "core": {k: str(getattr(prs.core_properties, k)) for k in ("author", "category", "comments", "content_status", "created", "identifier", "keywords", "language", "last_modified_by", "last_printed", "modified", "revision", "subject", "title", "version")},
        }
        notes = []
        for s in prs.slides:
            if s.has_notes_slide:
                tf = s.notes_slide.notes_text_frame
                notes.append(tf.text if tf is not None else None)
        info["notes"] = notes
        buf = io.BytesIO(); prs.save(buf)
        z = zipfile.ZipFile(buf)
        info["saved"] = {n: __import__("hashlib").sha1(z.read(n)).hexdigest() for n in z.namelist()}
        out[key] = info
    except Exception as e:
        out[key] = {"error": type(e).__name__ + ": " + str(e)[:120]}
json.dump(out, open(sys.argv[2], "w"), indent=1, default=str)
print(len(out))
