"""Record Pillow's text boxes and python-pptx's fit_text() sizes for test/text/fit.test.ts.

Usage: python scripts/text_golden.py test/test_files/calibriz.ttf test/fixtures/text_golden.json
"""
import json
import sys

from PIL import ImageFont
from pptx import Presentation
from pptx.util import Inches

font = sys.argv[1]
texts = ["Ty", "Hello world", "The quick brown fox", "jump", "AV To", "iiii", "WWW", "a", " lead", "x."]
bboxes = []
for size in [6, 8, 10, 11, 12, 14, 18, 24, 32, 48, 72]:
    f = ImageFont.truetype(font, size)
    for t in texts:
        bboxes.append([size, t, list(f.getbbox(t))])

fit_texts = ["Hello world", "The quick brown fox jumps over the lazy dog", "A considerably longer paragraph of text that will need to wrap across several lines in the box", "Ty", "Supercalifragilistic"]
boxes = [(1, 0.5), (2, 1), (3, 2), (4, 1), (6, 3), (1.5, 3)]
fits = []
prs = Presentation()
s = prs.slides.add_slide(prs.slide_layouts[6])
for t in fit_texts:
    for w, h in boxes:
        tb = s.shapes.add_textbox(0, 0, Inches(w), Inches(h))
        tb.text_frame.text = t
        try:
            tb.text_frame.fit_text(font_family="Calibri", max_size=40, bold=True, italic=True, font_file=font)
            fits.append([t, w, h, tb.text_frame.paragraphs[0].runs[0].font.size.pt])
        except TypeError:  # python-pptx fails when one word is wider than the box
            fits.append([t, w, h, "error"])
json.dump({"bboxes": bboxes, "fits": fits}, open(sys.argv[2], "w"), indent=0)
