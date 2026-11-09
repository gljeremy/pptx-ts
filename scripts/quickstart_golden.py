"""Write python-pptx's output for the examples in its docs/user/quickstart.rst.

test/quickstart.test.ts runs the same examples, as written in README.md's "Getting started",
with pptx-ts and compares the decks part by part.

Usage: python scripts/quickstart_golden.py <python-pptx checkout> test/fixtures/quickstart
"""
import os
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

checkout, out_dir = sys.argv[1], sys.argv[2]
img_path = os.path.join(checkout, "features", "steps", "test_files", "monty-truth.png")
os.makedirs(out_dir, exist_ok=True)


def hello_world():
    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "Hello, World!"
    subtitle.text = "python-pptx was here!"
    return prs


def bullet_slide():
    prs = Presentation()
    bullet_slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(bullet_slide_layout)
    shapes = slide.shapes
    title_shape = shapes.title
    body_shape = shapes.placeholders[1]
    title_shape.text = "Adding a Bullet Slide"
    tf = body_shape.text_frame
    tf.text = "Find the bullet slide layout"
    p = tf.add_paragraph()
    p.text = "Use _TextFrame.text for first bullet"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Use _TextFrame.add_paragraph() for subsequent bullets"
    p.level = 2
    return prs


def add_textbox():
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    left = top = width = height = Inches(1)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.text = "This is text inside a textbox"
    p = tf.add_paragraph()
    p.text = "This is a second paragraph that's bold"
    p.font.bold = True
    p = tf.add_paragraph()
    p.text = "This is a third paragraph that's big"
    p.font.size = Pt(40)
    return prs


def add_picture():
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    left = top = Inches(1)
    slide.shapes.add_picture(img_path, left, top)
    left = Inches(5)
    height = Inches(5.5)
    slide.shapes.add_picture(img_path, left, top, height=height)
    return prs


def add_shape():
    prs = Presentation()
    title_only_slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(title_only_slide_layout)
    shapes = slide.shapes
    shapes.title.text = "Adding an AutoShape"
    left = Inches(0.93)
    top = Inches(3.0)
    width = Inches(1.75)
    height = Inches(1.0)
    shape = shapes.add_shape(MSO_SHAPE.PENTAGON, left, top, width, height)
    shape.text = "Step 1"
    left = left + width - Inches(0.4)
    width = Inches(2.0)
    for n in range(2, 6):
        shape = shapes.add_shape(MSO_SHAPE.CHEVRON, left, top, width, height)
        shape.text = "Step %d" % n
        left = left + width - Inches(0.4)
    return prs


def add_table():
    prs = Presentation()
    title_only_slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(title_only_slide_layout)
    shapes = slide.shapes
    shapes.title.text = "Adding a Table"
    rows = cols = 2
    left = top = Inches(2.0)
    width = Inches(6.0)
    height = Inches(0.8)
    table = shapes.add_table(rows, cols, left, top, width, height).table
    table.columns[0].width = Inches(2.0)
    table.columns[1].width = Inches(4.0)
    table.cell(0, 0).text = "Foo"
    table.cell(0, 1).text = "Bar"
    table.cell(1, 0).text = "Baz"
    table.cell(1, 1).text = "Qux"
    return prs


for fn in (hello_world, bullet_slide, add_textbox, add_picture, add_shape, add_table):
    fn().save(os.path.join(out_dir, fn.__name__ + ".pptx"))
    print(fn.__name__)
