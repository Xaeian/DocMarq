"""End-to-end smoke tests for DOCX rendering.

Pass=no crash + valid ZIP-magic header on disk. Catches regressions in
the public API and `python-docx` integration without locking in
specific byte output.
"""
from docmarq import DOCX, Align, Styles, TableStyle, RichSegment
from docmarq.tests.conftest import assert_valid_docx

#--------------------------------------------------------------------------------- DOCX basic

def test_docx_empty(tmp_path):
  path = tmp_path / "empty.docx"
  with DOCX(str(path)) as doc:
    doc.para("hi")
  assert_valid_docx(path)

def test_docx_paragraphs_and_runs(tmp_path):
  path = tmp_path / "runs.docx"
  with DOCX(str(path)) as doc:
    doc.text("plain ").text("bold", bold=True).text(" italic", italic=True).enter()
    doc.text("code", code=True)
  assert_valid_docx(path)

def test_docx_headings(tmp_path):
  path = tmp_path / "head.docx"
  with DOCX(str(path)) as doc:
    for level in range(1, 7):
      doc.heading(f"Heading level {level}", level=level)
      doc.para(f"Body under h{level}.")
  assert_valid_docx(path)

def test_docx_table_simple(tmp_path):
  path = tmp_path / "tab.docx"
  with DOCX(str(path)) as doc:
    doc.table(
      [["a1", "b1"], ["a2", "b2"]],
      header=["A", "B"],
      aligns=[Align.LEFT, Align.RIGHT],
    )
  assert_valid_docx(path)

def test_docx_lists(tmp_path):
  path = tmp_path / "lists.docx"
  with DOCX(str(path)) as doc:
    doc.bullet("one").bullet("two").bullet("three")
    doc.ordered("first").ordered("second")
  assert_valid_docx(path)

def test_docx_blockquote_hr_code(tmp_path):
  path = tmp_path / "blocks.docx"
  with DOCX(str(path)) as doc:
    doc.blockquote("Quote text.")
    doc.hr()
    doc.code_block("def foo():\n  return 42", language="python")
  assert_valid_docx(path)

def test_docx_colors_hex_and_tuple(tmp_path):
  # Regression: `parse_color` now returns floats 0-1 in both libs; `rgb255`
  # converts at the OOXML boundary. Mixing hex strings and float tuples in
  # the same document must produce no crash.
  path = tmp_path / "colors.docx"
  with DOCX(str(path)) as doc:
    doc.text("hex", color="#1f2328").enter()
    doc.text("tuple", color=(0.03, 0.41, 0.85))
  assert_valid_docx(path)

def test_docx_metadata(tmp_path):
  path = tmp_path / "meta.docx"
  with DOCX(str(path)) as doc:
    doc.metadata(title="T", author="X", subject="S", keywords="k")
    doc.para("body")
  assert_valid_docx(path)

def test_style_presets_independent():
  # Each access must yield a fresh `Style` so users can't accidentally
  # mutate shared state by tweaking a preset.
  a = Styles.BOLD
  b = Styles.BOLD
  assert a is not b
  a.font_size = 99
  assert b.font_size != 99
