"""Unit tests for pure helpers - utils, fonts, layout.

Cheap and fast, no docx. Catches regressions in the bits the rest of
the library leans on (color parsing, margin parsing, EMU conversion).
"""
import pytest
from docmarq.utils import (
  to_mm, mm_to_emu, pt_to_emu, parse_color, rgb255, color_hex,
  parse_margin, smaller_size, tight_line_height,
)
from docmarq.constants import Align

#--------------------------------------------------------------------------------------- Units

def test_to_mm_identity():
  assert to_mm(10, "mm") == 10

def test_to_mm_cm():
  assert to_mm(1, "cm") == pytest.approx(10)

def test_to_mm_in():
  assert to_mm(1, "in") == pytest.approx(25.4)

def test_to_mm_bad_unit():
  with pytest.raises(ValueError):
    parse_margin  # silence unused import warning
    to_mm(1, "furlong")

def test_mm_to_emu():
  # 1 inch = 25.4 mm = 914400 EMU
  assert mm_to_emu(25.4) == 914400

def test_pt_to_emu():
  # 1 pt = 12700 EMU
  assert pt_to_emu(1) == 12700

#------------------------------------------------------------------------------------- Colors

def test_parse_color_returns_floats():
  # Regression: previously returned ints 0-255. Canonical form is now floats
  # 0-1 (matching `pdfmarq.parse_color`). Conversion to ints happens at the
  # OOXML boundary via `rgb255`.
  r, g, b = parse_color("#FF0000")
  assert (r, g, b) == pytest.approx((1.0, 0.0, 0.0))
  assert all(isinstance(c, float) for c in (r, g, b))

def test_parse_color_hex_short():
  assert parse_color("#F00") == pytest.approx((1.0, 0.0, 0.0))

def test_parse_color_tuple_passthrough():
  assert parse_color((0.2, 0.4, 0.8)) == pytest.approx((0.2, 0.4, 0.8))

def test_parse_color_none():
  assert parse_color(None) == (0.0, 0.0, 0.0)

def test_parse_color_invalid_hex():
  with pytest.raises(ValueError):
    parse_color("#GGGGGG")
  with pytest.raises(ValueError):
    parse_color("#12345")

def test_rgb255_from_floats():
  assert rgb255((1.0, 0.0, 0.5)) == (255, 0, 128)

def test_rgb255_from_hex():
  assert rgb255("#1f2328") == (0x1F, 0x23, 0x28)

def test_color_hex_from_floats():
  assert color_hex((1.0, 0.0, 0.0)) == "FF0000"
  assert color_hex((0.5, 0.5, 0.5)) == "808080"

def test_color_hex_from_hex():
  # Round-trip: hex in, hex out (uppercase).
  assert color_hex("#1f2328") == "1F2328"

#------------------------------------------------------------------------------------- Margin

def test_parse_margin_scalar():
  assert parse_margin(10) == (10, 10, 10, 10)

def test_parse_margin_2tuple():
  assert parse_margin((10, 20)) == (10, 20, 10, 20)

def test_parse_margin_3tuple():
  assert parse_margin((10, 20, 30)) == (10, 20, 30, 20)

def test_parse_margin_4tuple():
  assert parse_margin((1, 2, 3, 4)) == (1, 2, 3, 4)

#------------------------------------------------------------------------------------ Helpers

def test_smaller_size_steps():
  # Per typographic ladder: 11→10, 12→11, 14→12, 16→14.
  assert smaller_size(11) == 10
  assert smaller_size(12) == 11
  assert smaller_size(14) == 12
  assert smaller_size(16) == 14

def test_smaller_size_clamps():
  # Body smaller than the floor returns the floor.
  assert smaller_size(7, min_pt=7) == 7

def test_tight_line_height_body():
  assert tight_line_height(11) == 1.15

def test_tight_line_height_display():
  assert tight_line_height(24) == 1.0

#-------------------------------------------------------------------------------------- Align

def test_align_constants_match_pdfmarq():
  # Cross-lib reuse: both libs share these single-char codes.
  assert Align.LEFT == "L"
  assert Align.RIGHT == "R"
  assert Align.CENTER == "C"
  assert Align.JUSTIFY == "J"
