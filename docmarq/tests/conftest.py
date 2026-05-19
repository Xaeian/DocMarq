"""Shared test fixtures.

DOCX smoke tests confirm the resulting file exists and begins with the
ZIP magic (`PK\\x03\\x04`) since `.docx` is a ZIP container. No deeper
validation - keeps the suite zero-extra-deps.
"""
from pathlib import Path
import sys

# Make `docmarq` importable when running tests from inside the package dir.
_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
  sys.path.insert(0, str(_ROOT))

#----------------------------------------------------------------------------------- Assertions

def assert_valid_docx(path:str|Path, min_size:int=2000):
  """Validate that `path` points to a real-looking DOCX file."""
  p = Path(path)
  assert p.exists(), f"DOCX not created: {p}"
  size = p.stat().st_size
  assert size >= min_size, f"DOCX suspiciously small ({size} bytes): {p}"
  with p.open("rb") as f:
    head = f.read(4)
  assert head == b"PK\x03\x04", f"Not a DOCX/ZIP (header={head!r}): {p}"
