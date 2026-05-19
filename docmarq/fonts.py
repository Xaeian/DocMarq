# docmarq/fonts.py

"""Font configuration helpers.

Word fonts are referenced by family name only - actual rasterization is
done by the host system. We do NOT embed fonts here _(unlike PDFs where
embedding is the norm)_. If a target machine lacks the font, Word will
substitute. To embed fonts, save `settings.xml` with `<w:embedTrueTypeFonts/>`
- add later if needed.
"""

#-------------------------------------------------------------------------------------- Helpers

def is_safe_default(family:str) -> bool:
  """Returns `True` for fonts present on virtually all Word installs."""
  return family.lower() in {
    "calibri", "cambria", "arial", "times new roman", "verdana",
    "tahoma", "georgia", "courier new", "consolas", "segoe ui",
  }
