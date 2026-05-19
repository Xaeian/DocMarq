# docmarq/md/__init__.py

"""
Markdown to DOCX rendering for `docmarq`.

Mirrors `pdfmarq.md` API: drop in `md_to_docx` for one-shot conversion, or
construct `MarkdownRenderer(doc, style)` for embedding markdown into a
larger document.

  pip install docmarq[md]
  from docmarq.md import md_to_docx, MarkdownStyle, lang_style

Bundled dependencies:
  - markdown-it-py  # parser
"""

#------------------------------------------------------------------------- Extras for auto-toml

__extras__ = ("md", ["markdown-it-py", "mdit-py-plugins", "PyYAML"])

#----------------------------------------------------------------------------------- Public API

from .style import MarkdownStyle
from .renderer import MarkdownRenderer, md_to_docx
from .presets import lang_style, LANG_PRESETS

__all__ = [
  "MarkdownStyle",
  "MarkdownRenderer",
  "md_to_docx",
  "lang_style",
  "LANG_PRESETS",
]
