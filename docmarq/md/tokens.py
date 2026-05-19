"""
Pure helpers for walking markdown-it tokens.

`markdown-it-py` returns a flat token list with `*_open` / `*_close` pairs.
These helpers cover the two operations we need everywhere: pulling an HTML
attribute off an open token and finding the matching close for a range.
"""
import re
from markdown_it.token import Token

#-------------------------------------------------------------------------------------- Attrs

def get_attr(token:Token, name:str) -> str|None:
  """Get HTML attribute value from a token's attrs list/dict."""
  if not token.attrs:
    return None
  if isinstance(token.attrs, dict):
    return token.attrs.get(name)
  for k, v in token.attrs:
    if k == name:
      return v
  return None

#----------------------------------------------------------------------------- Range scanning

def find_close(tokens:list[Token], start:int, open_type:str, close_type:str) -> int:
  """Return index of the matching close token for a balanced open token.
  Falls back to last index when no close is found - safer than raising
  since `markdown-it` always produces balanced output for valid input.
  """
  depth = 0
  for j in range(start, len(tokens)):
    tt = tokens[j].type
    if tt == open_type:
      depth += 1
    elif tt == close_type:
      depth -= 1
      if depth == 0:
        return j
  return len(tokens) - 1

#---------------------------------------------------------------------------------- Callouts

# `> [!NOTE]` style marker matching GitHub-flavored markdown callouts.
CALLOUT_RE = re.compile(r"^\s*\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*$",
  re.IGNORECASE)

#---------------------------------------------------------------------------- Directives

# HTML-comment directives for layout control. Whitespace-tolerant inside
# the comment, case-insensitive on the name. Extra tokens inside the
# comment disqualify the match - `<!-- pagebreak xxx -->` is not a
# directive, just a regular HTML comment that gets dropped.
# Symmetric with `pdfmarq.md.md_html` detectors.
_PAGEBREAK_RE = re.compile(r"\s*<!--\s*pagebreak\s*-->\s*", re.IGNORECASE)
_GROUP_OPEN_RE = re.compile(r"\s*<!--\s*group\s*-->\s*", re.IGNORECASE)
_GROUP_CLOSE_RE = re.compile(r"\s*<!--\s*/\s*group\s*-->\s*", re.IGNORECASE)

def is_pagebreak_directive(content:str) -> bool:
  """True for `<!-- pagebreak -->` html_block content."""
  return bool(_PAGEBREAK_RE.fullmatch(content))

def is_group_open_directive(content:str) -> bool:
  """True for `<!-- group -->` opening directive."""
  return bool(_GROUP_OPEN_RE.fullmatch(content))

def is_group_close_directive(content:str) -> bool:
  """True for `<!-- /group -->` closing directive."""
  return bool(_GROUP_CLOSE_RE.fullmatch(content))
