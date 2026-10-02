"""
Stand-ins for the typefaces a deck asks for and a reader may not have.

Office's own typefaces, like Aptos, are usually not installed where a page is
read, and the typeface a browser falls back on sets wider or narrower, so
text runs under the pictures next to it or wraps where it did not. A stand-in
here is a common typeface scaled with ``size-adjust`` so that a line of text
sets the same width it does in the original.
"""

__all__ = [
    "font_faces",
    "font_stack",
]

_STYLES = [("normal", "normal"), ("bold", "normal"), ("normal", "italic"), ("bold", "italic")]
"""The faces of a typeface, as CSS ``font-weight`` and ``font-style``."""

_LOCAL = {
    "Segoe UI": [
        ("Segoe UI", "SegoeUI"),
        ("Segoe UI Bold", "SegoeUI-Bold"),
        ("Segoe UI Italic", "SegoeUI-Italic"),
        ("Segoe UI Bold Italic", "SegoeUI-BoldItalic"),
    ],
    "Arial": [
        ("Arial", "ArialMT"),
        ("Arial Bold", "Arial-BoldMT"),
        ("Arial Italic", "Arial-ItalicMT"),
        ("Arial Bold Italic", "Arial-BoldItalicMT"),
    ],
    "Liberation Sans": [
        ("Liberation Sans", "LiberationSans"),
        ("Liberation Sans Bold", "LiberationSans-Bold"),
        ("Liberation Sans Italic", "LiberationSans-Italic"),
        ("Liberation Sans Bold Italic", "LiberationSans-BoldItalic"),
    ],
}
"""The names ``local()`` finds each face of a stand-in by: full and PostScript."""

_SIZE_ADJUST = {
    "Aptos": {
        "Segoe UI": [96.9, 95.2, 99.0, 96.1],
        "Arial": [95.1, 94.3, 95.0, 94.3],
    },
    "Aptos Display": {
        "Segoe UI": [90.9, 89.2, 93.0, 90.2],
        "Arial": [89.3, 88.3, 89.2, 88.5],
    },
}
"""
How much to scale each stand-in, by face, for it to set a line of text as
wide as the typeface it stands in for. Measured on a paragraph of slide text
with the typefaces Office installs. Liberation Sans shares Arial's widths.
"""

_TWINS = {
    "Calibri": "Carlito",
    "Cambria": "Caladea",
}
"""Free typefaces drawn to the same widths as Office's older ones."""


def _stand_ins(typeface: str) -> "list[tuple[str, str]]":
    """Each stand-in for a typeface, with the measurements it borrows."""
    measured = _SIZE_ADJUST.get(typeface, {})
    stand_ins = [(name, name) for name in measured]
    if "Arial" in measured:
        stand_ins.append(("Liberation Sans", "Arial"))
    return stand_ins


def font_stack(typeface: str, quote: str = '"') -> str:
    """
    A CSS ``font-family`` list for a typeface: the typeface itself, then its
    stand-ins, ending with the generic families every reader has.
    """
    names = [typeface]
    names += [f"{typeface} ({stand_in})" for stand_in, _ in _stand_ins(typeface)]
    if typeface in _TWINS:
        names.append(_TWINS[typeface])
    quoted = [f"{quote}{name}{quote}" for name in names]
    return ", ".join(quoted + [f"{quote}Segoe UI{quote}", "system-ui", "sans-serif"])


def font_faces(page: str) -> str:
    """The ``@font-face`` rules for the stand-ins of every typeface a page uses."""
    rules = []
    for typeface in _SIZE_ADJUST:
        # Found by its stand-ins' names, which read the same however the
        # quotes around them were written or escaped.
        if f"{typeface} (" not in page:
            continue
        for stand_in, measured_as in _stand_ins(typeface):
            adjustments = _SIZE_ADJUST[typeface][measured_as]
            for (weight, style), names, adjust in zip(_STYLES, _LOCAL[stand_in], adjustments):
                sources = ", ".join(f'local("{name}")' for name in names)
                rules.append(
                    f'  @font-face {{ font-family: "{typeface} ({stand_in})"; src: {sources}; '
                    f"font-weight: {weight}; font-style: {style}; size-adjust: {adjust}%; }}"
                )
    return "\n".join(rules)
