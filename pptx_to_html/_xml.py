"""Namespaces and small helpers for reading the XML inside a deck."""

from lxml import etree

__all__ = [
    "NS",
    "find",
    "findall",
    "integer",
    "local",
    "qn",
    "require",
]

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "asvg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
}
"""The XML namespaces a deck uses, by the prefixes PowerPoint gives them."""


def qn(tag: str) -> str:
    """Expand a prefixed tag like ``a:off`` into Clark notation."""
    prefix, name = tag.split(":")
    return f"{{{NS[prefix]}}}{name}"


def local(element: etree._Element) -> str:
    """The tag of an element without its namespace."""
    return etree.QName(element).localname


def find(
    element: "etree._Element | None",
    path: str,
) -> "etree._Element | None":
    """Like :meth:`lxml.etree._Element.find`, but forgiving of a missing parent."""
    if element is None:
        return None
    return element.find(path, NS)


def findall(
    element: "etree._Element | None",
    path: str,
) -> "list[etree._Element]":
    """Like :meth:`lxml.etree._Element.findall`, but forgiving of a missing parent."""
    if element is None:
        return []
    return element.findall(path, NS)


def integer(
    element: "etree._Element | None",
    name: str,
    default: int = 0,
) -> int:
    """An integer attribute of an element, or ``default`` if it is not there."""
    if element is None:
        return default
    value = element.get(name)
    if value is None:
        return default
    return int(value)


def require(
    element: "etree._Element | None",
    path: str,
) -> etree._Element:
    """The element at ``path`` below ``element``, which the format promises is there."""
    found = find(element, path)
    if found is None:
        raise ValueError(f"the deck has no {path} where one is required")
    return found
