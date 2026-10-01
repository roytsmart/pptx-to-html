"""
Turn a PowerPoint deck into a single web page.

Every slide keeps its layout, with its pictures and text placed where
PowerPoint puts them, and its movies play in place instead of being frozen
on their first frame.
"""

from ._convert import ConversionWarning, Reencode, convert

__all__ = [
    "ConversionWarning",
    "Reencode",
    "convert",
]
