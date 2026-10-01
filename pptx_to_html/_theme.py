"""Colors and typefaces, as a deck's theme defines them."""

import colorsys
import dataclasses

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.slide import SlideMaster

from ._xml import find, findall, integer, local, require

__all__ = [
    "Color",
    "Theme",
    "fill_color",
    "line_style",
]

_PRESET_COLORS = {
    "black": "000000",
    "white": "FFFFFF",
    "red": "FF0000",
    "green": "008000",
    "blue": "0000FF",
    "yellow": "FFFF00",
    "gray": "808080",
}
"""The few preset colors decks actually use; anything else falls back to black."""


@dataclasses.dataclass(frozen=True)
class Color:
    """A color with an opacity, each channel between 0 and 1."""

    red: float
    green: float
    blue: float
    alpha: float = 1

    @classmethod
    def from_hex(cls, value: str) -> "Color":
        """A color from six hexadecimal digits, like ``"1F4E79"``."""
        return cls(*(int(value[i : i + 2], 16) / 255 for i in (0, 2, 4)))

    def css(self) -> str:
        """The color as CSS: a hex code when opaque, ``rgba()`` when not."""
        channels = [round(255 * c) for c in (self.red, self.green, self.blue)]
        if self.alpha >= 1:
            return "#" + "".join(f"{c:02X}" for c in channels)
        return f"rgba({channels[0]},{channels[1]},{channels[2]},{self.alpha:.3f})"


class Theme:
    """The color scheme, fonts, and line widths a slide master's theme defines."""

    def __init__(self, master: SlideMaster):
        root = etree.fromstring(master.part.part_related_by(RT.THEME).blob)

        self.colors: dict[str, str] = {}
        """The theme's colors as hex codes, by name (``dk1``, ``accent1``, ...)."""
        for slot in findall(root, "a:themeElements/a:clrScheme/*"):
            choice = slot[0]
            self.colors[local(slot)] = choice.get("lastClr") or choice.get("val", "000000")

        fonts = require(root, "a:themeElements/a:fontScheme")
        self.font_major: str = require(fonts, "a:majorFont/a:latin").get("typeface", "")
        """The typeface the theme gives headings."""
        self.font_minor: str = require(fonts, "a:minorFont/a:latin").get("typeface", "")
        """The typeface the theme gives body text."""

        self.line_widths: list[int] = [
            integer(line, "w", 9525)
            for line in findall(root, "a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln")
        ]
        """The widths of the theme's three line styles, in EMU."""

        clr_map = require(master.element, "p:clrMap")
        self.color_map: dict[str, str] = {str(k): str(v) for k, v in clr_map.attrib.items()}
        """How the names text uses (``tx1``, ``bg1``, ...) map onto the scheme."""

    def typeface(self, name: str) -> str:
        """Resolve a theme font reference like ``+mn-lt`` into a typeface."""
        if name.startswith("+mj"):
            return self.font_major
        if name.startswith("+mn"):
            return self.font_minor
        return name

    def color(
        self,
        parent: "etree._Element | None",
        placeholder: "Color | None" = None,
    ) -> "Color | None":
        """
        The color chosen inside ``parent``, such as an ``a:solidFill``.

        Parameters
        ----------
        parent
            An element whose first child is a color choice, like ``a:srgbClr``
            or ``a:schemeClr``.
        placeholder
            What ``phClr`` stands for, when the color comes from a style
            matrix.
        """
        if parent is None or len(parent) == 0:
            return None
        choice = parent[0]
        kind = local(choice)
        if kind == "srgbClr":
            color = Color.from_hex(choice.get("val", "000000"))
        elif kind == "sysClr":
            color = Color.from_hex(choice.get("lastClr", "000000"))
        elif kind == "prstClr":
            color = Color.from_hex(_PRESET_COLORS.get(choice.get("val", ""), "000000"))
        elif kind == "schemeClr":
            name = choice.get("val", "")
            if name == "phClr":
                if placeholder is None:
                    return None
                color = placeholder
            else:
                name = self.color_map.get(name, name)
                color = Color.from_hex(self.colors.get(name, "000000"))
        else:
            return None
        return _modify(color, choice)


def _modify(color: Color, choice: etree._Element) -> Color:
    """Apply the modifiers PowerPoint nests inside a color choice."""
    red, green, blue, alpha = color.red, color.green, color.blue, color.alpha
    for modifier in choice:
        kind = local(modifier)
        value = integer(modifier, "val", 100000) / 100000
        if kind == "alpha":
            alpha = value
        elif kind in ("lumMod", "lumOff"):
            hue, lightness, saturation = colorsys.rgb_to_hls(red, green, blue)
            if kind == "lumMod":
                lightness *= value
            else:
                lightness += value
            lightness = min(max(lightness, 0), 1)
            red, green, blue = colorsys.hls_to_rgb(hue, lightness, saturation)
        elif kind == "shade":
            red, green, blue = red * value, green * value, blue * value
        elif kind == "tint":
            red, green, blue = (1 - (1 - c) * value for c in (red, green, blue))
    return Color(red, green, blue, alpha)


def fill_color(
    theme: Theme,
    properties: "etree._Element | None",
    style: "etree._Element | None",
) -> "Color | None":
    """
    The solid color a shape is filled with, or `None` if it is not filled.

    A shape which says nothing about its fill takes the one its style points
    at, which in every theme in use is a solid fill of the style's color.
    """
    for child in properties if properties is not None else []:
        kind = local(child)
        if kind == "noFill":
            return None
        if kind == "solidFill":
            return theme.color(child)
        if kind == "gradFill":
            # The first stop stands in for the whole gradient.
            return theme.color(find(child, "a:gsLst/a:gs"))
        if kind in ("blipFill", "pattFill", "grpFill"):
            return None
    reference = find(style, "a:fillRef")
    if integer(reference, "idx") == 0:
        return None
    return theme.color(reference)


def line_style(
    theme: Theme,
    properties: "etree._Element | None",
    style: "etree._Element | None",
) -> "tuple[int, Color] | None":
    """The width in EMU and the color of a shape's outline, if it has one."""
    line = find(properties, "a:ln")
    reference = find(style, "a:lnRef")
    width = None
    index = integer(reference, "idx") - 1
    if 0 <= index < len(theme.line_widths):
        width = theme.line_widths[index]
    if line is not None:
        if find(line, "a:noFill") is not None:
            return None
        width = integer(line, "w", width if width is not None else 9525)
        color = theme.color(find(line, "a:solidFill"))
        if color is not None:
            return width, color
    if width is None:
        return None
    color = theme.color(reference)
    if color is None:
        return None
    return width, color

