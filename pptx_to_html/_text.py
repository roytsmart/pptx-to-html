"""Text, with the formatting PowerPoint lets it inherit from layouts and masters."""

import dataclasses
import html
from collections.abc import Callable

from lxml import etree

from ._fonts import font_stack
from ._math import mathml, text_of
from ._theme import Color, Theme
from ._xml import find, findall, integer, local, qn

__all__ = [
    "Inheritance",
    "Run",
    "TextRenderer",
]

_SINGLE_SPACING = 1.2
"""The height of a line set single-spaced, as a multiple of its font size."""

_DEFAULT_SIZE = 1800
"""The font size PowerPoint falls back on, in hundredths of a point."""

_WINGDINGS = {
    "§": "▪",
    "Ø": "➢",
    "ü": "✓",
    "q": "❑",
    "v": "❖",
    "n": "■",
    "l": "●",
}
"""What the bullets PowerPoint draws from Wingdings look like in Unicode."""

_BULLETS = ("buNone", "buChar", "buAutoNum", "buBlip")


@dataclasses.dataclass
class Inheritance:
    """
    Where a shape's text looks for the formatting it does not set itself.

    Each list runs from the most specific source to the least, so the first
    source to say something about a property is the one that counts.
    """

    own_styles: "list[etree._Element]"
    """The list styles that belong to the shape itself."""

    inherited_styles: "list[etree._Element]"
    """
    The list styles the shape inherits: those of the placeholders it is
    based on, then the master's text styles, then the presentation's.
    """

    bodies: "list[etree._Element]"
    """The ``a:bodyPr`` of the shape and of the placeholders it is based on."""

    font_color: "Color | None" = None
    """The text color the shape's style asks for, if it has one."""

    def body(self, name: str) -> "str | None":
        """An attribute of the text body, from wherever it is first set."""
        for body in self.bodies:
            value = body.get(name)
            if value is not None:
                return value
        return None


@dataclasses.dataclass(frozen=True)
class Run:
    """The formatting of a run of text, once everything it inherits is settled."""

    size: float
    """The font size, in points."""
    color: "Color | None"
    font: "str | None"
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    caps: bool = False
    baseline: int = 0
    """How far the run is raised (positive) or lowered, in thousandths of a percent."""


class TextRenderer:
    """Turns the text of a shape into the paragraphs of a web page."""

    def __init__(
        self,
        theme: Theme,
        slide_height: int,
        default_color: "Color | None",
        hyperlink: "Callable[[str], str | None]",
        warn: "Callable[[str], None] | None" = None,
    ):
        """
        Parameters
        ----------
        theme
            The theme of the slide the text is on.
        slide_height
            The height of the slide in EMU, which sizes are given as a share of.
        default_color
            The color the page gives text unless told otherwise.
        hyperlink
            Looks up the address a relationship id points to.
        warn
            Called with a message about anything that cannot be converted.
        """
        self.theme = theme
        self.slide_height = slide_height
        self.default_color = default_color
        self.hyperlink = hyperlink
        self.warn: Callable[[str], None] = warn or (lambda message: None)

    def cqh(self, points: float) -> float:
        """A length in points as a percentage of the slide's height."""
        return 100 * points / (self.slide_height / 12700)

    def paragraphs(
        self,
        body: etree._Element,
        inheritance: Inheritance,
    ) -> "list[str]":
        """
        The paragraphs of a text body as HTML, or an empty list if it holds
        no text at all.
        """
        paragraphs = findall(body, "a:p")
        if not any(_text_of(p).strip() for p in paragraphs):
            return []

        autofit = find(find(body, "a:bodyPr"), "a:normAutofit")
        scale = integer(autofit, "fontScale", 100000) / 100000
        reduction = integer(autofit, "lnSpcReduction", 0) / 100000

        result = []
        numbers: dict[int, int] = {}
        for index, paragraph in enumerate(paragraphs):
            result.append(
                self._paragraph(
                    paragraph,
                    inheritance,
                    scale=scale,
                    reduction=reduction,
                    first=index == 0,
                    numbers=numbers,
                )
            )
        return result

    def _paragraph(
        self,
        paragraph: etree._Element,
        inheritance: Inheritance,
        scale: float,
        reduction: float,
        first: bool,
        numbers: "dict[int, int]",
    ) -> str:
        properties = find(paragraph, "a:pPr")
        level = integer(properties, "lvl") + 1

        own = [properties] if properties is not None else []
        own += _levels(inheritance.own_styles, level)
        inherited = _levels(inheritance.inherited_styles, level)
        levels = own + inherited

        def attribute(name: str) -> "str | None":
            for element in levels:
                value = element.get(name)
                if value is not None:
                    return value
            return None

        def child(name: str) -> "etree._Element | None":
            for element in levels:
                found = find(element, name)
                if found is not None:
                    return found
            return None

        # The run formatting a paragraph inherits, most specific first. The
        # shape's style supplies a text color after anything the shape says
        # itself and before anything it inherits.
        defaults: list[etree._Element | Color] = [
            d for d in (find(e, "a:defRPr") for e in own) if d is not None
        ]
        if inheritance.font_color is not None:
            defaults.append(inheritance.font_color)
        defaults += [d for d in (find(e, "a:defRPr") for e in inherited) if d is not None]

        pieces: list[tuple[Run | None, str | None, etree._Element]] = []
        runs = []
        for element in paragraph:
            kind = local(element)
            if kind in ("r", "fld"):
                run = self._run(find(element, "a:rPr"), defaults)
                text = find(element, "a:t")
                runs.append(run)
                pieces.append((run, text.text if text is not None and text.text else "", element))
            elif kind == "br":
                run = self._run(find(element, "a:rPr"), defaults)
                pieces.append((run, None, element))
            elif kind == "m":
                # An equation, which comes through already written as MathML.
                pieces.append((None, mathml(element, self.warn), element))

        base = runs[0] if runs else self._run(find(paragraph, "a:endParaRPr"), defaults)

        styles = [f"font-size:{self.cqh(base.size * scale):.4f}cqh"]

        spacing = child("a:lnSpc")
        percent = find(spacing, "a:spcPct")
        points = find(spacing, "a:spcPts")
        if points is not None:
            styles.append(f"line-height:{self.cqh(integer(points, 'val') / 100):.4f}cqh")
        else:
            fraction = integer(percent, "val", 100000) / 100000
            styles.append(f"line-height:{_SINGLE_SPACING * (fraction - reduction):.3f}")

        align = {"ctr": "center", "r": "right", "just": "justify", "dist": "justify"}
        alignment = align.get(attribute("algn") or "l")
        if alignment is not None:
            styles.append(f"text-align:{alignment}")

        styles += self._run_styles(base, None, scale)

        if not first:
            styles += self._spacing(child("a:spcBef"), "margin-top", base.size * scale)
        styles += self._spacing(child("a:spcAft"), "margin-bottom", base.size * scale)

        margin = int(attribute("marL") or 0)
        indent = int(attribute("indent") or 0)
        classes = []
        bullet = self._bullet(levels, level, numbers)
        if margin:
            styles.append(f"padding-left:{self.cqh(margin / 12700):.4f}cqh")
        if bullet is not None and _text_of(paragraph).strip():
            classes.append("bullet")
            styles.append(f"--indent:{self.cqh(indent / 12700):.4f}cqh")
            styles.append(f"--bullet:'{_css_string(bullet)}'")
        elif indent:
            styles.append(f"text-indent:{self.cqh(indent / 12700):.4f}cqh")

        content = "".join(self._piece(run, text, element, base, scale) for run, text, element in pieces)
        if not _text_of(paragraph):
            content = "<br>"

        attributes = f' class="{" ".join(classes)}"' if classes else ""
        style = html.escape(";".join(styles))
        return f'<p{attributes} style="{style}">{content}</p>'

    def _piece(
        self,
        run: "Run | None",
        text: "str | None",
        element: etree._Element,
        base: Run,
        scale: float,
    ) -> str:
        if run is None:
            return text or ""
        if text is None:
            return "<br>"
        content = html.escape(text, quote=False)
        styles = self._run_styles(run, base, scale)
        if styles:
            content = f'<span style="{html.escape(";".join(styles))}">{content}</span>'
        link = find(find(element, "a:rPr"), "a:hlinkClick")
        if link is not None:
            address = self.hyperlink(link.get(qn("r:id"), ""))
            if address:
                content = f'<a href="{html.escape(address)}">{content}</a>'
        return content

    def _run(
        self,
        properties: "etree._Element | None",
        defaults: "list[etree._Element | Color]",
    ) -> Run:
        sources: list[etree._Element | Color] = [properties] if properties is not None else []
        sources += defaults
        elements = [s for s in sources if not isinstance(s, Color)]

        def attribute(name: str) -> "str | None":
            for element in elements:
                value = element.get(name)
                if value is not None:
                    return value
            return None

        color = None
        for source in sources:
            if isinstance(source, Color):
                color = source
                break
            fill = find(source, "a:solidFill")
            if fill is not None:
                color = self.theme.color(fill)
                break

        font = None
        for element in elements:
            latin = find(element, "a:latin")
            if latin is not None and latin.get("typeface"):
                font = self.theme.typeface(latin.get("typeface", ""))
                break

        def flag(name: str) -> bool:
            return attribute(name) in ("1", "true")

        return Run(
            size=int(attribute("sz") or _DEFAULT_SIZE) / 100,
            color=color,
            font=font,
            bold=flag("b"),
            italic=flag("i"),
            underline=(attribute("u") or "none") != "none",
            strike=(attribute("strike") or "noStrike") != "noStrike",
            caps=attribute("cap") == "all",
            baseline=int(attribute("baseline") or 0),
        )

    def _run_styles(self, run: Run, base: "Run | None", scale: float) -> "list[str]":
        """
        The CSS a run needs: everything about it, for the run a paragraph is
        styled after, or only how it differs, for the runs inside it.
        """
        styles = []
        if base is not None and run.size != base.size:
            styles.append(f"font-size:{self.cqh(run.size * scale):.4f}cqh")
        reference_color = base.color if base is not None else self.default_color
        if run.color is not None and run.color != reference_color:
            styles.append(f"color:{run.color.css()}")
        reference_font = base.font if base is not None else self.theme.font_minor
        if run.font is not None and run.font != reference_font:
            stack = font_stack(run.font, quote="'")
            styles.append(f"font-family:{stack}")
        if run.bold != (base.bold if base is not None else False):
            styles.append(f"font-weight:{'bold' if run.bold else 'normal'}")
        if run.italic != (base.italic if base is not None else False):
            styles.append(f"font-style:{'italic' if run.italic else 'normal'}")
        decorations = [n for n, on in (("underline", run.underline), ("line-through", run.strike)) if on]
        base_decorations = [] if base is None else [
            n for n, on in (("underline", base.underline), ("line-through", base.strike)) if on
        ]
        if decorations != base_decorations:
            styles.append(f"text-decoration:{' '.join(decorations) or 'none'}")
        if run.caps != (base.caps if base is not None else False):
            styles.append(f"text-transform:{'uppercase' if run.caps else 'none'}")
        if run.baseline and (base is None or run.baseline != base.baseline):
            styles.append(f"vertical-align:{'super' if run.baseline > 0 else 'sub'};font-size:70%")
        return styles

    def _spacing(
        self,
        spacing: "etree._Element | None",
        name: str,
        size: float,
    ) -> "list[str]":
        points = find(spacing, "a:spcPts")
        if points is not None:
            value = integer(points, "val") / 100
            return [f"{name}:{self.cqh(value):.4f}cqh"] if value else []
        percent = find(spacing, "a:spcPct")
        if percent is not None:
            value = integer(percent, "val") / 100000 * _SINGLE_SPACING * size
            return [f"{name}:{self.cqh(value):.4f}cqh"] if value else []
        return []

    def _bullet(
        self,
        levels: "list[etree._Element]",
        level: int,
        numbers: "dict[int, int]",
    ) -> "str | None":
        """The bullet a paragraph starts with, counting numbered ones as it goes."""
        chosen = None
        for element in levels:
            for kind in _BULLETS:
                chosen = find(element, f"a:{kind}")
                if chosen is not None:
                    break
            if chosen is not None:
                break
        # A paragraph at a shallower level ends any numbering deeper than it.
        for deeper in [k for k in numbers if k > level]:
            del numbers[deeper]
        if chosen is None or local(chosen) in ("buNone", "buBlip"):
            numbers.pop(level, None)
            return None
        if local(chosen) == "buChar":
            character = chosen.get("char", "•")
            typeface = ""
            for element in levels:
                font = find(element, "a:buFont")
                if font is not None:
                    typeface = font.get("typeface", "")
                    break
            if typeface.startswith("Wingdings"):
                return _WINGDINGS.get(character, "•")
            return character
        start = integer(chosen, "startAt", 1)
        numbers[level] = numbers.get(level, start - 1) + 1
        return _number(numbers[level], chosen.get("type", "arabicPeriod"))


def _levels(styles: "list[etree._Element]", level: int) -> "list[etree._Element]":
    found = (find(style, f"a:lvl{level}pPr") for style in styles)
    return [element for element in found if element is not None]


def _text_of(paragraph: etree._Element) -> str:
    return text_of(paragraph)


def _css_string(text: str) -> str:
    return text.replace("\\", "\\\\").replace("'", "\\'")


def _number(value: int, kind: str) -> str:
    """Write a list item's number the way an ``a:buAutoNum`` type asks."""
    for numeral in ("arabic", "alphaLc", "alphaUc", "romanLc", "romanUc"):
        if kind.startswith(numeral):
            style = kind[len(numeral):]
            break
    else:
        numeral, style = "arabic", "Period"
    if numeral == "arabic":
        text = str(value)
    elif numeral.startswith("alpha"):
        text = ""
        n = value
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            text = chr(ord("a") + remainder) + text
    else:
        text = _roman(value)
    if numeral.endswith("Uc"):
        text = text.upper()
    if style == "ParenBoth":
        return f"({text})"
    if style == "ParenR":
        return f"{text})"
    if style == "Plain":
        return text
    return f"{text}."


def _roman(value: int) -> str:
    numerals = [
        (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
        (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"),
    ]
    text = ""
    for amount, numeral in numerals:
        while value >= amount:
            text += numeral
            value -= amount
    return text
