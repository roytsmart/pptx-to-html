"""Office Math (OMML), as PowerPoint stores equations, written out as MathML."""

import html
import unicodedata
from collections.abc import Callable

from lxml import etree

from ._xml import local

__all__ = [
    "MATH",
    "mathml",
    "text_of",
]

MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math"
"""The namespace of Office Math."""

_OPERATORS = set("+-=<>()[]{}|/*,;:!?.^~'′″") | {
    "−", "±", "∓", "×", "÷", "·", "∙", "≈", "≠", "≡", "≤", "≥", "≪", "≫", "∝",
    "∼", "≃", "≅", "→", "←", "↔", "⇒", "⇐", "⇔", "∂", "∇", "∑", "∏", "∫", "∮",
    "∞", "⋅", "∘", "∈", "∉", "⊂", "⊃", "⊆", "⊇", "∪", "∩", "∧", "∨", "¬", "∀",
    "∃", "⟨", "⟩", "‖", "…", "⋯",
}
"""Characters set as operators rather than as identifiers."""


def _m(name: str) -> str:
    return f"{{{MATH}}}{name}"


def _property(element: "etree._Element | None", path: str, default: str) -> str:
    """The ``m:val`` of a property like ``m:begChr`` inside an ``m:*Pr`` element."""
    if element is None:
        return default
    found = element.find(path.replace("m:", f"{{{MATH}}}"))
    if found is None:
        return default
    return found.get(_m("val"), default)


def mathml(element: etree._Element, warn: "Callable[[str], None]") -> str:
    """
    An ``a14:m`` element as MathML.

    Parameters
    ----------
    element
        The ``a14:m`` element holding an ``m:oMath`` or ``m:oMathPara``.
    warn
        Called with a message for each structure that cannot be converted.
    """
    converter = _Converter(warn)
    pieces = []
    for child in element:
        if child.tag == _m("oMathPara"):
            justify = _property(child.find(_m("oMathParaPr")), "m:jc", "centerGroup")
            align = {"left": "left", "right": "right"}.get(justify)
            style = f' style="text-align:{align}"' if align else ""
            for equation in child.findall(_m("oMath")):
                body = converter.row(equation)
                pieces.append(f'<math display="block"{style}>{body}</math>')
        elif child.tag == _m("oMath"):
            pieces.append(f"<math>{converter.row(child)}</math>")
    return "".join(pieces)


class _Converter:
    def __init__(self, warn: "Callable[[str], None]"):
        self.warn = warn
        self.warned: set[str] = set()

    def row(self, element: "etree._Element | None") -> str:
        """The children of an argument like ``m:e``, as one ``mrow``."""
        if element is None:
            return "<mrow></mrow>"
        inner = "".join(self.node(child) for child in element)
        return f"<mrow>{inner}</mrow>"

    def node(self, element: etree._Element) -> str:
        if not isinstance(element.tag, str) or not element.tag.startswith(f"{{{MATH}}}"):
            return ""
        kind = local(element)
        if kind.endswith("Pr"):
            return ""
        argument = element.find
        if kind == "r":
            return self.run(element)
        if kind in ("e", "num", "den", "sub", "sup", "deg", "lim", "fName", "oMath"):
            return self.row(element)
        if kind == "sSub":
            return f"<msub>{self.row(argument(_m('e')))}{self.row(argument(_m('sub')))}</msub>"
        if kind == "sSup":
            return f"<msup>{self.row(argument(_m('e')))}{self.row(argument(_m('sup')))}</msup>"
        if kind == "sSubSup":
            return (
                f"<msubsup>{self.row(argument(_m('e')))}{self.row(argument(_m('sub')))}"
                f"{self.row(argument(_m('sup')))}</msubsup>"
            )
        if kind == "sPre":
            return (
                f"<mmultiscripts>{self.row(argument(_m('e')))}<none/><none/><mprescripts/>"
                f"{self.row(argument(_m('sub')))}{self.row(argument(_m('sup')))}</mmultiscripts>"
            )
        if kind == "f":
            numerator = self.row(argument(_m("num")))
            denominator = self.row(argument(_m("den")))
            style = _property(argument(_m("fPr")), "m:type", "bar")
            if style in ("lin", "skw"):
                return f"<mrow>{numerator}<mo>/</mo>{denominator}</mrow>"
            thickness = ' linethickness="0"' if style == "noBar" else ""
            return f"<mfrac{thickness}>{numerator}{denominator}</mfrac>"
        if kind == "rad":
            hidden = _property(argument(_m("radPr")), "m:degHide", "0") in ("1", "on", "true")
            degree = argument(_m("deg"))
            if hidden or degree is None or len(degree) == 0:
                return f"<msqrt>{self.row(argument(_m('e')))}</msqrt>"
            return f"<mroot>{self.row(argument(_m('e')))}{self.row(degree)}</mroot>"
        if kind == "d":
            properties = argument(_m("dPr"))
            begin = _property(properties, "m:begChr", "(")
            end = _property(properties, "m:endChr", ")")
            separator = _property(properties, "m:sepChr", "|")
            arguments = [self.row(e) for e in element.findall(_m("e"))]
            # PowerPoint grows brackets only around what is taller than a
            # line of text. Browsers draw a stretchy bracket with room around
            # it even when it does not stretch, so the rest are kept still.
            stretch = "true" if _tall(element) else "false"
            fence = f'<mo fence="true" stretchy="{stretch}">'
            inner = f"<mo>{_text(separator)}</mo>".join(arguments)
            opening = f"{fence}{_text(begin)}</mo>" if begin else ""
            closing = f"{fence}{_text(end)}</mo>" if end else ""
            return f"<mrow>{opening}{inner}{closing}</mrow>"
        if kind == "nary":
            properties = argument(_m("naryPr"))
            operator = f'<mo largeop="true">{_text(_property(properties, "m:chr", "∫"))}</mo>'
            location = _property(properties, "m:limLoc", "subSup")
            hide_sub = _property(properties, "m:subHide", "0") in ("1", "on", "true")
            hide_sup = _property(properties, "m:supHide", "0") in ("1", "on", "true")
            lower = self.row(argument(_m("sub")))
            upper = self.row(argument(_m("sup")))
            if location == "undOvr":
                tags = ("munderover", "munder", "mover")
            else:
                tags = ("msubsup", "msub", "msup")
            if not hide_sub and not hide_sup:
                base = f"<{tags[0]}>{operator}{lower}{upper}</{tags[0]}>"
            elif not hide_sub:
                base = f"<{tags[1]}>{operator}{lower}</{tags[1]}>"
            elif not hide_sup:
                base = f"<{tags[2]}>{operator}{upper}</{tags[2]}>"
            else:
                base = operator
            return f"<mrow>{base}{self.row(argument(_m('e')))}</mrow>"
        if kind == "func":
            return f"<mrow>{self.row(argument(_m('fName')))}<mo>&#x2061;</mo>{self.row(argument(_m('e')))}</mrow>"
        if kind == "acc":
            accent = _property(argument(_m("accPr")), "m:chr", "̂")
            return f'<mover accent="true">{self.row(argument(_m("e")))}<mo>{_text(accent)}</mo></mover>'
        if kind == "bar":
            if _property(argument(_m("barPr")), "m:pos", "bot") == "top":
                return f'<mover accent="true">{self.row(argument(_m("e")))}<mo>‾</mo></mover>'
            return f'<munder accentunder="true">{self.row(argument(_m("e")))}<mo>_</mo></munder>'
        if kind == "groupChr":
            properties = argument(_m("groupChrPr"))
            character = f"<mo>{_text(_property(properties, 'm:chr', '⏟'))}</mo>"
            if _property(properties, "m:pos", "bot") == "top":
                return f"<mover>{self.row(argument(_m('e')))}{character}</mover>"
            return f"<munder>{self.row(argument(_m('e')))}{character}</munder>"
        if kind == "limLow":
            return f"<munder>{self.row(argument(_m('e')))}{self.row(argument(_m('lim')))}</munder>"
        if kind == "limUpp":
            return f"<mover>{self.row(argument(_m('e')))}{self.row(argument(_m('lim')))}</mover>"
        if kind == "eqArr":
            rows = "".join(f"<mtr><mtd>{self.row(e)}</mtd></mtr>" for e in element.findall(_m("e")))
            return f"<mtable>{rows}</mtable>"
        if kind == "m":
            rows = "".join(
                "<mtr>" + "".join(f"<mtd>{self.row(e)}</mtd>" for e in row.findall(_m("e"))) + "</mtr>"
                for row in element.findall(_m("mr"))
            )
            return f"<mtable>{rows}</mtable>"
        if kind == "box":
            return self.row(argument(_m("e")))
        if kind == "borderBox":
            return f'<menclose notation="box">{self.row(argument(_m("e")))}</menclose>'
        if kind == "phant":
            return f"<mphantom>{self.row(argument(_m('e')))}</mphantom>"
        if kind not in self.warned:
            self.warned.add(kind)
            self.warn(f"an equation uses m:{kind}, which is shown as plain text")
        text = "".join(t.text or "" for t in element.iter(_m("t")))
        return f"<mtext>{_text(text)}</mtext>"

    def run(self, element: etree._Element) -> str:
        """A run of math text, split into numbers, operators, and identifiers."""
        text = "".join(t.text or "" for t in element.findall(_m("t")))
        properties = element.find(_m("rPr"))
        normal = properties is not None and (
            properties.find(_m("nor")) is not None or _property(properties, "m:sty", "i") == "p"
        )
        if normal:
            # Text written upright inside an equation, like the name of a
            # quantity, stays one piece.
            tokens = []
            word = ""
            for character in text:
                if character in _OPERATORS and character not in "'.":
                    if word:
                        tokens.append(_upright(word))
                        word = ""
                    tokens.append(_operator(character))
                else:
                    word += character
            if word:
                tokens.append(_upright(word))
            return "".join(tokens)

        tokens = []
        number = ""
        for character in text:
            if character.isdigit() or (character == "." and number):
                number += character
                continue
            if number:
                tokens.append(f"<mn>{number}</mn>")
                number = ""
            if character.isspace():
                continue
            if character in _OPERATORS or unicodedata.category(character).startswith(("S", "P")):
                tokens.append(_operator(character))
            else:
                tokens.append(f"<mi>{_text(character)}</mi>")
        if number:
            tokens.append(f"<mn>{number}</mn>")
        return "".join(tokens)


_BRACKETS = set("()[]{}|⟨⟩‖⌈⌉⌊⌋")
"""Characters that browsers would otherwise stretch, with room around them."""

_TALL = {"f", "nary", "m", "eqArr", "limLow", "limUpp", "rad", "groupChr"}
"""Structures taller than a line of text, which brackets around them grow to fit."""


def _tall(element: etree._Element) -> bool:
    """Whether anything inside a pair of brackets is taller than a line of text."""
    return any(
        isinstance(child.tag, str) and child.tag.startswith(f"{{{MATH}}}") and local(child) in _TALL
        for argument in element.findall(_m("e"))
        for child in argument.iter()
    )


def _operator(character: str) -> str:
    """An operator typed as text: a bracket written this way never stretches."""
    if character in _BRACKETS:
        return f'<mo stretchy="false">{_text(character)}</mo>'
    return f"<mo>{_text(character)}</mo>"


def _upright(word: str) -> str:
    if word.strip() == "":
        return f"<mtext>{_text(word)}</mtext>"
    return f'<mi mathvariant="normal">{_text(word)}</mi>' if len(word) == 1 else f"<mi>{_text(word)}</mi>"


def _text(value: str) -> str:
    return html.escape(value, quote=False)


def text_of(element: etree._Element) -> str:
    """All the text in an element, including the text of any equations in it."""
    return "".join(
        t.text or ""
        for t in element.iter()
        if isinstance(t.tag, str) and local(t) == "t"
    )
