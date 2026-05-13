"""lxml custom element classes for theme-related XML elements."""

from __future__ import annotations

from pptx_ng.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrOne, ZeroOrMore

from . import parse_from_template


class CT_OfficeStyleSheet(BaseOxmlElement):
    """``<a:theme>`` element, root of a theme part."""

    _tag_seq = (
        "a:themeElements",
        "a:objectDefaults",
        "a:extraClrSchemeLst",
        "a:custClrLst",
        "a:extLst",
    )

    themeElements = ZeroOrOne("a:themeElements", successors=_tag_seq[1:])
    objectDefaults = ZeroOrOne("a:objectDefaults", successors=_tag_seq[2:])
    extraClrSchemeLst = ZeroOrOne("a:extraClrSchemeLst", successors=_tag_seq[3:])
    custClrLst = ZeroOrOne("a:custClrLst", successors=_tag_seq[4:])

    del _tag_seq

    @classmethod
    def new_default(cls):
        """Return a new ``<a:theme>`` element containing default settings."""
        return parse_from_template("theme")


class CT_BaseStyles(BaseOxmlElement):
    """``<a:themeElements>`` element containing color, font, format schemes."""

    _tag_seq = (
        "a:clrScheme",
        "a:fontScheme",
        "a:fmtScheme",
        "a:extLst",
    )

    clrScheme = ZeroOrOne("a:clrScheme", successors=_tag_seq[1:])
    fontScheme = ZeroOrOne("a:fontScheme", successors=_tag_seq[2:])
    fmtScheme = ZeroOrOne("a:fmtScheme", successors=_tag_seq[3:])

    del _tag_seq


class CT_ColorScheme(BaseOxmlElement):
    """``<a:clrScheme>`` element defining 12 theme colors."""

    name = OptionalAttribute("name", str)


class CT_FontScheme(BaseOxmlElement):
    """``<a:fontScheme>`` element defining major and minor fonts."""

    name = OptionalAttribute("name", str)


class CT_FormatScheme(BaseOxmlElement):
    """``<a:fmtScheme>`` element defining fill, line, effect, bg styles."""

    name = OptionalAttribute("name", str)
