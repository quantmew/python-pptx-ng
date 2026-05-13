"""Theme proxy classes for accessing theme color, font, and format schemes."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.theme import CT_ColorScheme, CT_FontScheme, CT_OfficeStyleSheet


class Theme:
    """Proxy for the ``a:theme`` element in a theme part.

    Provides access to color scheme, font scheme, and format scheme.
    """

    def __init__(self, theme_element: CT_OfficeStyleSheet):
        super().__init__()
        self._element = theme_element

    @lazyproperty
    def color_scheme(self) -> ColorScheme:
        """Return a |ColorScheme| proxy for this theme."""
        clrScheme = self._element.themeElements.clrScheme
        return ColorScheme(clrScheme)

    @lazyproperty
    def font_scheme(self) -> FontScheme:
        """Return a |FontScheme| proxy for this theme."""
        fontScheme = self._element.themeElements.fontScheme
        return FontScheme(fontScheme)


class ColorScheme:
    """Proxy for an ``a:clrScheme`` element.

    Provides access to the 12 theme colors (dk1, lt1, dk2, lt2, accent1-6,
    hlink, folHlink).
    """

    _color_names = (
        "dk1",
        "lt1",
        "dk2",
        "lt2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hlink",
        "folHlink",
    )

    def __init__(self, clrScheme: CT_ColorScheme):
        super().__init__()
        self._element = clrScheme

    @property
    def name(self) -> str | None:
        """Return the name of this color scheme."""
        return self._element.get("name")

    def get_color(self, color_name: str) -> str | None:
        """Return the RGB value for a named theme color, if available.

        Returns the ``val`` attribute of the ``srgbClr`` child if present,
        otherwise returns |None|.
        """
        color_el = self._element.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}%s" % color_name
        )
        if color_el is None:
            return None
        srgbClr = color_el.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr"
        )
        if srgbClr is None:
            return None
        return srgbClr.get("val")


class FontScheme:
    """Proxy for an ``a:fontScheme`` element.

    Provides access to the major and minor font definitions.
    """

    def __init__(self, fontScheme: CT_FontScheme):
        super().__init__()
        self._element = fontScheme

    @property
    def name(self) -> str | None:
        """Return the name of this font scheme."""
        return self._element.get("name")

    @property
    def major_font(self) -> str | None:
        """Return the Latin typeface name of the major font."""
        return self._get_latin_typeface("majorFont")

    @property
    def minor_font(self) -> str | None:
        """Return the Latin typeface name of the minor font."""
        return self._get_latin_typeface("minorFont")

    def _get_latin_typeface(self, tag: str) -> str | None:
        el = self._element.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}%s" % tag
        )
        if el is None:
            return None
        latin = el.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}latin"
        )
        if latin is None:
            return None
        return latin.get("typeface")
