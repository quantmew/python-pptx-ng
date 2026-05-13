"""Custom element classes for table style XML elements.

The table styles part uses the drawingml namespace (``a:``) and contains a list
of table style definitions. Each table style defines formatting for different
table regions (whole table, banded rows, header row, etc.).
"""

from __future__ import annotations

from pptx.oxml.ns import nsdecls
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrMore


class CT_TableStyleList(BaseOxmlElement):
    """``a:tblStyleLst`` — Root element of the table styles part."""

    _tag_seq = ("a:tblStyle",)

    tblStyle = ZeroOrMore("a:tblStyle", successors=())

    def get_or_add_tblStyle(self):
        """Return the first ``a:tblStyle`` child, adding one if not present."""
        tblStyle = self.tblStyle
        if tblStyle is not None:
            return tblStyle
        return self._add_tblStyle()

    @classmethod
    def new(cls, default_style_id="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"):
        """Return a new ``CT_TableStyleList`` with a default style."""
        xml = (
            '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' def="%s"/>'
        ) % default_style_id
        from pptx.oxml import parse_xml

        return parse_xml(xml)


class CT_TableStyleEntry(BaseOxmlElement):
    """``a:tblStyle`` — A single table style definition."""

    styleName = OptionalAttribute("styleName", str)
    styleId = OptionalAttribute("styleId", str)
