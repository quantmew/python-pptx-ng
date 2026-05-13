"""Custom element classes for user-defined tags XML elements."""

from __future__ import annotations

from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    RequiredAttribute,
    ZeroOrMore,
)


class CT_TagList(BaseOxmlElement):
    """`p:tagLst` element, root of a UserDefinedTagsPart.

    Contains zero or more `p:tag` child elements.
    """

    tag_lst: list[CT_Tag]

    tag = ZeroOrMore("p:tag")

    @classmethod
    def new(cls) -> CT_TagList:
        """Return a new empty `p:tagLst` element."""
        from pptx.oxml import parse_xml

        return parse_xml("<p:tagLst %s/>" % nsdecls("p"))

    def add_tag(self, name: str, val: str) -> CT_Tag:
        """Add a new `p:tag` child element and return it."""
        tag = self._new_tag()
        tag.name = name
        tag.val = val
        self.append(tag)
        return tag

    def get_tag(self, name: str) -> CT_Tag | None:
        """Return the `p:tag` child element with matching *name*, or None."""
        for tag in self.tag_lst:
            if tag.name == name:
                return tag
        return None


class CT_Tag(BaseOxmlElement):
    """`p:tag` element representing a single user-defined name/value pair.

    Both `name` and `val` attributes are required.
    """

    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
    val: str = RequiredAttribute("val", XsdString)  # pyright: ignore[reportAssignmentType]
