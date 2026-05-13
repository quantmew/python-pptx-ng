"""Custom show XML element classes.

Corresponds to CT_CustomShowList, CT_CustomShow, CT_SlideRelationshipList,
CT_SlideRelationshipListEntry in the Open XML SDK
(schemas_openxmlformats_org_presentationml_2006_main).
"""

from __future__ import annotations

from typing import Callable

from lxml.etree import _Element

from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OxmlElement,
    OneAndOnlyOne,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_CustomShowList(BaseOxmlElement):
    """`p:custShowLst` element — list of custom shows.

    Contains zero or more `p:custShow` child elements.
    """

    custShow = ZeroOrMore("p:custShow", successors=())

    @classmethod
    def new(cls) -> CT_CustomShowList:
        """Return a new `<p:custShowLst>` element."""
        return OxmlElement("p:custShowLst")

    def add_custShow(self, id: int, name: str) -> CT_CustomShow:
        """Add a new `p:custShow` child and return it."""
        custShow = self._add_custShow()
        custShow.id = id
        custShow.name = name
        # Add required sldLst child
        sldLst = OxmlElement("p:sldLst")
        custShow.append(sldLst)
        return custShow


class CT_CustomShow(BaseOxmlElement):
    """`p:custShow` element — a single custom show definition.

    Attributes:
        name: Custom show name (required, String).
        id: Custom show ID (required, UInt32).

    Child elements (in order):
        p:sldLst — slide list (required, OneAndOnlyOne).
        p:extLst — extension list (optional).
    """

    _tag_seq = ("p:sldLst",)
    sldLst = OneAndOnlyOne("p:sldLst")  # pyright: ignore[reportAssignmentType]
    del _tag_seq

    name = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
    id = RequiredAttribute("id", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


class CT_SlideRelationshipList(BaseOxmlElement):
    """`p:sldLst` element (inside `p:custShow`) — list of slide references.

    Contains zero or more `p:sld` child elements.

    Note: `p:sld` entries are created manually via lxml because the `p:sld`
    tag is already registered to CT_Slide (the main slide element class).
    We use raw set/get to access the r:id attribute.
    """

    @classmethod
    def new(cls) -> CT_SlideRelationshipList:
        """Return a new `<p:sldLst>` element."""
        return OxmlElement("p:sldLst")

    def add_sld(self, rId: str) -> _Element:
        """Add a new `p:sld` child referencing slide by rId and return it."""
        sld = OxmlElement("p:sld")
        sld.set(qn("r:id"), rId)
        self.append(sld)
        return sld

    @property
    def sld_lst(self) -> list[_Element]:
        """Return list of `p:sld` child elements."""
        return self.findall(qn("p:sld"))


class CT_SlideRelationshipListEntry(BaseOxmlElement):
    """`p:sld` element (inside `p:custShow/p:sldLst`) — reference to a slide.

    Note: This class cannot be registered for the `p:sld` tag because that
    tag is already mapped to CT_Slide. It is provided as a type reference only.
    Access r:id via lxml's get/set methods directly.

    Attributes:
        rId: Relationship ID referencing the slide (required, String).
    """

    rId = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]
