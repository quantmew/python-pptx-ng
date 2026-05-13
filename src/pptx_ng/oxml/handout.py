"""Custom element classes for handout master-related XML elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx_ng.oxml import parse_from_template
from pptx_ng.oxml.ns import nsdecls
from pptx_ng.oxml.simpletypes import XsdString
from pptx_ng.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from pptx_ng.oxml.shapes.groupshape import CT_GroupShape
    from pptx_ng.oxml.slide import CT_CommonSlideData


class _BaseSlideElement(BaseOxmlElement):
    """Base class for the six slide types, providing common methods."""

    cSld: CT_CommonSlideData

    @property
    def spTree(self) -> CT_GroupShape:
        """Return required `p:cSld/p:spTree` grandchild."""
        return self.cSld.spTree


class CT_HandoutMaster(_BaseSlideElement):
    """`p:handoutMaster` element, root of a handout master part.

    Schema: CT_HandoutMaster
    Children: p:cSld (required), p:clrMap (required), p:hf (optional), p:extLst (optional)
    """

    _tag_seq = ("p:cSld", "p:clrMap", "p:hf", "p:extLst")
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    del _tag_seq

    @classmethod
    def new_default(cls) -> CT_HandoutMaster:
        """Return a new `p:handoutMaster` element based on the built-in default template."""
        return cast(CT_HandoutMaster, parse_from_template("handoutMaster"))


class CT_HandoutMasterIdList(BaseOxmlElement):
    """`p:handoutMasterIdLst` element.

    Contains zero or more references to handout master parts.
    """

    handoutMasterId_lst: list[CT_HandoutMasterIdListEntry]

    handoutMasterId = ZeroOrMore("p:handoutMasterId")


class CT_HandoutMasterIdListEntry(BaseOxmlElement):
    """`p:handoutMasterId` element.

    Contains a reference (rId) to a handout master part.
    """

    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]
