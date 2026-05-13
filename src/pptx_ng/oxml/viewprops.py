"""View properties XML element classes.

Corresponds to CT_ViewProperties, CT_NormalViewProperties, CT_SlideViewProperties,
CT_GridSpacing in the Open XML SDK (schemas_openxmlformats_org_presentationml_2006_main).
"""

from __future__ import annotations

from pptx_ng.oxml.simpletypes import (
    ST_Coordinate,
    XsdBoolean,
    XsdString,
)
from pptx_ng.oxml.xmlchemy import (
    BaseOxmlElement,
    OxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)


class CT_ViewProperties(BaseOxmlElement):
    """`p:viewPr` element — root of view properties part.

    Attributes:
        lastView: Last view type (optional, String enum).
        showComments: Show comments (optional, Boolean).

    Child elements (in order):
        p:normalViewPr (0..1)
        p:slideViewPr (0..1)
        p:outlineViewPr (0..1)
        p:notesTextViewPr (0..1)
        p:sorterViewPr (0..1)
        p:notesViewPr (0..1)
        p:gridSpacing (0..1)
        p:extLst (0..1)
    """

    gridSpacing = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:gridSpacing",
        successors=("p:extLst",),
    )

    lastView = OptionalAttribute("lastView", XsdString)  # pyright: ignore[reportAssignmentType]
    showComments = OptionalAttribute("showComments", XsdBoolean)  # pyright: ignore[reportAssignmentType]

    @classmethod
    def new(cls) -> CT_ViewProperties:
        """Return a new `<p:viewPr>` element."""
        return OxmlElement("p:viewPr")


class CT_GridSpacing(BaseOxmlElement):
    """`p:gridSpacing` element — grid spacing for view properties.

    Inherits CT_PositiveSize2D with cx/cy attributes.
    Corresponds to PositiveSize2DType in the SDK.

    Attributes:
        cx: Grid spacing width in EMU (required, Int64 >= 0).
        cy: Grid spacing height in EMU (required, Int64 >= 0).
    """

    cx = RequiredAttribute("cx", ST_Coordinate)  # pyright: ignore[reportAssignmentType]
    cy = RequiredAttribute("cy", ST_Coordinate)  # pyright: ignore[reportAssignmentType]
