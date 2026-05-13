"""Presentation properties XML element classes.

Corresponds to CT_PresentationProperties, CT_ShowProperties, CT_PrintProperties
in the Open XML SDK (schemas_openxmlformats_org_presentationml_2006_main).
"""

from __future__ import annotations

from pptx.oxml.simpletypes import (
    ST_Coordinate,
    XsdBoolean,
    XsdString,
    XsdUnsignedInt,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OxmlElement,
    OptionalAttribute,
    ZeroOrOne,
)


class CT_PresentationProperties(BaseOxmlElement):
    """`p:presentationPr` element — root of presentation properties part.

    Child elements (in order):
        p:htmlPubPr (0..1) — HTML publishing properties
        p:webPr (0..1) — web properties
        p:prnPr (0..1) — printing properties
        p:showPr (0..1) — show properties
        p:clrMru (0..1) — color most recently used
        p:extLst (0..1) — extension list
    """

    showPr = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:showPr",
        successors=(
            "p:clrMru",
            "p:extLst",
        ),
    )
    prnPr = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:prnPr",
        successors=(
            "p:showPr",
            "p:clrMru",
            "p:extLst",
        ),
    )

    @classmethod
    def new(cls) -> CT_PresentationProperties:
        """Return a new `<p:presentationPr>` element."""
        return OxmlElement("p:presentationPr")


class CT_ShowProperties(BaseOxmlElement):
    """`p:showPr` element — presentation show properties.

    Attributes:
        loop: Loop slide show (optional, Boolean).
        showNarration: Show narration in slide show (optional, Boolean).
        showAnimation: Show animation in slide show (optional, Boolean).
        useTimings: Use timings in slide show (optional, Boolean).

    Child elements (in order):
        choice group: p:present | p:browse | p:kiosk (0..1)
        choice group: p:sldAll | p:sldRg | p:custShow (0..1)
        p:penClr (0..1)
        p:extLst (0..1)
    """

    loop = OptionalAttribute("loop", XsdBoolean)  # pyright: ignore[reportAssignmentType]
    showNarration = OptionalAttribute("showNarration", XsdBoolean)  # pyright: ignore[reportAssignmentType]
    showAnimation = OptionalAttribute("showAnimation", XsdBoolean)  # pyright: ignore[reportAssignmentType]
    useTimings = OptionalAttribute("useTimings", XsdBoolean)  # pyright: ignore[reportAssignmentType]


class CT_PrintProperties(BaseOxmlElement):
    """`p:prnPr` element — printing properties.

    Attributes:
        prnWhat: Print output type (optional, String enum).
        clrMode: Print color mode (optional, String enum).
        hiddenSlides: Print hidden slides (optional, Boolean).
        scaleToFitPaper: Scale to fit paper (optional, Boolean).
        frameSlides: Frame slides when printing (optional, Boolean).

    Child elements:
        p:extLst (0..1)
    """

    prnWhat = OptionalAttribute("prnWhat", XsdString)  # pyright: ignore[reportAssignmentType]
    clrMode = OptionalAttribute("clrMode", XsdString)  # pyright: ignore[reportAssignmentType]
    hiddenSlides = OptionalAttribute("hiddenSlides", XsdBoolean)  # pyright: ignore[reportAssignmentType]
    scaleToFitPaper = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "scaleToFitPaper", XsdBoolean
    )
    frameSlides = OptionalAttribute("frameSlides", XsdBoolean)  # pyright: ignore[reportAssignmentType]
