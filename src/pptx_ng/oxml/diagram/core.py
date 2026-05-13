"""Core diagram (SmartArt) XML element classes.

Provides the minimum element classes needed for SmartArt round-trip support:
creating, reading, and saving SmartArt graphics in presentations.
"""

from __future__ import annotations

from typing import Callable, cast

from pptx_ng.oxml.ns import nsdecls, qn
from pptx_ng.oxml.simpletypes import XsdString
from pptx_ng.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


def _parse(xml: str) -> BaseOxmlElement:
    from pptx_ng.oxml import parse_xml

    return parse_xml(xml)


# ---------------------------------------------------------------------------
# DiagramData root element (dgm:dataModel)
# ---------------------------------------------------------------------------


class CT_DataModel(BaseOxmlElement):
    """`dgm:dataModel` — root element of diagram data part.

    Contains the point list and connection list that define the
    structure of a SmartArt graphic.
    """

    _tag_seq = ("dgm:cxnLst", "dgm:ptLst")
    cxnLst = ZeroOrOne("dgm:cxnLst", successors=("dgm:ptLst",))
    ptLst = ZeroOrOne("dgm:ptLst", successors=())
    del _tag_seq

    get_or_add_cxnLst: Callable[[], CT_ConnectionList]
    get_or_add_ptLst: Callable[[], CT_PointList]

    @classmethod
    def new(cls) -> CT_DataModel:
        return cast(
            CT_DataModel,
            _parse(
                "<dgm:dataModel %s>"
                "<dgm:cxnLst/>"
                "<dgm:ptLst>"
                '<dgm:pt modelId="0" type="doc">'
                "<dgm:prSet loCatId=\"\" qCatIdId=\"\" csCatId=\"\"/>"
                "<dgm:spPr/>"
                "<dgm:t/>"
                "</dgm:pt>"
                "</dgm:ptLst>"
                "</dgm:dataModel>" % nsdecls("dgm")
            ),
        )


class CT_PointList(BaseOxmlElement):
    """`dgm:ptLst` — list of points (nodes) in a diagram."""

    pt = ZeroOrMore("dgm:pt")

    def add_point(self, model_id: int, text: str = "") -> CT_Point:
        pt = cast(
            CT_Point,
            _parse(
                '<dgm:pt %s modelId="%d" type="node">'
                "<dgm:prSet/>"
                "<dgm:spPr/>"
                "<dgm:t>"
                "<a:bodyPr/><a:lstStyle/><a:p>"
                "<a:r><a:t>%s</a:t></a:r>"
                "</a:p>"
                "</dgm:t>"
                "</dgm:pt>" % (nsdecls("dgm", "a"), model_id, text)
            ),
        )
        self.append(pt)
        return pt


class CT_Point(BaseOxmlElement):
    """`dgm:pt` — a single point (node) in a diagram."""

    modelId = OptionalAttribute("modelId", XsdString)
    type = OptionalAttribute("type", XsdString)


class CT_ConnectionList(BaseOxmlElement):
    """`dgm:cxnLst` — list of connections between points."""

    cxn = ZeroOrMore("dgm:cxn")

    def add_connection(self, model_id: int, src_id: int, dest_id: int) -> CT_Connection:
        cxn = cast(
            CT_Connection,
            _parse(
                '<dgm:cxn %s modelId="%d" srcId="%d" destId="%d" type="parOf"/>'
                % (nsdecls("dgm"), model_id, src_id, dest_id)
            ),
        )
        self.append(cxn)
        return cxn


class CT_Connection(BaseOxmlElement):
    """`dgm:cxn` — a connection between two points."""

    modelId = OptionalAttribute("modelId", XsdString)
    srcId = OptionalAttribute("srcId", XsdString)
    destId = OptionalAttribute("destId", XsdString)
    type = OptionalAttribute("type", XsdString)


# ---------------------------------------------------------------------------
# DiagramLayout root element (dgm:layoutDef)
# ---------------------------------------------------------------------------

_LAYOUT_TEMPLATE = (
    '<dgm:layoutDef %(ns)s'
    ' uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/orgChart1"'
    ' minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
    '<dgm:title val="Organization Chart"/>'
    "<dgm:desc val=\"\"/>"
    '<dgm:catLst><dgm:cat type="hier" pri="1000"/></dgm:catLst>'
    '<dgm:layoutNode name="modelRoot" styleLbl="node0">'
    '<dgm:alg type="sp"><dgm:param type="horzAlign" val="ctr"/></dgm:alg>'
    "<dgm:constrLst/>"
    '<dgm:layoutNode name="node0" styleLbl="node0" moveWith="modelRoot">'
    '<dgm:alg type="tx"/>'
    '<dgm:shape type="rect"><dgm:adjstLst/></dgm:shape>'
    "<dgm:constrLst/>"
    '<dgm:presOf axis="desOrSelf" ptType="node"/>'
    "</dgm:layoutNode>"
    "</dgm:layoutNode>"
    "</dgm:layoutDef>"
)


class CT_LayoutDefinition(BaseOxmlElement):
    """`dgm:layoutDef` — root element of diagram layout part."""

    uniqueId = OptionalAttribute("uniqueId", XsdString)

    @classmethod
    def new(cls, layout_id: str = "urn:microsoft.com/office/officeart/2005/8/layout/orgChart1") -> CT_LayoutDefinition:
        return cast(
            CT_LayoutDefinition,
            _parse(_LAYOUT_TEMPLATE % {"ns": nsdecls("dgm")}),
        )


# ---------------------------------------------------------------------------
# DiagramColors root element (dgm:colorsDef)
# ---------------------------------------------------------------------------

_COLORS_TEMPLATE = (
    '<dgm:colorsDef %(ns)s'
    ' uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_2"'
    ' minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
    '<dgm:title val=""/>'
    "<dgm:desc val=\"\"/>"
    '<dgm:catLst><dgm:cat type="accent1" pri="1000"/></dgm:catLst>'
    '<dgm:styleLbl name="node0">'
    "<dgm:fill><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill></dgm:fill>"
    "<dgm:ln><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill></dgm:ln>"
    "<dgm:effectLst/>"
    "</dgm:styleLbl>"
    '<dgm:styleLbl name="node1">'
    "<dgm:fill><a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill></dgm:fill>"
    "<dgm:ln><a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill></dgm:ln>"
    "<dgm:effectLst/>"
    "</dgm:styleLbl>"
    "</dgm:colorsDef>"
)


class CT_ColorsDefinition(BaseOxmlElement):
    """`dgm:colorsDef` — root element of diagram colors part."""

    uniqueId = OptionalAttribute("uniqueId", XsdString)

    @classmethod
    def new(cls, colors_id: str = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2") -> CT_ColorsDefinition:
        return cast(
            CT_ColorsDefinition,
            _parse(_COLORS_TEMPLATE % {"ns": nsdecls("dgm", "a")}),
        )


# ---------------------------------------------------------------------------
# DiagramStyle root element (dgm:styleDef)
# ---------------------------------------------------------------------------

_STYLE_TEMPLATE = (
    '<dgm:styleDef %(ns)s'
    ' uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1"'
    ' minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
    '<dgm:title val=""/>'
    "<dgm:desc val=\"\"/>"
    '<dgm:catLst><dgm:cat type="simple" pri="1000"/></dgm:catLst>'
    '<dgm:styleLbl name="node0">'
    "<dgm:scene3d/>"
    "<dgm:sp3d/>"
    "<dgm:txPr><a:bodyPr/><a:lstStyle/><a:p/></dgm:txPr>"
    "<dgm:ln><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill></dgm:ln>"
    "<dgm:fill><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill></dgm:fill>"
    "</dgm:styleLbl>"
    '<dgm:styleLbl name="node1">'
    "<dgm:scene3d/>"
    "<dgm:sp3d/>"
    "<dgm:txPr><a:bodyPr/><a:lstStyle/><a:p/></dgm:txPr>"
    "<dgm:ln><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill></dgm:ln>"
    "<dgm:fill><a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill></dgm:fill>"
    "</dgm:styleLbl>"
    "</dgm:styleDef>"
)


class CT_StyleDefinition(BaseOxmlElement):
    """`dgm:styleDef` — root element of diagram style part."""

    uniqueId = OptionalAttribute("uniqueId", XsdString)

    @classmethod
    def new(cls, style_id: str = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1") -> CT_StyleDefinition:
        return cast(
            CT_StyleDefinition,
            _parse(_STYLE_TEMPLATE % {"ns": nsdecls("dgm", "a")}),
        )
