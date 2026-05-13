"""Core chartEx XML element classes.

Provides element classes for extended chart (chartEx) support. ChartEx is used
for modern chart types like treemap, sunburst, waterfall, histogram,
box-and-whisker, funnel, and region map charts (Office 2016+).

Reference: Open XML SDK, Office2016.Drawing.ChartDrawing namespace
"""

from __future__ import annotations

from typing import Callable, cast

from pptx_ng.oxml.ns import nsdecls
from pptx_ng.oxml.simpletypes import XsdString
from pptx_ng.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    ZeroOrOne,
)


def _parse(xml: str) -> BaseOxmlElement:
    from pptx_ng.oxml import parse_xml

    return parse_xml(xml)


class CT_ChartExSpace(BaseOxmlElement):
    """`cx:chartSpace` — root element of an extended chart part."""

    _tag_seq = ("cx:chart", "cx:chartData", "cx:fmtOvrs")
    chart = ZeroOrOne("cx:chart", successors=("cx:chartData", "cx:fmtOvrs"))
    chartData = ZeroOrOne("cx:chartData", successors=("cx:fmtOvrs",))
    del _tag_seq

    get_or_add_chart: Callable[[], CT_ChartEx]

    @classmethod
    def new(cls) -> CT_ChartExSpace:
        return cast(
            CT_ChartExSpace,
            _parse(
                '<cx:chartSpace %s xmlns:cxs="http://schemas.microsoft.com/office/drawing/2014/chartex">'
                "<cx:chart><cx:plotArea/></cx:chart>"
                "<cx:chartData/>"
                "</cx:chartSpace>" % nsdecls("cx")
            ),
        )


class CT_ChartEx(BaseOxmlElement):
    """`cx:chart` — chart definition within a chartEx part."""

    _tag_seq = ("cx:plotArea", "cx:extLst")
    plotArea: Callable[[], CT_PlotArea]
    plotArea = ZeroOrOne("cx:plotArea", successors=("cx:extLst",))
    del _tag_seq

    get_or_add_plotArea: Callable[[], CT_PlotArea]


class CT_ChartExData(BaseOxmlElement):
    """`cx:chartData` — chart data within a chartEx part."""

    pass


class CT_PlotArea(BaseOxmlElement):
    """`cx:plotArea` — chart plot area containing series and layout."""

    _tag_seq = ("cx:layoutPr", "cx:extLst")
    layoutPr: Callable[[], CT_SeriesLayoutPr]
    layoutPr = ZeroOrOne("cx:layoutPr", successors=("cx:extLst",))
    del _tag_seq

    get_or_add_layoutPr: Callable[[], CT_SeriesLayoutPr]


class CT_SeriesLayoutPr(BaseOxmlElement):
    """`cx:layoutPr` — series layout properties defining the chart type."""

    layout = OptionalAttribute("layout", XsdString)
