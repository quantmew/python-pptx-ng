"""Core chartEx XML element classes.

Provides minimal element classes for extended chart (chartEx) round-trip
support. ChartEx is used for modern chart types like treemap, sunburst,
waterfall, histogram, box-and-whisker, funnel, and region map charts.
"""

from __future__ import annotations

from typing import Callable, cast

from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    ZeroOrOne,
)


def _parse(xml: str) -> BaseOxmlElement:
    from pptx.oxml import parse_xml

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
                "<cx:chart/>"
                "<cx:chartData/>"
                "</cx:chartSpace>" % nsdecls("cx")
            ),
        )


class CT_ChartEx(BaseOxmlElement):
    """`cx:chart` — chart definition within a chartEx part."""

    pass


class CT_ChartExData(BaseOxmlElement):
    """`cx:chartData` — chart data within a chartEx part."""

    pass
