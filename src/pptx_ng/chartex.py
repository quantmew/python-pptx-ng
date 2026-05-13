"""High-level API for extended chart (chartEx) support.

Provides the |ChartEx| proxy class for creating and manipulating modern chart
types like treemap, sunburst, waterfall, funnel, and region map charts.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx_ng.shared import PartElementProxy

if TYPE_CHECKING:
    from pptx_ng.oxml.chartex.chartex import CT_ChartExSpace
    from pptx_ng.parts.chartex import ChartExPart


class ChartEx(PartElementProxy):
    """Proxy for a `cx:chartSpace` element providing a high-level API for
    extended chart manipulation.

    Use |SlideShapes.add_chartex()| to create a new |ChartEx| instance.
    """

    def __init__(self, chartSpace: CT_ChartExSpace, part: ChartExPart):
        super(ChartEx, self).__init__(chartSpace, part)
        self._chartSpace = chartSpace

    @property
    def chart_type(self) -> str | None:
        """The chart type string (e.g. 'treemap', 'sunburst', 'waterfall').

        Returns |None| if no layout property is set.
        """
        chart = self._chartSpace.chart
        if chart is None:
            return None
        plot_area = chart.plotArea
        if plot_area is None:
            return None
        layout_pr = plot_area.layoutPr
        if layout_pr is None:
            return None
        return layout_pr.layout

    @chart_type.setter
    def chart_type(self, value: str):
        chart = self._chartSpace.get_or_add_chart()
        plot_area = chart.get_or_add_plotArea()
        layout_pr = plot_area.get_or_add_layoutPr()
        layout_pr.layout = value

    @property
    def xml(self) -> str:
        """The XML of this chart element."""
        from lxml import etree

        return etree.tostring(self._chartSpace, encoding="unicode")
