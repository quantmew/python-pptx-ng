"""Unit tests for extended chart (chartEx) support (Phase 3.5)."""

from __future__ import annotations

import pytest

from pptx_ng import Presentation
from pptx_ng.chartex import ChartEx
from pptx_ng.oxml import parse_xml
from pptx_ng.oxml.chartex.chartex import (
    CT_ChartEx,
    CT_ChartExData,
    CT_ChartExSpace,
    CT_PlotArea,
    CT_SeriesLayoutPr,
)
from pptx_ng.oxml.ns import nsdecls, qn
from pptx_ng.parts.chartex import ChartExPart


class DescribeCT_ChartExSpace:
    def it_creates_empty_chart_space(self):
        cs = CT_ChartExSpace.new()
        assert cs.tag == qn("cx:chartSpace")
        assert cs.chart is not None
        assert cs.chartData is not None

    def it_parses_chart_space(self):
        xml = (
            '<cx:chartSpace %s>'
            '<cx:chart><cx:plotArea/></cx:chart>'
            '<cx:chartData/>'
            '</cx:chartSpace>'
        ) % nsdecls("cx")
        cs = parse_xml(xml)
        assert isinstance(cs, CT_ChartExSpace)
        assert cs.chart is not None

    def it_gets_or_adds_chart(self):
        cs = CT_ChartExSpace.new()
        chart = cs.get_or_add_chart()
        assert chart is not None
        assert isinstance(chart, CT_ChartEx)


class DescribeCT_ChartEx:
    def it_has_plot_area(self):
        cs = CT_ChartExSpace.new()
        chart = cs.chart
        assert chart.plotArea is not None
        assert isinstance(chart.plotArea, CT_PlotArea)

    def it_gets_or_adds_plot_area(self):
        xml = '<cx:chart %s/>' % nsdecls("cx")
        chart = parse_xml(xml)
        assert isinstance(chart, CT_ChartEx)
        pa = chart.get_or_add_plotArea()
        assert pa is not None
        assert isinstance(pa, CT_PlotArea)


class DescribeCT_PlotArea:
    def it_gets_or_adds_layout_pr(self):
        xml = '<cx:plotArea %s/>' % nsdecls("cx")
        pa = parse_xml(xml)
        assert isinstance(pa, CT_PlotArea)
        lp = pa.get_or_add_layoutPr()
        assert lp is not None
        assert isinstance(lp, CT_SeriesLayoutPr)

    def it_has_layout_pr(self):
        xml = (
            '<cx:plotArea %s>'
            '<cx:layoutPr layout="treemap"/>'
            '</cx:plotArea>'
        ) % nsdecls("cx")
        pa = parse_xml(xml)
        assert isinstance(pa, CT_PlotArea)
        assert pa.layoutPr is not None
        assert pa.layoutPr.layout == "treemap"


class DescribeCT_SeriesLayoutPr:
    def it_sets_and_gets_layout(self):
        xml = '<cx:layoutPr %s/>' % nsdecls("cx")
        lp = parse_xml(xml)
        assert isinstance(lp, CT_SeriesLayoutPr)
        lp.layout = "sunburst"
        assert lp.layout == "sunburst"

    def it_supports_chart_types(self):
        for chart_type in ("treemap", "sunburst", "waterfall", "funnel", "regionMap", "boxWhisker"):
            xml = '<cx:layoutPr %s/>' % nsdecls("cx")
            lp = parse_xml(xml)
            lp.layout = chart_type
            assert lp.layout == chart_type


class DescribeChartExPart:
    def it_creates_default_part(self):
        prs = Presentation()
        part = ChartExPart.new(prs.part.package)
        assert part._element is not None
        assert part._element.tag == qn("cx:chartSpace")

    def it_has_correct_content_type(self):
        from pptx_ng.opc.constants import CONTENT_TYPE as CT

        assert CT.OFC_CHART_EX == "application/vnd.ms-office.chartex+xml"

    def it_provides_chart_ex_proxy(self):
        prs = Presentation()
        part = ChartExPart.new(prs.part.package)
        chart_ex = part.chart_ex
        assert isinstance(chart_ex, ChartEx)


class DescribeChartExProxy:
    def it_gets_chart_type(self):
        cs = CT_ChartExSpace.new()
        chart = cs.get_or_add_chart()
        pa = chart.get_or_add_plotArea()
        lp = pa.get_or_add_layoutPr()
        lp.layout = "treemap"
        proxy = ChartEx(cs, part=None)
        assert proxy.chart_type == "treemap"

    def it_sets_chart_type(self):
        cs = CT_ChartExSpace.new()
        proxy = ChartEx(cs, part=None)
        proxy.chart_type = "waterfall"
        assert proxy.chart_type == "waterfall"

    def it_returns_none_when_no_chart_type(self):
        cs = CT_ChartExSpace.new()
        proxy = ChartEx(cs, part=None)
        assert proxy.chart_type is None

    def it_produces_xml(self):
        cs = CT_ChartExSpace.new()
        proxy = ChartEx(cs, part=None)
        xml = proxy.xml
        assert "cx:chartSpace" in xml


class DescribeChartExNamespace:
    def it_resolves_cx_namespace(self):
        from pptx_ng.oxml.ns import _nsmap

        assert "cx" in _nsmap
        assert _nsmap["cx"] == "http://schemas.microsoft.com/office/drawing/2014/chartex"

    def it_resolves_plot_area_element(self):
        xml = '<cx:plotArea %s/>' % nsdecls("cx")
        pa = parse_xml(xml)
        assert isinstance(pa, CT_PlotArea)

    def it_resolves_layout_pr_element(self):
        xml = '<cx:layoutPr %s/>' % nsdecls("cx")
        lp = parse_xml(xml)
        assert isinstance(lp, CT_SeriesLayoutPr)
