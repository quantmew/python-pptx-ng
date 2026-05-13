"""Unit tests for extended chart (chartEx) support (Phase 3.5)."""

from __future__ import annotations

import pytest

from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.oxml.chartex.chartex import CT_ChartEx, CT_ChartExData, CT_ChartExSpace
from pptx.oxml.ns import nsdecls, qn
from pptx.parts.chartex import ChartExPart


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


class DescribeChartExPart:
    def it_creates_default_part(self):
        prs = Presentation()
        part = ChartExPart.new(prs.part.package)
        assert part._element is not None
        assert part._element.tag == qn("cx:chartSpace")

    def it_has_correct_content_type(self):
        from pptx.opc.constants import CONTENT_TYPE as CT
        assert CT.OFC_CHART_EX == "application/vnd.ms-office.chartex+xml"


class DescribeChartExNamespace:
    def it_resolves_cx_namespace(self):
        from pptx.oxml.ns import _nsmap
        assert "cx" in _nsmap
        assert _nsmap["cx"] == "http://schemas.microsoft.com/office/drawing/2014/chartex"
