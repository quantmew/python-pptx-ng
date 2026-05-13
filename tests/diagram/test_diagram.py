"""Unit tests for SmartArt/Diagram support (Phase 3.4)."""

from __future__ import annotations

import pytest

from pptx_ng import Presentation
from pptx_ng.diagram import SmartArtData, SmartArtNode
from pptx_ng.oxml import parse_xml
from pptx_ng.oxml.diagram.core import (
    CT_ColorsDefinition,
    CT_ConnectionList,
    CT_DataModel,
    CT_LayoutDefinition,
    CT_PointList,
    CT_StyleDefinition,
)
from pptx_ng.oxml.ns import nsdecls, qn
from pptx_ng.parts.diagram import (
    DiagramColorsPart,
    DiagramDataPart,
    DiagramLayoutPart,
    DiagramStylePart,
)
from pptx_ng.util import Inches


# -- Oxml element tests --


class DescribeCT_DataModel:
    def it_creates_empty_model(self):
        dm = CT_DataModel.new()
        assert dm.tag == qn("dgm:dataModel")
        assert dm.ptLst is not None
        assert dm.cxnLst is not None

    def it_adds_point(self):
        dm = CT_DataModel.new()
        ptLst = dm.get_or_add_ptLst()
        pt = ptLst.add_point(0, "Hello")
        assert pt.modelId == "0"

    def it_adds_connection(self):
        dm = CT_DataModel.new()
        cxnLst = dm.get_or_add_cxnLst()
        cxn = cxnLst.add_connection(0, 0, 1)
        assert cxn.modelId == "0"
        assert cxn.srcId == "0"
        assert cxn.destId == "1"


class DescribeCT_LayoutDefinition:
    def it_creates_default_layout(self):
        ld = CT_LayoutDefinition.new()
        assert ld.tag == qn("dgm:layoutDef")
        assert ld.uniqueId is not None


class DescribeCT_ColorsDefinition:
    def it_creates_default_colors(self):
        cd = CT_ColorsDefinition.new()
        assert cd.tag == qn("dgm:colorsDef")
        assert cd.uniqueId is not None


class DescribeCT_StyleDefinition:
    def it_creates_default_style(self):
        sd = CT_StyleDefinition.new()
        assert sd.tag == qn("dgm:styleDef")
        assert sd.uniqueId is not None


# -- Part tests --


class DescribeDiagramDataPart:
    def it_creates_default_part(self):
        prs = Presentation()
        part = DiagramDataPart.new(prs.part.package)
        assert part.data_model is not None
        assert isinstance(part.data_model, SmartArtData)


class DescribeDiagramLayoutPart:
    def it_creates_default_part(self):
        prs = Presentation()
        part = DiagramLayoutPart.new(prs.part.package)
        assert part._element is not None
        assert part._element.tag == qn("dgm:layoutDef")


class DescribeDiagramColorsPart:
    def it_creates_default_part(self):
        prs = Presentation()
        part = DiagramColorsPart.new(prs.part.package)
        assert part._element is not None
        assert part._element.tag == qn("dgm:colorsDef")


class DescribeDiagramStylePart:
    def it_creates_default_part(self):
        prs = Presentation()
        part = DiagramStylePart.new(prs.part.package)
        assert part._element is not None
        assert part._element.tag == qn("dgm:styleDef")


# -- High-level API tests --


class DescribeSmartArtData:
    def it_adds_node(self):
        dm = CT_DataModel.new()
        data = SmartArtData(dm)
        node = data.add_node("First Node")
        assert isinstance(node, SmartArtNode)

    def it_lists_nodes(self):
        dm = CT_DataModel.new()
        data = SmartArtData(dm)
        data.add_node("A")
        data.add_node("B")
        assert len(data.nodes) == 2

    def it_assigns_incrementing_ids(self):
        dm = CT_DataModel.new()
        data = SmartArtData(dm)
        n0 = data.add_node("A")
        n1 = data.add_node("B")
        assert n0.model_id == "1"
        assert n1.model_id == "2"


# -- Integration tests --


class DescribeSlideShapesAddSmartArt:
    def it_adds_smartart_shape_to_slide(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        gf = slide.shapes.add_smartart(Inches(1), Inches(1), Inches(6), Inches(4))
        assert gf is not None
        assert gf.shape_id is not None

    def it_creates_diagram_parts(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.add_smartart(Inches(1), Inches(1), Inches(6), Inches(4))
        rels = slide.part.rels
        rel_types = {r.reltype for r in rels.values()}
        from pptx_ng.opc.constants import RELATIONSHIP_TYPE as RT
        assert RT.DIAGRAM_DATA in rel_types
        assert RT.DIAGRAM_LAYOUT in rel_types
        assert RT.DIAGRAM_QUICK_STYLE in rel_types
        assert RT.DIAGRAM_COLORS in rel_types

    def it_creates_graphic_frame_with_diagram(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.add_smartart(Inches(1), Inches(1), Inches(6), Inches(4))
        graphicFrames = slide._element.xpath(".//p:graphicFrame")
        assert len(graphicFrames) == 1
        relIds = graphicFrames[0].xpath(".//dgm:relIds")
        assert len(relIds) == 1

    def it_persists_through_save_load(self, tmp_path):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.add_smartart(Inches(1), Inches(1), Inches(6), Inches(4))
        path = tmp_path / "test_smartart.pptx"
        prs.save(str(path))
        prs2 = Presentation(str(path))
        slide2 = prs2.slides[0]
        graphicFrames = slide2._element.xpath(".//p:graphicFrame")
        assert len(graphicFrames) == 1
