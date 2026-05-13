"""Tests for custom show module."""

from __future__ import annotations

import os
import tempfile

import pytest

from pptx_ng import Presentation
from pptx_ng.customshow import CustomShow, CustomShows


class DescribeCustomShows:
    """Tests for CustomShows collection."""

    def it_provides_empty_collection_initially(self):
        prs = Presentation()
        shows = prs.custom_shows
        assert len(shows) == 0

    def it_adds_a_custom_show(self):
        prs = Presentation()
        shows = prs.custom_shows
        show = shows.add("Test Show")
        assert isinstance(show, CustomShow)
        assert show.name == "Test Show"
        assert show.id == 1

    def it_adds_custom_show_with_slides(self):
        prs = Presentation()
        slide_layout = prs.slide_layouts[0]
        prs.slides.add_slide(slide_layout)
        prs.slides.add_slide(slide_layout)
        rIds = [sldId.rId for sldId in prs._element.sldIdLst.sldId_lst]

        shows = prs.custom_shows
        show = shows.add("With Slides", [rIds[0], rIds[1]])
        assert show.slide_ids == [rIds[0], rIds[1]]

    def it_increments_ids(self):
        prs = Presentation()
        shows = prs.custom_shows
        show1 = shows.add("Show 1")
        show2 = shows.add("Show 2")
        show3 = shows.add("Show 3")
        assert show1.id == 1
        assert show2.id == 2
        assert show3.id == 3

    def it_iterates_custom_shows(self):
        prs = Presentation()
        shows = prs.custom_shows
        shows.add("Show A")
        shows.add("Show B")
        names = [s.name for s in shows]
        assert names == ["Show A", "Show B"]

    def it_reports_length(self):
        prs = Presentation()
        shows = prs.custom_shows
        assert len(shows) == 0
        shows.add("Show 1")
        shows.add("Show 2")
        assert len(shows) == 2

    def it_gets_by_id(self):
        prs = Presentation()
        shows = prs.custom_shows
        shows.add("Show 1")
        shows.add("Show 2")
        found = shows.get(2)
        assert found is not None
        assert found.name == "Show 2"

    def it_returns_none_for_missing_id(self):
        prs = Presentation()
        shows = prs.custom_shows
        assert shows.get(99) is None

    def it_gets_by_name(self):
        prs = Presentation()
        shows = prs.custom_shows
        shows.add("Alpha")
        shows.add("Beta")
        found = shows.get_by_name("Alpha")
        assert found is not None
        assert found.id == 1

    def it_returns_none_for_missing_name(self):
        prs = Presentation()
        shows = prs.custom_shows
        assert shows.get_by_name("Nonexistent") is None


class DescribeCustomShow:
    """Tests for CustomShow proxy."""

    def it_reads_id(self):
        prs = Presentation()
        show = prs.custom_shows.add("Test")
        assert show.id == 1

    def it_reads_name(self):
        prs = Presentation()
        show = prs.custom_shows.add("My Show")
        assert show.name == "My Show"

    def it_writes_name(self):
        prs = Presentation()
        show = prs.custom_shows.add("Old Name")
        show.name = "New Name"
        assert show.name == "New Name"

    def it_reads_empty_slide_ids(self):
        prs = Presentation()
        show = prs.custom_shows.add("Empty")
        assert show.slide_ids == []

    def it_adds_slide(self):
        prs = Presentation()
        slide_layout = prs.slide_layouts[0]
        prs.slides.add_slide(slide_layout)
        rIds = [sldId.rId for sldId in prs._element.sldIdLst.sldId_lst]

        show = prs.custom_shows.add("Test")
        show.add_slide(rIds[0])
        assert show.slide_ids == [rIds[0]]

    def it_adds_multiple_slides(self):
        prs = Presentation()
        slide_layout = prs.slide_layouts[0]
        prs.slides.add_slide(slide_layout)
        prs.slides.add_slide(slide_layout)
        prs.slides.add_slide(slide_layout)
        rIds = [sldId.rId for sldId in prs._element.sldIdLst.sldId_lst]

        show = prs.custom_shows.add("Test", [rIds[0], rIds[1], rIds[2]])
        assert show.slide_ids == [rIds[0], rIds[1], rIds[2]]

    def it_removes_slide(self):
        prs = Presentation()
        slide_layout = prs.slide_layouts[0]
        prs.slides.add_slide(slide_layout)
        prs.slides.add_slide(slide_layout)
        rIds = [sldId.rId for sldId in prs._element.sldIdLst.sldId_lst]

        show = prs.custom_shows.add("Test", [rIds[0], rIds[1]])
        show.remove_slide(rIds[0])
        assert show.slide_ids == [rIds[1]]

    def it_raises_on_remove_missing_slide(self):
        prs = Presentation()
        show = prs.custom_shows.add("Test")
        with pytest.raises(ValueError, match="not found"):
            show.remove_slide("rId999")


class DescribeCustomShowPersistence:
    """Tests that custom shows persist through save/load."""

    def it_persists_custom_shows_through_save_load(self):
        prs = Presentation()
        slide_layout = prs.slide_layouts[0]
        prs.slides.add_slide(slide_layout)
        prs.slides.add_slide(slide_layout)
        prs.slides.add_slide(slide_layout)
        rIds = [sldId.rId for sldId in prs._element.sldIdLst.sldId_lst]

        shows = prs.custom_shows
        shows.add("Show A", [rIds[0], rIds[2]])
        shows.add("Show B", [rIds[1]])

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name
        try:
            prs.save(path)

            prs2 = Presentation(path)
            shows2 = prs2.custom_shows
            assert len(shows2) == 2

            show_a = shows2.get_by_name("Show A")
            assert show_a is not None
            assert show_a.slide_ids == [rIds[0], rIds[2]]

            show_b = shows2.get_by_name("Show B")
            assert show_b is not None
            assert show_b.slide_ids == [rIds[1]]
        finally:
            os.unlink(path)

    def it_persists_name_change_through_save_load(self):
        prs = Presentation()
        show = prs.custom_shows.add("Original")
        show.name = "Renamed"

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name
        try:
            prs.save(path)
            prs2 = Presentation(path)
            assert prs2.custom_shows.get_by_name("Renamed") is not None
            assert prs2.custom_shows.get_by_name("Original") is None
        finally:
            os.unlink(path)

    def it_persists_empty_custom_show(self):
        prs = Presentation()
        prs.custom_shows.add("Empty Show")

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name
        try:
            prs.save(path)
            prs2 = Presentation(path)
            show = prs2.custom_shows.get_by_name("Empty Show")
            assert show is not None
            assert show.slide_ids == []
        finally:
            os.unlink(path)


class DescribeCustomShowXML:
    """Tests that verify XML structure of custom shows."""

    def it_produces_correct_xml(self):
        prs = Presentation()
        slide_layout = prs.slide_layouts[0]
        prs.slides.add_slide(slide_layout)
        rIds = [sldId.rId for sldId in prs._element.sldIdLst.sldId_lst]

        show = prs.custom_shows.add("My Show", [rIds[0]])

        from lxml import etree

        custShow = show._element
        assert custShow.get("id") == "1"
        assert custShow.get("name") == "My Show"

        nsmap = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
                 "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
        sldLst = custShow.find("p:sldLst", nsmap)
        assert sldLst is not None
        slds = sldLst.findall("p:sld", nsmap)
        assert len(slds) == 1
        assert slds[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id") == rIds[0]

    def it_verifies_xml_in_saved_file(self):
        prs = Presentation()
        slide_layout = prs.slide_layouts[0]
        prs.slides.add_slide(slide_layout)
        prs.slides.add_slide(slide_layout)
        rIds = [sldId.rId for sldId in prs._element.sldIdLst.sldId_lst]

        prs.custom_shows.add("Demo", [rIds[0], rIds[1]])

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name
        try:
            prs.save(path)
            import zipfile
            from lxml import etree

            with zipfile.ZipFile(path) as zf:
                xml_bytes = zf.read("ppt/presentation.xml")
                root = etree.fromstring(xml_bytes)
                nsmap = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
                         "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
                custShowLst = root.find("p:custShowLst", nsmap)
                assert custShowLst is not None
                custShows = custShowLst.findall("p:custShow", nsmap)
                assert len(custShows) == 1
                assert custShows[0].get("name") == "Demo"
        finally:
            os.unlink(path)


class DescribeOxmlElements:
    """Tests for low-level oxml custom show elements."""

    def it_creates_custShowLst(self):
        from pptx_ng.oxml.customshow import CT_CustomShowList

        el = CT_CustomShowList.new()
        assert el.tag.endswith("}custShowLst")

    def it_adds_custShow(self):
        from pptx_ng.oxml.customshow import CT_CustomShowList

        lst = CT_CustomShowList.new()
        cs = lst.add_custShow(id=1, name="Test")
        assert cs.id == 1
        assert cs.name == "Test"
        assert cs.sldLst is not None

    def it_creates_sldLst_and_adds_sld(self):
        from pptx_ng.oxml.customshow import CT_SlideRelationshipList
        from pptx_ng.oxml.ns import qn

        lst = CT_SlideRelationshipList.new()
        entry = lst.add_sld("rId3")
        assert entry.get(qn("r:id")) == "rId3"
        assert len(lst.sld_lst) == 1

    def it_lists_sld_entries(self):
        from pptx_ng.oxml.customshow import CT_SlideRelationshipList
        from pptx_ng.oxml.ns import qn

        lst = CT_SlideRelationshipList.new()
        lst.add_sld("rId1")
        lst.add_sld("rId2")
        lst.add_sld("rId3")
        ids = [s.get(qn("r:id")) for s in lst.sld_lst]
        assert ids == ["rId1", "rId2", "rId3"]
