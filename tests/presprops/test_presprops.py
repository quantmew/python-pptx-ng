"""Tests for presentation properties (presprops) module."""

from __future__ import annotations

import os
import tempfile

import pytest

from pptx_ng import Presentation
from pptx_ng.presprops import PresentationProperties, ShowProperties


class DescribePresentationProperties:
    """Tests for PresentationProperties proxy class."""

    def it_provides_access_to_show_properties(self):
        prs = Presentation()
        props = prs.properties
        assert isinstance(props, PresentationProperties)

    def it_creates_show_properties_on_demand(self):
        prs = Presentation()
        show = prs.properties.show_properties
        assert isinstance(show, ShowProperties)

    def it_returns_same_show_properties_instance(self):
        prs = Presentation()
        show1 = prs.properties.show_properties
        show2 = prs.properties.show_properties
        assert show1 is show2


class DescribeShowProperties:
    """Tests for ShowProperties proxy class."""

    def it_reads_loop_attribute(self):
        prs = Presentation()
        show = prs.properties.show_properties
        # Default is None (not set)
        assert show.loop is None

    def it_writes_loop_attribute(self):
        prs = Presentation()
        show = prs.properties.show_properties
        show.loop = True
        assert show.loop is True
        show.loop = False
        assert show.loop is False
        show.loop = None
        assert show.loop is None

    def it_reads_show_animation_attribute(self):
        prs = Presentation()
        show = prs.properties.show_properties
        assert show.show_animation is None

    def it_writes_show_animation_attribute(self):
        prs = Presentation()
        show = prs.properties.show_properties
        show.show_animation = True
        assert show.show_animation is True
        show.show_animation = False
        assert show.show_animation is False

    def it_reads_show_narration_attribute(self):
        prs = Presentation()
        show = prs.properties.show_properties
        assert show.show_narration is None

    def it_writes_show_narration_attribute(self):
        prs = Presentation()
        show = prs.properties.show_properties
        show.show_narration = True
        assert show.show_narration is True

    def it_reads_use_timings_attribute(self):
        prs = Presentation()
        show = prs.properties.show_properties
        assert show.use_timings is None

    def it_writes_use_timings_attribute(self):
        prs = Presentation()
        show = prs.properties.show_properties
        show.use_timings = True
        assert show.use_timings is True
        show.use_timings = False
        assert show.use_timings is False


class DescribeShowPropertiesPersistence:
    """Tests that show properties persist through save/load."""

    def it_persists_loop_through_save_load(self):
        prs = Presentation()
        show = prs.properties.show_properties
        show.loop = True
        show.show_animation = False
        show.show_narration = True
        show.use_timings = False

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name
        try:
            prs.save(path)
            prs2 = Presentation(path)
            show2 = prs2.properties.show_properties
            assert show2.loop is True
            assert show2.show_animation is False
            assert show2.show_narration is True
            assert show2.use_timings is False
        finally:
            os.unlink(path)

    def it_persists_none_values_through_save_load(self):
        prs = Presentation()
        show = prs.properties.show_properties
        # All defaults are None, don't set anything
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name
        try:
            prs.save(path)
            prs2 = Presentation(path)
            show2 = prs2.properties.show_properties
            assert show2.loop is None
            assert show2.show_animation is None
            assert show2.show_narration is None
            assert show2.use_timings is None
        finally:
            os.unlink(path)


class DescribePresentationPropertiesXML:
    """Tests that verify XML structure of presentation properties."""

    def it_creates_correct_xml_structure(self):
        prs = Presentation()
        show = prs.properties.show_properties
        show.loop = True
        show.show_animation = False

        # Access the underlying XML element
        from lxml import etree

        pres_props_el = prs.properties._element
        showPr = pres_props_el.find(
            "{http://schemas.openxmlformats.org/presentationml/2006/main}showPr"
        )
        assert showPr is not None
        assert showPr.get("loop") == "1"
        assert showPr.get("showAnimation") == "0"

    def it_preserves_showPr_element_on_reload(self):
        prs = Presentation()
        show = prs.properties.show_properties
        show.loop = True
        show.use_timings = True

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name
        try:
            prs.save(path)
            # Open saved file and check XML directly
            import zipfile

            with zipfile.ZipFile(path) as zf:
                xml_bytes = zf.read("ppt/presProps.xml")
                from lxml import etree

                root = etree.fromstring(xml_bytes)
                nsmap = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
                showPr = root.find("p:showPr", nsmap)
                assert showPr is not None
                assert showPr.get("loop") == "1"
                assert showPr.get("useTimings") == "1"
        finally:
            os.unlink(path)
