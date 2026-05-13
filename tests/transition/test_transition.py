"""Tests for slide transition module."""

from __future__ import annotations

import os
import tempfile

import pytest

from pptx_ng import Presentation
from pptx_ng.transition import Transition


class DescribeSlideTransition:
    """Tests for Slide.transition property."""

    def it_returns_none_when_no_transition(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        assert slide.transition is None

    def it_adds_transition(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        assert isinstance(t, Transition)
        assert slide.transition is not None

    def it_replaces_existing_transition(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t1 = slide.add_transition()
        t1.set_fade()
        t2 = slide.add_transition()
        t2.set_push()
        # Same element is reused (get_or_add)
        assert t2.transition_type == "push"


class DescribeTransitionProperties:
    """Tests for Transition read/write properties."""

    def it_reads_speed(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        assert t.speed is None

    def it_writes_speed(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.speed = "fast"
        assert t.speed == "fast"
        t.speed = "slow"
        assert t.speed == "slow"
        t.speed = None
        assert t.speed is None

    def it_reads_advance_on_click(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        assert t.advance_on_click is None

    def it_writes_advance_on_click(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.advance_on_click = False
        assert t.advance_on_click is False
        t.advance_on_click = True
        assert t.advance_on_click is True
        t.advance_on_click = None
        assert t.advance_on_click is None

    def it_reads_advance_after_time(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        assert t.advance_after_time is None

    def it_writes_advance_after_time(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.advance_after_time = 5000
        assert t.advance_after_time == 5000
        t.advance_after_time = None
        assert t.advance_after_time is None


class DescribeTransitionTypes:
    """Tests for all transition type set methods."""

    def it_sets_blinds(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_blinds("vert")
        assert t.transition_type == "blinds"

    def it_sets_checker(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_checker()
        assert t.transition_type == "checker"

    def it_sets_comb(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_comb()
        assert t.transition_type == "comb"

    def it_sets_cover(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_cover("r")
        assert t.transition_type == "cover"

    def it_sets_pull(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_pull()
        assert t.transition_type == "pull"

    def it_sets_cut(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_cut(thruBlk=True)
        assert t.transition_type == "cut"

    def it_sets_fade(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_fade()
        assert t.transition_type == "fade"

    def it_sets_push(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_push("r")
        assert t.transition_type == "push"

    def it_sets_wipe(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_wipe()
        assert t.transition_type == "wipe"

    def it_sets_split(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_split(orient="vert", dir="in")
        assert t.transition_type == "split"

    def it_sets_strips(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_strips()
        assert t.transition_type == "strips"

    def it_sets_random_bar(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_random_bar()
        assert t.transition_type == "randomBar"

    def it_sets_wheel(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_wheel(spokes=8)
        assert t.transition_type == "wheel"

    def it_sets_wedge(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_wedge()
        assert t.transition_type == "wedge"

    def it_sets_zoom(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_zoom()
        assert t.transition_type == "zoom"

    def it_sets_dissolve(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_dissolve()
        assert t.transition_type == "dissolve"

    def it_sets_random(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_random()
        assert t.transition_type == "random"

    def it_sets_circle(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_circle()
        assert t.transition_type == "circle"

    def it_sets_diamond(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_diamond()
        assert t.transition_type == "diamond"

    def it_sets_newsflash(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_newsflash()
        assert t.transition_type == "newsflash"

    def it_sets_plus(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_plus()
        assert t.transition_type == "plus"

    def it_replaces_transition_type(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_fade()
        assert t.transition_type == "fade"
        t.set_push()
        assert t.transition_type == "push"


class DescribeTransitionPersistence:
    """Tests that transitions persist through save/load."""

    def it_persists_fade_through_save_load(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_fade()
        t.speed = "fast"
        t.advance_after_time = 5000
        t.advance_on_click = False

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name
        try:
            prs.save(path)
            prs2 = Presentation(path)
            slide2 = prs2.slides[0]
            t2 = slide2.transition
            assert t2 is not None
            assert t2.transition_type == "fade"
            assert t2.speed == "fast"
            assert t2.advance_after_time == 5000
            assert t2.advance_on_click is False
        finally:
            os.unlink(path)

    def it_persists_wheel_through_save_load(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_wheel(spokes=8)
        t.speed = "slow"

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name
        try:
            prs.save(path)
            prs2 = Presentation(path)
            slide2 = prs2.slides[0]
            t2 = slide2.transition
            assert t2 is not None
            assert t2.transition_type == "wheel"
            assert t2.speed == "slow"
        finally:
            os.unlink(path)

    def it_persists_push_with_direction(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_push("r")

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name
        try:
            prs.save(path)
            prs2 = Presentation(path)
            slide2 = prs2.slides[0]
            t2 = slide2.transition
            assert t2.transition_type == "push"
        finally:
            os.unlink(path)


class DescribeTransitionXML:
    """Tests that verify XML structure of transitions."""

    def it_produces_correct_fade_xml(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_fade(thruBlk=True)
        t.speed = "fast"

        from lxml import etree

        trans_el = t._element
        assert trans_el.get("spd") == "fast"
        nsmap = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
        fade = trans_el.find("p:fade", nsmap)
        assert fade is not None
        assert fade.get("thruBlk") == "1"

    def it_verifies_transition_xml_in_saved_file(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        t = slide.add_transition()
        t.set_wipe("r")
        t.speed = "med"
        t.advance_after_time = 3000

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name
        try:
            prs.save(path)
            import zipfile
            from lxml import etree

            with zipfile.ZipFile(path) as zf:
                xml_bytes = zf.read("ppt/slides/slide1.xml")
                root = etree.fromstring(xml_bytes)
                nsmap = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
                trans = root.find("p:transition", nsmap)
                assert trans is not None
                assert trans.get("spd") == "med"
                assert trans.get("advTm") == "3000"
                wipe = trans.find("p:wipe", nsmap)
                assert wipe is not None
                assert wipe.get("dir") == "r"
        finally:
            os.unlink(path)


class DescribeOxmlElements:
    """Tests for low-level oxml transition element classes."""

    def it_creates_transition_element(self):
        from pptx_ng.oxml.xmlchemy import OxmlElement

        trans = OxmlElement("p:transition")
        trans.set("spd", "fast")
        assert trans.get("spd") == "fast"

    def it_detects_transition_type(self):
        from pptx_ng.oxml.xmlchemy import OxmlElement

        trans = OxmlElement("p:transition")
        assert trans.transition_type is None
        trans.set_fade()
        assert trans.transition_type == "fade"

    def it_handles_split_attributes(self):
        from pptx_ng.oxml.xmlchemy import OxmlElement
        from lxml import etree

        trans = OxmlElement("p:transition")
        trans.set_split(orient="vert", dir="in")
        nsmap = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
        split = trans.find("p:split", nsmap)
        assert split is not None
        assert split.get("orient") == "vert"
        assert split.get("dir") == "in"

    def it_handles_wheel_spokes(self):
        from pptx_ng.oxml.xmlchemy import OxmlElement
        from lxml import etree

        trans = OxmlElement("p:transition")
        trans.set_wheel(spokes=8)
        nsmap = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
        wheel = trans.find("p:wheel", nsmap)
        assert wheel is not None
        assert wheel.get("spokes") == "8"
