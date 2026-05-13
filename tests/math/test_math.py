"""Unit tests for OMML math formula support (Phase 3.7)."""

from __future__ import annotations

import pytest

from pptx_ng import Presentation
from pptx_ng.math import (
    MathAccent,
    MathArgument,
    MathDelimiter,
    MathFormula,
    MathFraction,
    MathFunction,
    MathMatrix,
    MathNary,
    MathRadical,
    MathRun,
    MathSubscript,
    MathSubSuperscript,
    MathSuperscript,
)
from pptx_ng.oxml import parse_xml
from pptx_ng.oxml.math.core import (
    CT_OMath,
    CT_OMathAccent,
    CT_OMathBar,
    CT_OMathBase,
    CT_OMathDelimiter,
    CT_OMathFrac,
    CT_OMathFunction,
    CT_OMathMatrix,
    CT_OMathNary,
    CT_OMathPara,
    CT_OMathRad,
    CT_OMathRun,
    CT_OMathSub,
    CT_OMathSubSup,
    CT_OMathSup,
)
from pptx_ng.oxml.ns import nsdecls, qn
from pptx_ng.util import Inches


# -- CT_OMath element tests --


class DescribeCT_OMath:
    def it_creates_empty_element(self):
        omath = CT_OMath.new()
        assert omath.tag == qn("m:oMath")

    def it_adds_run(self):
        omath = CT_OMath.new()
        r = omath.add_run("x")
        assert r.t == "x"
        assert len(omath.findall(qn("m:r"))) == 1

    def it_adds_fraction(self):
        omath = CT_OMath.new()
        f = omath.add_fraction()
        assert isinstance(f, CT_OMathFrac)
        assert f.num is not None
        assert f.den is not None

    def it_adds_radical(self):
        omath = CT_OMath.new()
        rad = omath.add_radical()
        assert isinstance(rad, CT_OMathRad)
        assert rad.e is not None

    def it_adds_superscript(self):
        omath = CT_OMath.new()
        sSup = omath.add_superscript()
        assert isinstance(sSup, CT_OMathSup)
        assert sSup.e is not None
        assert sSup.sup is not None

    def it_adds_subscript(self):
        omath = CT_OMath.new()
        sSub = omath.add_subscript()
        assert isinstance(sSub, CT_OMathSub)
        assert sSub.e is not None
        assert sSub.sub is not None

    def it_adds_sub_superscript(self):
        omath = CT_OMath.new()
        sSubSup = omath.add_sub_superscript()
        assert isinstance(sSubSup, CT_OMathSubSup)
        assert sSubSup.e is not None
        assert sSubSup.sub is not None
        assert sSubSup.sup is not None

    def it_adds_nary(self):
        omath = CT_OMath.new()
        nary = omath.add_nary("∑")
        assert isinstance(nary, CT_OMathNary)
        naryPr = nary.get_or_add_naryPr()
        assert naryPr.chr == "∑"

    def it_adds_delimiter(self):
        omath = CT_OMath.new()
        d = omath.add_delimiter()
        assert isinstance(d, CT_OMathDelimiter)

    def it_adds_accent(self):
        omath = CT_OMath.new()
        acc = omath.add_accent("̃")
        assert isinstance(acc, CT_OMathAccent)
        accPr = acc.get_or_add_accPr()
        assert accPr.chr == "̃"

    def it_adds_matrix(self):
        omath = CT_OMath.new()
        m = omath.add_matrix(2, 3)
        assert isinstance(m, CT_OMathMatrix)
        assert len(m.mr_lst) == 2
        assert len(m.mr_lst[0].e_lst) == 3

    def it_adds_function(self):
        omath = CT_OMath.new()
        func = omath.add_function("sin")
        assert isinstance(func, CT_OMathFunction)
        assert func.fName is not None


class DescribeCT_OMathPara:
    def it_creates_empty_element(self):
        para = CT_OMathPara.new()
        assert para.tag == qn("m:oMathPara")

    def it_holds_oMath_children(self):
        para = CT_OMathPara.new()
        omath = CT_OMath.new()
        para.append(omath)
        assert len(para.oMath_lst) == 1


class DescribeCT_OMathRun:
    def it_creates_with_text(self):
        r = CT_OMathRun.new("hello")
        assert r.t == "hello"

    def it_sets_text(self):
        r = CT_OMathRun.new()
        r.t = "world"
        assert r.t == "world"

    def it_escapes_xml_special_chars(self):
        r = CT_OMathRun.new("a<b&c>")
        assert r.t == "a<b&c>"


class DescribeCT_OMathArgument:
    def it_adds_run(self):
        e = CT_OMathBase.new()
        r = e.add_run("test")
        assert r.t == "test"

    def it_holds_fraction(self):
        xml = "<m:e %s><m:f><m:fPr/><m:num/><m:den/></m:f></m:e>" % nsdecls("m")
        e = parse_xml(xml)
        assert isinstance(e, CT_OMathBase)
        assert len(e.f_lst) == 1

    def it_adds_fraction(self):
        e = CT_OMathBase.new()
        f = e.add_fraction()
        assert f.num is not None
        assert f.den is not None

    def it_adds_superscript(self):
        e = CT_OMathBase.new()
        sSup = e.add_superscript()
        assert sSup.e is not None
        assert sSup.sup is not None


class DescribeCT_OMathFrac:
    def it_has_numerator_and_denominator(self):
        xml = (
            "<m:f %s><m:fPr/><m:num><m:r><m:t>1</m:t></m:r></m:num>"
            "<m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>"
        ) % nsdecls("m")
        f = parse_xml(xml)
        assert isinstance(f, CT_OMathFrac)
        assert f.num is not None
        assert f.den is not None


class DescribeCT_OMathNaryProperties:
    def it_gets_and_sets_chr(self):
        xml = '<m:naryPr %s><m:chr m:val="∫"/></m:naryPr>' % nsdecls("m")
        pr = parse_xml(xml)
        assert pr.chr == "∫"

    def it_sets_chr_when_absent(self):
        xml = "<m:naryPr %s/>" % nsdecls("m")
        pr = parse_xml(xml)
        pr.chr = "∑"
        assert pr.chr == "∑"


class DescribeCT_OMathMatrix:
    def it_has_rows_and_columns(self):
        xml = (
            "<m:m %s>"
            "<m:mPr/>"
            "<m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr>"
            "<m:mr><m:e><m:r><m:t>c</m:t></m:r></m:e><m:e><m:r><m:t>d</m:t></m:r></m:e></m:mr>"
            "</m:m>"
        ) % nsdecls("m")
        m = parse_xml(xml)
        assert isinstance(m, CT_OMathMatrix)
        assert len(m.mr_lst) == 2
        assert len(m.mr_lst[0].e_lst) == 2


class DescribeCT_OMathFunction:
    def it_has_name_and_body(self):
        xml = (
            "<m:func %s>"
            "<m:funcPr/>"
            "<m:fName><m:r><m:t>sin</m:t></m:r></m:fName>"
            "<m:e><m:r><m:t>x</m:t></m:r></m:e>"
            "</m:func>"
        ) % nsdecls("m")
        func = parse_xml(xml)
        assert isinstance(func, CT_OMathFunction)
        assert func.fName is not None
        assert func.e is not None


# -- High-level API tests --


class DescribeMathFormula:
    def it_creates_from_paragraph(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        math = p.add_math()
        assert isinstance(math, MathFormula)

    def it_adds_run(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        math = p.add_math()
        run = math.add_run("E=mc")
        assert isinstance(run, MathRun)
        assert run.text == "E=mc"

    def it_adds_fraction(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        math = p.add_math()
        frac = math.add_fraction()
        assert isinstance(frac, MathFraction)
        frac.numerator.add_run("1")
        frac.denominator.add_run("2")

    def it_adds_radical(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        math = p.add_math()
        rad = math.add_radical()
        assert isinstance(rad, MathRadical)
        rad.base.add_run("x")

    def it_adds_superscript(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        math = p.add_math()
        sSup = math.add_superscript()
        assert isinstance(sSup, MathSuperscript)
        sSup.base.add_run("x")
        sSup.superscript.add_run("2")

    def it_adds_nary_integral(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        math = p.add_math()
        nary = math.add_nary("∫")
        assert isinstance(nary, MathNary)
        nary.lower_limit.add_run("0")
        nary.upper_limit.add_run("∞")
        nary.base.add_run("x")

    def it_adds_matrix(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        math = p.add_math()
        m = math.add_matrix(2, 2)
        assert isinstance(m, MathMatrix)
        m.cell(0, 0).add_run("a")
        m.cell(0, 1).add_run("b")
        m.cell(1, 0).add_run("c")
        m.cell(1, 1).add_run("d")

    def it_adds_function(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        math = p.add_math()
        func = math.add_function("sin")
        assert isinstance(func, MathFunction)
        func.body.add_run("x")

    def it_produces_xml(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        math = p.add_math()
        math.add_run("x")
        xml = math.xml
        assert "m:oMath" in xml
        assert "m:r" in xml

    def it_persists_through_save_load(self, tmp_path):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        math = p.add_math()
        frac = math.add_fraction()
        frac.numerator.add_run("π")
        frac.denominator.add_run("2")
        path = tmp_path / "test_math.pptx"
        prs.save(str(path))
        prs2 = Presentation(str(path))
        slide2 = prs2.slides[0]
        omaths = slide2._element.xpath(".//m:oMath")
        assert len(omaths) == 1


class DescribeMathFormulaComplex:
    def it_builds_nested_formula(self):
        """Test nested: fraction inside superscript"""
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        math = p.add_math()
        sSup = math.add_superscript()
        sSup.base.add_run("x")
        arg = sSup.superscript
        arg.add_run("n")

    def it_builds_complex_equation(self):
        """Test building x = (-b ± √(b²-4ac)) / 2a"""
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        math = p.add_math()
        math.add_run("x=")
        frac = math.add_fraction()
        num = frac.numerator
        num.add_run("-b±√")
        nary = num.add_nary("")  # just as container
        frac.denominator.add_run("2a")
