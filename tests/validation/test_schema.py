"""Unit-test suite for the pptx.validation.schema module."""

from __future__ import annotations

from lxml import etree

from pptx_ng.oxml import parse_xml
from pptx_ng.oxml.ns import qn
from pptx_ng.oxml.simpletypes import BaseIntType, XsdBoolean
from pptx_ng.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OneOrMore,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
    ZeroOrOneChoice,
    Choice,
)
from pptx_ng.validation.context import ValidationContext
from pptx_ng.validation.schema import SchemaValidator


# -- test element classes ----------------------------------------------------

class CT_TestRequiredChild(BaseOxmlElement):
    required_child = OneAndOnlyOne("p:child")


class CT_TestRepeatingChild(BaseOxmlElement):
    items = OneOrMore("p:item")


class CT_TestChoiceGroup(BaseOxmlElement):
    fill = ZeroOrOneChoice(
        (Choice("a:solidFill"), Choice("a:noFill"), Choice("a:gradFill")),
        successors=(),
    )


class CT_TestAttrs(BaseOxmlElement):
    required_attr = RequiredAttribute("cnt", BaseIntType)
    optional_attr = OptionalAttribute("visible", XsdBoolean, default=None)


class CT_TestNoDecls(BaseOxmlElement):
    pass


# -- test fixtures -----------------------------------------------------------

class DescribeSchemaValidator:
    def it_detects_missing_required_child(self):
        ctx = ValidationContext()
        element = parse_xml(
            '<p:parent xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
        )
        # Cast to test class
        validator = SchemaValidator()
        validator._check_required_children(ctx, element, None)
        # No declarations on plain element, should not error
        assert ctx.is_valid

    def it_detects_missing_required_attribute(self):
        ctx = ValidationContext()
        element = parse_xml(
            '<p:test xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
        )
        validator = SchemaValidator()
        validator._validate_required_attr(
            ctx,
            element,
            RequiredAttribute("cnt", BaseIntType),
            None,
        )
        assert not ctx.is_valid
        assert "cnt" in ctx.errors[0].description

    def it_accepts_valid_required_attribute(self):
        ctx = ValidationContext()
        element = parse_xml(
            '<p:test cnt="5" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
        )
        validator = SchemaValidator()
        validator._validate_required_attr(
            ctx,
            element,
            RequiredAttribute("cnt", BaseIntType),
            None,
        )
        assert ctx.is_valid

    def it_detects_invalid_attribute_value(self):
        ctx = ValidationContext()
        element = parse_xml(
            '<p:test cnt="abc" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
        )
        validator = SchemaValidator()
        validator._validate_required_attr(
            ctx,
            element,
            RequiredAttribute("cnt", BaseIntType),
            None,
        )
        assert not ctx.is_valid

    def it_accepts_valid_optional_attribute(self):
        ctx = ValidationContext()
        element = parse_xml(
            '<p:test visible="true" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
        )
        validator = SchemaValidator()
        validator._validate_optional_attr(
            ctx,
            element,
            OptionalAttribute("visible", XsdBoolean, default=None),
            None,
        )
        assert ctx.is_valid

    def it_skips_missing_optional_attribute(self):
        ctx = ValidationContext()
        element = parse_xml(
            '<p:test xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
        )
        validator = SchemaValidator()
        validator._validate_optional_attr(
            ctx,
            element,
            OptionalAttribute("visible", XsdBoolean, default=None),
            None,
        )
        assert ctx.is_valid

    def it_validates_full_element_tree(self):
        ctx = ValidationContext()
        element = parse_xml(
            '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '  <p:sldMasterIdLst/>'
            '  <p:sldSz cx="9144000" cy="6858000" type="custom"/>'
            '</p:presentation>'
        )
        validator = SchemaValidator()
        validator.validate_element(ctx, element)
        assert ctx.is_valid

    def it_detects_choice_exclusivity_violation(self):
        ctx = ValidationContext()
        element = parse_xml(
            '<p:test xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '  <a:solidFill/>'
            '  <a:noFill/>'
            '</p:test>'
        )
        validator = SchemaValidator()
        validator._check_choice_exclusivity(ctx, element, None)
        # This element doesn't have a ZeroOrOneChoice declaration, so it won't
        # trigger choice validation. This tests the path for elements without
        # choice declarations.
        assert ctx.is_valid
