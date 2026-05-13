"""Unit-test suite for the pptx.validation.constraints module."""

from __future__ import annotations

from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.validation.constraints import (
    AttributeCannotOmitConstraint,
    AttributePairConstraint,
    AttributeValueRangeConstraint,
    ReferenceExistConstraint,
    UniqueAttributeValueConstraint,
)
from pptx.validation.context import ValidationContext

NSMAP = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


class DescribeAttributeCannotOmitConstraint:
    def it_passes_when_attribute_present(self):
        ctx = ValidationContext()
        element = parse_xml(f'<p:test cnt="5" {NSMAP}/>')
        constraint = AttributeCannotOmitConstraint("cnt")
        constraint.validate(ctx, element)
        assert ctx.is_valid

    def it_fails_when_attribute_missing(self):
        ctx = ValidationContext()
        element = parse_xml(f'<p:test {NSMAP}/>')
        constraint = AttributeCannotOmitConstraint("cnt")
        constraint.validate(ctx, element)
        assert not ctx.is_valid
        assert "cnt" in ctx.errors[0].description


class DescribeAttributeValueRangeConstraint:
    def it_passes_for_value_in_range(self):
        ctx = ValidationContext()
        element = parse_xml(f'<p:test cnt="5" {NSMAP}/>')
        constraint = AttributeValueRangeConstraint(
            attr_name="cnt", min_value=0, min_inclusive=True, max_value=10, max_inclusive=True
        )
        constraint.validate(ctx, element)
        assert ctx.is_valid

    def it_fails_for_value_below_min(self):
        ctx = ValidationContext()
        element = parse_xml(f'<p:test cnt="-1" {NSMAP}/>')
        constraint = AttributeValueRangeConstraint(
            attr_name="cnt", min_value=0, min_inclusive=True, max_value=10, max_inclusive=True
        )
        constraint.validate(ctx, element)
        assert not ctx.is_valid

    def it_fails_for_value_above_max(self):
        ctx = ValidationContext()
        element = parse_xml(f'<p:test cnt="15" {NSMAP}/>')
        constraint = AttributeValueRangeConstraint(
            attr_name="cnt", min_value=0, min_inclusive=True, max_value=10, max_inclusive=True
        )
        constraint.validate(ctx, element)
        assert not ctx.is_valid

    def it_skips_missing_attribute(self):
        ctx = ValidationContext()
        element = parse_xml(f'<p:test {NSMAP}/>')
        constraint = AttributeValueRangeConstraint(
            attr_name="cnt", min_value=0, min_inclusive=True, max_value=10, max_inclusive=True
        )
        constraint.validate(ctx, element)
        assert ctx.is_valid

    def it_respects_exclusive_bounds(self):
        ctx = ValidationContext()
        element = parse_xml(f'<p:test cnt="0" {NSMAP}/>')
        constraint = AttributeValueRangeConstraint(
            attr_name="cnt", min_value=0, min_inclusive=False, max_value=10, max_inclusive=True
        )
        constraint.validate(ctx, element)
        assert not ctx.is_valid


class DescribeUniqueAttributeValueConstraint:
    def it_passes_for_unique_values(self):
        ctx = ValidationContext()
        element = parse_xml(
            f'<p:root {NSMAP}>'
            '  <p:item id="1"/>'
            '  <p:item id="2"/>'
            '  <p:item id="3"/>'
            '</p:root>'
        )
        item = element[0]
        constraint = UniqueAttributeValueConstraint(attr_name="id")
        constraint.validate(ctx, item)
        assert ctx.is_valid

    def it_detects_duplicate_values(self):
        ctx = ValidationContext()
        element = parse_xml(
            f'<p:root {NSMAP}>'
            '  <p:item id="1"/>'
            '  <p:item id="2"/>'
            '  <p:item id="1"/>'
            '</p:root>'
        )
        first_item = element[0]
        constraint = UniqueAttributeValueConstraint(attr_name="id")
        constraint.validate(ctx, first_item)
        assert not ctx.is_valid

    def it_skips_missing_attribute(self):
        ctx = ValidationContext()
        element = parse_xml(
            f'<p:root {NSMAP}>'
            '  <p:item/>'
            '  <p:item/>'
            '</p:root>'
        )
        item = element[0]
        constraint = UniqueAttributeValueConstraint(attr_name="id")
        constraint.validate(ctx, item)
        assert ctx.is_valid


class DescribeReferenceExistConstraint:
    def it_skips_when_no_rId_attr(self):
        ctx = ValidationContext()
        element = parse_xml(f'<p:test {NSMAP}/>')
        constraint = ReferenceExistConstraint(ref_attr_name="r:id")
        constraint.validate(ctx, element)
        assert ctx.is_valid

    def it_skips_when_no_part_context(self):
        ctx = ValidationContext()
        element = parse_xml(f'<p:test r:id="rId2" {NSMAP}/>')
        constraint = ReferenceExistConstraint(ref_attr_name="r:id")
        constraint.validate(ctx, element)
        assert ctx.is_valid


class DescribeAttributePairConstraint:
    def it_passes_when_both_present(self):
        ctx = ValidationContext()
        element = parse_xml(f'<p:test attr1="a" attr2="b" {NSMAP}/>')
        constraint = AttributePairConstraint(attr1_name="attr1", attr2_name="attr2")
        constraint.validate(ctx, element)
        assert ctx.is_valid

    def it_passes_when_both_absent(self):
        ctx = ValidationContext()
        element = parse_xml(f'<p:test {NSMAP}/>')
        constraint = AttributePairConstraint(attr1_name="attr1", attr2_name="attr2")
        constraint.validate(ctx, element)
        assert ctx.is_valid

    def it_fails_when_only_one_present(self):
        ctx = ValidationContext()
        element = parse_xml(f'<p:test attr1="a" {NSMAP}/>')
        constraint = AttributePairConstraint(attr1_name="attr1", attr2_name="attr2")
        constraint.validate(ctx, element)
        assert not ctx.is_valid
