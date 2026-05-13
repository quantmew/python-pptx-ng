"""Schema validation for XML elements.

Validates child element cardinality and attribute values against the declarative
constraints defined on BaseOxmlElement subclasses.

Corresponds to DocumentFormat.OpenXml.Validation.Schema.SchemaTypeValidator
in the Open XML SDK.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml.etree import _Element

from pptx_ng.oxml.ns import qn
from pptx_ng.oxml.xmlchemy import (
    BaseAttribute,
    BaseOxmlElement,
    OneAndOnlyOne,
    OneOrMore,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOneChoice,
)
from pptx_ng.validation.context import ValidationContext

if TYPE_CHECKING:
    from pptx_ng.opc.package import Part


class SchemaValidator:
    """Validates XML elements against their declared schema constraints.

    Walks the element tree recursively, checking child cardinality
    (OneAndOnlyOne, OneOrMore, ZeroOrOneChoice) and attribute
    presence/values (RequiredAttribute, OptionalAttribute).
    """

    def validate_element(
        self,
        ctx: ValidationContext,
        element: _Element,
        part: Part | None = None,
    ):
        if not isinstance(element, BaseOxmlElement):
            return

        self._validate_child_cardinality(ctx, element, part)
        self._validate_attributes(ctx, element, part)

        for child in element:
            self.validate_element(ctx, child, part)

    # -- child cardinality checks -------------------------------------------

    def _validate_child_cardinality(
        self,
        ctx: ValidationContext,
        element: BaseOxmlElement,
        part: Part | None,
    ):
        self._check_required_children(ctx, element, part)
        self._check_one_or_more_children(ctx, element, part)
        self._check_choice_exclusivity(ctx, element, part)

    def _check_required_children(
        self,
        ctx: ValidationContext,
        element: BaseOxmlElement,
        part: Part | None,
    ):
        tag = element.tag
        for cls in type(element).__mro__:
            if cls is BaseOxmlElement or cls is _Element:
                break
            for prop_name, value in cls.__dict__.items():
                if not isinstance(value, OneAndOnlyOne):
                    continue
                nsptagname = value._nsptagname
                if element.find(qn(nsptagname)) is None:
                    ctx.create_schema_error(
                        id="Sch_MissingRequiredChild",
                        description=f"The required child element <{nsptagname}> "
                        f"is missing from <{_short_tag(tag)}>.",
                        part=part,
                        element=element,
                    )

    def _check_one_or_more_children(
        self,
        ctx: ValidationContext,
        element: BaseOxmlElement,
        part: Part | None,
    ):
        tag = element.tag
        for cls in type(element).__mro__:
            if cls is BaseOxmlElement or cls is _Element:
                break
            for prop_name, value in cls.__dict__.items():
                if not isinstance(value, OneOrMore):
                    continue
                nsptagname = value._nsptagname
                if not element.findall(qn(nsptagname)):
                    ctx.create_schema_error(
                        id="Sch_MissingRequiredChild",
                        description=f"At least one <{nsptagname}> child is "
                        f"required in <{_short_tag(tag)}>, but none found.",
                        part=part,
                        element=element,
                    )

    def _check_choice_exclusivity(
        self,
        ctx: ValidationContext,
        element: BaseOxmlElement,
        part: Part | None,
    ):
        tag = element.tag
        for cls in type(element).__mro__:
            if cls is BaseOxmlElement or cls is _Element:
                break
            for prop_name, value in cls.__dict__.items():
                if not isinstance(value, ZeroOrOneChoice):
                    continue
                present = []
                for choice in value._choices:
                    if element.find(qn(choice.nsptagname)) is not None:
                        present.append(choice.nsptagname)
                if len(present) > 1:
                    ctx.create_schema_error(
                        id="Sch_MultipleChoiceMembers",
                        description=f"Choice group in <{_short_tag(tag)}> has "
                        f"multiple members: {present}. At most one is allowed.",
                        part=part,
                        element=element,
                    )

    # -- attribute checks ---------------------------------------------------

    def _validate_attributes(
        self,
        ctx: ValidationContext,
        element: BaseOxmlElement,
        part: Part | None,
    ):
        for cls in type(element).__mro__:
            if cls is BaseOxmlElement or cls is _Element:
                break
            for prop_name, value in cls.__dict__.items():
                if isinstance(value, RequiredAttribute):
                    self._validate_required_attr(ctx, element, value, part)
                elif isinstance(value, OptionalAttribute):
                    self._validate_optional_attr(ctx, element, value, part)

    def _validate_required_attr(
        self,
        ctx: ValidationContext,
        element: BaseOxmlElement,
        attr_decl: RequiredAttribute,
        part: Part | None,
    ):
        attr_value = element.get(attr_decl._clark_name)
        if attr_value is None:
            ctx.create_schema_error(
                id="Sch_MissRequiredAttribute",
                description=f"The required attribute '{attr_decl._attr_name}' "
                f"is missing on <{_short_tag(element.tag)}>.",
                part=part,
                element=element,
            )
            return
        self._validate_attr_value(ctx, element, attr_decl, attr_value, part)

    def _validate_optional_attr(
        self,
        ctx: ValidationContext,
        element: BaseOxmlElement,
        attr_decl: OptionalAttribute,
        part: Part | None,
    ):
        attr_value = element.get(attr_decl._clark_name)
        if attr_value is None:
            return
        self._validate_attr_value(ctx, element, attr_decl, attr_value, part)

    def _validate_attr_value(
        self,
        ctx: ValidationContext,
        element: BaseOxmlElement,
        attr_decl: BaseAttribute,
        attr_value: str,
        part: Part | None,
    ):
        simple_type = attr_decl._simple_type
        attr_name = attr_decl._attr_name
        try:
            simple_type.from_xml(attr_value)
        except (ValueError, TypeError) as e:
            ctx.create_schema_error(
                id="Sch_AttributeValueDataTypeDetailed",
                description=f"Attribute '{attr_name}' on <{_short_tag(element.tag)}> "
                f"has invalid value '{attr_value}': {e}",
                part=part,
                element=element,
            )
        except Exception:
            pass


def _short_tag(tag: str) -> str:
    """Strip namespace from a Clark-notation tag for display."""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag
