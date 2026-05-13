"""Semantic validation constraints for Open XML elements.

Provides constraint classes that check semantic correctness beyond schema
structure: value ranges, uniqueness, reference integrity, etc.

Corresponds to DocumentFormat.OpenXml.Validation.Semantic constraints in
the Open XML SDK.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.enum.validation import SemanticValidationLevel, ValidationErrorType
from pptx.validation.context import ValidationErrorInfo, ValidationContext

if TYPE_CHECKING:
    from lxml.etree import _Element

    from pptx.opc.package import Part


class SemanticConstraint:
    """Base class for semantic validation constraints.

    Corresponds to DocumentFormat.OpenXml.Validation.Semantic.SemanticConstraint.
    """

    def __init__(self, level: SemanticValidationLevel = SemanticValidationLevel.ELEMENT):
        self._level = level

    @property
    def level(self) -> SemanticValidationLevel:
        return self._level

    def validate(self, ctx: ValidationContext, element: _Element, part: Part | None = None):
        err = self.validate_core(ctx, element, part)
        if err is not None:
            ctx.add_error(err)

    def validate_core(
        self,
        ctx: ValidationContext,
        element: _Element,
        part: Part | None = None,
    ) -> ValidationErrorInfo | None:
        raise NotImplementedError


class AttributeCannotOmitConstraint(SemanticConstraint):
    """Ensures a required attribute is present on the element.

    Corresponds to DocumentFormat.OpenXml.Validation.Semantic.AttributeCannotOmitConstraint.
    """

    def __init__(self, attr_name: str):
        super().__init__(SemanticValidationLevel.ELEMENT)
        self._attr_name = attr_name

    def validate_core(self, ctx, element, part=None):
        from pptx.oxml.ns import qn

        clark_name = qn(self._attr_name) if ":" in self._attr_name else self._attr_name
        value = element.get(clark_name)
        if value is None:
            return ValidationErrorInfo(
                id="Sem_MissRequiredAttribute",
                error_type=ValidationErrorType.SCHEMA,
                description=f"Required attribute '{self._attr_name}' is missing.",
                part=part,
                element=element,
            )
        return None


class AttributeValueRangeConstraint(SemanticConstraint):
    """Validates that a numeric attribute value falls within a range.

    Corresponds to DocumentFormat.OpenXml.Validation.Semantic.AttributeValueRangeConstraint.
    """

    def __init__(
        self,
        attr_name: str,
        is_valid_range: bool = True,
        min_value: float = 0.0,
        min_inclusive: bool = True,
        max_value: float = 0.0,
        max_inclusive: bool = True,
    ):
        super().__init__(SemanticValidationLevel.ELEMENT)
        self._attr_name = attr_name
        self._is_valid_range = is_valid_range
        self._min_value = min_value
        self._min_inclusive = min_inclusive
        self._max_value = max_value
        self._max_inclusive = max_inclusive

    def validate_core(self, ctx, element, part=None):
        from pptx.oxml.ns import qn

        clark_name = qn(self._attr_name) if ":" in self._attr_name else self._attr_name
        attr_value = element.get(clark_name)

        if attr_value is None or attr_value == "":
            return None

        try:
            value = float(attr_value)
        except (ValueError, TypeError):
            return None

        sub_msg = None

        if self._min_inclusive:
            if not (self._min_value <= value):
                sub_msg = f"must be >= {self._min_value}"
        else:
            if not (self._min_value < value):
                sub_msg = f"must be > {self._min_value}"

        if sub_msg is None:
            if self._max_inclusive:
                if not (value <= self._max_value):
                    sub_msg = f"must be <= {self._max_value}"
            else:
                if not (value < self._max_value):
                    sub_msg = f"must be < {self._max_value}"

        if sub_msg is None:
            return None

        return ValidationErrorInfo(
            id="Sem_AttributeValueDataTypeDetailed",
            error_type=ValidationErrorType.SCHEMA,
            description=f"Attribute '{self._attr_name}' value '{attr_value}' "
            f"is invalid: {sub_msg}.",
            part=part,
            element=element,
        )


class UniqueAttributeValueConstraint(SemanticConstraint):
    """Validates that an attribute value is unique across all elements of a
    specified type within a scope.

    Corresponds to DocumentFormat.OpenXml.Validation.Semantic.UniqueAttributeValueConstraint.
    """

    def __init__(
        self,
        attr_name: str,
        case_sensitive: bool = True,
        parent_tag: str | None = None,
    ):
        super().__init__(SemanticValidationLevel.PART)
        self._attr_name = attr_name
        self._case_sensitive = case_sensitive
        self._parent_tag = parent_tag

    def validate_core(self, ctx, element, part=None):
        from pptx.oxml.ns import qn

        clark_name = qn(self._attr_name) if ":" in self._attr_name else self._attr_name
        attr_value = element.get(clark_name)

        if attr_value is None or attr_value == "":
            return None

        scope_root = self._find_scope_root(element)
        if scope_root is None:
            return None

        element_tag = element.tag
        seen_values: set[str] = set()
        is_duplicate = False

        for child in scope_root.iter(element_tag):
            val = child.get(clark_name)
            if val is None or val == "":
                continue
            compare_val = val if self._case_sensitive else val.lower()
            if compare_val in seen_values:
                if val == attr_value or (
                    not self._case_sensitive and val.lower() == attr_value.lower()
                ):
                    is_duplicate = True
                    break
            seen_values.add(compare_val)

        if not is_duplicate:
            return None

        return ValidationErrorInfo(
            id="Sem_UniqueAttributeValue",
            error_type=ValidationErrorType.SEMANTIC,
            description=f"Attribute '{self._attr_name}' value '{attr_value}' "
            f"is not unique within the scope.",
            part=part,
            element=element,
        )

    def _find_scope_root(self, element: _Element):
        if self._parent_tag is None:
            root = element
            while root.getparent() is not None:
                root = root.getparent()
            return root

        from pptx.oxml.ns import qn

        parent_clark = qn(self._parent_tag)
        current = element.getparent()
        while current is not None:
            if current.tag == parent_clark:
                return current
            current = current.getparent()
        return None


class ReferenceExistConstraint(SemanticConstraint):
    """Validates that a relationship reference (rId) resolves to an existing
    part with the referenced element.

    Corresponds to DocumentFormat.OpenXml.Validation.Semantic.ReferenceExistConstraint.
    """

    def __init__(
        self,
        ref_attr_name: str = "r:id",
        element_tag: str | None = None,
        target_attr_name: str | None = None,
    ):
        super().__init__(SemanticValidationLevel.PACKAGE)
        self._ref_attr_name = ref_attr_name
        self._element_tag = element_tag
        self._target_attr_name = target_attr_name

    def validate_core(self, ctx, element, part=None):
        from pptx.oxml.ns import qn

        clark_name = qn(self._ref_attr_name) if ":" in self._ref_attr_name else self._ref_attr_name
        ref_value = element.get(clark_name)

        if ref_value is None or ref_value == "":
            return None

        if part is None:
            return None

        try:
            part.related_part(ref_value)
        except KeyError:
            return ValidationErrorInfo(
                id="Sem_MissingReferenceElement",
                error_type=ValidationErrorType.SEMANTIC,
                description=f"Referenced part with rId='{ref_value}' not found.",
                part=part,
                element=element,
            )

        return None


class AttributePairConstraint(SemanticConstraint):
    """Validates that two attributes must appear as a pair — either both
    present or both absent.

    Corresponds to DocumentFormat.OpenXml.Validation.Semantic.AttributePairConstraint.
    """

    def __init__(self, attr1_name: str, attr2_name: str):
        super().__init__(SemanticValidationLevel.ELEMENT)
        self._attr1_name = attr1_name
        self._attr2_name = attr2_name

    def validate_core(self, ctx, element, part=None):
        from pptx.oxml.ns import qn

        clark1 = qn(self._attr1_name) if ":" in self._attr1_name else self._attr1_name
        clark2 = qn(self._attr2_name) if ":" in self._attr2_name else self._attr2_name

        attr1_exists = element.get(clark1) is not None
        attr2_exists = element.get(clark2) is not None

        if attr1_exists != attr2_exists:
            return ValidationErrorInfo(
                id="Sem_AttributePairConstraint",
                error_type=ValidationErrorType.SEMANTIC,
                description=f"Attributes '{self._attr1_name}' and "
                f"'{self._attr2_name}' must appear as a pair.",
                part=part,
                element=element,
            )

        return None
