"""Validation framework for python-pptx presentations.

Provides the PresentationValidator class for validating .pptx files against
Open XML schema and semantic constraints.

Usage::

    from pptx import Presentation
    from pptx.validation import PresentationValidator

    prs = Presentation('test.pptx')
    validator = PresentationValidator()
    errors = validator.validate(prs)
    for error in errors:
        print(error)

Corresponds to DocumentFormat.OpenXml.Validation.OpenXmlValidator in the
Open XML SDK.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.validation.context import ValidationErrorInfo, ValidationContext
from pptx.validation.package_validator import PackageValidator
from pptx.validation.schema import SchemaValidator
from pptx.validation.semantic import SemanticValidator

if TYPE_CHECKING:
    from pptx import Presentation


class PresentationValidator:
    """Validates a Presentation against schema, semantic, and package constraints.

    Corresponds to DocumentFormat.OpenXml.Validation.OpenXmlValidator.

    The validation flow mirrors the Open XML SDK:
    1. Package structure validation (required parts, relationship integrity)
    2. Schema validation (element cardinality, attribute presence/values)
    3. Semantic validation (value ranges, uniqueness, reference integrity)
    """

    def __init__(self, max_errors: int = 1000):
        if max_errors < 0:
            raise ValueError("max_errors must be >= 0")
        self._max_errors = max_errors
        self._schema_validator = SchemaValidator()
        self._package_validator = PackageValidator()
        self._semantic_validator = SemanticValidator()

    @property
    def max_errors(self) -> int:
        return self._max_errors

    @max_errors.setter
    def max_errors(self, value: int):
        if value < 0:
            raise ValueError("max_errors must be >= 0")
        self._max_errors = value

    def validate(self, presentation: Presentation) -> list[ValidationErrorInfo]:
        """Validate a Presentation and return a list of errors."""
        ctx = ValidationContext(max_errors=self._max_errors)

        package = presentation.part.package

        self._package_validator.validate(ctx, package)

        for part in package.iter_parts():
            element = getattr(part, "_element", None)
            if element is not None:
                self._schema_validator.validate_element(ctx, element, part)
                self._semantic_validator.validate_element(ctx, element, part)

        return ctx.errors

    def validate_part(self, part) -> list[ValidationErrorInfo]:
        """Validate a single part."""
        ctx = ValidationContext(max_errors=self._max_errors)

        element = getattr(part, "_element", None)
        if element is not None:
            self._schema_validator.validate_element(ctx, element, part)
            self._semantic_validator.validate_element(ctx, element, part)

        return ctx.errors

    def validate_element(self, element) -> list[ValidationErrorInfo]:
        """Validate a single XML element (schema + semantic)."""
        ctx = ValidationContext(max_errors=self._max_errors)

        self._schema_validator.validate_element(ctx, element)
        self._semantic_validator.validate_element(ctx, element)

        return ctx.errors

    def register_semantic_constraint(self, constraint):
        """Register a custom semantic constraint."""
        self._semantic_validator.register(constraint)


__all__ = [
    "PresentationValidator",
    "ValidationErrorInfo",
    "ValidationContext",
]
