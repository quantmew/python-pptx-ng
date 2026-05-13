"""Enumerations used for validation of presentations and their parts."""

from __future__ import annotations

from pptx_ng.enum.base import BaseEnum


class ValidationErrorType(BaseEnum):
    """The type of the validation error.

    Corresponds to DocumentFormat.OpenXml.Validation.ValidationErrorType
    in the Open XML SDK.
    """

    SCHEMA = (
        1,
        "Schema validation error.",
    )
    SEMANTIC = (
        2,
        "Semantic validation error.",
    )
    PACKAGE = (
        3,
        "Package structure validation error.",
    )
    MARKUP_COMPATIBILITY = (
        4,
        "Markup Compatibility validation error.",
    )


class SemanticValidationLevel(BaseEnum):
    """Scope at which a semantic constraint is evaluated.

    Corresponds to DocumentFormat.OpenXml.Validation.Semantic.SemanticValidationLevel
    in the Open XML SDK.
    """

    ELEMENT = (1, "Constraint evaluated at the element level.")
    PART = (2, "Constraint evaluated at the part level.")
    PACKAGE = (3, "Constraint evaluated at the package level.")
