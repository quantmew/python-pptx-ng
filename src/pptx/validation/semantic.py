"""Semantic validation engine.

Applies semantic constraints to elements, checking correctness beyond
schema structure.

Corresponds to DocumentFormat.OpenXml.Validation.DocumentValidator's
semantic validation pass in the Open XML SDK.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml.etree import _Element

from pptx.oxml.xmlchemy import BaseOxmlElement
from pptx.validation.constraints import SemanticConstraint
from pptx.validation.context import ValidationContext

if TYPE_CHECKING:
    from pptx.opc.package import Part


class SemanticValidator:
    """Applies registered semantic constraints during validation."""

    def __init__(self):
        self._constraints: list[SemanticConstraint] = []

    def register(self, constraint: SemanticConstraint):
        self._constraints.append(constraint)

    def validate_element(
        self,
        ctx: ValidationContext,
        element: _Element,
        part: Part | None = None,
    ):
        if not isinstance(element, BaseOxmlElement):
            return

        for constraint in self._constraints:
            constraint.validate(ctx, element, part)

        for child in element:
            self.validate_element(ctx, child, part)
