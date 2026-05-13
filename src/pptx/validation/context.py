"""Validation context and error info classes.

Corresponds to DocumentFormat.OpenXml.Validation.ValidationContext and
ValidationErrorInfo in the Open XML SDK.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.enum.validation import ValidationErrorType

if TYPE_CHECKING:
    from lxml.etree import _Element

    from pptx.opc.package import Part


class ValidationErrorInfo:
    """An individual validation finding with location and description.

    Corresponds to DocumentFormat.OpenXml.Validation.ValidationErrorInfo.
    """

    def __init__(
        self,
        id: str,
        error_type: ValidationErrorType,
        description: str,
        part: Part | None = None,
        element: _Element | None = None,
        related_element: _Element | None = None,
        related_part: Part | None = None,
    ):
        self.id = id
        self.error_type = error_type
        self.description = description
        self.part = part
        self.element = element
        self.related_element = related_element
        self.related_part = related_part

    def __repr__(self):
        return (
            f"ValidationErrorInfo(id={self.id!r}, "
            f"error_type={self.error_type!r}, "
            f"description={self.description!r})"
        )

    def __str__(self):
        parts = [f"[{self.error_type.name}] {self.description}"]
        if self.part:
            parts.append(f"  Part: {self.part.partname}")
        if self.element is not None:
            tag = self.element.tag
            if isinstance(tag, str) and "}" in tag:
                tag = tag.split("}", 1)[1]
            parts.append(f"  Element: <{tag}>")
        return "\n".join(parts)


class ValidationContext:
    """Collects validation errors during a validation run.

    Corresponds to DocumentFormat.OpenXml.Validation.ValidationContext.
    """

    def __init__(self, max_errors: int = 1000):
        self._errors: list[ValidationErrorInfo] = []
        self._max_errors = max_errors

    @property
    def errors(self) -> list[ValidationErrorInfo]:
        """All validation errors collected so far."""
        return list(self._errors)

    @property
    def error_count(self) -> int:
        return len(self._errors)

    @property
    def is_valid(self) -> bool:
        return self.error_count == 0

    @property
    def max_errors(self) -> int:
        return self._max_errors

    @max_errors.setter
    def max_errors(self, value: int):
        if value < 0:
            raise ValueError("max_errors must be >= 0")
        self._max_errors = value

    def add_error(self, error: ValidationErrorInfo) -> None:
        """Add a validation error, respecting the max_errors limit."""
        if error is not None and not self._is_full():
            self._errors.append(error)

    def _is_full(self) -> bool:
        return self._max_errors > 0 and self.error_count >= self._max_errors

    def clear(self) -> None:
        self._errors.clear()

    def create_schema_error(
        self,
        id: str,
        description: str,
        part: Part | None = None,
        element: _Element | None = None,
    ) -> None:
        self.add_error(
            ValidationErrorInfo(
                id=id,
                error_type=ValidationErrorType.SCHEMA,
                description=description,
                part=part,
                element=element,
            )
        )

    def create_semantic_error(
        self,
        id: str,
        description: str,
        part: Part | None = None,
        element: _Element | None = None,
        related_element: _Element | None = None,
        related_part: Part | None = None,
    ) -> None:
        self.add_error(
            ValidationErrorInfo(
                id=id,
                error_type=ValidationErrorType.SEMANTIC,
                description=description,
                part=part,
                element=element,
                related_element=related_element,
                related_part=related_part,
            )
        )

    def create_package_error(
        self,
        id: str,
        description: str,
        part: Part | None = None,
        related_part: Part | None = None,
    ) -> None:
        self.add_error(
            ValidationErrorInfo(
                id=id,
                error_type=ValidationErrorType.PACKAGE,
                description=description,
                part=part,
                related_part=related_part,
            )
        )
