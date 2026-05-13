"""Unit-test suite for the pptx.validation.context module."""

from __future__ import annotations

import pytest

from pptx.enum.validation import ValidationErrorType
from pptx.validation.context import ValidationErrorInfo, ValidationContext


class DescribeValidationErrorInfo:
    def it_stores_error_properties(self):
        error = ValidationErrorInfo(
            id="Sch_Test",
            error_type=ValidationErrorType.SCHEMA,
            description="Test error description",
        )
        assert error.id == "Sch_Test"
        assert error.error_type == ValidationErrorType.SCHEMA
        assert error.description == "Test error description"
        assert error.part is None
        assert error.element is None
        assert error.related_element is None
        assert error.related_part is None

    def it_has_a_repr(self):
        error = ValidationErrorInfo(
            id="Sem_Test",
            error_type=ValidationErrorType.SEMANTIC,
            description="test",
        )
        assert "Sem_Test" in repr(error)
        assert "SEMANTIC" in repr(error)

    def it_has_a_str_including_part_info(self):
        error = ValidationErrorInfo(
            id="Pkg_Test",
            error_type=ValidationErrorType.PACKAGE,
            description="test error",
        )
        s = str(error)
        assert "PACKAGE" in s
        assert "test error" in s


class DescribeValidationContext:
    def it_starts_empty(self):
        ctx = ValidationContext()
        assert ctx.is_valid
        assert ctx.error_count == 0
        assert ctx.errors == []

    def it_tracks_errors(self):
        ctx = ValidationContext()
        ctx.add_error(
            ValidationErrorInfo(
                id="Test",
                error_type=ValidationErrorType.SCHEMA,
                description="test error",
            )
        )
        assert not ctx.is_valid
        assert ctx.error_count == 1

    def it_respects_max_errors(self):
        ctx = ValidationContext(max_errors=2)
        ctx.add_error(
            ValidationErrorInfo(
                id="Test1",
                error_type=ValidationErrorType.SCHEMA,
                description="error 1",
            )
        )
        ctx.add_error(
            ValidationErrorInfo(
                id="Test2",
                error_type=ValidationErrorType.SCHEMA,
                description="error 2",
            )
        )
        ctx.add_error(
            ValidationErrorInfo(
                id="Test3",
                error_type=ValidationErrorType.SCHEMA,
                description="error 3",
            )
        )
        assert ctx.error_count == 2

    def it_clears_errors(self):
        ctx = ValidationContext()
        ctx.add_error(
            ValidationErrorInfo(
                id="Test",
                error_type=ValidationErrorType.SCHEMA,
                description="test",
            )
        )
        ctx.clear()
        assert ctx.is_valid

    def it_creates_schema_errors(self):
        ctx = ValidationContext()
        ctx.create_schema_error(id="Sch_Test", description="schema error")
        assert ctx.error_count == 1
        assert ctx.errors[0].error_type == ValidationErrorType.SCHEMA
        assert ctx.errors[0].id == "Sch_Test"

    def it_creates_semantic_errors(self):
        ctx = ValidationContext()
        ctx.create_semantic_error(id="Sem_Test", description="semantic error")
        assert ctx.error_count == 1
        assert ctx.errors[0].error_type == ValidationErrorType.SEMANTIC

    def it_creates_package_errors(self):
        ctx = ValidationContext()
        ctx.create_package_error(id="Pkg_Test", description="package error")
        assert ctx.error_count == 1
        assert ctx.errors[0].error_type == ValidationErrorType.PACKAGE

    def it_rejects_negative_max_errors(self):
        ctx = ValidationContext()
        with pytest.raises(ValueError):
            ctx.max_errors = -1
