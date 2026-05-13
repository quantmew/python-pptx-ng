"""Default semantic constraints for PPTX validation.

Provides a set of commonly-needed constraints for validating PowerPoint
presentations. These are automatically registered by PresentationValidator.
"""

from __future__ import annotations

from pptx_ng.validation.constraints import (
    SemanticConstraint,
    UniqueAttributeValueConstraint,
)


def get_default_constraints() -> list[SemanticConstraint]:
    """Return the default set of semantic constraints for PPTX validation."""
    return [
        # r:id must be unique within sldIdLst (slide references)
        UniqueAttributeValueConstraint("r:id", parent_tag="p:sldIdLst"),
        # r:id must be unique within sldLayoutIdLst
        UniqueAttributeValueConstraint("r:id", parent_tag="p:sldLayoutIdLst"),
        # r:id must be unique within sldMasterIdLst
        UniqueAttributeValueConstraint("r:id", parent_tag="p:sldMasterIdLst"),
    ]
