"""HandoutMaster proxy class."""

from __future__ import annotations

from pptx_ng.shared import PartElementProxy


class HandoutMaster(PartElementProxy):
    """Proxy for the handout master part.

    Provides access to the shapes and properties of the handout master.
    """

    @property
    def shapes(self):
        """Return the shape collection for this handout master."""
        from pptx_ng.shapes.shapetree import MasterShapes

        return MasterShapes(self._element.cSld.spTree, self)
