"""View properties proxy classes.

Provides high-level API for view-related properties like grid spacing.
"""

from __future__ import annotations

from pptx_ng.oxml.viewprops import CT_GridSpacing, CT_ViewProperties
from pptx_ng.shared import PartElementProxy
from pptx_ng.util import lazyproperty


class ViewProperties(PartElementProxy):
    """Presentation view properties.

    Accessed via `Presentation.view_properties`.
    """

    _element: CT_ViewProperties

    @property
    def last_view(self) -> str | None:
        """The last view type used.

        Read/write. Common values: 'sldView', 'outlineView', 'notesView',
        'handoutView', 'sldSorterView'. Returns |None| if not set.
        """
        return self._element.lastView

    @last_view.setter
    def last_view(self, value: str | None):
        self._element.lastView = value

    @property
    def show_comments(self) -> bool | None:
        """Whether to show comments in the view.

        Read/write. Returns |None| if not set.
        """
        return self._element.showComments

    @show_comments.setter
    def show_comments(self, value: bool | None):
        self._element.showComments = value

    @lazyproperty
    def grid_spacing(self) -> GridSpacing:
        """The |GridSpacing| for this view.

        If no grid spacing is defined, a default is created.
        """
        from pptx_ng.oxml.xmlchemy import OxmlElement

        gs = self._element.gridSpacing
        if gs is None:
            gs = OxmlElement("p:gridSpacing")
            gs.set("cx", "914400")
            gs.set("cy", "914400")
            self._element.append(gs)
        return GridSpacing(gs, self)


class GridSpacing:
    """Grid spacing settings for the view.

    Accessed via `ViewProperties.grid_spacing`.
    """

    def __init__(self, gs: CT_GridSpacing, parent: ViewProperties):
        self._element = gs
        self._parent = parent

    @property
    def cx(self) -> int:
        """Grid spacing width in EMU."""
        return self._element.cx

    @cx.setter
    def cx(self, value: int):
        self._element.cx = value

    @property
    def cy(self) -> int:
        """Grid spacing height in EMU."""
        return self._element.cy

    @cy.setter
    def cy(self, value: int):
        self._element.cy = value
