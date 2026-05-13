"""High-level API for SmartArt diagram support.

Provides the |SmartArtData| proxy for manipulating SmartArt diagram data.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pptx_ng.oxml.diagram.core import CT_DataModel, CT_Point


class SmartArtData:
    """Proxy for a `dgm:dataModel` element providing high-level access
    to SmartArt diagram data.
    """

    def __init__(self, data_model: CT_DataModel, part=None):
        super(SmartArtData, self).__init__()
        self._element = self._data_model = data_model
        self._part = part

    def add_node(self, text: str = "") -> SmartArtNode:
        """Add a new node to this diagram and return its proxy."""
        ptLst = self._data_model.get_or_add_ptLst()
        model_id = self._next_id()
        pt = ptLst.add_point(model_id, text)
        return SmartArtNode(pt)

    @property
    def nodes(self) -> tuple[SmartArtNode, ...]:
        """Tuple of |SmartArtNode| proxies for the content nodes in this diagram."""
        ptLst = self._data_model.ptLst
        if ptLst is None:
            return ()
        return tuple(
            SmartArtNode(pt) for pt in ptLst.pt_lst if pt.type != "doc"
        )

    def _next_id(self) -> int:
        """Return the next available model ID."""
        ptLst = self._data_model.ptLst
        if ptLst is None:
            return 0
        max_id = -1
        for pt in ptLst.pt_lst:
            try:
                mid = int(pt.modelId or "0")
            except (ValueError, TypeError):
                mid = 0
            if mid > max_id:
                max_id = mid
        return max_id + 1


class SmartArtNode:
    """Proxy for a `dgm:pt` element representing a single node in a diagram."""

    def __init__(self, pt: CT_Point):
        self._pt = pt

    @property
    def model_id(self) -> str | None:
        return self._pt.modelId

    @property
    def text(self) -> str:
        t_el = self._pt.find("{http://schemas.openxmlformats.org/drawingml/2006/diagram}t")
        if t_el is None:
            return ""
        paragraphs = t_el.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}t")
        return "".join(p.text or "" for p in paragraphs)
