"""Table styles proxy classes."""

from __future__ import annotations

from pptx.oxml.tablestyle import CT_TableStyleList
from pptx.util import lazyproperty


class TableStyles:
    """Proxy for the table styles collection in a presentation.

    Provides access to the ``a:tblStyleLst`` element in the table styles part.
    """

    def __init__(self, tblStyleLst: CT_TableStyleList):
        super().__init__()
        self._tblStyleLst = tblStyleLst

    @property
    def default_style_id(self) -> str | None:
        """Return the GUID of the default table style, or |None|."""
        return self._tblStyleLst.get("def")

    @default_style_id.setter
    def default_style_id(self, value: str):
        self._tblStyleLst.set("def", value)
