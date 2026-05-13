"""Custom show proxy classes.

Provides high-level API for working with custom shows in a presentation.
"""

from __future__ import annotations

from typing import Iterator

from pptx.oxml.customshow import CT_CustomShow, CT_CustomShowList, CT_SlideRelationshipList
from pptx.oxml.ns import qn
from pptx.shared import PartElementProxy


class CustomShows(PartElementProxy):
    """Collection of custom shows in a presentation.

    Accessed via `Presentation.custom_shows`.
    """

    _element: CT_CustomShowList

    def __iter__(self) -> Iterator[CustomShow]:
        for custShow in self._element.custShow_lst:
            yield CustomShow(custShow, self)

    def __len__(self) -> int:
        return len(self._element.custShow_lst)

    def add(self, name: str, slide_ids: list[str] | None = None) -> CustomShow:
        """Add a new custom show and return it.

        Args:
            name: Display name for the custom show.
            slide_ids: List of slide relationship IDs (rId strings) to include.
                If None, the custom show starts empty.
        """
        next_id = self._next_id
        custShow = self._element.add_custShow(id=next_id, name=name)
        if slide_ids:
            sldLst = custShow.sldLst
            for rId in slide_ids:
                sldLst.add_sld(rId)
        return CustomShow(custShow, self)

    def get(self, id: int) -> CustomShow | None:
        """Return the custom show with the given ID, or None."""
        for custShow in self._element.custShow_lst:
            if custShow.id == id:
                return CustomShow(custShow, self)
        return None

    def get_by_name(self, name: str) -> CustomShow | None:
        """Return the custom show with the given name, or None."""
        for custShow in self._element.custShow_lst:
            if custShow.name == name:
                return CustomShow(custShow, self)
        return None

    @property
    def _next_id(self) -> int:
        """Return the next available custom show ID."""
        if not self._element.custShow_lst:
            return 1
        return max(cs.id for cs in self._element.custShow_lst) + 1


class CustomShow:
    """A single custom show definition."""

    def __init__(self, custShow: CT_CustomShow, parent: CustomShows):
        self._element = custShow
        self._parent = parent

    @property
    def id(self) -> int:
        """Custom show ID."""
        return self._element.id

    @property
    def name(self) -> str:
        """Custom show display name."""
        return self._element.name

    @name.setter
    def name(self, value: str):
        self._element.name = value

    @property
    def slide_ids(self) -> list[str]:
        """List of slide relationship IDs (rId strings) in this custom show."""
        sldLst = self._element.sldLst
        return [sld.get(qn("r:id")) for sld in sldLst.sld_lst]

    def add_slide(self, rId: str):
        """Add a slide reference to this custom show by its relationship ID."""
        self._element.sldLst.add_sld(rId)

    def remove_slide(self, rId: str):
        """Remove a slide reference from this custom show by its relationship ID."""
        sldLst = self._element.sldLst
        for sld in sldLst.sld_lst:
            if sld.get(qn("r:id")) == rId:
                sldLst.remove(sld)
                return
        raise ValueError("slide rId '%s' not found in custom show" % rId)
