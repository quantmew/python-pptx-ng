"""Presentation part, the main part in a .pptx package."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, Iterable

from pptx_ng.opc.constants import RELATIONSHIP_TYPE as RT
from pptx_ng.opc.package import XmlPart
from pptx_ng.opc.packuri import PackURI
from pptx_ng.parts.comment import CommentAuthorsPart
from pptx_ng.parts.handout import HandoutMasterPart
from pptx_ng.parts.presprops import PresentationPropertiesPart, ViewPropertiesPart
from pptx_ng.parts.slide import NotesMasterPart, SlidePart
from pptx_ng.parts.tablestyle import TableStylesPart
from pptx_ng.presentation import Presentation
from pptx_ng.util import lazyproperty

if TYPE_CHECKING:
    from pptx_ng.parts.coreprops import CorePropertiesPart
    from pptx_ng.slide import NotesMaster, Slide, SlideLayout, SlideMaster


class PresentationPart(XmlPart):
    """Top level class in object model.

    Represents the contents of the /ppt directory of a .pptx file.
    """

    def add_slide(self, slide_layout: SlideLayout):
        """Return (rId, slide) pair of a newly created blank slide.

        New slide inherits appearance from `slide_layout`.
        """
        partname = self._next_slide_partname
        slide_layout_part = slide_layout.part
        slide_part = SlidePart.new(partname, self.package, slide_layout_part)
        rId = self.relate_to(slide_part, RT.SLIDE)
        return rId, slide_part.slide

    @property
    def core_properties(self) -> CorePropertiesPart:
        """A |CoreProperties| object for the presentation.

        Provides read/write access to the Dublin Core properties of this presentation.
        """
        return self.package.core_properties

    def get_slide(self, slide_id: int) -> Slide | None:
        """Return optional related |Slide| object identified by `slide_id`.

        Returns |None| if no slide with `slide_id` is related to this presentation.
        """
        for sldId in self._element.sldIdLst:
            if sldId.id == slide_id:
                return self.related_part(sldId.rId).slide
        return None

    @lazyproperty
    def notes_master(self) -> NotesMaster:
        """
        Return the |NotesMaster| object for this presentation. If the
        presentation does not have a notes master, one is created from
        a default template. The same single instance is returned on each
        call.
        """
        return self.notes_master_part.notes_master

    @lazyproperty
    def notes_master_part(self) -> NotesMasterPart:
        """Return the |NotesMasterPart| object for this presentation.

        If the presentation does not have a notes master, one is created from a default template.
        The same single instance is returned on each call.
        """
        try:
            return self.part_related_by(RT.NOTES_MASTER)
        except KeyError:
            notes_master_part = NotesMasterPart.create_default(self.package)
            self.relate_to(notes_master_part, RT.NOTES_MASTER)
            return notes_master_part

    @lazyproperty
    def handout_master_part(self) -> HandoutMasterPart:
        """Return the |HandoutMasterPart| object for this presentation.

        If the presentation does not have a handout master, one is created from a default template.
        """
        try:
            return self.part_related_by(RT.HANDOUT_MASTER)
        except KeyError:
            handout_master_part = HandoutMasterPart.create_default(self.package)
            rId = self.relate_to(handout_master_part, RT.HANDOUT_MASTER)
            handoutMasterIdLst = self._element.get_or_add_handoutMasterIdLst()
            handoutMasterIdLst._add_handoutMasterId(rId=rId)
            return handout_master_part

    @lazyproperty
    def comment_authors_part(self) -> CommentAuthorsPart:
        """Return the |CommentAuthorsPart| for this presentation.

        If the presentation does not have a comment authors part, one is created.
        """
        try:
            return self.part_related_by(RT.COMMENT_AUTHORS)
        except KeyError:
            comment_authors_part = CommentAuthorsPart.new(self.package)
            self.relate_to(comment_authors_part, RT.COMMENT_AUTHORS)
            return comment_authors_part

    @lazyproperty
    def table_styles_part(self) -> TableStylesPart:
        """Return the |TableStylesPart| for this presentation.

        If the presentation does not have one, it is created.
        """
        try:
            return self.part_related_by(RT.TABLE_STYLES)
        except KeyError:
            table_styles_part = TableStylesPart.new(self.package)
            self.relate_to(table_styles_part, RT.TABLE_STYLES)
            return table_styles_part

    @lazyproperty
    def presentation(self):
        """
        A |Presentation| object providing access to the content of this
        presentation.
        """
        return Presentation(self._element, self)

    @lazyproperty
    def pres_props_part(self) -> PresentationPropertiesPart:
        """Return the PresentationPropertiesPart for this presentation.

        If the presentation does not have one, it is created.
        """
        try:
            return self.part_related_by(RT.PRES_PROPS)
        except KeyError:
            pres_props_part = PresentationPropertiesPart.new(self.package)
            self.relate_to(pres_props_part, RT.PRES_PROPS)
            return pres_props_part

    @lazyproperty
    def view_props_part(self) -> ViewPropertiesPart:
        """Return the ViewPropertiesPart for this presentation.

        If the presentation does not have one, it is created.
        """
        try:
            return self.part_related_by(RT.VIEW_PROPS)
        except KeyError:
            view_props_part = ViewPropertiesPart.new(self.package)
            self.relate_to(view_props_part, RT.VIEW_PROPS)
            return view_props_part

    def related_slide(self, rId: str) -> Slide:
        """Return |Slide| object for related |SlidePart| related by `rId`."""
        return self.related_part(rId).slide

    def related_slide_master(self, rId: str) -> SlideMaster:
        """Return |SlideMaster| object for |SlideMasterPart| related by `rId`."""
        return self.related_part(rId).slide_master

    def rename_slide_parts(self, rIds: Iterable[str]):
        """Assign incrementing partnames to the slide parts identified by `rIds`.

        Partnames are like `/ppt/slides/slide9.xml` and are assigned in the order their id appears
        in the `rIds` sequence. The name portion is always `slide`. The number part forms a
        continuous sequence starting at 1 (e.g. 1, 2, ... 10, ...). The extension is always
        `.xml`.
        """
        for idx, rId in enumerate(rIds):
            slide_part = self.related_part(rId)
            slide_part.partname = PackURI("/ppt/slides/slide%d.xml" % (idx + 1))

    def save(self, path_or_stream: str | IO[bytes]):
        """Save this presentation package to `path_or_stream`.

        `path_or_stream` can be either a path to a filesystem location (a string) or a
        file-like object.
        """
        self.package.save(path_or_stream)

    def slide_id(self, slide_part):
        """Return the slide-id associated with `slide_part`."""
        for sldId in self._element.sldIdLst:
            if self.related_part(sldId.rId) is slide_part:
                return sldId.id
        raise ValueError("matching slide_part not found")

    @property
    def _next_slide_partname(self):
        """Return |PackURI| instance containing next available slide partname."""
        sldIdLst = self._element.get_or_add_sldIdLst()
        partname_str = "/ppt/slides/slide%d.xml" % (len(sldIdLst) + 1)
        return PackURI(partname_str)
