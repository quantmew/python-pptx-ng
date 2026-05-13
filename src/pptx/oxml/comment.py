"""Comment-related XML element classes.

Corresponds to CT_CommentAuthorList, CT_CommentAuthor, CT_CommentList, CT_Comment
in the Open XML SDK (schemas_openxmlformats_org_presentationml_2006_main).
"""

from __future__ import annotations

from typing import Callable

from lxml.etree import _Element

from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import (
    ST_Coordinate,
    XsdString,
    XsdUnsignedInt,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
)
from pptx.util import Emu


class CT_CommentAuthorList(BaseOxmlElement):
    """`p:cmAuthorLst` element — root element of the comment authors part.

    Contains zero or more `p:cmAuthor` child elements.
    """

    _tag_seq = ("p:cmAuthor",)
    cmAuthor = ZeroOrMore("p:cmAuthor", successors=())
    del _tag_seq

    @classmethod
    def new(cls) -> CT_CommentAuthorList:
        """Return a new `<p:cmAuthorLst>` element."""
        return OxmlElement("p:cmAuthorLst")

    def add_cmAuthor(self, id: int, name: str, initials: str) -> CT_CommentAuthor:
        """Add a new `p:cmAuthor` child and return it."""
        cmAuthor = self._add_cmAuthor()
        cmAuthor.id = id
        cmAuthor.name = name
        cmAuthor.initials = initials
        cmAuthor.lastIdx = 1
        cmAuthor.clrIdx = id
        return cmAuthor


class CT_CommentAuthor(BaseOxmlElement):
    """`p:cmAuthor` element — represents a comment author.

    Attributes:
        id: Unique author identifier (required, UInt32).
        name: Author display name (required, String).
        initials: Author initials (required, String).
        lastIdx: Last comment index used by this author (required, UInt32).
        clrIdx: Color index for this author (required, UInt32).
    """

    id = RequiredAttribute("id", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]
    name = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
    initials = RequiredAttribute("initials", XsdString)  # pyright: ignore[reportAssignmentType]
    lastIdx = RequiredAttribute("lastIdx", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]
    clrIdx = RequiredAttribute("clrIdx", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


class CT_CommentList(BaseOxmlElement):
    """`p:cmLst` element — root element of a slide comments part.

    Contains zero or more `p:cm` child elements.
    """

    _tag_seq = ("p:cm",)
    cm = ZeroOrMore("p:cm", successors=())
    del _tag_seq

    @classmethod
    def new(cls) -> CT_CommentList:
        """Return a new `<p:cmLst>` element."""
        return OxmlElement("p:cmLst")

    def add_cm(
        self,
        author_id: int,
        idx: int,
        x: int,
        y: int,
        text: str,
        dt: str | None = None,
    ) -> CT_Comment:
        """Add a new `p:cm` child and return it."""
        cm = self._add_cm()
        cm.authorId = author_id
        cm.idx = idx
        if dt is not None:
            cm.dt = dt
        pos = OxmlElement("p:pos")
        pos.set("x", str(x))
        pos.set("y", str(y))
        cm.insert(0, pos)
        text_elem = OxmlElement("p:text")
        text_elem.text = text
        cm.append(text_elem)
        return cm


class CT_Comment(BaseOxmlElement):
    """`p:cm` element — represents a single comment on a slide.

    Attributes:
        authorId: ID of the comment author (required, UInt32).
        dt: Date/time when comment was created (optional, DateTime).
        idx: Comment index (required, UInt32).

    Child elements (in order):
        p:pos — position on slide (required).
        p:text — comment text (required).
        p:extLst — extension list (optional, not yet implemented).
    """

    _tag_seq = ("p:pos", "p:text")
    pos = OneAndOnlyOne("p:pos")  # pyright: ignore[reportAssignmentType]
    text = OneAndOnlyOne("p:text")  # pyright: ignore[reportAssignmentType]
    del _tag_seq

    authorId = RequiredAttribute("authorId", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]
    dt = OptionalAttribute("dt", XsdString)  # pyright: ignore[reportAssignmentType]
    idx = RequiredAttribute("idx", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]

    get_or_add_pos: Callable[[], CT_CommentPosition]
    get_or_add_text: Callable[[], CT_CommentText]


class CT_CommentPosition(BaseOxmlElement):
    """`p:pos` element — comment position on a slide.

    Corresponds to CT_Point2D with x/y as ST_Coordinate (Int64/Emu).
    """

    x = RequiredAttribute("x", ST_Coordinate)  # pyright: ignore[reportAssignmentType]
    y = RequiredAttribute("y", ST_Coordinate)  # pyright: ignore[reportAssignmentType]


class CT_CommentText(BaseOxmlElement):
    """`p:text` element — text content of a comment.

    A leaf text element. The comment string is stored as the lxml `.text`
    attribute (the text content of the element).
    """

    @property
    def content(self) -> str:
        """The comment text content."""
        return _Element.text.__get__(self) or ""

    @content.setter
    def content(self, value: str):
        """Set the comment text content."""
        _Element.text.__set__(self, value)
