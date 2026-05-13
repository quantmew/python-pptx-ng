"""Comment-related part classes.

CommentAuthorsPart — presentation-level part containing all comment authors.
CommentPart — slide-level part containing comments for a single slide.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.comment import CT_CommentAuthorList, CT_CommentList
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.comment import CommentAuthors, Comments


class CommentAuthorsPart(XmlPart):
    """Comment authors part.

    Contains the list of all comment authors in the presentation.
    Corresponds to package file `ppt/commentAuthors.xml`.
    """

    @classmethod
    def new(cls, package) -> CommentAuthorsPart:
        """Return a new CommentAuthorsPart with an empty author list."""
        return cls(
            PackURI("/ppt/commentAuthors.xml"),
            CT.PML_COMMENT_AUTHORS,
            package,
            CT_CommentAuthorList.new(),
        )

    @lazyproperty
    def comment_authors(self) -> CommentAuthors:
        """Return the CommentAuthors proxy for this part."""
        from pptx.comment import CommentAuthors

        return CommentAuthors(self._element, self)


class CommentPart(XmlPart):
    """Slide comments part.

    Contains the list of comments for a single slide.
    Corresponds to package files `ppt/comments/commentN.xml`.
    """

    @classmethod
    def new(cls, package) -> CommentPart:
        """Return a new CommentPart with an empty comment list."""
        partname = package.next_partname("/ppt/comments/comment%d.xml")
        return cls(
            partname,
            CT.PML_COMMENTS,
            package,
            CT_CommentList.new(),
        )

    @lazyproperty
    def comments(self) -> Comments:
        """Return the Comments proxy for this part."""
        from pptx.comment import Comments

        return Comments(self._element, self)
