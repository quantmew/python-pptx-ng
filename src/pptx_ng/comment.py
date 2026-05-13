"""Comment-related proxy classes.

Provides high-level API for working with PowerPoint comments and comment authors.
"""

from __future__ import annotations

from datetime import datetime
from typing import Iterator

from pptx_ng.oxml.comment import (
    CT_Comment,
    CT_CommentAuthor,
    CT_CommentAuthorList,
    CT_CommentList,
)
from pptx_ng.shared import PartElementProxy


class CommentAuthors(PartElementProxy):
    """Collection of comment authors in a presentation.

    Accessed via `Presentation.comment_authors`.
    """

    _element: CT_CommentAuthorList

    def __iter__(self) -> Iterator[CommentAuthor]:
        for cmAuthor in self._element.cmAuthor_lst:
            yield CommentAuthor(cmAuthor, self)

    def __len__(self) -> int:
        return len(self._element.cmAuthor_lst)

    def add(self, name: str, initials: str) -> CommentAuthor:
        """Add a new comment author and return it.

        The author is assigned the next available ID.
        """
        next_id = self._next_id
        cmAuthor = self._element.add_cmAuthor(
            id=next_id, name=name, initials=initials
        )
        return CommentAuthor(cmAuthor, self)

    def get(self, id: int) -> CommentAuthor | None:
        """Return the author with the given ID, or None."""
        for cmAuthor in self._element.cmAuthor_lst:
            if cmAuthor.id == id:
                return CommentAuthor(cmAuthor, self)
        return None

    @property
    def _next_id(self) -> int:
        """Return the next available author ID."""
        if not self._element.cmAuthor_lst:
            return 0
        return max(a.id for a in self._element.cmAuthor_lst) + 1


class CommentAuthor:
    """A single comment author."""

    def __init__(self, cmAuthor: CT_CommentAuthor, parent: CommentAuthors):
        self._element = cmAuthor
        self._parent = parent

    @property
    def id(self) -> int:
        return self._element.id

    @property
    def name(self) -> str:
        return self._element.name

    @name.setter
    def name(self, value: str):
        self._element.name = value

    @property
    def initials(self) -> str:
        return self._element.initials

    @initials.setter
    def initials(self, value: str):
        self._element.initials = value


class Comments(PartElementProxy):
    """Collection of comments on a slide.

    Accessed via `Slide.comments`.
    """

    _element: CT_CommentList

    def __iter__(self) -> Iterator[Comment]:
        for cm in self._element.cm_lst:
            yield Comment(cm, self)

    def __len__(self) -> int:
        return len(self._element.cm_lst)

    def add(
        self,
        text: str,
        author: CommentAuthor,
        left: int = 914400,
        top: int = 914400,
    ) -> Comment:
        """Add a new comment and return it.

        Args:
            text: Comment text content.
            author: The CommentAuthor who wrote this comment.
            left: X position in EMU (default 1 inch).
            top: Y position in EMU (default 1 inch).
        """
        idx = self._next_idx
        dt = datetime.now().strftime("%Y-%m-%dT%H:%M:%S.%fZ")
        cm = self._element.add_cm(
            author_id=author.id,
            idx=idx,
            x=left,
            y=top,
            text=text,
            dt=dt,
        )
        # Update author's lastIdx
        self._update_author_last_idx(author.id, idx)
        return Comment(cm, self)

    @property
    def _next_idx(self) -> int:
        """Return the next available comment index."""
        if not self._element.cm_lst:
            return 1
        return max(cm.idx for cm in self._element.cm_lst) + 1

    def _update_author_last_idx(self, author_id: int, idx: int):
        """Update the author's lastIdx to reflect the new comment index."""
        comment_part = self.part
        package = comment_part.package
        try:
            pres_part = package.presentation_part
            authors_part = pres_part.comment_authors_part
            for cmAuthor in authors_part._element.cmAuthor_lst:
                if cmAuthor.id == author_id:
                    if idx > cmAuthor.lastIdx:
                        cmAuthor.lastIdx = idx
                    break
        except (AttributeError, KeyError):
            pass


class Comment:
    """A single comment on a slide."""

    def __init__(self, cm: CT_Comment, parent: Comments):
        self._element = cm
        self._parent = parent

    @property
    def author_id(self) -> int:
        return self._element.authorId

    @property
    def text(self) -> str:
        return self._element.text.content

    @text.setter
    def text(self, value: str):
        self._element.text.content = value

    @property
    def idx(self) -> int:
        return self._element.idx

    @property
    def dt(self) -> str | None:
        return self._element.dt

    @property
    def x(self) -> int:
        return self._element.pos.x

    @property
    def y(self) -> int:
        return self._element.pos.y
