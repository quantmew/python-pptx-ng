"""User-defined tags proxy classes."""

from __future__ import annotations

from pptx.shared import PartElementProxy


class UserDefinedTags(PartElementProxy):
    """Proxy for the user-defined tags part.

    Provides access to the tag collection on a slide.
    """

    def get(self, name: str) -> str | None:
        """Return the value of the tag with *name*, or None if not found."""
        tag = self._element.get_tag(name)
        if tag is None:
            return None
        return tag.val

    def set(self, name: str, val: str) -> None:
        """Set a tag with *name* to *val*, creating it if it doesn't exist."""
        tag = self._element.get_tag(name)
        if tag is not None:
            tag.val = val
        else:
            self._element.add_tag(name, val)

    def items(self) -> list[tuple[str, str]]:
        """Return all tags as a list of (name, value) tuples."""
        return [(tag.name, tag.val) for tag in self._element.tag_lst]

    def __len__(self) -> int:
        return len(self._element.tag_lst)

    def __contains__(self, name: str) -> bool:
        return self._element.get_tag(name) is not None
