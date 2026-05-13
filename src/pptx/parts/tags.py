"""UserDefinedTagsPart and related objects."""

from __future__ import annotations

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.oxml.tags import CT_TagList
from pptx.util import lazyproperty


class UserDefinedTagsPart(XmlPart):
    """Part containing user-defined tags for a slide.

    Corresponds to package files `ppt/tags/tag[1-9][0-9]*.xml`.
    """

    @classmethod
    def new(cls, package):
        """Return a new |UserDefinedTagsPart| with an empty tag list."""
        return UserDefinedTagsPart(
            package.next_partname("/ppt/tags/tag%d.xml"),
            CT.PML_TAGS,
            package,
            CT_TagList.new(),
        )

    @lazyproperty
    def tag_list(self):
        """Return the |UserDefinedTags| proxy for this part."""
        from pptx.tags import UserDefinedTags

        return UserDefinedTags(self._element, self)
