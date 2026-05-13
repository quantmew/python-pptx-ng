"""Theme XML part class."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.oxml.theme import CT_OfficeStyleSheet
from pptx.theme import Theme
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.package import Package


class ThemePart(XmlPart):
    """Theme XML part.

    Corresponds to parts matching ppt/theme/theme%d.xml.
    """

    partname_template = "/ppt/theme/theme%d.xml"

    @classmethod
    def new(cls, package: Package) -> ThemePart:
        """Return a new |ThemePart| with default content."""
        return cls(
            package.next_partname(cls.partname_template),
            CT.OFC_THEME,
            package,
            CT_OfficeStyleSheet.new_default(),
        )

    @lazyproperty
    def theme(self) -> Theme:
        """Return a |Theme| proxy for this part."""
        return Theme(cast(CT_OfficeStyleSheet, self._element))
