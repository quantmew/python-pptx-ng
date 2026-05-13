"""HandoutMasterPart and related objects."""

from __future__ import annotations

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.handout import CT_HandoutMaster
from pptx.oxml.theme import CT_OfficeStyleSheet
from pptx.util import lazyproperty


class HandoutMasterPart(XmlPart):
    """Handout master part.

    Corresponds to package file `ppt/handoutMasters/handoutMaster1.xml`.
    """

    @classmethod
    def create_default(cls, package):
        """Create and return a default handout master part with theme."""
        handout_master_part = cls._new(package)
        theme_part = cls._new_theme_part(package)
        handout_master_part.relate_to(theme_part, RT.THEME)
        return handout_master_part

    @lazyproperty
    def handout_master(self):
        """Return the |HandoutMaster| proxy for this part."""
        from pptx.handout import HandoutMaster

        return HandoutMaster(self._element, self)

    @classmethod
    def _new(cls, package):
        """Return a standalone, default handout master part."""
        return HandoutMasterPart(
            PackURI("/ppt/handoutMasters/handoutMaster1.xml"),
            CT.PML_HANDOUT_MASTER,
            package,
            CT_HandoutMaster.new_default(),
        )

    @classmethod
    def _new_theme_part(cls, package):
        """Return new default theme-part suitable for use with a handout master."""
        return XmlPart(
            package.next_partname("/ppt/theme/theme%d.xml"),
            CT.OFC_THEME,
            package,
            CT_OfficeStyleSheet.new_default(),
        )
