"""Slide sync data XML part class."""

from __future__ import annotations

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart


class SlideSyncDataPart(XmlPart):
    """Slide synchronization data XML part.

    Corresponds to parts matching ppt/slideUpdateInfo/slideUpdateInfo%d.xml.
    Used for slide library synchronization scenarios.
    """

    partname_template = "/ppt/slideUpdateInfo/slideUpdateInfo%d.xml"

    @classmethod
    def new(cls, package) -> SlideSyncDataPart:
        """Return a new |SlideSyncDataPart| with minimal content."""
        xml = (
            '<p:si xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
        )
        return cls(
            package.next_partname(cls.partname_template),
            CT.PML_SLIDE_UPDATE_INFO,
            package,
            xml,
        )
