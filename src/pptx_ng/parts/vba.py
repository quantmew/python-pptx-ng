"""VBA project binary part class."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx_ng.opc.constants import CONTENT_TYPE as CT
from pptx_ng.opc.package import Part

if TYPE_CHECKING:
    from pptx_ng.package import Package


class VbaProjectPart(Part):
    """VBA project binary part.

    Corresponds to parts matching ppt/vbaProject.bin.
    """

    partname_template = "/ppt/vbaProject.bin"

    @classmethod
    def new(cls, blob: bytes, package: Package) -> VbaProjectPart:
        """Return a new |VbaProjectPart| containing *blob*."""
        return cls(
            package.next_partname(cls.partname_template),
            CT.VBA_PROJECT,
            package,
            blob,
        )
