"""ActiveX control part class for passthrough support."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.package import Part

if TYPE_CHECKING:
    from pptx.package import Package


class ActiveXControlPart(Part):
    """ActiveX control binary part.

    Corresponds to embeddings/control%d.bin.
    """

    partname_template = "/embeddings/control%d.bin"

    @classmethod
    def new(cls, blob: bytes, package: Package) -> ActiveXControlPart:
        return cls(
            package.next_partname(cls.partname_template),
            "application/vnd.ms-office.activeX",
            package,
            blob,
        )
