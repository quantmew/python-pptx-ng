"""3D model (GLTF Binary) part class."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import Part

if TYPE_CHECKING:
    from pptx.package import Package


class Model3DBlobPart(Part):
    """Embedded 3D model binary (.glb) part.

    Corresponds to parts matching ppt/3dObjects/model3d%d.glb.
    """

    partname_template = "/ppt/3dObjects/model3d%d.glb"

    @classmethod
    def new(cls, blob: bytes, package: Package) -> Model3DBlobPart:
        """Return a new |Model3DBlobPart| containing *blob*."""
        return cls(
            package.next_partname(cls.partname_template),
            CT.MODEL_GLB,
            package,
            blob,
        )
