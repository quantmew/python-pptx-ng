"""High-level API for 3D model support in PowerPoint presentations.

Provides the |Model3D| proxy class for creating and manipulating embedded 3D
models (GLTF Binary format) on slides.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.shared import ElementProxy

if TYPE_CHECKING:
    from pptx.oxml.model3d import CT_Model3D, CT_Model3DCamera, CT_Model3DTransform


class Model3D(ElementProxy):
    """Proxy for an `m3d:model3D` element providing a high-level API for
    3D model manipulation.
    """

    def __init__(self, model3d: CT_Model3D):
        super(Model3D, self).__init__(model3d)
        self._model3d = model3d

    @property
    def camera(self) -> Model3DCamera:
        """|Model3DCamera| proxy for this model's camera settings."""
        cam = self._model3d.get_or_add_camera()
        return Model3DCamera(cam)

    @property
    def transform(self) -> Model3DTransform:
        """|Model3DTransform| proxy for this model's 3D transform."""
        trans = self._model3d.get_or_add_trans()
        return Model3DTransform(trans)

    @property
    def embed_rId(self) -> str | None:
        """The relationship ID for the embedded .glb file, or None."""
        return self._model3d.r_embed

    @property
    def xml(self) -> str:
        """The XML of this 3D model element."""
        from lxml import etree

        return etree.tostring(self._model3d, encoding="unicode")


class Model3DCamera(ElementProxy):
    """Proxy for an `m3d:camera` element."""

    def __init__(self, camera: CT_Model3DCamera):
        super(Model3DCamera, self).__init__(camera)
        self._camera = camera


class Model3DTransform(ElementProxy):
    """Proxy for an `m3d:trans` element."""

    def __init__(self, trans: CT_Model3DTransform):
        super(Model3DTransform, self).__init__(trans)
        self._trans = trans
