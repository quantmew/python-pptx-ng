"""Custom element classes for 3D model-related XML elements.

3D models use the m3d namespace:
  http://schemas.microsoft.com/office/drawing/2017/model3d
"""

from __future__ import annotations

from pptx.oxml.xmlchemy import BaseOxmlElement


class CT_Model3D(BaseOxmlElement):
    """`m3d:model3D` element, root element for 3D model content."""
