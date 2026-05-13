"""Diagram (SmartArt) part classes.

Provides the four diagram-related parts: DiagramDataPart, DiagramLayoutPart,
DiagramColorsPart, and DiagramStylePart.
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.package import Package


def _read_template_bytes(template_name: str) -> bytes:
    thisdir = os.path.split(__file__)[0]
    filename = os.path.join(thisdir, "..", "templates", "%s.xml" % template_name)
    with open(filename, "rb") as f:
        return f.read()


class DiagramDataPart(XmlPart):
    """Diagram data part containing the SmartArt data model.

    Corresponds to parts matching ppt/diagrams/data%d.xml.
    """

    partname_template = "/ppt/diagrams/data%d.xml"

    @classmethod
    def new(cls, package: Package) -> DiagramDataPart:
        return cls.load(
            package.next_partname(cls.partname_template),
            CT.DML_DIAGRAM_DATA,
            package,
            _read_template_bytes("diagramData"),
        )

    @lazyproperty
    def data_model(self):
        from pptx.diagram import SmartArtData

        return SmartArtData(self._element, self)


class DiagramLayoutPart(XmlPart):
    """Diagram layout definition part.

    Corresponds to parts matching ppt/diagrams/layout%d.xml.
    """

    partname_template = "/ppt/diagrams/layout%d.xml"

    @classmethod
    def new(cls, package: Package, layout_id: str | None = None) -> DiagramLayoutPart:
        return cls.load(
            package.next_partname(cls.partname_template),
            CT.DML_DIAGRAM_LAYOUT,
            package,
            _read_template_bytes("diagramLayout"),
        )


class DiagramColorsPart(XmlPart):
    """Diagram colors part.

    Corresponds to parts matching ppt/diagrams/colors%d.xml.
    """

    partname_template = "/ppt/diagrams/colors%d.xml"

    @classmethod
    def new(cls, package: Package, colors_id: str | None = None) -> DiagramColorsPart:
        return cls.load(
            package.next_partname(cls.partname_template),
            CT.DML_DIAGRAM_COLORS,
            package,
            _read_template_bytes("diagramColors"),
        )


class DiagramStylePart(XmlPart):
    """Diagram style part.

    Corresponds to parts matching ppt/diagrams/quickStyle%d.xml.
    """

    partname_template = "/ppt/diagrams/quickStyle%d.xml"

    @classmethod
    def new(cls, package: Package, style_id: str | None = None) -> DiagramStylePart:
        return cls.load(
            package.next_partname(cls.partname_template),
            CT.DML_DIAGRAM_STYLE,
            package,
            _read_template_bytes("diagramStyle"),
        )
