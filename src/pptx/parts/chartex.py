"""Extended chart (chartEx) part class.

Supports modern chart types like treemap, sunburst, waterfall, etc.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.chartex import ChartEx
    from pptx.package import Package


class ChartExPart(XmlPart):
    """Extended chart part.

    Corresponds to parts matching ppt/charts/chartEx%d.xml.
    """

    partname_template = "/ppt/charts/chartEx%d.xml"

    @classmethod
    def new(cls, package: Package) -> ChartExPart:
        from pptx.oxml.chartex.chartex import CT_ChartExSpace
        from lxml import etree

        chart_space = CT_ChartExSpace.new()
        return cls.load(
            package.next_partname(cls.partname_template),
            CT.OFC_CHART_EX,
            package,
            etree.tostring(chart_space, xml_declaration=True, encoding="UTF-8", standalone=True),
        )

    @lazyproperty
    def chart_ex(self) -> ChartEx:
        """|ChartEx| object representing the extended chart in this part."""
        from pptx.chartex import ChartEx

        return ChartEx(self._element, self)
