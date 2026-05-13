"""Table styles XML part class."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx_ng.opc.constants import CONTENT_TYPE as CT
from pptx_ng.opc.package import XmlPart
from pptx_ng.oxml.tablestyle import CT_TableStyleList
from pptx_ng.tablestyle import TableStyles
from pptx_ng.util import lazyproperty

if TYPE_CHECKING:
    from pptx_ng.package import Package


class TableStylesPart(XmlPart):
    """Table styles XML part.

    Corresponds to parts matching ppt/tableStyles.xml.
    """

    partname_template = "/ppt/tableStyles%d.xml"

    @classmethod
    def new(cls, package: Package) -> TableStylesPart:
        """Return a new |TableStylesPart| with default content."""
        tblStyleLst = CT_TableStyleList.new()
        return cls(
            package.next_partname(cls.partname_template),
            CT.PML_TABLE_STYLES,
            package,
            tblStyleLst,
        )

    @lazyproperty
    def table_styles(self) -> TableStyles:
        """Return a |TableStyles| proxy for this part."""
        return TableStyles(self._element)
