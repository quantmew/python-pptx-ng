"""Web extension part classes for passthrough support."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx_ng.opc.package import XmlPart

if TYPE_CHECKING:
    from pptx_ng.package import Package


class WebExtensionPart(XmlPart):
    """Web extension XML part.

    Corresponds to webextensions/webextension%d.xml.
    """

    partname_template = "/ppt/webextensions/webextension%d.xml"

    @classmethod
    def new(cls, package: Package) -> WebExtensionPart:
        return cls(
            package.next_partname(cls.partname_template),
            "application/vnd.ms-office.webextension+xml",
            package,
            '<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"/>',
        )


class WebExTaskpanesPart(XmlPart):
    """Web extension task panes XML part.

    Corresponds to ppt/webextensions/taskpanes%d.xml.
    """

    partname_template = "/ppt/webextensions/taskpanes%d.xml"

    @classmethod
    def new(cls, package: Package) -> WebExTaskpanesPart:
        return cls(
            package.next_partname(cls.partname_template),
            "application/vnd.ms-office.webextensiontaskpanes+xml",
            package,
            '<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"/>',
        )
