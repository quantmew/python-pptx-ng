"""Ribbon/customUI XML part classes for passthrough support."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.package import XmlPart

if TYPE_CHECKING:
    from pptx.package import Package


class RibbonExtensibilityPart(XmlPart):
    """Ribbon extensibility (customUI) XML part.

    Corresponds to customUI/customUI%d.xml.
    Relationship type: http://schemas.microsoft.com/office/2006/relationships/ui/extensibility
    """

    partname_template = "/customUI/customUI%d.xml"

    @classmethod
    def new(cls, package: Package) -> RibbonExtensibilityPart:
        return cls(
            package.next_partname(cls.partname_template),
            "application/xml",
            package,
            '<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>',
        )


class QuickAccessToolbarCustomizationsPart(XmlPart):
    """Quick access toolbar customization XML part.

    Corresponds to userCustomization/userCustomization%d.xml.
    """

    partname_template = "/userCustomization/userCustomization%d.xml"

    @classmethod
    def new(cls, package: Package) -> QuickAccessToolbarCustomizationsPart:
        return cls(
            package.next_partname(cls.partname_template),
            "application/xml",
            package,
            '<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>',
        )
