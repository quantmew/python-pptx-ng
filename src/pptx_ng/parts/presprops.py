"""Presentation properties and view properties part classes."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx_ng.opc.constants import CONTENT_TYPE as CT
from pptx_ng.opc.constants import RELATIONSHIP_TYPE as RT
from pptx_ng.opc.package import XmlPart
from pptx_ng.opc.packuri import PackURI
from pptx_ng.oxml.presprops import CT_PresentationProperties
from pptx_ng.oxml.viewprops import CT_ViewProperties
from pptx_ng.util import lazyproperty

if TYPE_CHECKING:
    from pptx_ng.presprops import PresentationProperties
    from pptx_ng.viewprops import ViewProperties


class PresentationPropertiesPart(XmlPart):
    """Presentation properties part.

    Contains presentation-wide settings like show properties and printing properties.
    Corresponds to package file `ppt/presProps.xml`.
    """

    @classmethod
    def new(cls, package) -> PresentationPropertiesPart:
        """Return a new PresentationPropertiesPart with default content."""
        return cls(
            PackURI("/ppt/presProps.xml"),
            CT.PML_PRES_PROPS,
            package,
            CT_PresentationProperties.new(),
        )

    @lazyproperty
    def presentation_properties(self) -> PresentationProperties:
        """Return the PresentationProperties proxy for this part."""
        from pptx_ng.presprops import PresentationProperties

        return PresentationProperties(self._element, self)


class ViewPropertiesPart(XmlPart):
    """View properties part.

    Contains view settings like grid spacing and last view type.
    Corresponds to package file `ppt/viewProps.xml`.
    """

    @classmethod
    def new(cls, package) -> ViewPropertiesPart:
        """Return a new ViewPropertiesPart with default content."""
        return cls(
            PackURI("/ppt/viewProps.xml"),
            CT.PML_VIEW_PROPS,
            package,
            CT_ViewProperties.new(),
        )

    @lazyproperty
    def view_properties(self) -> ViewProperties:
        """Return the ViewProperties proxy for this part."""
        from pptx_ng.viewprops import ViewProperties

        return ViewProperties(self._element, self)
