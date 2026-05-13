"""Digital signature part classes for passthrough support."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx_ng.opc.constants import CONTENT_TYPE as CT
from pptx_ng.opc.package import Part, XmlPart

if TYPE_CHECKING:
    from pptx_ng.package import Package


class DigitalSignatureOriginPart(Part):
    """Digital signature origin part.

    Corresponds to _xmlsignatures/origin.sigs.
    """

    @classmethod
    def new(cls, blob: bytes, package: Package) -> DigitalSignatureOriginPart:
        return cls(
            package.next_partname("/_xmlsignatures/origin.sigs"),
            CT.OPC_DIGITAL_SIGNATURE_ORIGIN,
            package,
            blob,
        )


class DigitalSignatureXmlSignaturePart(XmlPart):
    """Individual digital signature XML part.

    Corresponds to _xmlsignatures/sig%d.xml.
    """

    partname_template = "/_xmlsignatures/sig%d.xml"

    @classmethod
    def new(cls, package: Package) -> DigitalSignatureXmlSignaturePart:
        return cls(
            package.next_partname(cls.partname_template),
            CT.OPC_DIGITAL_SIGNATURE_XMLSIGNATURE,
            package,
            '<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>',
        )
