"""Custom element classes for 3D model-related XML elements.

3D models use the m3d namespace (am3d in the ECMA schema):
  http://schemas.microsoft.com/office/drawing/2017/model3d

Reference: Open XML SDK, Office2019.Drawing.Model3D namespace
"""

from __future__ import annotations

from typing import Callable, cast

from pptx_ng.oxml.ns import nsdecls, qn
from pptx_ng.oxml.simpletypes import XsdString
from pptx_ng.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    ZeroOrOne,
)


def _parse(xml: str) -> BaseOxmlElement:
    from pptx_ng.oxml import parse_xml

    return parse_xml(xml)


class CT_Model3D(BaseOxmlElement):
    """`m3d:model3D` element — root element for a 3D model embedded in a slide."""

    _tag_seq = (
        "m3d:spPr",
        "m3d:camera",
        "m3d:trans",
        "m3d:objViewport",
        "m3d:winViewport",
        "m3d:ambientLight",
        "m3d:ptLight",
        "m3d:spotLight",
        "m3d:dirLight",
        "m3d:unkLight",
        "m3d:extLst",
    )
    camera: Callable[[], CT_Model3DCamera]
    camera = ZeroOrOne("m3d:camera", successors=_tag_seq[2:])
    trans: Callable[[], CT_Model3DTransform]
    trans = ZeroOrOne("m3d:trans", successors=_tag_seq[3:])
    del _tag_seq

    r_embed = OptionalAttribute("r:embed", XsdString)
    r_link = OptionalAttribute("r:link", XsdString)

    get_or_add_camera: Callable[[], CT_Model3DCamera]
    get_or_add_trans: Callable[[], CT_Model3DTransform]

    @classmethod
    def new(cls, r_embed: str | None = None) -> CT_Model3D:
        el = cast(
            CT_Model3D,
            _parse(
                "<m3d:model3D %s %s>"
                "<m3d:spPr/>"
                "<m3d:camera><m3d:pos/><m3d:up/><m3d:lookAt/><m3d:perspective/></m3d:camera>"
                "<m3d:trans><m3d:rot/></m3d:trans>"
                "</m3d:model3D>" % (nsdecls("m3d", "a", "r"), "")
            ),
        )
        if r_embed is not None:
            el.r_embed = r_embed
        return el


class CT_Model3DCamera(BaseOxmlElement):
    """`m3d:camera` element — defines camera position and projection for a 3D model."""

    _tag_seq = (
        "m3d:pos",
        "m3d:up",
        "m3d:lookAt",
        "m3d:orthographic",
        "m3d:perspective",
        "m3d:extLst",
    )
    del _tag_seq


class CT_Model3DTransform(BaseOxmlElement):
    """`m3d:trans` element — defines 3D transformation (scale, rotation, translation)."""

    _tag_seq = (
        "m3d:meterPerModelUnit",
        "m3d:preTrans",
        "m3d:scale",
        "m3d:rot",
        "m3d:postTrans",
        "m3d:extLst",
    )
    del _tag_seq


class CT_ObjectViewport(BaseOxmlElement):
    """`m3d:objViewport` element — object viewport settings."""


class CT_WindowViewport(BaseOxmlElement):
    """`m3d:winViewport` element — window viewport settings."""


class CT_AmbientLight(BaseOxmlElement):
    """`m3d:ambientLight` element — ambient light settings."""


class CT_PointLight(BaseOxmlElement):
    """`m3d:ptLight` element — point light settings."""


class CT_SpotLight(BaseOxmlElement):
    """`m3d:spotLight` element — spot light settings."""


class CT_DirectionalLight(BaseOxmlElement):
    """`m3d:dirLight` element — directional light settings."""


class CT_UnknownLight(BaseOxmlElement):
    """`m3d:unkLight` element — unknown light type."""


class CT_OrthographicProjection(BaseOxmlElement):
    """`m3d:orthographic` element — orthographic camera projection."""


class CT_PerspectiveProjection(BaseOxmlElement):
    """`m3d:perspective` element — perspective camera projection."""
