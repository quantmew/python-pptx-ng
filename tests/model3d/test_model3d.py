"""Unit tests for 3D model support (Phase 3.6)."""

from __future__ import annotations

import pytest

from pptx import Presentation
from pptx.model3d import Model3D, Model3DCamera, Model3DTransform
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.model3d import (
    CT_AmbientLight,
    CT_DirectionalLight,
    CT_Model3D,
    CT_Model3DCamera,
    CT_Model3DTransform,
    CT_ObjectViewport,
    CT_WindowViewport,
    CT_PointLight,
    CT_SpotLight,
    CT_UnknownLight,
    CT_OrthographicProjection,
    CT_PerspectiveProjection,
)
from pptx.oxml import parse_xml
from pptx.parts.model3d import Model3DBlobPart


# -- CT_Model3D element tests --


class DescribeCT_Model3D:
    def it_creates_with_embed_reference(self):
        m = CT_Model3D.new("rId42")
        assert m.tag == qn("m3d:model3D")
        assert m.r_embed == "rId42"

    def it_creates_without_embed_reference(self):
        m = CT_Model3D.new()
        assert m.r_embed is None

    def it_has_camera_child(self):
        m = CT_Model3D.new()
        assert m.camera is not None
        assert isinstance(m.camera, CT_Model3DCamera)

    def it_has_transform_child(self):
        m = CT_Model3D.new()
        assert m.trans is not None
        assert isinstance(m.trans, CT_Model3DTransform)

    def it_sets_embed_attribute(self):
        m = CT_Model3D.new()
        m.r_embed = "rId5"
        assert m.r_embed == "rId5"


class DescribeCT_Model3DCamera:
    def it_parses_from_xml(self):
        xml = (
            "<m3d:camera %s>"
            "<m3d:pos/>"
            "<m3d:up/>"
            "<m3d:lookAt/>"
            "<m3d:perspective/>"
            "</m3d:camera>" % nsdecls("m3d")
        )
        cam = parse_xml(xml)
        assert isinstance(cam, CT_Model3DCamera)


class DescribeCT_Model3DTransform:
    def it_parses_from_xml(self):
        xml = (
            "<m3d:trans %s>"
            "<m3d:rot/>"
            "</m3d:trans>" % nsdecls("m3d")
        )
        trans = parse_xml(xml)
        assert isinstance(trans, CT_Model3DTransform)


class DescribeLightElements:
    def it_creates_ambient_light(self):
        xml = "<m3d:ambientLight %s/>" % nsdecls("m3d")
        el = parse_xml(xml)
        assert isinstance(el, CT_AmbientLight)

    def it_creates_directional_light(self):
        xml = "<m3d:dirLight %s/>" % nsdecls("m3d")
        el = parse_xml(xml)
        assert isinstance(el, CT_DirectionalLight)

    def it_creates_point_light(self):
        xml = "<m3d:ptLight %s/>" % nsdecls("m3d")
        el = parse_xml(xml)
        assert isinstance(el, CT_PointLight)

    def it_creates_spot_light(self):
        xml = "<m3d:spotLight %s/>" % nsdecls("m3d")
        el = parse_xml(xml)
        assert isinstance(el, CT_SpotLight)

    def it_creates_unknown_light(self):
        xml = "<m3d:unkLight %s/>" % nsdecls("m3d")
        el = parse_xml(xml)
        assert isinstance(el, CT_UnknownLight)


class DescribeViewportElements:
    def it_creates_object_viewport(self):
        xml = "<m3d:objViewport %s/>" % nsdecls("m3d")
        el = parse_xml(xml)
        assert isinstance(el, CT_ObjectViewport)

    def it_creates_window_viewport(self):
        xml = "<m3d:winViewport %s/>" % nsdecls("m3d")
        el = parse_xml(xml)
        assert isinstance(el, CT_WindowViewport)


class DescribeProjectionElements:
    def it_creates_orthographic(self):
        xml = "<m3d:orthographic %s/>" % nsdecls("m3d")
        el = parse_xml(xml)
        assert isinstance(el, CT_OrthographicProjection)

    def it_creates_perspective(self):
        xml = "<m3d:perspective %s/>" % nsdecls("m3d")
        el = parse_xml(xml)
        assert isinstance(el, CT_PerspectiveProjection)


# -- Model3DBlobPart tests --


class DescribeModel3DBlobPart:
    def it_creates_from_blob(self):
        prs = Presentation()
        blob = b"fake glb data"
        part = Model3DBlobPart.new(blob, prs.part.package)
        assert part._blob == blob
        assert "model3d" in part.partname

    def it_has_correct_content_type(self):
        from pptx.opc.constants import CONTENT_TYPE as CT

        assert CT.MODEL_GLB == "model/gltf-binary"


# -- High-level Model3D proxy tests --


class DescribeModel3DProxy:
    def it_creates_from_element(self):
        m = CT_Model3D.new("rId1")
        proxy = Model3D(m)
        assert isinstance(proxy, Model3D)

    def it_provides_camera_proxy(self):
        m = CT_Model3D.new()
        proxy = Model3D(m)
        cam = proxy.camera
        assert isinstance(cam, Model3DCamera)

    def it_provides_transform_proxy(self):
        m = CT_Model3D.new()
        proxy = Model3D(m)
        trans = proxy.transform
        assert isinstance(trans, Model3DTransform)

    def it_returns_embed_rId(self):
        m = CT_Model3D.new("rId42")
        proxy = Model3D(m)
        assert proxy.embed_rId == "rId42"

    def it_produces_xml(self):
        m = CT_Model3D.new()
        proxy = Model3D(m)
        xml = proxy.xml
        assert "m3d:model3D" in xml
        assert "m3d:camera" in xml


# -- Namespace registration tests --


class DescribeModel3DNamespace:
    def it_resolves_m3d_namespace(self):
        from pptx.oxml.ns import _nsmap

        assert "m3d" in _nsmap
        assert _nsmap["m3d"] == "http://schemas.microsoft.com/office/drawing/2017/model3d"

    def it_resolves_model3d_elements(self):
        from pptx.oxml import register_element_cls

        m = parse_xml("<m3d:model3D %s/>" % nsdecls("m3d"))
        assert isinstance(m, CT_Model3D)
