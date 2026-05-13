"""Unit tests for handout master support (Phase 3.2)."""

from __future__ import annotations

import pytest

from pptx_ng import Presentation
from pptx_ng.handout import HandoutMaster
from pptx_ng.oxml.handout import CT_HandoutMaster, CT_HandoutMasterIdList, CT_HandoutMasterIdListEntry
from pptx_ng.oxml.ns import qn


class DescribeCT_HandoutMaster:
    def it_creates_default_element(self):
        hm = CT_HandoutMaster.new_default()
        assert hm.tag == qn("p:handoutMaster")
        assert hm.cSld is not None
        assert hm.cSld.spTree is not None

    def it_has_spTree_property(self):
        hm = CT_HandoutMaster.new_default()
        assert hm.spTree is not None


class DescribeCT_HandoutMasterIdList:
    def it_creates_element(self):
        from pptx_ng.oxml import parse_xml

        xml = '<p:handoutMasterIdLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
        lst = parse_xml(xml)
        assert isinstance(lst, CT_HandoutMasterIdList)


class DescribeHandoutMasterPart:
    def it_creates_default_part(self):
        from pptx_ng.parts.handout import HandoutMasterPart

        prs = Presentation()
        part = HandoutMasterPart.create_default(prs.part.package)
        assert part.handout_master is not None
        assert isinstance(part.handout_master, HandoutMaster)


class DescribePresentationHandoutMaster:
    def it_provides_handout_master(self):
        prs = Presentation()
        hm = prs.handout_master
        assert isinstance(hm, HandoutMaster)

    def it_persists_through_save_load(self, tmp_path):
        prs = Presentation()
        _ = prs.handout_master  # trigger creation
        path = tmp_path / "test_handout.pptx"
        prs.save(str(path))
        prs2 = Presentation(str(path))
        hm2 = prs2.handout_master
        assert isinstance(hm2, HandoutMaster)
        assert hm2.shapes is not None
