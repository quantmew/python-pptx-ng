"""Animation proxy classes.

Provides high-level API for working with slide animations.
"""

from __future__ import annotations

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.xmlchemy import OxmlElement
from pptx.shared import PartElementProxy


class AnimationTimeline(PartElementProxy):
    """Animation timeline for a slide.

    Accessed via `Slide.timing`.
    """

    def add_appear_effect(self, shape_id: int, delay: str = "0") -> None:
        """Add an 'appear' entrance animation to a shape.

        Args:
            shape_id: The spid (shape ID) of the target shape.
            delay: Delay before the animation starts (e.g. "0", "500").
        """
        self._ensure_timing()
        cTn_id = self._next_cTn_id
        xml = (
            "<p:par %s>\n"
            "  <p:cTn id=\"%d\" fill=\"hold\">\n"
            "    <p:stCondLst>\n"
            "      <p:cond delay=\"%s\"/>\n"
            "    </p:stCondLst>\n"
            "    <p:childTnLst>\n"
            "      <p:par>\n"
            "        <p:cTn id=\"%d\" fill=\"hold\">\n"
            "          <p:stCondLst>\n"
            "            <p:cond delay=\"0\"/>\n"
            "          </p:stCondLst>\n"
            "          <p:childTnLst>\n"
            "            <p:par>\n"
            "              <p:cTn id=\"%d\" presetID=\"1\" presetClass=\"entr\" "
            "presetSubtype=\"0\" fill=\"hold\" nodeType=\"afterEffect\">\n"
            "                <p:stCondLst>\n"
            "                  <p:cond delay=\"0\"/>\n"
            "                </p:stCondLst>\n"
            "                <p:childTnLst>\n"
            "                  <p:set>\n"
            "                    <p:cBhvr>\n"
            "                      <p:cTn id=\"%d\" dur=\"1\" fill=\"hold\">\n"
            "                        <p:stCondLst>\n"
            "                          <p:cond delay=\"0\"/>\n"
            "                        </p:stCondLst>\n"
            "                      </p:cTn>\n"
            "                      <p:tgtEl>\n"
            "                        <p:spTgt spid=\"%d\"/>\n"
            "                      </p:tgtEl>\n"
            "                      <p:attrNameLst>\n"
            "                        <p:attrName>style.visibility</p:attrName>\n"
            "                      </p:attrNameLst>\n"
            "                    </p:cBhvr>\n"
            "                    <p:to><p:strVal val=\"visible\"/></p:to>\n"
            "                  </p:set>\n"
            "                </p:childTnLst>\n"
            "              </p:cTn>\n"
            "            </p:par>\n"
            "          </p:childTnLst>\n"
            "        </p:cTn>\n"
            "      </p:par>\n"
            "    </p:childTnLst>\n"
            "  </p:cTn>\n"
            "</p:par>" % (nsdecls("p"), cTn_id, delay, cTn_id + 1, cTn_id + 2, cTn_id + 3, shape_id)
        )
        par = parse_xml(xml)
        self._element.tnLst.append(par)

    def add_fade_effect(self, shape_id: int, delay: str = "0", dur: str = "500") -> None:
        """Add a 'fade' entrance animation to a shape.

        Args:
            shape_id: The spid (shape ID) of the target shape.
            delay: Delay before the animation starts (e.g. "0", "500").
            dur: Duration in milliseconds (e.g. "500", "1000").
        """
        self._ensure_timing()
        cTn_id = self._next_cTn_id
        xml = (
            "<p:par %s>\n"
            "  <p:cTn id=\"%d\" fill=\"hold\">\n"
            "    <p:stCondLst>\n"
            "      <p:cond delay=\"%s\"/>\n"
            "    </p:stCondLst>\n"
            "    <p:childTnLst>\n"
            "      <p:par>\n"
            "        <p:cTn id=\"%d\" fill=\"hold\">\n"
            "          <p:stCondLst>\n"
            "            <p:cond delay=\"0\"/>\n"
            "          </p:stCondLst>\n"
            "          <p:childTnLst>\n"
            "            <p:par>\n"
            "              <p:cTn id=\"%d\" presetID=\"10\" presetClass=\"entr\" "
            "presetSubtype=\"0\" fill=\"hold\" nodeType=\"afterEffect\">\n"
            "                <p:stCondLst>\n"
            "                  <p:cond delay=\"0\"/>\n"
            "                </p:stCondLst>\n"
            "                <p:childTnLst>\n"
            "                  <p:animEffect transition=\"in\" filter=\"fade\">\n"
            "                    <p:cBhvr>\n"
            "                      <p:cTn id=\"%d\" dur=\"%s\"/>\n"
            "                      <p:tgtEl>\n"
            "                        <p:spTgt spid=\"%d\"/>\n"
            "                      </p:tgtEl>\n"
            "                    </p:cBhvr>\n"
            "                  </p:animEffect>\n"
            "                  <p:set>\n"
            "                    <p:cBhvr>\n"
            "                      <p:cTn id=\"%d\" dur=\"1\" fill=\"hold\">\n"
            "                        <p:stCondLst>\n"
            "                          <p:cond delay=\"0\"/>\n"
            "                        </p:stCondLst>\n"
            "                      </p:cTn>\n"
            "                      <p:tgtEl>\n"
            "                        <p:spTgt spid=\"%d\"/>\n"
            "                      </p:tgtEl>\n"
            "                      <p:attrNameLst>\n"
            "                        <p:attrName>style.visibility</p:attrName>\n"
            "                      </p:attrNameLst>\n"
            "                    </p:cBhvr>\n"
            "                    <p:to><p:strVal val=\"visible\"/></p:to>\n"
            "                  </p:set>\n"
            "                </p:childTnLst>\n"
            "              </p:cTn>\n"
            "            </p:par>\n"
            "          </p:childTnLst>\n"
            "        </p:cTn>\n"
            "      </p:par>\n"
            "    </p:childTnLst>\n"
            "  </p:cTn>\n"
            "</p:par>" % (nsdecls("p"), cTn_id, delay, cTn_id + 1, cTn_id + 2, cTn_id + 3, dur, shape_id, cTn_id + 4, shape_id)
        )
        par = parse_xml(xml)
        self._element.tnLst.append(par)

    def _ensure_timing(self):
        """Ensure the tnLst child exists in the timing element."""
        if self._element.tnLst is None:
            tnLst = OxmlElement("p:tnLst")
            self._element.append(tnLst)

    @property
    def _next_cTn_id(self) -> int:
        """Return the next available cTn ID."""
        ids = self._element.xpath("//p:cTn/@id")
        if not ids:
            return 1
        return max(int(i) for i in ids) + 1
