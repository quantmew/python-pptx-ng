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

    # -- Entrance effects --

    def add_appear_effect(self, shape_id: int, delay: str = "0") -> None:
        """Add an 'appear' entrance animation to a shape.

        Args:
            shape_id: The spid (shape ID) of the target shape.
            delay: Delay before the animation starts (e.g. "0", "500").
        """
        self._add_set_effect(
            shape_id=shape_id,
            delay=delay,
            preset_id=1,
            preset_class="entr",
            dur="1",
        )

    def add_fade_effect(self, shape_id: int, delay: str = "0", dur: str = "500") -> None:
        """Add a 'fade' entrance animation to a shape.

        Args:
            shape_id: The spid (shape ID) of the target shape.
            delay: Delay before the animation starts (e.g. "0", "500").
            dur: Duration in milliseconds (e.g. "500", "1000").
        """
        self._add_filter_effect(
            shape_id=shape_id,
            delay=delay,
            dur=dur,
            preset_id=10,
            preset_class="entr",
            transition="in",
            filter_name="fade",
        )

    def add_fly_in_effect(
        self,
        shape_id: int,
        delay: str = "0",
        dur: str = "500",
        direction: str = "bottom",
    ) -> None:
        """Add a 'fly in' entrance animation.

        Args:
            shape_id: The spid (shape ID) of the target shape.
            delay: Delay before the animation starts.
            dur: Duration in milliseconds.
            direction: Fly-in direction (e.g. "bottom", "left", "right", "top").
        """
        subtype_map = {
            "left": 8,
            "top": 2,
            "right": 6,
            "bottom": 4,
            "top-left": 1,
            "top-right": 3,
            "bottom-left": 5,
            "bottom-right": 7,
        }
        self._add_anim_effect(
            shape_id=shape_id,
            delay=delay,
            dur=dur,
            preset_id=2,
            preset_class="entr",
            preset_subtype=str(subtype_map.get(direction, 4)),
            attr_name="ppt_x",
            to_val="visible",
        )

    def add_zoom_effect(self, shape_id: int, delay: str = "0", dur: str = "500") -> None:
        """Add a 'zoom' entrance animation.

        Args:
            shape_id: The spid (shape ID) of the target shape.
            delay: Delay before the animation starts.
            dur: Duration in milliseconds.
        """
        self._add_filter_effect(
            shape_id=shape_id,
            delay=delay,
            dur=dur,
            preset_id=53,
            preset_class="entr",
            transition="in",
            filter_name="wipe",
        )

    def add_grow_turn_effect(self, shape_id: int, delay: str = "0", dur: str = "500") -> None:
        """Add a 'grow & turn' entrance animation.

        Args:
            shape_id: The spid (shape ID) of the target shape.
            delay: Delay before the animation starts.
            dur: Duration in milliseconds.
        """
        self._add_set_effect(
            shape_id=shape_id,
            delay=delay,
            preset_id=3,
            preset_class="entr",
            dur=dur,
        )

    # -- Exit effects --

    def add_fade_out_effect(self, shape_id: int, delay: str = "0", dur: str = "500") -> None:
        """Add a 'fade out' exit animation.

        Args:
            shape_id: The spid (shape ID) of the target shape.
            delay: Delay before the animation starts.
            dur: Duration in milliseconds.
        """
        self._add_filter_effect(
            shape_id=shape_id,
            delay=delay,
            dur=dur,
            preset_id=10,
            preset_class="exit",
            transition="out",
            filter_name="fade",
        )

    def add_disappear_effect(self, shape_id: int, delay: str = "0") -> None:
        """Add a 'disappear' exit animation.

        Args:
            shape_id: The spid (shape ID) of the target shape.
            delay: Delay before the animation starts.
        """
        self._add_set_effect(
            shape_id=shape_id,
            delay=delay,
            preset_id=1,
            preset_class="exit",
            dur="1",
            visibility_val="hidden",
        )

    # -- Emphasis effects --

    def add_spin_effect(self, shape_id: int, delay: str = "0", dur: str = "500") -> None:
        """Add a 'spin' emphasis animation.

        Args:
            shape_id: The spid (shape ID) of the target shape.
            delay: Delay before the animation starts.
            dur: Duration in milliseconds.
        """
        self._add_anim_effect(
            shape_id=shape_id,
            delay=delay,
            dur=dur,
            preset_id=22,
            preset_class="emph",
            preset_subtype="0",
            attr_name="style.rotation",
            by_val="360",
        )

    def add_grow_shrink_effect(
        self,
        shape_id: int,
        delay: str = "0",
        dur: str = "500",
        scale: str = "150",
    ) -> None:
        """Add a 'grow/shrink' emphasis animation.

        Args:
            shape_id: The spid (shape ID) of the target shape.
            delay: Delay before the animation starts.
            dur: Duration in milliseconds.
            scale: Target scale percentage (e.g. "150" for 150%).
        """
        self._add_anim_effect(
            shape_id=shape_id,
            delay=delay,
            dur=dur,
            preset_id=26,
            preset_class="emph",
            preset_subtype="0",
            attr_name="style.width",
            by_val=scale,
        )

    def add_transparency_effect(
        self,
        shape_id: int,
        delay: str = "0",
        dur: str = "500",
        transparency: str = "50",
    ) -> None:
        """Add a 'transparency' emphasis animation.

        Args:
            shape_id: The spid (shape ID) of the target shape.
            delay: Delay before the animation starts.
            dur: Duration in milliseconds.
            transparency: Target transparency percentage (0-100).
        """
        self._add_anim_effect(
            shape_id=shape_id,
            delay=delay,
            dur=dur,
            preset_id=39,
            preset_class="emph",
            preset_subtype="0",
            attr_name="style.opacity",
            to_val=transparency,
        )

    # -- Internal helpers --

    def _add_set_effect(
        self,
        shape_id: int,
        delay: str,
        preset_id: int,
        preset_class: str,
        dur: str = "1",
        visibility_val: str = "visible",
    ) -> None:
        """Add a p:set-based animation effect."""
        self._ensure_timing()
        cTn_id = self._next_cTn_id
        xml = (
            "<p:par %s>\n"
            "  <p:cTn id=\"%d\" fill=\"hold\">\n"
            "    <p:stCondLst><p:cond delay=\"%s\"/></p:stCondLst>\n"
            "    <p:childTnLst>\n"
            "      <p:par>\n"
            "        <p:cTn id=\"%d\" fill=\"hold\">\n"
            "          <p:stCondLst><p:cond delay=\"0\"/></p:stCondLst>\n"
            "          <p:childTnLst>\n"
            "            <p:par>\n"
            '              <p:cTn id=\"%d\" presetID=\"%d\" presetClass=\"%s\" '
            "presetSubtype=\"0\" fill=\"hold\" nodeType=\"afterEffect\">\n"
            "                <p:stCondLst><p:cond delay=\"0\"/></p:stCondLst>\n"
            "                <p:childTnLst>\n"
            "                  <p:set>\n"
            "                    <p:cBhvr>\n"
            '                      <p:cTn id=\"%d\" dur=\"%s\" fill=\"hold\">'
            "<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\n"
            "                      <p:tgtEl><p:spTgt spid=\"%d\"/></p:tgtEl>\n"
            "                      <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>\n"
            "                    </p:cBhvr>\n"
            "                    <p:to><p:strVal val=\"%s\"/></p:to>\n"
            "                  </p:set>\n"
            "                </p:childTnLst>\n"
            "              </p:cTn>\n"
            "            </p:par>\n"
            "          </p:childTnLst>\n"
            "        </p:cTn>\n"
            "      </p:par>\n"
            "    </p:childTnLst>\n"
            "  </p:cTn>\n"
            "</p:par>"
            % (nsdecls("p"), cTn_id, delay, cTn_id + 1, cTn_id + 2, preset_id, preset_class, cTn_id + 3, dur, shape_id, visibility_val)
        )
        self._element.tnLst.append(parse_xml(xml))

    def _add_filter_effect(
        self,
        shape_id: int,
        delay: str,
        dur: str,
        preset_id: int,
        preset_class: str,
        transition: str,
        filter_name: str,
    ) -> None:
        """Add a p:animEffect-based animation with filter and visibility set."""
        self._ensure_timing()
        cTn_id = self._next_cTn_id
        xml = (
            "<p:par %s>\n"
            "  <p:cTn id=\"%d\" fill=\"hold\">\n"
            "    <p:stCondLst><p:cond delay=\"%s\"/></p:stCondLst>\n"
            "    <p:childTnLst>\n"
            "      <p:par>\n"
            "        <p:cTn id=\"%d\" fill=\"hold\">\n"
            "          <p:stCondLst><p:cond delay=\"0\"/></p:stCondLst>\n"
            "          <p:childTnLst>\n"
            "            <p:par>\n"
            '              <p:cTn id=\"%d\" presetID=\"%d\" presetClass=\"%s\" '
            "presetSubtype=\"0\" fill=\"hold\" nodeType=\"afterEffect\">\n"
            "                <p:stCondLst><p:cond delay=\"0\"/></p:stCondLst>\n"
            "                <p:childTnLst>\n"
            '                  <p:animEffect transition=\"%s\" filter=\"%s\">\n'
            "                    <p:cBhvr>\n"
            '                      <p:cTn id=\"%d\" dur=\"%s\"/>\n'
            "                      <p:tgtEl><p:spTgt spid=\"%d\"/></p:tgtEl>\n"
            "                    </p:cBhvr>\n"
            "                  </p:animEffect>\n"
            "                  <p:set>\n"
            "                    <p:cBhvr>\n"
            '                      <p:cTn id=\"%d\" dur=\"1\" fill=\"hold\">'
            "<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\n"
            "                      <p:tgtEl><p:spTgt spid=\"%d\"/></p:tgtEl>\n"
            "                      <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>\n"
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
            "</p:par>"
            % (
                nsdecls("p"), cTn_id, delay, cTn_id + 1, cTn_id + 2,
                preset_id, preset_class, transition, filter_name,
                cTn_id + 3, dur, shape_id, cTn_id + 4, shape_id,
            )
        )
        self._element.tnLst.append(parse_xml(xml))

    def _add_anim_effect(
        self,
        shape_id: int,
        delay: str,
        dur: str,
        preset_id: int,
        preset_class: str,
        preset_subtype: str,
        attr_name: str,
        by_val: str | None = None,
        to_val: str | None = None,
    ) -> None:
        """Add a p:anim-based animation effect with optional by/to values."""
        self._ensure_timing()
        cTn_id = self._next_cTn_id
        anim_attrs = ""
        if by_val is not None:
            anim_attrs += ' by="%s"' % by_val
        if to_val is not None:
            anim_attrs += ' to="%s"' % to_val
        xml = (
            "<p:par %s>\n"
            "  <p:cTn id=\"%d\" fill=\"hold\">\n"
            "    <p:stCondLst><p:cond delay=\"%s\"/></p:stCondLst>\n"
            "    <p:childTnLst>\n"
            "      <p:par>\n"
            "        <p:cTn id=\"%d\" fill=\"hold\">\n"
            "          <p:stCondLst><p:cond delay=\"0\"/></p:stCondLst>\n"
            "          <p:childTnLst>\n"
            "            <p:par>\n"
            '              <p:cTn id=\"%d\" presetID=\"%d\" presetClass=\"%s\" '
            "presetSubtype=\"%s\" fill=\"hold\" nodeType=\"afterEffect\">\n"
            "                <p:stCondLst><p:cond delay=\"0\"/></p:stCondLst>\n"
            "                <p:childTnLst>\n"
            '                  <p:anim%s>\n'
            "                    <p:cBhvr>\n"
            '                      <p:cTn id=\"%d\" dur=\"%s\"/>\n'
            "                      <p:tgtEl><p:spTgt spid=\"%d\"/></p:tgtEl>\n"
            "                      <p:attrNameLst><p:attrName>%s</p:attrName></p:attrNameLst>\n"
            "                    </p:cBhvr>\n"
            "                  </p:anim>\n"
            "                </p:childTnLst>\n"
            "              </p:cTn>\n"
            "            </p:par>\n"
            "          </p:childTnLst>\n"
            "        </p:cTn>\n"
            "      </p:par>\n"
            "    </p:childTnLst>\n"
            "  </p:cTn>\n"
            "</p:par>"
            % (
                nsdecls("p"), cTn_id, delay, cTn_id + 1, cTn_id + 2,
                preset_id, preset_class, preset_subtype, anim_attrs,
                cTn_id + 3, dur, shape_id, attr_name,
            )
        )
        self._element.tnLst.append(parse_xml(xml))

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
