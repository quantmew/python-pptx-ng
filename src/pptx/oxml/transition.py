"""Slide transition XML element classes.

Corresponds to CT_SlideTransition and related transition type classes
in the Open XML SDK (schemas_openxmlformats_org_presentationml_2006_main).

Transition type hierarchy:
    CT_Transition (p:transition) — composite, choice of 1 transition child
      Orientation-type (blinds, checker, comb, randomBar): dir attr
      EightDirection-type (cover, pull): dir attr
      SideDirection-type (push, wipe): dir attr
      OptionalBlack-type (cut, fade): thruBlk attr
      Split: orient + dir attrs
      Strips: dir attr (corner direction)
      Wheel: spokes attr
      Zoom: dir attr (in/out)
      Empty-type (random, circle, diamond, newsflash, plus, wedge, dissolve): no attrs
"""

from __future__ import annotations

from pptx.oxml.simpletypes import XsdBoolean, XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)


class CT_Transition(BaseOxmlElement):
    """`p:transition` element — slide transition settings.

    Attributes:
        spd: Transition speed (optional, String enum: "slow", "med", "fast").
        advClick: Advance on click (optional, Boolean, default true).
        advTm: Advance after time in milliseconds (optional, String/UInt32).

    Child elements (choice of 0..1):
        One transition type element, e.g. p:fade, p:push, etc.
    """

    spd = OptionalAttribute("spd", XsdString)  # pyright: ignore[reportAssignmentType]
    advClick = OptionalAttribute("advClick", XsdBoolean)  # pyright: ignore[reportAssignmentType]
    advTm = OptionalAttribute("advTm", XsdString)  # pyright: ignore[reportAssignmentType]

    def _set_transition_type(self, tag: str) -> BaseOxmlElement:
        """Remove any existing transition child and add a new one."""
        # Remove existing transition type child
        for child in list(self):
            if child.tag != self.tag and not child.tag.endswith("}sndAc") and not child.tag.endswith("}extLst"):
                self.remove(child)
        return OxmlElement(tag)

    def set_blinds(self, dir: str = "horz") -> None:
        el = self._set_transition_type("p:blinds")
        el.set("dir", dir)
        self.insert(0, el)

    def set_checker(self, dir: str = "horz") -> None:
        el = self._set_transition_type("p:checker")
        el.set("dir", dir)
        self.insert(0, el)

    def set_comb(self, dir: str = "horz") -> None:
        el = self._set_transition_type("p:comb")
        el.set("dir", dir)
        self.insert(0, el)

    def set_cover(self, dir: str = "l") -> None:
        el = self._set_transition_type("p:cover")
        el.set("dir", dir)
        self.insert(0, el)

    def set_pull(self, dir: str = "l") -> None:
        el = self._set_transition_type("p:pull")
        el.set("dir", dir)
        self.insert(0, el)

    def set_cut(self, thruBlk: bool = False) -> None:
        el = self._set_transition_type("p:cut")
        if thruBlk:
            el.set("thruBlk", "1")
        self.insert(0, el)

    def set_fade(self, thruBlk: bool = False) -> None:
        el = self._set_transition_type("p:fade")
        if thruBlk:
            el.set("thruBlk", "1")
        self.insert(0, el)

    def set_push(self, dir: str = "l") -> None:
        el = self._set_transition_type("p:push")
        el.set("dir", dir)
        self.insert(0, el)

    def set_wipe(self, dir: str = "l") -> None:
        el = self._set_transition_type("p:wipe")
        el.set("dir", dir)
        self.insert(0, el)

    def set_split(self, orient: str = "horz", dir: str = "out") -> None:
        el = self._set_transition_type("p:split")
        el.set("orient", orient)
        el.set("dir", dir)
        self.insert(0, el)

    def set_strips(self, dir: str = "ld") -> None:
        el = self._set_transition_type("p:strips")
        el.set("dir", dir)
        self.insert(0, el)

    def set_random_bar(self, dir: str = "horz") -> None:
        el = self._set_transition_type("p:randomBar")
        el.set("dir", dir)
        self.insert(0, el)

    def set_wheel(self, spokes: int = 4) -> None:
        el = self._set_transition_type("p:wheel")
        el.set("spokes", str(spokes))
        self.insert(0, el)

    def set_wedge(self) -> None:
        el = self._set_transition_type("p:wedge")
        self.insert(0, el)

    def set_zoom(self, dir: str = "out") -> None:
        el = self._set_transition_type("p:zoom")
        el.set("dir", dir)
        self.insert(0, el)

    def set_dissolve(self) -> None:
        el = self._set_transition_type("p:dissolve")
        self.insert(0, el)

    def set_random(self) -> None:
        el = self._set_transition_type("p:random")
        self.insert(0, el)

    def set_circle(self) -> None:
        el = self._set_transition_type("p:circle")
        self.insert(0, el)

    def set_diamond(self) -> None:
        el = self._set_transition_type("p:diamond")
        self.insert(0, el)

    def set_newsflash(self) -> None:
        el = self._set_transition_type("p:newsflash")
        self.insert(0, el)

    def set_plus(self) -> None:
        el = self._set_transition_type("p:plus")
        self.insert(0, el)

    @property
    def transition_type(self) -> str | None:
        """Return the local tag name of the active transition child, or None."""
        for child in self:
            tag = child.tag
            if "}" in tag:
                local = tag.split("}")[1]
            else:
                local = tag
            if local in (
                "blinds", "checker", "comb", "cover", "cut", "dissolve",
                "fade", "pull", "push", "random", "randomBar", "split",
                "strips", "wedge", "wheel", "wipe", "zoom", "circle",
                "diamond", "newsflash", "plus",
            ):
                return local
        return None
