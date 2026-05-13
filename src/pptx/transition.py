"""Slide transition proxy class.

Provides high-level API for working with slide transitions.
"""

from __future__ import annotations

from pptx.oxml.transition import CT_Transition
from pptx.oxml.xmlchemy import OxmlElement


class Transition:
    """Slide transition settings.

    Accessed via `Slide.transition`. Use `Slide.add_transition()` to create one.
    """

    def __init__(self, transition: CT_Transition):
        self._element = transition

    @property
    def speed(self) -> str | None:
        """Transition speed: "slow", "med", or "fast". Read/write."""
        return self._element.spd

    @speed.setter
    def speed(self, value: str | None):
        self._element.spd = value

    @property
    def advance_on_click(self) -> bool | None:
        """Whether clicking advances the slide. Read/write."""
        return self._element.advClick

    @advance_on_click.setter
    def advance_on_click(self, value: bool | None):
        self._element.advClick = value

    @property
    def advance_after_time(self) -> int | None:
        """Auto-advance after this many milliseconds. Read/write."""
        val = self._element.advTm
        return int(val) if val is not None else None

    @advance_after_time.setter
    def advance_after_time(self, value: int | None):
        self._element.advTm = str(value) if value is not None else None

    @property
    def transition_type(self) -> str | None:
        """The active transition type name (e.g. 'fade', 'push')."""
        return self._element.transition_type

    def set_blinds(self, dir: str = "horz"):
        self._element.set_blinds(dir)

    def set_checker(self, dir: str = "horz"):
        self._element.set_checker(dir)

    def set_comb(self, dir: str = "horz"):
        self._element.set_comb(dir)

    def set_cover(self, dir: str = "l"):
        self._element.set_cover(dir)

    def set_pull(self, dir: str = "l"):
        self._element.set_pull(dir)

    def set_cut(self, thruBlk: bool = False):
        self._element.set_cut(thruBlk)

    def set_fade(self, thruBlk: bool = False):
        self._element.set_fade(thruBlk)

    def set_push(self, dir: str = "l"):
        self._element.set_push(dir)

    def set_wipe(self, dir: str = "l"):
        self._element.set_wipe(dir)

    def set_split(self, orient: str = "horz", dir: str = "out"):
        self._element.set_split(orient, dir)

    def set_strips(self, dir: str = "ld"):
        self._element.set_strips(dir)

    def set_random_bar(self, dir: str = "horz"):
        self._element.set_random_bar(dir)

    def set_wheel(self, spokes: int = 4):
        self._element.set_wheel(spokes)

    def set_wedge(self):
        self._element.set_wedge()

    def set_zoom(self, dir: str = "out"):
        self._element.set_zoom(dir)

    def set_dissolve(self):
        self._element.set_dissolve()

    def set_random(self):
        self._element.set_random()

    def set_circle(self):
        self._element.set_circle()

    def set_diamond(self):
        self._element.set_diamond()

    def set_newsflash(self):
        self._element.set_newsflash()

    def set_plus(self):
        self._element.set_plus()
