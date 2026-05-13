"""Presentation properties proxy classes.

Provides high-level API for presentation show and printing properties.
"""

from __future__ import annotations

from pptx_ng.oxml.presprops import CT_ShowProperties
from pptx_ng.shared import PartElementProxy
from pptx_ng.util import lazyproperty


class PresentationProperties(PartElementProxy):
    """Presentation-wide properties.

    Accessed via `Presentation.properties`.
    """

    @lazyproperty
    def show_properties(self) -> ShowProperties:
        """The |ShowProperties| for this presentation.

        If show properties do not exist, they are created.
        """
        showPr = self._element.get_or_add_showPr()
        return ShowProperties(showPr, self)


class ShowProperties:
    """Slide show properties.

    Accessed via `PresentationProperties.show_properties`.
    """

    def __init__(self, showPr: CT_ShowProperties, parent: PresentationProperties):
        self._element = showPr
        self._parent = parent

    @property
    def loop(self) -> bool | None:
        """Whether to loop the slide show continuously.

        Read/write. Returns |None| if not set.
        """
        return self._element.loop

    @loop.setter
    def loop(self, value: bool | None):
        self._element.loop = value

    @property
    def show_animation(self) -> bool | None:
        """Whether to show animations during the slide show.

        Read/write. Returns |None| if not set.
        """
        return self._element.showAnimation

    @show_animation.setter
    def show_animation(self, value: bool | None):
        self._element.showAnimation = value

    @property
    def show_narration(self) -> bool | None:
        """Whether to show narration during the slide show.

        Read/write. Returns |None| if not set.
        """
        return self._element.showNarration

    @show_narration.setter
    def show_narration(self, value: bool | None):
        self._element.showNarration = value

    @property
    def use_timings(self) -> bool | None:
        """Whether to use timings during the slide show.

        Read/write. Returns |None| if not set.
        """
        return self._element.useTimings

    @use_timings.setter
    def use_timings(self, value: bool | None):
        self._element.useTimings = value
