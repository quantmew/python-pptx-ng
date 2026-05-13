"""Enumerations used by slide transition objects."""

from __future__ import annotations

from pptx_ng.enum.base import BaseXmlEnum


class PP_TRANSITION_SPEED(BaseXmlEnum):
    """Transition speed.

    Corresponds to the ``spd`` attribute on ``p:transition``.

    MS API Name: `MsoAnimTransitionSpeed`

    Example::

        from pptx_ng.enum.transition import PP_TRANSITION_SPEED

        slide.transition.speed = PP_TRANSITION_SPEED.FAST
    """

    SLOW = (0, "slow", "Slow transition speed.")
    MEDIUM = (1, "med", "Medium transition speed.")
    FAST = (2, "fast", "Fast transition speed.")


class PP_TRANSITION_SIDE_DIRECTION(BaseXmlEnum):
    """Side direction for transitions like push, wipe.

    Corresponds to the ``dir`` attribute on ``p:push``, ``p:wipe``, etc.

    MS API Name: `TransitionSlideDirectionValues`
    """

    LEFT = (0, "l", "Left.")
    UP = (1, "u", "Up.")
    RIGHT = (2, "r", "Right.")
    DOWN = (3, "d", "Down.")


class PP_TRANSITION_CORNER_DIRECTION(BaseXmlEnum):
    """Corner direction for cover, pull, strips transitions.

    Corresponds to the ``dir`` attribute on ``p:cover``, ``p:pull``, ``p:strips``.

    MS API Name: `TransitionCornerDirectionValues`
    """

    LEFT_UP = (0, "lu", "Left-up.")
    RIGHT_UP = (1, "ru", "Right-up.")
    LEFT_DOWN = (2, "ld", "Left-down.")
    RIGHT_DOWN = (3, "rd", "Right-down.")


class PP_TRANSITION_ORIENTATION(BaseXmlEnum):
    """Orientation for blinds, checker, comb, randomBar, split transitions.

    Corresponds to the ``dir`` or ``orient`` attribute.

    MS API Name: `DirectionValues`
    """

    HORIZONTAL = (0, "horz", "Horizontal orientation.")
    VERTICAL = (1, "vert", "Vertical orientation.")


class PP_TRANSITION_IN_OUT_DIRECTION(BaseXmlEnum):
    """In/out direction for zoom and split transitions.

    Corresponds to the ``dir`` attribute on ``p:zoom`` and ``p:split``.

    MS API Name: `TransitionInOutDirectionValues`
    """

    OUT = (0, "out", "Outward direction.")
    IN = (1, "in", "Inward direction.")
