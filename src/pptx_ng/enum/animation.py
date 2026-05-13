"""Enumerations used by animation objects."""

from __future__ import annotations

from pptx_ng.enum.base import BaseXmlEnum


class PP_TIMING_NODE_TYPE(BaseXmlEnum):
    """Time node type.

    Corresponds to the ``nodeType`` attribute on ``p:cTn``.

    MS API Name: `TimeNodeValues`
    """

    CLICK_EFFECT = (0, "clickEffect", "Click effect.")
    WITH_EFFECT = (1, "withEffect", "With effect.")
    AFTER_EFFECT = (2, "afterEffect", "After effect.")
    MAIN_SEQUENCE = (3, "mainSeq", "Main sequence.")
    INTERACTIVE_SEQUENCE = (4, "interactiveSeq", "Interactive sequence.")
    CLICK_PARAGRAPH = (5, "clickPar", "Click paragraph.")
    WITH_GROUP = (6, "withGroup", "With group.")
    AFTER_GROUP = (7, "afterGroup", "After group.")
    TIMING_ROOT = (8, "tmRoot", "Timing root.")


class PP_ANIMATION_PRESET_CLASS(BaseXmlEnum):
    """Animation preset class.

    Corresponds to the ``presetClass`` attribute on ``p:cTn``.

    MS API Name: Used in ``presetClass`` attribute.
    """

    ENTRANCE = (0, "entr", "Entrance animation.")
    EXIT = (1, "exit", "Exit animation.")
    EMPHASIS = (2, "emph", "Emphasis animation.")
    MOTION = (3, "motion", "Motion path animation.")
    MEDIA = (4, "medi", "Media animation.")


class PP_TIME_NODE_FILL(BaseXmlEnum):
    """Time node fill mode.

    Corresponds to the ``fill`` attribute on ``p:cTn``.

    MS API Name: `TimeNodeFillValues`
    """

    FREEZE = (0, "freeze", "Freeze.")
    HOLD = (1, "hold", "Hold.")
    REMOVE = (2, "remove", "Remove.")
    TRANSITION = (3, "transition", "Transition.")


class PP_ANIMATION_EFFECT_TRANSITION(BaseXmlEnum):
    """Effect transition type (in/out).

    Corresponds to the ``transition`` attribute on ``p:animEffect``.

    MS API Name: `AnimateEffectTransitionValues`
    """

    IN = (0, "in", "In transition.")
    OUT = (1, "out", "Out transition.")
    NONE = (2, "none", "No transition.")


class PP_PARAGRAPH_BUILD_TYPE(BaseXmlEnum):
    """Paragraph build type for animation.

    Corresponds to the ``build`` attribute on ``p:bldP``.

    MS API Name: `ParagraphBuildValues`
    """

    ALL_AT_ONCE = (0, "allAtOnce", "All at once.")
    PARAGRAPH = (1, "p", "By paragraph.")
    CUSTOM = (2, "cust", "Custom build.")
    WHOLE = (3, "whole", "Whole.")


class PP_BEHAVIOR_ADDITIVE(BaseXmlEnum):
    """Behavior additive mode.

    Corresponds to the ``additive`` attribute on ``p:cBhvr``.

    MS API Name: `BehaviorAdditiveValues`
    """

    BASE = (0, "base", "Base.")
    SUM = (1, "sum", "Sum.")
    REPLACE = (2, "repl", "Replace.")
    MULTIPLY = (3, "mult", "Multiply.")
    NONE = (4, "none", "None.")
