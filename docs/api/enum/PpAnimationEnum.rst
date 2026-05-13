.. _PpAnimationEnum:

Animation Enumerations
======================

``PP_TIMING_NODE_TYPE``
------------------------

Time node type for animations.

Example::

    from pptx_ng.enum.animation import PP_TIMING_NODE_TYPE

----

CLICK_EFFECT
    Click effect.

WITH_EFFECT
    With effect.

AFTER_EFFECT
    After effect.

MAIN_SEQUENCE
    Main sequence.

INTERACTIVE_SEQUENCE
    Interactive sequence.

CLICK_PARAGRAPH
    Click paragraph.

WITH_GROUP
    With group.

AFTER_GROUP
    After group.

TIMING_ROOT
    Timing root.


``PP_ANIMATION_PRESET_CLASS``
-------------------------------

Animation preset class.

----

ENTRANCE
    Entrance animation.

EXIT
    Exit animation.

EMPHASIS
    Emphasis animation.

MOTION
    Motion path animation.

MEDIA
    Media animation.


``PP_TIME_NODE_FILL``
-----------------------

Time node fill mode.

----

FREEZE
    Freeze.

HOLD
    Hold.

REMOVE
    Remove.

TRANSITION
    Transition.


``PP_ANIMATION_EFFECT_TRANSITION``
------------------------------------

Effect transition type (in/out).

----

IN
    In transition.

OUT
    Out transition.

NONE
    No transition.


``PP_PARAGRAPH_BUILD_TYPE``
-----------------------------

Paragraph build type for animation.

----

ALL_AT_ONCE
    All at once.

PARAGRAPH
    By paragraph.

CUSTOM
    Custom build.

WHOLE
    Whole.
