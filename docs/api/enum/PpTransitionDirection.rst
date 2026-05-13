.. _PpTransitionDirection:

Transition Direction Enumerations
==================================

``PP_TRANSITION_SIDE_DIRECTION``
--------------------------------

Side direction for push, wipe transitions.

Example::

    from pptx_ng.enum.transition import PP_TRANSITION_SIDE_DIRECTION

    slide.transition.set_push(PP_TRANSITION_SIDE_DIRECTION.LEFT)

----

LEFT
    Left direction.

UP
    Up direction.

RIGHT
    Right direction.

DOWN
    Down direction.


``PP_TRANSITION_CORNER_DIRECTION``
-----------------------------------

Corner direction for cover, pull, strips transitions.

----

LEFT_UP
    Left-up corner direction.

RIGHT_UP
    Right-up corner direction.

LEFT_DOWN
    Left-down corner direction.

RIGHT_DOWN
    Right-down corner direction.


``PP_TRANSITION_ORIENTATION``
-----------------------------

Orientation for blinds, checker, comb, split transitions.

----

HORIZONTAL
    Horizontal orientation.

VERTICAL
    Vertical orientation.


``PP_TRANSITION_IN_OUT_DIRECTION``
-----------------------------------

In/out direction for zoom and split transitions.

----

OUT
    Outward direction.

IN
    Inward direction.
