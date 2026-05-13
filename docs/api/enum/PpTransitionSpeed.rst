.. _PpTransitionSpeed:

``PP_TRANSITION_SPEED``
=======================

Transition speed.

Corresponds to the ``spd`` attribute on ``p:transition``.

Example::

    from pptx_ng.enum.transition import PP_TRANSITION_SPEED

    slide.transition.speed = PP_TRANSITION_SPEED.FAST

----

SLOW
    Slow transition speed.

MEDIUM
    Medium transition speed.

FAST
    Fast transition speed.
