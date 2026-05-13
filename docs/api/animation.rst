.. _animation_api:

Animations
==========


|AnimationTimeline| objects
----------------------------

The |AnimationTimeline| for a slide is accessed via
:attr:`~pptx.slide.Slide.timing`::

    slide = prs.slides[0]
    timing = slide.timing
    timing.add_appear_effect(shape_id)
    timing.add_fade_effect(shape_id, delay="500", dur="1000")

.. autoclass:: pptx.animation.AnimationTimeline()
   :members:
   :undoc-members:
