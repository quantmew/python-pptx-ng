.. _transition_api:

Transitions
============


|Transition| objects
--------------------

A |Transition| object is created via :meth:`~pptx.slide.Slide.add_transition`
and read via :attr:`~pptx.slide.Slide.transition`::

    slide = prs.slides[0]
    transition = slide.add_transition()
    transition.set_fade()
    transition.speed = "med"

.. autoclass:: pptx.transition.Transition()
   :members:
   :undoc-members:
