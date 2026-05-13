.. _diagram_api:

SmartArt Diagrams
=================

python-pptx-ng supports adding SmartArt graphics to slides using the
:meth:`add_smartart` method of |SlideShapes|.

Usage
-----

Add a SmartArt graphic to a slide::

    from pptx_ng import Presentation
    from pptx_ng.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    gf = slide.shapes.add_smartart(
        Inches(1), Inches(1), Inches(6), Inches(4)
    )


|SmartArtData| objects
----------------------

.. autoclass:: pptx.diagram.SmartArtData()
   :members:
   :undoc-members:


|SmartArtNode| objects
----------------------

.. autoclass:: pptx.diagram.SmartArtNode()
   :members:
   :undoc-members:
