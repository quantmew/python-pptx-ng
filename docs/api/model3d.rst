.. _model3d_api:

3D Models
=========

python-pptx-ng supports embedding 3D models (GLTF Binary format, ``.glb``
files) in slides. This feature requires Office 2019 or later.

Usage
-----

Embed a 3D model file on a slide::

    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    gf = slide.shapes.add_3d_model(
        "model.glb", Inches(1), Inches(1), Inches(4), Inches(4)
    )


|Model3D| objects
-----------------

.. autoclass:: pptx.model3d.Model3D()
   :members:
   :undoc-members:


|Model3DCamera| objects
-----------------------

.. autoclass:: pptx.model3d.Model3DCamera()
   :members:
   :undoc-members:


|Model3DTransform| objects
--------------------------

.. autoclass:: pptx.model3d.Model3DTransform()
   :members:
   :undoc-members:
