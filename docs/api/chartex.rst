.. _chartex_api:

Extended Charts (chartEx)
=========================

python-pptx-ng supports adding modern chart types (treemap, sunburst, waterfall,
funnel, region map, box-and-whisker) to slides using the chartEx format
introduced in Office 2016.

Usage
-----

Add a treemap chart to a slide::

    from pptx_ng import Presentation
    from pptx_ng.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    gf = slide.shapes.add_chartex(
        "treemap", Inches(1), Inches(1), Inches(6), Inches(4)
    )

Supported chart types: ``treemap``, ``sunburst``, ``waterfall``, ``funnel``,
``regionMap``, ``boxWhisker``.


|ChartEx| objects
-----------------

.. autoclass:: pptx.chartex.ChartEx()
   :members:
   :undoc-members:
