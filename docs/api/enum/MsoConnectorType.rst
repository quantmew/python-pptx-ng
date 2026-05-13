.. _MsoConnectorType:

``MSO_CONNECTOR_TYPE``
======================

Specifies a type of connector.

Alias: ``MSO_CONNECTOR``

Example::

    from pptx_ng.enum.shapes import MSO_CONNECTOR
    from pptx_ng.util import Cm

    shapes = prs.slides[0].shapes
    connector = shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Cm(2), Cm(2), Cm(10), Cm(10)
    )
    assert connector.left.cm == 2

----

CURVE
    Curved connector.

ELBOW
    Elbow connector.

STRAIGHT
    Straight line connector.

MIXED
    Return value only; indicates a combination of other states.
