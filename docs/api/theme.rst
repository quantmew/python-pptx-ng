.. _theme:

Theme
=====

.. module:: pptx.theme

The theme API provides read access to color and font schemes defined in a
presentation's theme part.

Theme
-----

.. class:: Theme(theme_element)

   Proxy for the ``a:theme`` element. Access via
   ``slide_master.part.part_related_by(RT.THEME).theme``.

   .. attribute:: color_scheme

      Return a |ColorScheme| proxy for this theme.

   .. attribute:: font_scheme

      Return a |FontScheme| proxy for this theme.

ColorScheme
-----------

.. class:: ColorScheme(clrScheme)

   Proxy for an ``a:clrScheme`` element.

   .. attribute:: name

      Return the name of this color scheme.

   .. method:: get_color(color_name)

      Return the RGB hex string for a named theme color (e.g. ``"dk1"``,
      ``"accent1"``), if available as ``srgbClr``. Returns |None| for
      system colors.

FontScheme
----------

.. class:: FontScheme(fontScheme)

   Proxy for an ``a:fontScheme`` element.

   .. attribute:: name

      Return the name of this font scheme.

   .. attribute:: major_font

      Return the Latin typeface name of the major (heading) font.

   .. attribute:: minor_font

      Return the Latin typeface name of the minor (body) font.
