.. _tablestyle:

Table Styles
============

.. module:: pptx.tablestyle

The table styles API provides access to the table style collection stored in
a presentation's table styles part.

TableStyles
-----------

.. class:: TableStyles(tblStyleLst)

   Proxy for the table styles collection (``a:tblStyleLst`` element).

   .. attribute:: default_style_id

      Return or set the GUID string of the default table style. This value
      is used when a table does not specify an explicit style.
