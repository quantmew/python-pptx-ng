.. _math_api:

Math Formulas (OMML)
====================

python-pptx-ng supports embedding mathematical formulas in slide text using
Office Math Markup Language (OMML). Formulas are added to paragraphs and can
contain fractions, radicals, superscripts, subscripts, integrals, matrices,
and more.

**PowerPoint compatibility note:** python-pptx-ng automatically wraps OMML
formulas in the ``<a14:m>`` element required by PowerPoint for proper
rendering in DrawingML text frames. You do not need to handle this manually.

Usage
-----

Add an inline math formula to a paragraph::

    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(4))
    tf = txBox.text_frame

Build E = mc² (superscript)::

    p = tf.paragraphs[0]
    math = p.add_math()
    math.add_run("E=m")
    sSup = math.add_superscript()
    sSup.base.add_run("c")
    sSup.superscript.add_run("2")

Build a fraction::

    p2 = tf.add_paragraph()
    math = p2.add_math()
    frac = math.add_fraction()
    frac.numerator.add_run("a")
    frac.denominator.add_run("b")

Build an integral ∫₀¹ x dx::

    p3 = tf.add_paragraph()
    math = p3.add_math()
    nary = math.add_nary("∫")
    nary.lower_limit.add_run("0")
    nary.upper_limit.add_run("1")
    nary.base.add_run("x dx")

Build a radical √(x+1)::

    p4 = tf.add_paragraph()
    math = p4.add_math()
    rad = math.add_radical()
    rad.base.add_run("x+1")

Build a 2×2 matrix::

    p5 = tf.add_paragraph()
    math = p5.add_math()
    m = math.add_matrix(2, 2)
    m.cell(0, 0).add_run("a")
    m.cell(0, 1).add_run("b")
    m.cell(1, 0).add_run("c")
    m.cell(1, 1).add_run("d")

Build x_i^n (sub-superscript)::

    p6 = tf.add_paragraph()
    math = p6.add_math()
    sSubSup = math.add_sub_superscript()
    sSubSup.base.add_run("x")
    sSubSup.subscript.add_run("i")
    sSubSup.superscript.add_run("n")

Build sin(x) (function)::

    p7 = tf.add_paragraph()
    math = p7.add_math()
    func = math.add_function("sin")
    func.body.add_run("x")

Build (a+b) (delimiter)::

    p8 = tf.add_paragraph()
    math = p8.add_math()
    delim = math.add_delimiter()
    arg = delim.add_argument()
    arg.add_run("a+b")

Build the quadratic formula x = (-b ± √(b²-4ac)) / 2a::

    p9 = tf.add_paragraph()
    math = p9.add_math()
    math.add_run("x=")
    frac = math.add_fraction()
    frac.numerator.add_run("-b±")
    rad = frac.numerator.add_radical()
    ss = rad.base.add_superscript()
    ss.base.add_run("b")
    ss.superscript.add_run("2")
    rad.base.add_run("-4ac")
    frac.denominator.add_run("2a")


|MathFormula| objects
---------------------

.. autoclass:: pptx.math.MathFormula()
   :members:
   :undoc-members:


|MathRun| objects
-----------------

.. autoclass:: pptx.math.MathRun()
   :members:
   :undoc-members:


|MathFraction| objects
----------------------

.. autoclass:: pptx.math.MathFraction()
   :members:
   :undoc-members:


|MathRadical| objects
---------------------

.. autoclass:: pptx.math.MathRadical()
   :members:
   :undoc-members:


|MathSuperscript| objects
-------------------------

.. autoclass:: pptx.math.MathSuperscript()
   :members:
   :undoc-members:


|MathSubscript| objects
-----------------------

.. autoclass:: pptx.math.MathSubscript()
   :members:
   :undoc-members:


|MathSubSuperscript| objects
----------------------------

.. autoclass:: pptx.math.MathSubSuperscript()
   :members:
   :undoc-members:


|MathNary| objects
------------------

.. autoclass:: pptx.math.MathNary()
   :members:
   :undoc-members:


|MathMatrix| objects
--------------------

.. autoclass:: pptx.math.MathMatrix()
   :members:
   :undoc-members:


|MathDelimiter| objects
-----------------------

.. autoclass:: pptx.math.MathDelimiter()
   :members:
   :undoc-members:


|MathAccent| objects
--------------------

.. autoclass:: pptx.math.MathAccent()
   :members:
   :undoc-members:


|MathFunction| objects
----------------------

.. autoclass:: pptx.math.MathFunction()
   :members:
   :undoc-members:


|MathArgument| objects
----------------------

.. autoclass:: pptx.math.MathArgument()
   :members:
   :undoc-members:
