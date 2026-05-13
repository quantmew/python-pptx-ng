"""High-level API for mathematical formula support (OMML).

Provides the |MathFormula| proxy class for creating and manipulating
mathematical formulas in PowerPoint presentations using Office Math
Markup Language (OMML).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pptx_ng.oxml.math.core import (
        CT_OMath,
        CT_OMathAccent,
        CT_OMathBar,
        CT_OMathBase,
        CT_OMathBorderBox,
        CT_OMathBox,
        CT_OMathDelimiter,
        CT_OMathEqArr,
        CT_OMathFrac,
        CT_OMathFunction,
        CT_OMathGroupChar,
        CT_OMathLimitLower,
        CT_OMathLimitUpper,
        CT_OMathMatrix,
        CT_OMathNary,
        CT_OMathRad,
        CT_OMathRun,
        CT_OMathSub,
        CT_OMathSubSup,
        CT_OMathSup,
    )


class MathFormula:
    """Proxy for an `m:oMath` element providing a high-level API for
    constructing mathematical formulas.

    Use |Paragraph.add_math()| to create a new |MathFormula| instance.
    """

    def __init__(self, oMath: CT_OMath):
        super(MathFormula, self).__init__()
        self._element = self._oMath = oMath

    def add_run(self, text: str = "") -> MathRun:
        """Append a text run to this formula and return its proxy."""
        r = self._oMath.add_run(text)
        return MathRun(r)

    def add_fraction(self) -> MathFraction:
        """Append a fraction (a/b) structure and return its proxy."""
        f = self._oMath.add_fraction()
        return MathFraction(f)

    def add_radical(self) -> MathRadical:
        """Append a radical (square root) structure and return its proxy."""
        rad = self._oMath.add_radical()
        return MathRadical(rad)

    def add_superscript(self) -> MathSuperscript:
        """Append a superscript (x^2) structure and return its proxy."""
        sSup = self._oMath.add_superscript()
        return MathSuperscript(sSup)

    def add_subscript(self) -> MathSubscript:
        """Append a subscript (x_i) structure and return its proxy."""
        sSub = self._oMath.add_subscript()
        return MathSubscript(sSub)

    def add_sub_superscript(self) -> MathSubSuperscript:
        """Append a sub-superscript (x_i^2) structure and return its proxy."""
        sSubSup = self._oMath.add_sub_superscript()
        return MathSubSuperscript(sSubSup)

    def add_nary(self, chr_val: str = "∫") -> MathNary:
        """Append an n-ary operator (integral, sum, product) structure.

        The default operator is the integral sign (∫). Use `chr_val` to
        specify a different operator character.
        """
        nary = self._oMath.add_nary(chr_val)
        return MathNary(nary)

    def add_delimiter(self) -> MathDelimiter:
        """Append a delimiter (parentheses) structure and return its proxy."""
        d = self._oMath.add_delimiter()
        return MathDelimiter(d)

    def add_accent(self, chr_val: str = "̂") -> MathAccent:
        """Append an accent (hat, tilde) structure.

        The default accent is the combining circumflex (^).
        """
        acc = self._oMath.add_accent(chr_val)
        return MathAccent(acc)

    def add_matrix(self, rows: int = 1, cols: int = 1) -> MathMatrix:
        """Append a matrix with the specified dimensions and return its proxy."""
        m = self._oMath.add_matrix(rows, cols)
        return MathMatrix(m)

    def add_function(self, name: str = "") -> MathFunction:
        """Append a function structure and return its proxy."""
        func = self._oMath.add_function(name)
        return MathFunction(func)

    @property
    def xml(self) -> str:
        """The XML of this formula element."""
        from lxml import etree

        return etree.tostring(self._oMath, encoding="unicode")


class MathRun:
    """Proxy for an `m:r` (math run) element."""

    def __init__(self, r: CT_OMathRun):
        self._r = r

    @property
    def text(self) -> str:
        return self._r.t

    @text.setter
    def text(self, value: str):
        self._r.t = value


class MathFraction:
    """Proxy for an `m:f` (fraction) element."""

    def __init__(self, f: CT_OMathFrac):
        self._f = f

    @property
    def numerator(self) -> MathArgument:
        return MathArgument(self._f.num)

    @property
    def denominator(self) -> MathArgument:
        return MathArgument(self._f.den)


class MathRadical:
    """Proxy for an `m:rad` (radical) element."""

    def __init__(self, rad: CT_OMathRad):
        self._rad = rad

    @property
    def base(self) -> MathArgument:
        return MathArgument(self._rad.e)

    @property
    def degree(self) -> MathArgument:
        return MathArgument(self._rad.deg)


class MathSuperscript:
    """Proxy for an `m:sSup` (superscript) element."""

    def __init__(self, sSup: CT_OMathSup):
        self._sSup = sSup

    @property
    def base(self) -> MathArgument:
        return MathArgument(self._sSup.e)

    @property
    def superscript(self) -> MathArgument:
        return MathArgument(self._sSup.sup)


class MathSubscript:
    """Proxy for an `m:sSub` (subscript) element."""

    def __init__(self, sSub: CT_OMathSub):
        self._sSub = sSub

    @property
    def base(self) -> MathArgument:
        return MathArgument(self._sSub.e)

    @property
    def subscript(self) -> MathArgument:
        return MathArgument(self._sSub.sub)


class MathSubSuperscript:
    """Proxy for an `m:sSubSup` element."""

    def __init__(self, sSubSup: CT_OMathSubSup):
        self._sSubSup = sSubSup

    @property
    def base(self) -> MathArgument:
        return MathArgument(self._sSubSup.e)

    @property
    def subscript(self) -> MathArgument:
        return MathArgument(self._sSubSup.sub)

    @property
    def superscript(self) -> MathArgument:
        return MathArgument(self._sSubSup.sup)


class MathNary:
    """Proxy for an `m:nary` (n-ary operator) element."""

    def __init__(self, nary: CT_OMathNary):
        self._nary = nary

    @property
    def base(self) -> MathArgument:
        return MathArgument(self._nary.e)

    @property
    def lower_limit(self) -> MathArgument:
        return MathArgument(self._nary.sub)

    @property
    def upper_limit(self) -> MathArgument:
        return MathArgument(self._nary.sup)


class MathDelimiter:
    """Proxy for an `m:d` (delimiter) element."""

    def __init__(self, d: CT_OMathDelimiter):
        self._d = d

    def add_argument(self) -> MathArgument:
        """Add a new argument (e) to this delimiter and return its proxy."""
        from pptx_ng.oxml.math.core import CT_OMathBase

        e = CT_OMathBase.new()  # type: ignore[attr-defined]
        self._d.append(e)
        return MathArgument(e)


class MathAccent:
    """Proxy for an `m:acc` (accent) element."""

    def __init__(self, acc: CT_OMathAccent):
        self._acc = acc

    @property
    def base(self) -> MathArgument:
        return MathArgument(self._acc.e)


class MathMatrix:
    """Proxy for an `m:m` (matrix) element."""

    def __init__(self, m: CT_OMathMatrix):
        self._m = m

    def cell(self, row: int, col: int) -> MathArgument:
        """Return a |MathArgument| proxy for the cell at (*row*, *col*)."""
        rows = self._m.mr_lst
        if row >= len(rows):
            raise IndexError(f"row index {row} out of range (have {len(rows)} rows)")
        cols = rows[row].e_lst
        if col >= len(cols):
            raise IndexError(f"col index {col} out of range (have {len(cols)} cols)")
        return MathArgument(cols[col])


class MathFunction:
    """Proxy for an `m:func` (function) element."""

    def __init__(self, func: CT_OMathFunction):
        self._func = func

    @property
    def name(self) -> MathArgument:
        return MathArgument(self._func.fName)

    @property
    def body(self) -> MathArgument:
        return MathArgument(self._func.e)


class MathArgument:
    """Proxy for math argument elements (e, num, den, deg, sub, sup, etc.).

    These containers can hold any math content (runs, fractions, etc.).
    """

    def __init__(self, arg: CT_OMathBase):
        self._arg = arg

    def add_run(self, text: str = "") -> MathRun:
        """Append a text run to this argument."""
        r = self._arg.add_run(text)
        return MathRun(r)

    def add_fraction(self) -> MathFraction:
        """Append a fraction to this argument."""
        f = self._arg.add_fraction()
        return MathFraction(f)

    def add_superscript(self) -> MathSuperscript:
        """Append a superscript to this argument."""
        sSup = self._arg.add_superscript()
        return MathSuperscript(sSup)

    def add_subscript(self) -> MathSubscript:
        """Append a subscript to this argument."""
        sSub = self._arg.add_subscript()
        return MathSubscript(sSub)

    def add_nary(self, chr_val: str = "∫") -> MathNary:
        """Append an n-ary operator to this argument."""
        nary = self._arg.add_nary(chr_val)
        return MathNary(nary)

    def add_radical(self) -> MathRadical:
        """Append a radical to this argument."""
        rad = self._arg.add_radical()
        return MathRadical(rad)
