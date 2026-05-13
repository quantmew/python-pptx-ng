"""Core OMML math element classes.

Contains the main container elements (oMath, oMathPara), run/text elements,
and argument containers (e, num, den, deg, sub, sup).
"""

from __future__ import annotations

from typing import Callable, cast

from lxml import etree

from pptx_ng.oxml.ns import nsdecls, qn
from pptx_ng.oxml.simpletypes import XsdString
from pptx_ng.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


def _parse(xml: str) -> BaseOxmlElement:
    from pptx_ng.oxml import parse_xml

    return parse_xml(xml)


# ---------------------------------------------------------------------------
# Container elements
# ---------------------------------------------------------------------------


class CT_OMath(BaseOxmlElement):
    """`m:oMath` element — inline math formula container."""

    def add_run(self, text: str = "") -> CT_OMathRun:
        r = cast(CT_OMathRun, _parse("<m:r %s><m:t/></m:r>" % nsdecls("m")))
        if text:
            r.t = text
        self.append(r)
        return r

    def add_fraction(self) -> CT_OMathFrac:
        f = cast(
            CT_OMathFrac,
            _parse("<m:f %s><m:fPr/><m:num/><m:den/></m:f>" % nsdecls("m")),
        )
        self.append(f)
        return f

    def add_radical(self) -> CT_OMathRad:
        rad = cast(
            CT_OMathRad,
            _parse(
                "<m:rad %s>"
                "<m:radPr/>"
                "<m:e/>"
                "</m:rad>" % nsdecls("m")
            ),
        )
        self.append(rad)
        return rad

    def add_superscript(self) -> CT_OMathSup:
        sSup = cast(
            CT_OMathSup,
            _parse("<m:sSup %s><m:sSupPr/><m:e/><m:sup/></m:sSup>" % nsdecls("m")),
        )
        self.append(sSup)
        return sSup

    def add_subscript(self) -> CT_OMathSub:
        sSub = cast(
            CT_OMathSub,
            _parse("<m:sSub %s><m:sSubPr/><m:e/><m:sub/></m:sSub>" % nsdecls("m")),
        )
        self.append(sSub)
        return sSub

    def add_sub_superscript(self) -> CT_OMathSubSup:
        sSubSup = cast(
            CT_OMathSubSup,
            _parse(
                "<m:sSubSup %s><m:sSubSupPr/><m:e/><m:sub/><m:sup/></m:sSubSup>"
                % nsdecls("m")
            ),
        )
        self.append(sSubSup)
        return sSubSup

    def add_nary(self, chr_val: str = "∫") -> CT_OMathNary:
        nary = cast(
            CT_OMathNary,
            _parse(
                "<m:nary %s><m:naryPr/><m:sub/><m:sup/><m:e/></m:nary>" % nsdecls("m")
            ),
        )
        nary_pr = nary.get_or_add_naryPr()
        nary_pr.chr = chr_val
        self.append(nary)
        return nary

    def add_delimiter(self) -> CT_OMathDelimiter:
        d = cast(
            CT_OMathDelimiter,
            _parse("<m:d %s><m:dPr/><m:e/></m:d>" % nsdecls("m")),
        )
        self.append(d)
        return d

    def add_accent(self, chr_val: str = "̂") -> CT_OMathAccent:
        acc = cast(
            CT_OMathAccent,
            _parse("<m:acc %s><m:accPr/><m:e/></m:acc>" % nsdecls("m")),
        )
        acc_pr = acc.get_or_add_accPr()
        acc_pr.chr = chr_val
        self.append(acc)
        return acc

    def add_matrix(self, rows: int = 1, cols: int = 1) -> CT_OMathMatrix:
        m = cast(
            CT_OMathMatrix,
            _parse(
                "<m:m %s>"
                "<m:mPr>"
                "<m:baseJc m:val=\"center\"/>"
                "<m:mcs>"
                "<m:mc><m:mcPr><m:count m:val=\"%d\"/><m:mcJc m:val=\"center\"/></m:mcPr></m:mc>"
                "</m:mcs>"
                "</m:mPr>"
                "</m:m>" % (nsdecls("m"), cols)
            ),
        )
        for _ in range(rows):
            mr = cast(CT_OMathMatrixRow, _parse("<m:mr %s/>" % nsdecls("m")))
            for _ in range(cols):
                e = cast(CT_OMathBase, _parse("<m:e %s/>" % nsdecls("m")))
                mr.append(e)
            m.append(mr)
        self.append(m)
        return m

    def add_function(self, name: str = "") -> CT_OMathFunction:
        func = cast(
            CT_OMathFunction,
            _parse("<m:func %s><m:funcPr/><m:fName/><m:e/></m:func>" % nsdecls("m")),
        )
        if name:
            fname = func.fName
            fname.add_run(name)
        self.append(func)
        return func

    # -- children --
    r = ZeroOrMore("m:r")
    f = ZeroOrMore("m:f")
    rad = ZeroOrMore("m:rad")
    sSup = ZeroOrMore("m:sSup")
    sSub = ZeroOrMore("m:sSub")
    sSubSup = ZeroOrMore("m:sSubSup")
    nary = ZeroOrMore("m:nary")
    d = ZeroOrMore("m:d")
    acc = ZeroOrMore("m:acc")
    bar = ZeroOrMore("m:bar")
    box = ZeroOrMore("m:box")
    borderBox = ZeroOrMore("m:borderBox")
    eqArr = ZeroOrMore("m:eqArr")
    func = ZeroOrMore("m:func")
    groupChr = ZeroOrMore("m:groupChr")
    limLow = ZeroOrMore("m:limLow")
    limUpp = ZeroOrMore("m:limUpp")
    m = ZeroOrMore("m:m")

    @classmethod
    def new(cls) -> CT_OMath:
        return cast(CT_OMath, _parse("<m:oMath %s/>" % nsdecls("m")))


class CT_OMathPara(BaseOxmlElement):
    """`m:oMathPara` element — math paragraph container."""

    oMath = ZeroOrMore("m:oMath")

    @classmethod
    def new(cls) -> CT_OMathPara:
        return cast(CT_OMathPara, _parse("<m:oMathPara %s/>" % nsdecls("m")))


# ---------------------------------------------------------------------------
# Run and text elements
# ---------------------------------------------------------------------------


class CT_OMathRun(BaseOxmlElement):
    """`m:r` element — math run containing text."""

    _tag_seq = ("m:rPr", "m:t")
    rPr = ZeroOrOne("m:rPr", successors=("m:t",))
    del _tag_seq

    @property
    def t(self) -> str:
        t_el = self.find(qn("m:t"))
        if t_el is None:
            return ""
        return t_el.text or ""

    @t.setter
    def t(self, value: str):
        t_el = self.find(qn("m:t"))
        if t_el is None:
            t_el = etree.SubElement(self, qn("m:t"))
        t_el.text = value

    @classmethod
    def new(cls, text: str = "") -> CT_OMathRun:
        escaped = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        xml = "<m:r %s><m:t>%s</m:t></m:r>" % (nsdecls("m"), escaped)
        return cast(CT_OMathRun, _parse(xml))


class CT_OMathText(BaseOxmlElement):
    """`m:t` element — text content within a math run."""

    pass


class CT_OMathRunProperties(BaseOxmlElement):
    """`m:rPr` element — run formatting properties."""

    pass


# ---------------------------------------------------------------------------
# Argument containers (shared base for num, den, deg, sub, sup, e, fName)
# ---------------------------------------------------------------------------


class CT_OMathArgument(BaseOxmlElement):
    """Base class for math argument elements.

    These containers can hold any math content (runs, fractions, etc.).
    """

    r = ZeroOrMore("m:r")
    f = ZeroOrMore("m:f")
    rad = ZeroOrMore("m:rad")
    sSup = ZeroOrMore("m:sSup")
    sSub = ZeroOrMore("m:sSub")
    sSubSup = ZeroOrMore("m:sSubSup")
    nary = ZeroOrMore("m:nary")
    d = ZeroOrMore("m:d")
    acc = ZeroOrMore("m:acc")
    bar = ZeroOrMore("m:bar")
    box = ZeroOrMore("m:box")
    borderBox = ZeroOrMore("m:borderBox")
    eqArr = ZeroOrMore("m:eqArr")
    func = ZeroOrMore("m:func")
    groupChr = ZeroOrMore("m:groupChr")
    limLow = ZeroOrMore("m:limLow")
    limUpp = ZeroOrMore("m:limUpp")
    m = ZeroOrMore("m:m")

    def add_run(self, text: str = "") -> CT_OMathRun:
        r = cast(CT_OMathRun, _parse("<m:r %s><m:t/></m:r>" % nsdecls("m")))
        if text:
            r.t = text
        self.append(r)
        return r

    def add_fraction(self) -> CT_OMathFrac:
        f = cast(
            CT_OMathFrac,
            _parse("<m:f %s><m:fPr/><m:num/><m:den/></m:f>" % nsdecls("m")),
        )
        self.append(f)
        return f

    def add_radical(self) -> CT_OMathRad:
        rad = cast(
            CT_OMathRad,
            _parse(
                "<m:rad %s>"
                "<m:radPr/>"
                "<m:e/>"
                "</m:rad>" % nsdecls("m")
            ),
        )
        self.append(rad)
        return rad

    def add_superscript(self) -> CT_OMathSup:
        sSup = cast(
            CT_OMathSup,
            _parse("<m:sSup %s><m:sSupPr/><m:e/><m:sup/></m:sSup>" % nsdecls("m")),
        )
        self.append(sSup)
        return sSup

    def add_subscript(self) -> CT_OMathSub:
        sSub = cast(
            CT_OMathSub,
            _parse("<m:sSub %s><m:sSubPr/><m:e/><m:sub/></m:sSub>" % nsdecls("m")),
        )
        self.append(sSub)
        return sSub

    def add_nary(self, chr_val: str = "∫") -> CT_OMathNary:
        nary = cast(
            CT_OMathNary,
            _parse(
                "<m:nary %s><m:naryPr/><m:sub/><m:sup/><m:e/></m:nary>" % nsdecls("m")
            ),
        )
        nary_pr = nary.get_or_add_naryPr()
        nary_pr.chr = chr_val
        self.append(nary)
        return nary

    def add_delimiter(self) -> CT_OMathDelimiter:
        d = cast(
            CT_OMathDelimiter,
            _parse("<m:d %s><m:dPr/><m:e/></m:d>" % nsdecls("m")),
        )
        self.append(d)
        return d


class CT_OMathBase(CT_OMathArgument):
    """`m:e` element — base argument."""

    @classmethod
    def new(cls) -> CT_OMathBase:
        return cast(CT_OMathBase, _parse("<m:e %s/>" % nsdecls("m")))


class CT_OMathNumerator(CT_OMathArgument):
    """`m:num` element — numerator."""

    pass


class CT_OMathDenominator(CT_OMathArgument):
    """`m:den` element — denominator."""

    pass


class CT_OMathDegree(CT_OMathArgument):
    """`m:deg` element — degree (for radicals)."""

    pass


class CT_OMathSubArgument(CT_OMathArgument):
    """`m:sub` element — subscript argument."""

    pass


class CT_OMathSupArgument(CT_OMathArgument):
    """`m:sup` element — superscript argument."""

    pass


class CT_OMathFunctionName(CT_OMathArgument):
    """`m:fName` element — function name."""

    pass


class CT_OMathLimit(CT_OMathArgument):
    """`m:lim` element — limit argument."""

    pass


# ---------------------------------------------------------------------------
# Operator elements
# ---------------------------------------------------------------------------


class CT_OMathFrac(BaseOxmlElement):
    """`m:f` element — fraction."""

    _tag_seq = ("m:fPr", "m:num", "m:den")
    fPr = ZeroOrOne("m:fPr", successors=("m:num", "m:den"))
    num = ZeroOrOne("m:num", successors=("m:den",))
    den = ZeroOrOne("m:den", successors=())
    del _tag_seq

    get_or_add_fPr: Callable[[], CT_OMathFracProperties]
    get_or_add_num: Callable[[], CT_OMathNumerator]
    get_or_add_den: Callable[[], CT_OMathDenominator]


class CT_OMathRad(BaseOxmlElement):
    """`m:rad` element — radical."""

    _tag_seq = ("m:radPr", "m:deg", "m:e")
    radPr = ZeroOrOne("m:radPr", successors=("m:deg", "m:e"))
    deg = ZeroOrOne("m:deg", successors=("m:e",))
    e = ZeroOrOne("m:e", successors=())
    del _tag_seq

    get_or_add_radPr: Callable[[], CT_OMathRadProperties]
    get_or_add_deg: Callable[[], CT_OMathDegree]
    get_or_add_e: Callable[[], CT_OMathBase]


class CT_OMathSup(BaseOxmlElement):
    """`m:sSup` element — superscript."""

    _tag_seq = ("m:sSupPr", "m:e", "m:sup")
    sSupPr = ZeroOrOne("m:sSupPr", successors=("m:e", "m:sup"))
    e = ZeroOrOne("m:e", successors=("m:sup",))
    sup = ZeroOrOne("m:sup", successors=())
    del _tag_seq

    get_or_add_sSupPr: Callable[[], CT_OMathSupProperties]
    get_or_add_e: Callable[[], CT_OMathBase]
    get_or_add_sup: Callable[[], CT_OMathSupArgument]


class CT_OMathSub(BaseOxmlElement):
    """`m:sSub` element — subscript."""

    _tag_seq = ("m:sSubPr", "m:e", "m:sub")
    sSubPr = ZeroOrOne("m:sSubPr", successors=("m:e", "m:sub"))
    e = ZeroOrOne("m:e", successors=("m:sub",))
    sub = ZeroOrOne("m:sub", successors=())
    del _tag_seq

    get_or_add_sSubPr: Callable[[], CT_OMathSubProperties]
    get_or_add_e: Callable[[], CT_OMathBase]
    get_or_add_sub: Callable[[], CT_OMathSubArgument]


class CT_OMathSubSup(BaseOxmlElement):
    """`m:sSubSup` element — subscript-superscript."""

    _tag_seq = ("m:sSubSupPr", "m:e", "m:sub", "m:sup")
    sSubSupPr = ZeroOrOne("m:sSubSupPr", successors=("m:e", "m:sub", "m:sup"))
    e = ZeroOrOne("m:e", successors=("m:sub",))
    sub = ZeroOrOne("m:sub", successors=("m:sup",))
    sup = ZeroOrOne("m:sup", successors=())
    del _tag_seq

    get_or_add_sSubSupPr: Callable[[], CT_OMathSubSupProperties]
    get_or_add_e: Callable[[], CT_OMathBase]
    get_or_add_sub: Callable[[], CT_OMathSubArgument]
    get_or_add_sup: Callable[[], CT_OMathSupArgument]


class CT_OMathNary(BaseOxmlElement):
    """`m:nary` element — n-ary operator."""

    _tag_seq = ("m:naryPr", "m:sub", "m:sup", "m:e")
    naryPr = ZeroOrOne("m:naryPr", successors=("m:sub", "m:sup", "m:e"))
    sub = ZeroOrOne("m:sub", successors=("m:sup",))
    sup = ZeroOrOne("m:sup", successors=("m:e",))
    e = ZeroOrOne("m:e", successors=())
    del _tag_seq

    get_or_add_naryPr: Callable[[], CT_OMathNaryProperties]
    get_or_add_sub: Callable[[], CT_OMathSubArgument]
    get_or_add_sup: Callable[[], CT_OMathSupArgument]
    get_or_add_e: Callable[[], CT_OMathBase]


class CT_OMathDelimiter(BaseOxmlElement):
    """`m:d` element — delimiter."""

    _tag_seq = ("m:dPr", "m:e")
    dPr = ZeroOrOne("m:dPr", successors=("m:e",))
    e = ZeroOrMore("m:e")
    del _tag_seq

    get_or_add_dPr: Callable[[], CT_OMathDelimiterProperties]


class CT_OMathAccent(BaseOxmlElement):
    """`m:acc` element — accent."""

    _tag_seq = ("m:accPr", "m:e")
    accPr = ZeroOrOne("m:accPr", successors=("m:e",))
    e = ZeroOrOne("m:e", successors=())
    del _tag_seq

    get_or_add_accPr: Callable[[], CT_OMathAccentProperties]
    get_or_add_e: Callable[[], CT_OMathBase]


class CT_OMathBar(BaseOxmlElement):
    """`m:bar` element — overbar/underbar."""

    _tag_seq = ("m:barPr", "m:e")
    barPr = ZeroOrOne("m:barPr", successors=("m:e",))
    e = ZeroOrOne("m:e", successors=())
    del _tag_seq

    get_or_add_barPr: Callable[[], CT_OMathBarProperties]
    get_or_add_e: Callable[[], CT_OMathBase]


class CT_OMathBox(BaseOxmlElement):
    """`m:box` element — box."""

    _tag_seq = ("m:boxPr", "m:e")
    boxPr = ZeroOrOne("m:boxPr", successors=("m:e",))
    e = ZeroOrOne("m:e", successors=())
    del _tag_seq

    get_or_add_boxPr: Callable[[], CT_OMathBoxProperties]
    get_or_add_e: Callable[[], CT_OMathBase]


class CT_OMathBorderBox(BaseOxmlElement):
    """`m:borderBox` element — border box."""

    _tag_seq = ("m:borderBoxPr", "m:e")
    borderBoxPr = ZeroOrOne("m:borderBoxPr", successors=("m:e",))
    e = ZeroOrOne("m:e", successors=())
    del _tag_seq

    get_or_add_borderBoxPr: Callable[[], CT_OMathBorderBoxProperties]
    get_or_add_e: Callable[[], CT_OMathBase]


class CT_OMathEqArr(BaseOxmlElement):
    """`m:eqArr` element — equation array."""

    _tag_seq = ("m:eqArrPr", "m:e")
    eqArrPr = ZeroOrOne("m:eqArrPr", successors=("m:e",))
    e = ZeroOrMore("m:e")
    del _tag_seq

    get_or_add_eqArrPr: Callable[[], CT_OMathEqArrProperties]


class CT_OMathFunction(BaseOxmlElement):
    """`m:func` element — function."""

    _tag_seq = ("m:funcPr", "m:fName", "m:e")
    funcPr = ZeroOrOne("m:funcPr", successors=("m:fName", "m:e"))
    fName = ZeroOrOne("m:fName", successors=("m:e",))
    e = ZeroOrOne("m:e", successors=())
    del _tag_seq

    get_or_add_funcPr: Callable[[], CT_OMathFunctionProperties]
    get_or_add_fName: Callable[[], CT_OMathFunctionName]
    get_or_add_e: Callable[[], CT_OMathBase]


class CT_OMathGroupChar(BaseOxmlElement):
    """`m:groupChr` element — group character."""

    _tag_seq = ("m:groupChrPr", "m:e")
    groupChrPr = ZeroOrOne("m:groupChrPr", successors=("m:e",))
    e = ZeroOrOne("m:e", successors=())
    del _tag_seq

    get_or_add_groupChrPr: Callable[[], CT_OMathGroupCharProperties]
    get_or_add_e: Callable[[], CT_OMathBase]


class CT_OMathLimitLower(BaseOxmlElement):
    """`m:limLow` element — lower limit."""

    _tag_seq = ("m:limLowPr", "m:e", "m:lim")
    limLowPr = ZeroOrOne("m:limLowPr", successors=("m:e", "m:lim"))
    e = ZeroOrOne("m:e", successors=("m:lim",))
    lim = ZeroOrOne("m:lim", successors=())
    del _tag_seq

    get_or_add_limLowPr: Callable[[], CT_OMathLimitLowerProperties]
    get_or_add_e: Callable[[], CT_OMathBase]
    get_or_add_lim: Callable[[], CT_OMathLimit]


class CT_OMathLimitUpper(BaseOxmlElement):
    """`m:limUpp` element — upper limit."""

    _tag_seq = ("m:limUppPr", "m:e", "m:lim")
    limUppPr = ZeroOrOne("m:limUppPr", successors=("m:e", "m:lim"))
    e = ZeroOrOne("m:e", successors=("m:lim",))
    lim = ZeroOrOne("m:lim", successors=())
    del _tag_seq

    get_or_add_limUppPr: Callable[[], CT_OMathLimitUpperProperties]
    get_or_add_e: Callable[[], CT_OMathBase]
    get_or_add_lim: Callable[[], CT_OMathLimit]


# ---------------------------------------------------------------------------
# Matrix elements
# ---------------------------------------------------------------------------


class CT_OMathMatrix(BaseOxmlElement):
    """`m:m` element — matrix."""

    _tag_seq = ("m:mPr", "m:mr")
    mPr = ZeroOrOne("m:mPr", successors=("m:mr",))
    mr = ZeroOrMore("m:mr")
    del _tag_seq


class CT_OMathMatrixRow(BaseOxmlElement):
    """`m:mr` element — matrix row.

    Contains `<m:e>` (argument) elements for each column cell.
    """

    e = ZeroOrMore("m:e")


class CT_OMathMatrixColumn(BaseOxmlElement):
    """`m:mc` element — matrix column/cell."""

    _tag_seq = ("m:mcPr", "m:e")
    mcPr = ZeroOrOne("m:mcPr", successors=("m:e",))
    e = ZeroOrOne("m:e", successors=())
    del _tag_seq

    get_or_add_mcPr: Callable[[], CT_OMathMatrixColumnProperties]
    get_or_add_e: Callable[[], CT_OMathBase]


# ---------------------------------------------------------------------------
# Property elements
# ---------------------------------------------------------------------------


class CT_OMathFracProperties(BaseOxmlElement):
    """`m:fPr` element — fraction properties."""

    pass


class CT_OMathRadProperties(BaseOxmlElement):
    """`m:radPr` element — radical properties."""

    pass


class CT_OMathNaryProperties(BaseOxmlElement):
    """`m:naryPr` element — n-ary operator properties."""

    chr_val = OptionalAttribute("m:chr", XsdString)

    @property
    def chr(self) -> str | None:
        el = self.find(qn("m:chr"))
        if el is None:
            return None
        return el.get(qn("m:val"))

    @chr.setter
    def chr(self, value: str):
        el = self.find(qn("m:chr"))
        if el is None:
            el = etree.SubElement(self, qn("m:chr"))
        el.set(qn("m:val"), value)


class CT_OMathDelimiterProperties(BaseOxmlElement):
    """`m:dPr` element — delimiter properties."""

    pass


class CT_OMathAccentProperties(BaseOxmlElement):
    """`m:accPr` element — accent properties."""

    @property
    def chr(self) -> str | None:
        el = self.find(qn("m:chr"))
        if el is None:
            return None
        return el.get(qn("m:val"))

    @chr.setter
    def chr(self, value: str):
        el = self.find(qn("m:chr"))
        if el is None:
            el = etree.SubElement(self, qn("m:chr"))
        el.set(qn("m:val"), value)


class CT_OMathBarProperties(BaseOxmlElement):
    """`m:barPr` element — bar properties."""

    pass


class CT_OMathSupProperties(BaseOxmlElement):
    """`m:sSupPr` element — superscript properties."""

    pass


class CT_OMathSubProperties(BaseOxmlElement):
    """`m:sSubPr` element — subscript properties."""

    pass


class CT_OMathSubSupProperties(BaseOxmlElement):
    """`m:sSubSupPr` element — subscript-superscript properties."""

    pass


class CT_OMathBoxProperties(BaseOxmlElement):
    """`m:boxPr` element — box properties."""

    pass


class CT_OMathBorderBoxProperties(BaseOxmlElement):
    """`m:borderBoxPr` element — border box properties."""

    pass


class CT_OMathEqArrProperties(BaseOxmlElement):
    """`m:eqArrPr` element — equation array properties."""

    pass


class CT_OMathFunctionProperties(BaseOxmlElement):
    """`m:funcPr` element — function properties."""

    pass


class CT_OMathGroupCharProperties(BaseOxmlElement):
    """`m:groupChrPr` element — group character properties."""

    pass


class CT_OMathLimitLowerProperties(BaseOxmlElement):
    """`m:limLowPr` element — lower limit properties."""

    pass


class CT_OMathLimitUpperProperties(BaseOxmlElement):
    """`m:limUppPr` element — upper limit properties."""

    pass


class CT_OMathMatrixProperties(BaseOxmlElement):
    """`m:mPr` element — matrix properties."""

    pass


class CT_OMathMatrixColumnProperties(BaseOxmlElement):
    """`m:mcPr` element — matrix column properties."""

    pass


class CT_OMathControlProperties(BaseOxmlElement):
    """`m:ctrlPr` element — control properties."""

    pass
