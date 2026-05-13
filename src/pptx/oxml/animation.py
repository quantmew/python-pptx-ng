"""Animation XML element classes.

Corresponds to CT_TLCommonTimeNodeData, CT_TLTimeNodeParallel,
CT_TLTimeNodeSequence, CT_TLTimeNodeExclusive, CT_TLAnimateEffect,
CT_TLCommonBehavior, CT_TLTargetElement, CT_TLShapeTarget, etc.
in the Open XML SDK (schemas_openxmlformats_org_presentationml_2006_main).
"""

from __future__ import annotations

from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import XsdBoolean, XsdInt, XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_TLCommonTimeNodeData(BaseOxmlElement):
    """`p:cTn` element — common time node data.

    Contains timing and sequencing attributes shared by all time node types.

    Attributes:
        id: Unique time node identifier (optional, UInt32).
        presetID: Preset animation ID (optional, Int32).
        presetClass: Preset class: "entr", "exit", "emph", "motion", "medi" (optional, String).
        presetSubtype: Preset subtype (optional, Int32).
        dur: Duration in ms or "indefinite" (optional, String).
        repeatCount: Repeat count or "indefinite" (optional, String).
        repeatDur: Repeat duration (optional, String).
        fill: Fill mode: "freeze", "hold", "remove", "transition" (optional, String).
        nodeType: Node type: "withEffect", "afterEffect", "mainSeq", etc. (optional, String).

    Child elements (in order):
        p:stCondLst — start condition list (optional)
        p:endCondLst — end condition list (optional)
        p:endSync — end sync (optional)
        p:iter — iteration (optional)
        p:childTnLst — child time node list (optional)
        p:seq — sequence (optional, only if parent is certain elements)
    """

    id = OptionalAttribute("id", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]
    presetID = OptionalAttribute("presetID", XsdInt)  # pyright: ignore[reportAssignmentType]
    presetClass = OptionalAttribute("presetClass", XsdString)  # pyright: ignore[reportAssignmentType]
    presetSubtype = OptionalAttribute("presetSubtype", XsdInt)  # pyright: ignore[reportAssignmentType]
    dur = OptionalAttribute("dur", XsdString)  # pyright: ignore[reportAssignmentType]
    repeatCount = OptionalAttribute("repeatCount", XsdString)  # pyright: ignore[reportAssignmentType]
    repeatDur = OptionalAttribute("repeatDur", XsdString)  # pyright: ignore[reportAssignmentType]
    fill = OptionalAttribute("fill", XsdString)  # pyright: ignore[reportAssignmentType]
    nodeType = OptionalAttribute("nodeType", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_TLTimeNodeParallel(BaseOxmlElement):
    """`p:par` element — parallel time node.

    Contains a single p:cTn child which holds the timing data and child nodes.
    """

    _tag_seq = ("p:cTn",)
    cTn = OneAndOnlyOne("p:cTn")  # pyright: ignore[reportAssignmentType]
    del _tag_seq


class CT_TLTimeNodeSequence(BaseOxmlElement):
    """`p:seq` element — sequence time node.

    Attributes:
        concurrent: Whether children can play concurrently (optional, Boolean).
        nextAc: Next action condition: "seek", "none" (optional, String).

    Child elements (in order):
        p:cTn (required)
        p:prevCondLst (optional)
        p:nextCondLst (optional)
    """

    concurrent = OptionalAttribute("concurrent", XsdBoolean)  # pyright: ignore[reportAssignmentType]
    nextAc = OptionalAttribute("nextAc", XsdString)  # pyright: ignore[reportAssignmentType]

    _tag_seq = ("p:cTn",)
    cTn = OneAndOnlyOne("p:cTn")  # pyright: ignore[reportAssignmentType]
    del _tag_seq


class CT_TLTimeNodeExclusive(BaseOxmlElement):
    """`p:excl` element — exclusive time node.

    Child elements (in order):
        p:cTn (required)
    """

    _tag_seq = ("p:cTn",)
    cTn = OneAndOnlyOne("p:cTn")  # pyright: ignore[reportAssignmentType]
    del _tag_seq


class CT_TLAnimateEffect(BaseOxmlElement):
    """`p:animEffect` element — animation effect.

    Attributes:
        transition: Transition type: "in", "out" (optional, String).
        filter: Filter string (optional, String).

    Child elements (in order):
        p:cBhvr (required)
        p:extLst (optional)
    """

    transition = OptionalAttribute("transition", XsdString)  # pyright: ignore[reportAssignmentType]
    filter = OptionalAttribute("filter", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_TLCommonBehavior(BaseOxmlElement):
    """`p:cBhvr` element — common behavior.

    Child elements (in order):
        p:cTn (required)
        p:tgtEl (required)
        p:attrNameLst (optional)
    """

    _tag_seq = ("p:cTn", "p:tgtEl")
    cTn = OneAndOnlyOne("p:cTn")  # pyright: ignore[reportAssignmentType]
    tgtEl = OneAndOnlyOne("p:tgtEl")  # pyright: ignore[reportAssignmentType]
    del _tag_seq


class CT_TLTargetElement(BaseOxmlElement):
    """`p:tgtEl` element — target element.

    Child elements (choice of one):
        p:spTgt — shape target
        p:inkTgt — ink target
        p:sndTgt — sound target
        p:mediaTgt — media target
    """

    spTgt = ZeroOrOne("p:spTgt", successors=())  # pyright: ignore[reportAssignmentType]


class CT_TLShapeTarget(BaseOxmlElement):
    """`p:spTgt` element — shape target.

    Attributes:
        spid: Shape ID (optional, String).
    """

    spid = OptionalAttribute("spid", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_TLSetEffect(BaseOxmlElement):
    """`p:set` element — set effect.

    Child elements (in order):
        p:cBhvr (required)
        p:to (optional)
    """

    _tag_seq = ("p:cBhvr",)
    cBhvr = OneAndOnlyOne("p:cBhvr")  # pyright: ignore[reportAssignmentType]
    del _tag_seq


class CT_TLBuildList(BaseOxmlElement):
    """`p:bldLst` element — build list for slide animations.

    Contains zero or more p:bldP (build paragraph) children.
    """

    bldP = ZeroOrMore("p:bldP", successors=())  # pyright: ignore[reportAssignmentType]


class CT_TLBuildParagraph(BaseOxmlElement):
    """`p:bldP` element — build paragraph.

    Attributes:
        spid: Shape ID (optional, String).
        grpId: Group ID (optional, UInt32).
        uiExpand: Whether to expand UI (optional, Boolean).
        build: Build type: "allAtOnce", "p", "l" (optional, String).

    Child elements:
        p:tmTx (optional)
    """

    spid = OptionalAttribute("spid", XsdString)  # pyright: ignore[reportAssignmentType]
    grpId = OptionalAttribute("grpId", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]
    uiExpand = OptionalAttribute("uiExpand", XsdBoolean)  # pyright: ignore[reportAssignmentType]
    build = OptionalAttribute("build", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_TLAnimate(BaseOxmlElement):
    """`p:anim` element — animate.

    Attributes:
        by: By value (optional, String).
        from: From value (optional, String).
        to: To value (optional, String).
        calcmode: Calculation mode (optional, String).
        valueType: Value type (optional, String).

    Child elements (in order):
        p:cBhvr (required)
        p:tavLst (optional)
    """

    by = OptionalAttribute("by", XsdString)  # pyright: ignore[reportAssignmentType]
    from_ = OptionalAttribute("from", XsdString)  # pyright: ignore[reportAssignmentType]
    to = OptionalAttribute("to", XsdString)  # pyright: ignore[reportAssignmentType]
    calcmode = OptionalAttribute("calcmode", XsdString)  # pyright: ignore[reportAssignmentType]
    valueType = OptionalAttribute("valueType", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_TLTimeAnimateValueList(BaseOxmlElement):
    """`p:tavLst` element — time animate value list."""

    tav = ZeroOrMore("p:tav", successors=())  # pyright: ignore[reportAssignmentType]


class CT_TLTimeAnimateValue(BaseOxmlElement):
    """`p:tav` element — time animate value.

    Attributes:
        tm: Time value (optional, String).
        fmla: Formula (optional, String).
    """

    tm = OptionalAttribute("tm", XsdString)  # pyright: ignore[reportAssignmentType]
    fmla = OptionalAttribute("fmla", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_TLConditionList(BaseOxmlElement):
    """`p:stCondLst` / `p:cond` list — condition list."""

    cond = ZeroOrMore("p:cond", successors=())  # pyright: ignore[reportAssignmentType]


class CT_TLCondition(BaseOxmlElement):
    """`p:cond` element — condition.

    Attributes:
        delay: Delay value (optional, String).
        evt: Event type (optional, String).
        op: Operator (optional, String).
    """

    delay = OptionalAttribute("delay", XsdString)  # pyright: ignore[reportAssignmentType]
    evt = OptionalAttribute("evt", XsdString)  # pyright: ignore[reportAssignmentType]
    op = OptionalAttribute("op", XsdString)  # pyright: ignore[reportAssignmentType]
