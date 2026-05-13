"""Initializes lxml parser, particularly the custom element classes.

Also makes available a handful of functions that wrap its typical uses.
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, Type

from lxml import etree

from pptx.oxml.ns import NamespacePrefixedTag

if TYPE_CHECKING:
    from pptx.oxml.xmlchemy import BaseOxmlElement


# -- configure etree XML parser ----------------------------
element_class_lookup = etree.ElementNamespaceClassLookup()
oxml_parser = etree.XMLParser(remove_blank_text=True, resolve_entities=False)
oxml_parser.set_element_class_lookup(element_class_lookup)


def parse_from_template(template_file_name: str):
    """Return an element loaded from the XML in the template file identified by `template_name`."""
    thisdir = os.path.split(__file__)[0]
    filename = os.path.join(thisdir, "..", "templates", "%s.xml" % template_file_name)
    with open(filename, "rb") as f:
        xml = f.read()
    return parse_xml(xml)


def parse_xml(xml: str | bytes):
    """Return root lxml element obtained by parsing XML character string in `xml`."""
    return etree.fromstring(xml, oxml_parser)


def register_element_cls(nsptagname: str, cls: Type[BaseOxmlElement]):
    """Register `cls` to be constructed when oxml parser encounters element having `nsptag_name`.

    `nsptag_name` is a string of the form `nspfx:tagroot`, e.g. `"w:document"`.
    """
    nsptag = NamespacePrefixedTag(nsptagname)
    namespace = element_class_lookup.get_namespace(nsptag.nsuri)
    namespace[nsptag.local_part] = cls


from pptx.oxml.action import CT_Hyperlink  # noqa: E402

register_element_cls("a:hlinkClick", CT_Hyperlink)
register_element_cls("a:hlinkHover", CT_Hyperlink)


from pptx.oxml.animation import (  # noqa: E402
    CT_TLAnimate,
    CT_TLAnimateEffect,
    CT_TLBuildList,
    CT_TLBuildParagraph,
    CT_TLCommonBehavior,
    CT_TLCommonTimeNodeData,
    CT_TLCondition,
    CT_TLConditionList,
    CT_TLSetEffect,
    CT_TLShapeTarget,
    CT_TLTargetElement,
    CT_TLTimeAnimateValue,
    CT_TLTimeAnimateValueList,
    CT_TLTimeNodeExclusive,
    CT_TLTimeNodeParallel,
    CT_TLTimeNodeSequence,
)

register_element_cls("p:anim", CT_TLAnimate)
register_element_cls("p:animEffect", CT_TLAnimateEffect)
register_element_cls("p:bldLst", CT_TLBuildList)
register_element_cls("p:bldP", CT_TLBuildParagraph)
register_element_cls("p:cBhvr", CT_TLCommonBehavior)
register_element_cls("p:cTn", CT_TLCommonTimeNodeData)
register_element_cls("p:cond", CT_TLCondition)
register_element_cls("p:excl", CT_TLTimeNodeExclusive)
register_element_cls("p:par", CT_TLTimeNodeParallel)
register_element_cls("p:seq", CT_TLTimeNodeSequence)
register_element_cls("p:set", CT_TLSetEffect)
register_element_cls("p:spTgt", CT_TLShapeTarget)
register_element_cls("p:stCondLst", CT_TLConditionList)
register_element_cls("p:tav", CT_TLTimeAnimateValue)
register_element_cls("p:tavLst", CT_TLTimeAnimateValueList)
register_element_cls("p:tgtEl", CT_TLTargetElement)

from pptx.oxml.comment import (  # noqa: E402
    CT_Comment,
    CT_CommentAuthor,
    CT_CommentAuthorList,
    CT_CommentList,
    CT_CommentPosition,
    CT_CommentText,
)

register_element_cls("p:cm", CT_Comment)
register_element_cls("p:cmAuthor", CT_CommentAuthor)
register_element_cls("p:cmAuthorLst", CT_CommentAuthorList)
register_element_cls("p:cmLst", CT_CommentList)
register_element_cls("p:pos", CT_CommentPosition)
register_element_cls("p:text", CT_CommentText)


from pptx.oxml.customshow import (  # noqa: E402
    CT_CustomShow,
    CT_CustomShowList,
    CT_SlideRelationshipList,
    CT_SlideRelationshipListEntry,
)

register_element_cls("p:custShow", CT_CustomShow)
register_element_cls("p:custShowLst", CT_CustomShowList)
register_element_cls("p:sldLst", CT_SlideRelationshipList)


from pptx.oxml.handout import (  # noqa: E402
    CT_HandoutMaster,
    CT_HandoutMasterIdList,
    CT_HandoutMasterIdListEntry,
)

register_element_cls("p:handoutMaster", CT_HandoutMaster)
register_element_cls("p:handoutMasterId", CT_HandoutMasterIdListEntry)
register_element_cls("p:handoutMasterIdLst", CT_HandoutMasterIdList)


from pptx.oxml.tags import CT_Tag, CT_TagList  # noqa: E402

register_element_cls("p:tag", CT_Tag)
register_element_cls("p:tagLst", CT_TagList)


from pptx.oxml.chartex.chartex import (  # noqa: E402
    CT_ChartEx,
    CT_ChartExData,
    CT_ChartExSpace,
    CT_PlotArea,
    CT_SeriesLayoutPr,
)

register_element_cls("cx:chart", CT_ChartEx)
register_element_cls("cx:chartData", CT_ChartExData)
register_element_cls("cx:chartSpace", CT_ChartExSpace)
register_element_cls("cx:plotArea", CT_PlotArea)
register_element_cls("cx:layoutPr", CT_SeriesLayoutPr)


from pptx.oxml.model3d import (  # noqa: E402
    CT_AmbientLight,
    CT_DirectionalLight,
    CT_Model3D,
    CT_Model3DCamera,
    CT_Model3DTransform,
    CT_ObjectViewport,
    CT_OrthographicProjection,
    CT_PerspectiveProjection,
    CT_PointLight,
    CT_SpotLight,
    CT_UnknownLight,
    CT_WindowViewport,
)

register_element_cls("m3d:model3D", CT_Model3D)
register_element_cls("m3d:camera", CT_Model3DCamera)
register_element_cls("m3d:trans", CT_Model3DTransform)
register_element_cls("m3d:objViewport", CT_ObjectViewport)
register_element_cls("m3d:winViewport", CT_WindowViewport)
register_element_cls("m3d:ambientLight", CT_AmbientLight)
register_element_cls("m3d:ptLight", CT_PointLight)
register_element_cls("m3d:spotLight", CT_SpotLight)
register_element_cls("m3d:dirLight", CT_DirectionalLight)
register_element_cls("m3d:unkLight", CT_UnknownLight)
register_element_cls("m3d:orthographic", CT_OrthographicProjection)
register_element_cls("m3d:perspective", CT_PerspectiveProjection)


from pptx.oxml.chart.axis import (  # noqa: E402
    CT_AxisUnit,
    CT_CatAx,
    CT_ChartLines,
    CT_Crosses,
    CT_DateAx,
    CT_LblOffset,
    CT_Orientation,
    CT_Scaling,
    CT_TickLblPos,
    CT_TickMark,
    CT_ValAx,
)

register_element_cls("c:catAx", CT_CatAx)
register_element_cls("c:crosses", CT_Crosses)
register_element_cls("c:dateAx", CT_DateAx)
register_element_cls("c:lblOffset", CT_LblOffset)
register_element_cls("c:majorGridlines", CT_ChartLines)
register_element_cls("c:majorTickMark", CT_TickMark)
register_element_cls("c:majorUnit", CT_AxisUnit)
register_element_cls("c:minorTickMark", CT_TickMark)
register_element_cls("c:minorUnit", CT_AxisUnit)
register_element_cls("c:orientation", CT_Orientation)
register_element_cls("c:scaling", CT_Scaling)
register_element_cls("c:tickLblPos", CT_TickLblPos)
register_element_cls("c:valAx", CT_ValAx)


from pptx.oxml.chart.chart import (  # noqa: E402
    CT_Chart,
    CT_ChartSpace,
    CT_ExternalData,
    CT_PlotArea,
    CT_Style,
)

register_element_cls("c:chart", CT_Chart)
register_element_cls("c:chartSpace", CT_ChartSpace)
register_element_cls("c:externalData", CT_ExternalData)
register_element_cls("c:plotArea", CT_PlotArea)
register_element_cls("c:style", CT_Style)


from pptx.oxml.chart.datalabel import CT_DLbl, CT_DLblPos, CT_DLbls  # noqa: E402

register_element_cls("c:dLbl", CT_DLbl)
register_element_cls("c:dLblPos", CT_DLblPos)
register_element_cls("c:dLbls", CT_DLbls)


from pptx.oxml.chart.legend import CT_Legend, CT_LegendPos  # noqa: E402

register_element_cls("c:legend", CT_Legend)
register_element_cls("c:legendPos", CT_LegendPos)


from pptx.oxml.chart.marker import CT_Marker, CT_MarkerSize, CT_MarkerStyle  # noqa: E402

register_element_cls("c:marker", CT_Marker)
register_element_cls("c:size", CT_MarkerSize)
register_element_cls("c:symbol", CT_MarkerStyle)


from pptx.oxml.chart.plot import (  # noqa: E402
    CT_Area3DChart,
    CT_AreaChart,
    CT_BarChart,
    CT_BarDir,
    CT_BubbleChart,
    CT_BubbleScale,
    CT_DoughnutChart,
    CT_GapAmount,
    CT_Grouping,
    CT_LineChart,
    CT_Overlap,
    CT_PieChart,
    CT_RadarChart,
    CT_ScatterChart,
)

register_element_cls("c:area3DChart", CT_Area3DChart)
register_element_cls("c:areaChart", CT_AreaChart)
register_element_cls("c:barChart", CT_BarChart)
register_element_cls("c:barDir", CT_BarDir)
register_element_cls("c:bubbleChart", CT_BubbleChart)
register_element_cls("c:bubbleScale", CT_BubbleScale)
register_element_cls("c:doughnutChart", CT_DoughnutChart)
register_element_cls("c:gapWidth", CT_GapAmount)
register_element_cls("c:grouping", CT_Grouping)
register_element_cls("c:lineChart", CT_LineChart)
register_element_cls("c:overlap", CT_Overlap)
register_element_cls("c:pieChart", CT_PieChart)
register_element_cls("c:radarChart", CT_RadarChart)
register_element_cls("c:scatterChart", CT_ScatterChart)


from pptx.oxml.chart.series import (  # noqa: E402
    CT_AxDataSource,
    CT_DPt,
    CT_Lvl,
    CT_NumDataSource,
    CT_SeriesComposite,
    CT_StrVal_NumVal_Composite,
)

register_element_cls("c:bubbleSize", CT_NumDataSource)
register_element_cls("c:cat", CT_AxDataSource)
register_element_cls("c:dPt", CT_DPt)
register_element_cls("c:lvl", CT_Lvl)
register_element_cls("c:pt", CT_StrVal_NumVal_Composite)
register_element_cls("c:ser", CT_SeriesComposite)
register_element_cls("c:val", CT_NumDataSource)
register_element_cls("c:xVal", CT_NumDataSource)
register_element_cls("c:yVal", CT_NumDataSource)


from pptx.oxml.chart.shared import (  # noqa: E402
    CT_Boolean,
    CT_Boolean_Explicit,
    CT_Double,
    CT_Layout,
    CT_LayoutMode,
    CT_ManualLayout,
    CT_NumFmt,
    CT_Title,
    CT_Tx,
    CT_UnsignedInt,
)

register_element_cls("c:autoTitleDeleted", CT_Boolean_Explicit)
register_element_cls("c:autoUpdate", CT_Boolean)
register_element_cls("c:bubble3D", CT_Boolean)
register_element_cls("c:crossAx", CT_UnsignedInt)
register_element_cls("c:crossesAt", CT_Double)
register_element_cls("c:date1904", CT_Boolean)
register_element_cls("c:delete", CT_Boolean)
register_element_cls("c:idx", CT_UnsignedInt)
register_element_cls("c:invertIfNegative", CT_Boolean_Explicit)
register_element_cls("c:layout", CT_Layout)
register_element_cls("c:manualLayout", CT_ManualLayout)
register_element_cls("c:max", CT_Double)
register_element_cls("c:min", CT_Double)
register_element_cls("c:numFmt", CT_NumFmt)
register_element_cls("c:order", CT_UnsignedInt)
register_element_cls("c:overlay", CT_Boolean_Explicit)
register_element_cls("c:ptCount", CT_UnsignedInt)
register_element_cls("c:showCatName", CT_Boolean_Explicit)
register_element_cls("c:showLegendKey", CT_Boolean_Explicit)
register_element_cls("c:showPercent", CT_Boolean_Explicit)
register_element_cls("c:showSerName", CT_Boolean_Explicit)
register_element_cls("c:showVal", CT_Boolean_Explicit)
register_element_cls("c:smooth", CT_Boolean)
register_element_cls("c:title", CT_Title)
register_element_cls("c:tx", CT_Tx)
register_element_cls("c:varyColors", CT_Boolean)
register_element_cls("c:x", CT_Double)
register_element_cls("c:xMode", CT_LayoutMode)


from pptx.oxml.coreprops import CT_CoreProperties  # noqa: E402

register_element_cls("cp:coreProperties", CT_CoreProperties)


from pptx.oxml.dml.color import (  # noqa: E402
    CT_Color,
    CT_HslColor,
    CT_Percentage,
    CT_PresetColor,
    CT_SchemeColor,
    CT_ScRgbColor,
    CT_SRgbColor,
    CT_SystemColor,
)

register_element_cls("a:bgClr", CT_Color)
register_element_cls("a:fgClr", CT_Color)
register_element_cls("a:hslClr", CT_HslColor)
register_element_cls("a:lumMod", CT_Percentage)
register_element_cls("a:lumOff", CT_Percentage)
register_element_cls("a:prstClr", CT_PresetColor)
register_element_cls("a:schemeClr", CT_SchemeColor)
register_element_cls("a:scrgbClr", CT_ScRgbColor)
register_element_cls("a:srgbClr", CT_SRgbColor)
register_element_cls("a:sysClr", CT_SystemColor)


from pptx.oxml.dml.fill import (  # noqa: E402
    CT_Blip,
    CT_BlipFillProperties,
    CT_GradientFillProperties,
    CT_GradientStop,
    CT_GradientStopList,
    CT_GroupFillProperties,
    CT_LinearShadeProperties,
    CT_NoFillProperties,
    CT_PatternFillProperties,
    CT_RelativeRect,
    CT_SolidColorFillProperties,
)

register_element_cls("a:blip", CT_Blip)
register_element_cls("a:blipFill", CT_BlipFillProperties)
register_element_cls("a:gradFill", CT_GradientFillProperties)
register_element_cls("a:grpFill", CT_GroupFillProperties)
register_element_cls("a:gs", CT_GradientStop)
register_element_cls("a:gsLst", CT_GradientStopList)
register_element_cls("a:lin", CT_LinearShadeProperties)
register_element_cls("a:noFill", CT_NoFillProperties)
register_element_cls("a:pattFill", CT_PatternFillProperties)
register_element_cls("a:solidFill", CT_SolidColorFillProperties)
register_element_cls("a:srcRect", CT_RelativeRect)


from pptx.oxml.dml.line import CT_PresetLineDashProperties  # noqa: E402

register_element_cls("a:prstDash", CT_PresetLineDashProperties)


from pptx.oxml.presentation import (  # noqa: E402
    CT_Presentation,
    CT_SlideId,
    CT_SlideIdList,
    CT_SlideMasterIdList,
    CT_SlideMasterIdListEntry,
    CT_SlideSize,
)

register_element_cls("p:presentation", CT_Presentation)
register_element_cls("p:sldId", CT_SlideId)
register_element_cls("p:sldIdLst", CT_SlideIdList)
register_element_cls("p:sldMasterId", CT_SlideMasterIdListEntry)
register_element_cls("p:sldMasterIdLst", CT_SlideMasterIdList)
register_element_cls("p:sldSz", CT_SlideSize)


from pptx.oxml.presprops import (  # noqa: E402
    CT_PresentationProperties,
    CT_PrintProperties,
    CT_ShowProperties,
)

register_element_cls("p:presentationPr", CT_PresentationProperties)
register_element_cls("p:prnPr", CT_PrintProperties)
register_element_cls("p:showPr", CT_ShowProperties)


from pptx.oxml.viewprops import CT_GridSpacing, CT_ViewProperties  # noqa: E402

register_element_cls("p:gridSpacing", CT_GridSpacing)
register_element_cls("p:viewPr", CT_ViewProperties)


from pptx.oxml.shapes.autoshape import (  # noqa: E402
    CT_AdjPoint2D,
    CT_CustomGeometry2D,
    CT_GeomGuide,
    CT_GeomGuideList,
    CT_NonVisualDrawingShapeProps,
    CT_Path2D,
    CT_Path2DClose,
    CT_Path2DLineTo,
    CT_Path2DList,
    CT_Path2DMoveTo,
    CT_PresetGeometry2D,
    CT_Shape,
    CT_ShapeNonVisual,
)

register_element_cls("a:avLst", CT_GeomGuideList)
register_element_cls("a:custGeom", CT_CustomGeometry2D)
register_element_cls("a:gd", CT_GeomGuide)
register_element_cls("a:close", CT_Path2DClose)
register_element_cls("a:lnTo", CT_Path2DLineTo)
register_element_cls("a:moveTo", CT_Path2DMoveTo)
register_element_cls("a:path", CT_Path2D)
register_element_cls("a:pathLst", CT_Path2DList)
register_element_cls("a:prstGeom", CT_PresetGeometry2D)
register_element_cls("a:pt", CT_AdjPoint2D)
register_element_cls("p:cNvSpPr", CT_NonVisualDrawingShapeProps)
register_element_cls("p:nvSpPr", CT_ShapeNonVisual)
register_element_cls("p:sp", CT_Shape)


from pptx.oxml.shapes.connector import (  # noqa: E402
    CT_Connection,
    CT_Connector,
    CT_ConnectorNonVisual,
    CT_NonVisualConnectorProperties,
)

register_element_cls("a:endCxn", CT_Connection)
register_element_cls("a:stCxn", CT_Connection)
register_element_cls("p:cNvCxnSpPr", CT_NonVisualConnectorProperties)
register_element_cls("p:cxnSp", CT_Connector)
register_element_cls("p:nvCxnSpPr", CT_ConnectorNonVisual)


from pptx.oxml.shapes.graphfrm import (  # noqa: E402
    CT_GraphicalObject,
    CT_GraphicalObjectData,
    CT_GraphicalObjectFrame,
    CT_GraphicalObjectFrameNonVisual,
    CT_OleObject,
)

register_element_cls("a:graphic", CT_GraphicalObject)
register_element_cls("a:graphicData", CT_GraphicalObjectData)
register_element_cls("p:graphicFrame", CT_GraphicalObjectFrame)
register_element_cls("p:nvGraphicFramePr", CT_GraphicalObjectFrameNonVisual)
register_element_cls("p:oleObj", CT_OleObject)


from pptx.oxml.shapes.groupshape import (  # noqa: E402
    CT_GroupShape,
    CT_GroupShapeNonVisual,
    CT_GroupShapeProperties,
)

register_element_cls("p:grpSp", CT_GroupShape)
register_element_cls("p:grpSpPr", CT_GroupShapeProperties)
register_element_cls("p:nvGrpSpPr", CT_GroupShapeNonVisual)
register_element_cls("p:spTree", CT_GroupShape)


from pptx.oxml.shapes.picture import CT_Picture, CT_PictureNonVisual  # noqa: E402

register_element_cls("p:blipFill", CT_BlipFillProperties)
register_element_cls("p:nvPicPr", CT_PictureNonVisual)
register_element_cls("p:pic", CT_Picture)


from pptx.oxml.shapes.shared import (  # noqa: E402
    CT_ApplicationNonVisualDrawingProps,
    CT_LineProperties,
    CT_NonVisualDrawingProps,
    CT_Placeholder,
    CT_Point2D,
    CT_PositiveSize2D,
    CT_ShapeProperties,
    CT_Transform2D,
)

register_element_cls("a:chExt", CT_PositiveSize2D)
register_element_cls("a:chOff", CT_Point2D)
register_element_cls("a:ext", CT_PositiveSize2D)
register_element_cls("a:ln", CT_LineProperties)
register_element_cls("a:off", CT_Point2D)
register_element_cls("a:xfrm", CT_Transform2D)
register_element_cls("c:spPr", CT_ShapeProperties)
register_element_cls("p:cNvPr", CT_NonVisualDrawingProps)
register_element_cls("p:nvPr", CT_ApplicationNonVisualDrawingProps)
register_element_cls("p:ph", CT_Placeholder)
register_element_cls("p:spPr", CT_ShapeProperties)
register_element_cls("p:xfrm", CT_Transform2D)


from pptx.oxml.transition import CT_Transition  # noqa: E402

register_element_cls("p:transition", CT_Transition)


from pptx.oxml.slide import (  # noqa: E402
    CT_Background,
    CT_BackgroundProperties,
    CT_CommonSlideData,
    CT_NotesMaster,
    CT_NotesSlide,
    CT_Slide,
    CT_SlideLayout,
    CT_SlideLayoutIdList,
    CT_SlideLayoutIdListEntry,
    CT_SlideMaster,
    CT_SlideTiming,
    CT_TimeNodeList,
    CT_TLMediaNodeAudio,
    CT_TLMediaNodeVideo,
)

register_element_cls("p:audio", CT_TLMediaNodeAudio)
register_element_cls("p:bg", CT_Background)
register_element_cls("p:bgPr", CT_BackgroundProperties)
register_element_cls("p:childTnLst", CT_TimeNodeList)
register_element_cls("p:cSld", CT_CommonSlideData)
register_element_cls("p:notes", CT_NotesSlide)
register_element_cls("p:notesMaster", CT_NotesMaster)
register_element_cls("p:sld", CT_Slide)
register_element_cls("p:sldLayout", CT_SlideLayout)
register_element_cls("p:sldLayoutId", CT_SlideLayoutIdListEntry)
register_element_cls("p:sldLayoutIdLst", CT_SlideLayoutIdList)
register_element_cls("p:sldMaster", CT_SlideMaster)
register_element_cls("p:timing", CT_SlideTiming)
register_element_cls("p:video", CT_TLMediaNodeVideo)


from pptx.oxml.table import (  # noqa: E402
    CT_Table,
    CT_TableCell,
    CT_TableCellProperties,
    CT_TableCol,
    CT_TableGrid,
    CT_TableProperties,
    CT_TableRow,
)

register_element_cls("a:gridCol", CT_TableCol)
register_element_cls("a:tbl", CT_Table)
register_element_cls("a:tblGrid", CT_TableGrid)
register_element_cls("a:tblPr", CT_TableProperties)
register_element_cls("a:tc", CT_TableCell)
register_element_cls("a:tcPr", CT_TableCellProperties)
register_element_cls("a:tr", CT_TableRow)


from pptx.oxml.tablestyle import (  # noqa: E402
    CT_TableStyleEntry,
    CT_TableStyleList,
)

register_element_cls("a:tblStyle", CT_TableStyleEntry)
register_element_cls("a:tblStyleLst", CT_TableStyleList)


from pptx.oxml.text import (  # noqa: E402
    CT_RegularTextRun,
    CT_TextBody,
    CT_TextBodyProperties,
    CT_TextCharacterProperties,
    CT_TextField,
    CT_TextFont,
    CT_TextLineBreak,
    CT_TextNormalAutofit,
    CT_TextParagraph,
    CT_TextParagraphProperties,
    CT_TextSpacing,
    CT_TextSpacingPercent,
    CT_TextSpacingPoint,
)

register_element_cls("a:bodyPr", CT_TextBodyProperties)
register_element_cls("a:br", CT_TextLineBreak)
register_element_cls("a:defRPr", CT_TextCharacterProperties)
register_element_cls("a:endParaRPr", CT_TextCharacterProperties)
register_element_cls("a:fld", CT_TextField)
register_element_cls("a:latin", CT_TextFont)
register_element_cls("a:lnSpc", CT_TextSpacing)
register_element_cls("a:normAutofit", CT_TextNormalAutofit)
register_element_cls("a:r", CT_RegularTextRun)
register_element_cls("a:p", CT_TextParagraph)
register_element_cls("a:pPr", CT_TextParagraphProperties)
register_element_cls("c:rich", CT_TextBody)
register_element_cls("a:rPr", CT_TextCharacterProperties)
register_element_cls("a:spcAft", CT_TextSpacing)
register_element_cls("a:spcBef", CT_TextSpacing)
register_element_cls("a:spcPct", CT_TextSpacingPercent)
register_element_cls("a:spcPts", CT_TextSpacingPoint)
register_element_cls("a:txBody", CT_TextBody)
register_element_cls("c:txPr", CT_TextBody)
register_element_cls("p:txBody", CT_TextBody)


from pptx.oxml.diagram.core import (  # noqa: E402
    CT_ColorsDefinition,
    CT_Connection,
    CT_ConnectionList,
    CT_DataModel,
    CT_LayoutDefinition,
    CT_Point,
    CT_PointList,
    CT_StyleDefinition,
)

register_element_cls("dgm:colorsDef", CT_ColorsDefinition)
register_element_cls("dgm:cxn", CT_Connection)
register_element_cls("dgm:cxnLst", CT_ConnectionList)
register_element_cls("dgm:dataModel", CT_DataModel)
register_element_cls("dgm:layoutDef", CT_LayoutDefinition)
register_element_cls("dgm:pt", CT_Point)
register_element_cls("dgm:ptLst", CT_PointList)
register_element_cls("dgm:styleDef", CT_StyleDefinition)


from pptx.oxml.math.core import (  # noqa: E402
    CT_OMath,
    CT_OMathAccent,
    CT_OMathAccentProperties,
    CT_OMathArgument,
    CT_OMathBar,
    CT_OMathBarProperties,
    CT_OMathBase,
    CT_OMathBorderBox,
    CT_OMathBorderBoxProperties,
    CT_OMathBox,
    CT_OMathBoxProperties,
    CT_OMathControlProperties,
    CT_OMathDelimiter,
    CT_OMathDelimiterProperties,
    CT_OMathDenominator,
    CT_OMathDegree,
    CT_OMathEqArr,
    CT_OMathEqArrProperties,
    CT_OMathFrac,
    CT_OMathFracProperties,
    CT_OMathFunction,
    CT_OMathFunctionName,
    CT_OMathFunctionProperties,
    CT_OMathGroupChar,
    CT_OMathGroupCharProperties,
    CT_OMathLimit,
    CT_OMathLimitLower,
    CT_OMathLimitLowerProperties,
    CT_OMathLimitUpper,
    CT_OMathLimitUpperProperties,
    CT_OMathMatrix,
    CT_OMathMatrixColumn,
    CT_OMathMatrixColumnProperties,
    CT_OMathMatrixProperties,
    CT_OMathMatrixRow,
    CT_OMathNary,
    CT_OMathNaryProperties,
    CT_OMathNumerator,
    CT_OMathPara,
    CT_OMathRad,
    CT_OMathRadProperties,
    CT_OMathRun,
    CT_OMathRunProperties,
    CT_OMathSub,
    CT_OMathSubArgument,
    CT_OMathSubProperties,
    CT_OMathSubSup,
    CT_OMathSubSupProperties,
    CT_OMathSup,
    CT_OMathSupArgument,
    CT_OMathSupProperties,
    CT_OMathText,
)

register_element_cls("m:acc", CT_OMathAccent)
register_element_cls("m:accPr", CT_OMathAccentProperties)
register_element_cls("m:bar", CT_OMathBar)
register_element_cls("m:barPr", CT_OMathBarProperties)
register_element_cls("m:borderBox", CT_OMathBorderBox)
register_element_cls("m:borderBoxPr", CT_OMathBorderBoxProperties)
register_element_cls("m:box", CT_OMathBox)
register_element_cls("m:boxPr", CT_OMathBoxProperties)
register_element_cls("m:d", CT_OMathDelimiter)
register_element_cls("m:deg", CT_OMathDegree)
register_element_cls("m:den", CT_OMathDenominator)
register_element_cls("m:dPr", CT_OMathDelimiterProperties)
register_element_cls("m:e", CT_OMathBase)
register_element_cls("m:eqArr", CT_OMathEqArr)
register_element_cls("m:eqArrPr", CT_OMathEqArrProperties)
register_element_cls("m:f", CT_OMathFrac)
register_element_cls("m:fName", CT_OMathFunctionName)
register_element_cls("m:fPr", CT_OMathFracProperties)
register_element_cls("m:func", CT_OMathFunction)
register_element_cls("m:funcPr", CT_OMathFunctionProperties)
register_element_cls("m:groupChr", CT_OMathGroupChar)
register_element_cls("m:groupChrPr", CT_OMathGroupCharProperties)
register_element_cls("m:lim", CT_OMathLimit)
register_element_cls("m:limLow", CT_OMathLimitLower)
register_element_cls("m:limLowPr", CT_OMathLimitLowerProperties)
register_element_cls("m:limUpp", CT_OMathLimitUpper)
register_element_cls("m:limUppPr", CT_OMathLimitUpperProperties)
register_element_cls("m:m", CT_OMathMatrix)
register_element_cls("m:mc", CT_OMathMatrixColumn)
register_element_cls("m:mcPr", CT_OMathMatrixColumnProperties)
register_element_cls("m:mPr", CT_OMathMatrixProperties)
register_element_cls("m:mr", CT_OMathMatrixRow)
register_element_cls("m:nary", CT_OMathNary)
register_element_cls("m:naryPr", CT_OMathNaryProperties)
register_element_cls("m:num", CT_OMathNumerator)
register_element_cls("m:oMath", CT_OMath)
register_element_cls("m:oMathPara", CT_OMathPara)
register_element_cls("m:rad", CT_OMathRad)
register_element_cls("m:radPr", CT_OMathRadProperties)
register_element_cls("m:r", CT_OMathRun)
register_element_cls("m:rPr", CT_OMathRunProperties)
register_element_cls("m:sSub", CT_OMathSub)
register_element_cls("m:sSubPr", CT_OMathSubProperties)
register_element_cls("m:sSubSup", CT_OMathSubSup)
register_element_cls("m:sSubSupPr", CT_OMathSubSupProperties)
register_element_cls("m:sSup", CT_OMathSup)
register_element_cls("m:sSupPr", CT_OMathSupProperties)
register_element_cls("m:sub", CT_OMathSubArgument)
register_element_cls("m:sup", CT_OMathSupArgument)
register_element_cls("m:t", CT_OMathText)
register_element_cls("m:ctrlPr", CT_OMathControlProperties)

register_element_cls("m:acc", CT_OMathAccent)
register_element_cls("m:accPr", CT_OMathAccentProperties)
register_element_cls("m:bar", CT_OMathBar)
register_element_cls("m:barPr", CT_OMathBarProperties)
register_element_cls("m:borderBox", CT_OMathBorderBox)
register_element_cls("m:borderBoxPr", CT_OMathBorderBoxProperties)
register_element_cls("m:box", CT_OMathBox)
register_element_cls("m:boxPr", CT_OMathBoxProperties)
register_element_cls("m:d", CT_OMathDelimiter)
register_element_cls("m:deg", CT_OMathDegree)
register_element_cls("m:den", CT_OMathDenominator)
register_element_cls("m:dPr", CT_OMathDelimiterProperties)
register_element_cls("m:e", CT_OMathBase)
register_element_cls("m:eqArr", CT_OMathEqArr)
register_element_cls("m:eqArrPr", CT_OMathEqArrProperties)
register_element_cls("m:f", CT_OMathFrac)
register_element_cls("m:fName", CT_OMathFunctionName)
register_element_cls("m:fPr", CT_OMathFracProperties)
register_element_cls("m:func", CT_OMathFunction)
register_element_cls("m:funcPr", CT_OMathFunctionProperties)
register_element_cls("m:groupChr", CT_OMathGroupChar)
register_element_cls("m:groupChrPr", CT_OMathGroupCharProperties)
register_element_cls("m:lim", CT_OMathLimit)
register_element_cls("m:limLow", CT_OMathLimitLower)
register_element_cls("m:limLowPr", CT_OMathLimitLowerProperties)
register_element_cls("m:limUpp", CT_OMathLimitUpper)
register_element_cls("m:limUppPr", CT_OMathLimitUpperProperties)
register_element_cls("m:m", CT_OMathMatrix)
register_element_cls("m:mc", CT_OMathMatrixColumn)
register_element_cls("m:mcPr", CT_OMathMatrixColumnProperties)
register_element_cls("m:mPr", CT_OMathMatrixProperties)
register_element_cls("m:mr", CT_OMathMatrixRow)
register_element_cls("m:nary", CT_OMathNary)
register_element_cls("m:naryPr", CT_OMathNaryProperties)
register_element_cls("m:num", CT_OMathNumerator)
register_element_cls("m:oMath", CT_OMath)
register_element_cls("m:oMathPara", CT_OMathPara)
register_element_cls("m:rad", CT_OMathRad)
register_element_cls("m:radPr", CT_OMathRadProperties)
register_element_cls("m:r", CT_OMathRun)
register_element_cls("m:rPr", CT_OMathRunProperties)
register_element_cls("m:sSub", CT_OMathSub)
register_element_cls("m:sSubPr", CT_OMathSubProperties)
register_element_cls("m:sSubSup", CT_OMathSubSup)
register_element_cls("m:sSubSupPr", CT_OMathSubSupProperties)
register_element_cls("m:sSup", CT_OMathSup)
register_element_cls("m:sSupPr", CT_OMathSupProperties)
register_element_cls("m:sub", CT_OMathSubArgument)
register_element_cls("m:sup", CT_OMathSupArgument)
register_element_cls("m:t", CT_OMathText)
register_element_cls("m:ctrlPr", CT_OMathControlProperties)


from pptx.oxml.theme import (  # noqa: E402
    CT_BaseStyles,
    CT_ColorScheme,
    CT_FormatScheme,
    CT_FontScheme,
    CT_OfficeStyleSheet,
)

register_element_cls("a:theme", CT_OfficeStyleSheet)
register_element_cls("a:themeElements", CT_BaseStyles)
register_element_cls("a:clrScheme", CT_ColorScheme)
register_element_cls("a:fontScheme", CT_FontScheme)
register_element_cls("a:fmtScheme", CT_FormatScheme)
