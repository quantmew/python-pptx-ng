"""Initialization module for python-pptx package."""

from __future__ import annotations

import sys
from typing import TYPE_CHECKING

import pptx.exc as exceptions
from pptx.api import Presentation
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import PartFactory
from pptx.parts.chart import ChartPart
from pptx.parts.chartex import ChartExPart
from pptx.parts.comment import CommentAuthorsPart, CommentPart
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.diagram import (
    DiagramColorsPart,
    DiagramDataPart,
    DiagramLayoutPart,
    DiagramStylePart,
)
from pptx.parts.handout import HandoutMasterPart
from pptx.parts.image import ImagePart
from pptx.parts.media import MediaPart
from pptx.parts.model3d import Model3DBlobPart
from pptx.parts.presprops import PresentationPropertiesPart, ViewPropertiesPart
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import (
    NotesMasterPart,
    NotesSlidePart,
    SlideLayoutPart,
    SlideMasterPart,
    SlidePart,
)
from pptx.parts.tags import UserDefinedTagsPart

if TYPE_CHECKING:
    from pptx.opc.package import Part

__version__ = "1.0.2"

sys.modules["pptx.exceptions"] = exceptions
del sys

__all__ = ["Presentation"]

content_type_to_part_class_map: dict[str, type[Part]] = {
    CT.PML_PRESENTATION_MAIN: PresentationPart,
    CT.PML_PRES_MACRO_MAIN: PresentationPart,
    CT.PML_TEMPLATE_MAIN: PresentationPart,
    CT.PML_SLIDESHOW_MAIN: PresentationPart,
    CT.OPC_CORE_PROPERTIES: CorePropertiesPart,
    CT.PML_NOTES_MASTER: NotesMasterPart,
    CT.PML_NOTES_SLIDE: NotesSlidePart,
    CT.PML_PRES_PROPS: PresentationPropertiesPart,
    CT.PML_VIEW_PROPS: ViewPropertiesPart,
    CT.PML_COMMENTS: CommentPart,
    CT.PML_COMMENT_AUTHORS: CommentAuthorsPart,
    CT.PML_HANDOUT_MASTER: HandoutMasterPart,
    CT.PML_SLIDE: SlidePart,
    CT.PML_SLIDE_LAYOUT: SlideLayoutPart,
    CT.PML_SLIDE_MASTER: SlideMasterPart,
    CT.DML_CHART: ChartPart,
    CT.DML_DIAGRAM_COLORS: DiagramColorsPart,
    CT.DML_DIAGRAM_DATA: DiagramDataPart,
    CT.DML_DIAGRAM_LAYOUT: DiagramLayoutPart,
    CT.DML_DIAGRAM_STYLE: DiagramStylePart,
    CT.OFC_CHART_EX: ChartExPart,
    CT.AIFF: MediaPart,
    CT.AUDIO: MediaPart,
    CT.AUDIO_M4A: MediaPart,
    CT.AUDIO_MIDI: MediaPart,
    CT.AUDIO_MP3: MediaPart,
    CT.AUDIO_OGG: MediaPart,
    CT.AUDIO_WAV: MediaPart,
    CT.AUDIO_WMA: MediaPart,
    CT.BMP: ImagePart,
    CT.GIF: ImagePart,
    CT.JPEG: ImagePart,
    CT.MS_PHOTO: ImagePart,
    CT.PNG: ImagePart,
    CT.TIFF: ImagePart,
    CT.X_EMF: ImagePart,
    CT.X_WMF: ImagePart,
    CT.ASF: MediaPart,
    CT.AVI: MediaPart,
    CT.MOV: MediaPart,
    CT.MP4: MediaPart,
    CT.MPG: MediaPart,
    CT.MS_VIDEO: MediaPart,
    CT.SWF: MediaPart,
    CT.VIDEO: MediaPart,
    CT.WMV: MediaPart,
    CT.X_MS_VIDEO: MediaPart,
    CT.PML_TAGS: UserDefinedTagsPart,
    CT.MODEL_GLB: Model3DBlobPart,
    # -- accommodate "image/jpg" as an alias for "image/jpeg" --
    "image/jpg": ImagePart,
}

PartFactory.part_type_for.update(content_type_to_part_class_map)

del (
    ChartPart,
    ChartExPart,
    CommentAuthorsPart,
    CommentPart,
    CorePropertiesPart,
    DiagramColorsPart,
    DiagramDataPart,
    DiagramLayoutPart,
    DiagramStylePart,
    HandoutMasterPart,
    ImagePart,
    MediaPart,
    Model3DBlobPart,
    PresentationPropertiesPart,
    ViewPropertiesPart,
    SlidePart,
    SlideLayoutPart,
    SlideMasterPart,
    PresentationPart,
    UserDefinedTagsPart,
    CT,
    PartFactory,
)
