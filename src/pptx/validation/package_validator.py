"""Package-level validation for OPC package structure.

Validates that a .pptx package has the required parts and relationships.

Corresponds to DocumentFormat.OpenXml.Validation.PackageValidator in the
Open XML SDK.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.validation.context import ValidationContext

if TYPE_CHECKING:
    from pptx.opc.package import OpcPackage, Part


class PackageValidator:
    """Validates the package structure of a Presentation."""

    def validate(self, ctx: ValidationContext, package: OpcPackage):
        self._check_main_document_part(ctx, package)
        self._check_relationship_targets(ctx, package)

    def _check_main_document_part(self, ctx: ValidationContext, package: OpcPackage):
        try:
            main_part = package.main_document_part
        except KeyError:
            ctx.create_package_error(
                id="Pkg_RequiredPartDoNotExist",
                description="Package is missing the main document part "
                "(presentation.xml). No OFFICE_DOCUMENT relationship found.",
            )
            return

        self._check_presentation_structure(ctx, main_part)

    def _check_presentation_structure(self, ctx: ValidationContext, pres_part: Part):
        try:
            sldIdLst = pres_part._element.sldIdLst
        except Exception:
            ctx.create_package_error(
                id="Sch_MissingRequiredChild",
                description="Presentation part is missing required <p:sldIdLst> element.",
                part=pres_part,
            )
            return

        if sldIdLst is None:
            ctx.create_package_error(
                id="Sch_MissingRequiredChild",
                description="Presentation part has no <p:sldIdLst> element.",
                part=pres_part,
            )
            return

        slide_count = 0
        for sldId in sldIdLst:
            try:
                slide_part = pres_part.related_part(sldId.rId)
                slide_count += 1
                self._check_slide_structure(ctx, slide_part)
            except KeyError:
                ctx.create_package_error(
                    id="Pkg_InvalidPartUri",
                    description=f"Slide relationship rId='{sldId.rId}' "
                    f"references a part that does not exist.",
                    part=pres_part,
                )

        if slide_count == 0:
            ctx.create_package_error(
                id="Pkg_RequiredPartDoNotExist",
                description="Presentation has no slides. At least one slide is required.",
                part=pres_part,
            )

    def _check_slide_structure(self, ctx: ValidationContext, slide_part: Part):
        try:
            slide_part.part_related_by(RT.SLIDE_LAYOUT)
        except KeyError:
            ctx.create_package_error(
                id="Pkg_RequiredPartDoNotExist",
                description=f"Slide part '{slide_part.partname}' is missing "
                f"required slideLayout relationship.",
                part=slide_part,
            )

    def _check_relationship_targets(self, ctx: ValidationContext, package: OpcPackage):
        for rel in package.iter_rels():
            if rel.is_external:
                continue
            try:
                rel.target_part
            except KeyError:
                ctx.create_package_error(
                    id="Pkg_InvalidPartUri",
                    description=f"Relationship rId='{rel.rId}' references "
                    f"non-existent target.",
                )
