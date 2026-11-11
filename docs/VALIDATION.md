# Validation record

## Review completed

Source inspection and official API documentation review were performed on 2026-11-11. The original macro was retrieved from the preceding project conversation. Corrections are listed in `CHANGELOG.md`.

Checked property fallback status, sheet-property indices, view-creation error paths, annotation-import argument meanings, VBA continuation/object assignment style, note failures, and save error/warning reporting. Repository files were checked for whitespace errors and common credential patterns before upload.

This is **not** a VBA compilation or a SOLIDWORKS execution result. SOLIDWORKS 2024, its COM type libraries, the company template and representative parts are unavailable in this environment. The Windows checks below remain pending.

## Official references

- [2024 custom-property result codes](https://help.solidworks.com/2024/english/api/swconst/SolidWorks.Interop.swconst~SolidWorks.Interop.swconst.swCustomInfoGetResult_e.html?id=2.10.1.216): missing property is `1`, cached success is `0`.
- [Get6](https://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ICustomPropertyManager~Get6.html): evaluated values and resolution status; available since 2018.
- [GetProperties2](https://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isheet~getproperties2.html): sheet array indices; available since 2016.
- [CreateDrawViewFromModelView3](https://help.solidworks.com/2021/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IDrawingDoc~CreateDrawViewFromModelView3.html): exact model-view names, metre coordinates, automatic-scaling preference.
- [InsertModelAnnotations3](https://help.solidworks.com/2016/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IDrawingDoc~InsertModelAnnotations3.html): `DuplicateDims=True` requests elimination of duplicates.
- [Official VBA annotation example](https://help.solidworks.com/2021/English/api/sldworksapi/Insert_Model_Annotations_Example_VB.htm): view activation before import.
- [2024 note example](https://help.solidworks.com/2024/English/api/sldworksapi/Insert_a_Note_Example_VB.htm): note annotation and `SetPosition2`.

Older official help versions are linked where the 2024 web endpoint returned only a help shell. These are longstanding APIs; signatures still need confirmation against the installed 2024 libraries through compilation.

## Windows acceptance checklist — pending

1. Import module with both 2024 references, run **Debug > Compile**, and record SOLIDWORKS service pack and template used.
2. Use a saved simple bracket with holes, several dimensions marked for drawing, configuration material and approved custom-property text. Confirm four views reference the captured configuration and both exports open.
3. Check configuration-specific values override document properties; absent values fall back; present blank values and unresolved expressions result in warning/default note text. Confirm material database fallback separately.
4. Run with missing material/process/tolerance; confirm explicit review text. Check title-block model links independently.
5. Test first-angle and third-angle templates/settings. A mismatch must stop before view generation/export. Check scale with automatic scaling enabled and disabled.
6. Test imported geometry with no marked dimensions. Drawing may have no dimensions; inspect rather than treating save success as completeness.
7. Try no active document, an assembly, an unsaved part, missing template, and existing output files. Confirm no existing output is changed.
8. Test view-name failure and an unwritable output folder on disposable data. Confirm visible failure and inspect partial documents/files.
9. Disable PDF, then test PDF export separately. Confirm error/warning codes are shown and drawing survives PDF failure.
10. Inspect view overlap, off-sheet geometry, note clearance, units, precision, projection symbol, title block, material, tolerance wording, all marked dimensions and PDF contents before any release.

Record pass/fail results and screenshots in a separate local validation report. Do not commit proprietary models or templates as test fixtures without explicit authorization.
