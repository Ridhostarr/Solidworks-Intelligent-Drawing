# Solidworks Intelligent Drawing

SOLIDWORKS 2024 VBA drawing generator v0.1. Creates a preliminary `.SLDDRW` and optional PDF from the active saved `.SLDPRT`, using your existing `.DRWDOT` template. Derived from the v0.1 macro in the preceding project conversation, with reviewed fixes listed in [CHANGELOG.md](CHANGELOG.md).

**Prototype: engineering review required.** Source/API review has been performed; this environment cannot compile VBA against the SOLIDWORKS type libraries or run SOLIDWORKS. No production validation is claimed.

## Setup and run

1. On Windows, open SOLIDWORKS 2024 and a saved test part. Save current edits first.
2. Choose **Tools > Macro > New**, and save a local `SW_DrawingGenerator_V01.swp`.
3. In the VBA editor, use **File > Import File** to import [src/SW_DrawingGenerator_V01.bas](src/SW_DrawingGenerator_V01.bas). Remove the newly created empty macro module if it also has a `main` procedure, or explicitly select the imported module's `main` when running.
4. Under **Tools > References**, enable the **SOLIDWORKS 2024 Type Library** and **SOLIDWORKS 2024 Constant type library**; resolve any `MISSING` references.
5. Edit `TEMPLATE_PATH` in the imported module to your `.DRWDOT` location. Use a template with one blank drawing sheet, its sheet format, and title block; predefined model views are not populated by this macro.
6. Set `FIRST_ANGLE` to match both the template's sheet projection setting and projection symbol. Set `IMPORT_MARKED_DIMS` and `EXPORT_PDF` as needed.
7. In SOLIDWORKS system options, clear **Drawings > Automatically scale new drawing views** if you want the template sheet scale preserved. The macro leaves this user preference unchanged; otherwise SOLIDWORKS may change the scale on insertion.
8. Choose **Debug > Compile** in the VBA editor, then save the `.swp`. Run its `SW_DrawingGenerator_V01.main` entry point using **Tools > Macro > Run**.
9. Inspect all views, dimensions, material, tolerances, title block, scale, and PDF using the [validation checklist](docs/VALIDATION.md).

The `.bas` is importable source, not a directly executable `.swp`. If pasting code instead, omit the `Attribute VB_Name` line; VBA's editor manages this hidden module attribute.

Outputs are placed alongside the part as `<part>_AUTO_DRAFT.SLDDRW` and `<part>_AUTO_DRAFT.pdf`. The macro refuses to run if either path already exists, even when PDF export is disabled. Move or rename prior drafts deliberately. Failures can leave a partial drawing open or partial files on disk; inspect these before retrying. A PDF failure does not remove the saved drawing. Saving is not transactional; avoid concurrent runs against the same output paths.

## Rules configuration

v0.1 uses module constants and model custom properties. It does **not** load an external rules file or infer approved tolerances from material.

| Setting | Default | Purpose |
| --- | --- | --- |
| `TEMPLATE_PATH` | Example local path | Existing drawing template |
| `FIRST_ANGLE` | `True` | View placement and projection validation |
| `IMPORT_MARKED_DIMS` | `True` | Import dimensions marked for drawing |
| `EXPORT_PDF` | `True` | Export PDF after saving drawing |

| Part property | v0.1 behavior |
| --- | --- |
| `MaterialSpec` | Preferred material note; falls back to assigned configuration material, then `NOT SPECIFIED` |
| `ManufacturingProcess` | Process note; defaults to `NOT SPECIFIED` |
| `GeneralTolerance` | Copies your approved wording verbatim; defaults to `GENERAL TOLERANCES NOT APPROVED` |
| `PartNumber`, `Description`, `Revision` | Available to your template's existing model-property links; no explicit title-block mapping |
| `SurfaceFinish` | Reserved for a later version; not read or applied |

Property lookup checks the active configuration first and falls back to document properties only when the configuration property is absent. A present blank property intentionally overrides the document value. Only evaluated values are used; missing or unresolved values produce the note fallbacks above. Material and tolerance notes are snapshots taken when the macro runs, not live property links. A nonempty tolerance property is not proof of engineering approval.

## API workflow

1. Validate active saved part, template path, and output paths; capture active configuration.
2. Read `ICustomPropertyManager.Get6` and `IPartDoc.GetMaterialPropertyName2`.
3. Create drawing via `ISldWorks.NewDocument`; read `ISheet.GetProperties2` and verify projection.
4. Use `IDrawingDoc.CreateDrawViewFromModelView3` for front, top, right and isometric views. Explicitly set each `IView.ReferencedConfiguration` and check creation.
5. Activate the front view and call `InsertModelAnnotations3` for marked dimensions across views, requesting duplicate elimination.
6. Activate the sheet, insert the review/material/process/tolerance note, and rebuild.
7. Save `.SLDDRW` and optionally PDF using `IModelDocExtension.SaveAs`; report errors and warning codes.

View centres and note positions are sheet fractions, expressed in metres from the bottom-left origin. Edit the `AddView` call coordinates and note position for your template's usable drawing area. First-angle top/right views are below/left of the front view; third-angle views are above/right. The views are independent model views, not associatively aligned projected views.

## Limitations

- Fixed views and positions; no geometry recognition, fit-to-sheet algorithm, overlap detection, title-block clearance, or automatic sheet selection.
- Scales depend on the template and SOLIDWORKS automatic-scaling preference described above.
- Imports existing marked model dimensions only. Imported STEP/Parasolid bodies may have none. No generated overall dimensions, hole callouts, datums, fits, GD&T, or individual tolerances.
- Duplicate elimination is requested from SOLIDWORKS; annotation completeness and placement still require review. No annotations returned is allowed and does not imply a fully dimensioned drawing.
- Mixed parts receive the same workflow. No sheet-metal flat pattern, bend table, cut list, per-body material treatment, or additive-specific logic.
- Intended for one active part configuration and one blank sheet. Assemblies, batch processing, multibody detail drawings, and multi-sheet templates are outside validation scope. PDF uses default `SaveAs` export behavior, with no explicit sheet selection or PDF options.
- Standard model-view names must exist exactly (`*Front`, `*Top`, `*Right`, `*Isometric`); localized or customized names may require edits.
- Existing `$PRPSHEET` title-block links must already refer to the appropriate model/configuration. Unlinked fields are not populated.
- The macro does not save or deliberately edit the source part. Review any SOLIDWORKS rebuild prompts or dirty-document state yourself.

## Roadmap

| Version | Planned work |
| --- | --- |
| v0.2 | Feature/geometry recognition, overall and hole dimensions, view selection, scale fitting, annotation placement |
| v0.3 | External approved material/process/tolerance tables; critical-feature tags; traceable rules and surface-finish requirements |
| v1.0 | Batch generation, quality reports, revision control, controlled release, possible C# add-in |

For later rules, separate material specifications, manufacturing processes, dimension schemes, and approved tolerance policies. Evaluate material, process, nominal size and functional criticality together; do not assign fits or GD&T from material alone.

## Repository contents and privacy

- `src/`: VBA source to import into SOLIDWORKS.
- `docs/VALIDATION.md`: review evidence and Windows acceptance checks.
- `CHANGELOG.md`: changes from the preceding prototype.
- `.gitignore`: excludes credentials, local settings, native CAD/template files and generated artifacts.

No company templates, CAD parts, tokens, passwords, or account configuration are included. Keep real templates and model files outside this repository. No license grant has been added; the owner can choose a license later.
