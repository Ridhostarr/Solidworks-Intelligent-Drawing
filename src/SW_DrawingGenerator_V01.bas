Attribute VB_Name = "SW_DrawingGenerator_V01"
Option Explicit

'=========================================
' SOLIDWORKS DRAWING GENERATOR - V0.1
' Target: SOLIDWORKS 2024
' Source: preceding ChatGPT v0.1 prototype; reviewed fixes documented in CHANGELOG.md
' Input: Active saved SLDPRT
' Output: SLDDRW and PDF
'=========================================

Private Const TEMPLATE_PATH As String = _
    "C:\SW_Templates\Company_A3.drwdot"

Private Const IMPORT_MARKED_DIMS As Boolean = True
Private Const EXPORT_PDF As Boolean = True

' Change to False for third-angle layout
Private Const FIRST_ANGLE As Boolean = True

Sub main()

    On Error GoTo ErrorHandler

    Dim swApp As SldWorks.SldWorks
    Dim swPartModel As SldWorks.ModelDoc2
    Dim swPart As SldWorks.PartDoc
    Dim swModel As SldWorks.ModelDoc2
    Dim swDraw As SldWorks.DrawingDoc
    Dim swSheet As SldWorks.Sheet
    Dim swView As SldWorks.View

    Dim partPath As String
    Dim basePath As String
    Dim drawingPath As String
    Dim pdfPath As String
    Dim cfg As String
    Dim frontViewName As String

    Dim material As String
    Dim database As String
    Dim process As String
    Dim toleranceNote As String
    Dim materialSpec As String

    Dim props As Variant
    Dim sheetW As Double
    Dim sheetH As Double
    Dim ok As Boolean
    Dim errCode As Long
    Dim warnCode As Long

    Set swApp = Application.SldWorks
    Set swPartModel = swApp.ActiveDoc

    If swPartModel Is Nothing Then
        MsgBox "Open a SOLIDWORKS part first."
        Exit Sub
    End If

    If swPartModel.GetType <> swDocPART Then
        MsgBox "Only part documents are supported."
        Exit Sub
    End If

    partPath = swPartModel.GetPathName

    If Len(partPath) = 0 Then
        MsgBox "Save the part before running."
        Exit Sub
    End If

    If Dir(TEMPLATE_PATH) = "" Then
        MsgBox "Drawing template not found:" & _
            vbCrLf & TEMPLATE_PATH
        Exit Sub
    End If

    basePath = Left$(partPath, _
        InStrRev(partPath, ".") - 1)

    drawingPath = basePath & "_AUTO_DRAFT.SLDDRW"
    pdfPath = basePath & "_AUTO_DRAFT.pdf"

    ' Prevent accidental overwriting
    If Dir(drawingPath) <> "" Or _
       Dir(pdfPath) <> "" Then
        MsgBox "Draft output already exists." & _
            vbCrLf & "Rename or move it first."
        Exit Sub
    End If

    cfg = swPartModel.ConfigurationManager.ActiveConfiguration.Name

    Set swPart = swPartModel

    '---------- READ MODEL DATA ----------

    material = swPart.GetMaterialPropertyName2( _
        cfg, database)

    materialSpec = GetModelProperty( _
        swPartModel, cfg, "MaterialSpec")

    If Len(materialSpec) = 0 Then
        materialSpec = material
    End If

    If Len(materialSpec) = 0 Then
        materialSpec = "NOT SPECIFIED"
    End If

    process = GetModelProperty( _
        swPartModel, cfg, "ManufacturingProcess")

    If Len(process) = 0 Then
        process = "NOT SPECIFIED"
    End If

    toleranceNote = GetModelProperty( _
        swPartModel, cfg, "GeneralTolerance")

    If Len(toleranceNote) = 0 Then
        toleranceNote = _
            "GENERAL TOLERANCES NOT APPROVED"
    End If

    '---------- CREATE DRAWING ----------

    Set swModel = swApp.NewDocument( _
        TEMPLATE_PATH, 0, 0, 0)

    If swModel Is Nothing Then
        Err.Raise vbObjectError + 100, , _
            "Failed to create drawing."
    End If

    If swModel.GetType <> swDocDRAWING Then
        Err.Raise vbObjectError + 101, , _
            "Template did not create a drawing."
    End If

    Set swDraw = swModel
    Set swSheet = swDraw.GetCurrentSheet

    If swSheet Is Nothing Then
        Err.Raise vbObjectError + 104, , "No current drawing sheet."
    End If

    props = swSheet.GetProperties2
    sheetW = CDbl(props(5))
    sheetH = CDbl(props(6))
    If sheetW <= 0# Or sheetH <= 0# Then
        Err.Raise vbObjectError + 105, , "Invalid sheet dimensions."
    End If

    ' Check projection matches the template
    If CBool(props(4)) <> FIRST_ANGLE Then
        MsgBox "Projection convention mismatch." & _
            vbCrLf & _
            "Check FIRST_ANGLE and template settings."
        Exit Sub
    End If

    '---------- CREATE VIEWS ----------

    ' Positions are in metres, relative to
    ' the drawing sheet's bottom-left origin.

    Set swView = AddView(swDraw, partPath, _
        "*Front", sheetW * 0.28, _
        sheetH * 0.64, cfg)

    If swView Is Nothing Then
        Err.Raise vbObjectError + 102, , _
            "Failed to create front view."
    End If

    frontViewName = swView.GetName2

    If FIRST_ANGLE Then
        ' Top view below front; right view left
        Set swView = AddView(swDraw, partPath, _
            "*Top", sheetW * 0.28, _
            sheetH * 0.32, cfg)

        Set swView = AddView(swDraw, partPath, _
            "*Right", sheetW * 0.09, _
            sheetH * 0.64, cfg)
    Else
        ' Third-angle layout
        Set swView = AddView(swDraw, partPath, _
            "*Top", sheetW * 0.28, _
            sheetH * 0.88, cfg)

        Set swView = AddView(swDraw, partPath, _
            "*Right", sheetW * 0.51, _
            sheetH * 0.64, cfg)
    End If

    Set swView = AddView(swDraw, partPath, _
        "*Isometric", sheetW * 0.76, _
        sheetH * 0.67, cfg)

    '---------- IMPORT MODEL DIMENSIONS ----------

    If IMPORT_MARKED_DIMS Then
        ImportMarkedDimensions swModel, swDraw, frontViewName
    End If

    '---------- ADD DRAWING NOTES ----------

    swModel.ClearSelection2 True
    If Not swDraw.ActivateView("") Then
        Err.Raise vbObjectError + 109, , "Cannot activate sheet for drawing note."
    End If

    Dim notes As String

    notes = "AUTOMATED DRAWING - REVIEW REQUIRED" & _
        vbCrLf & _
        "MATERIAL: " & materialSpec & vbCrLf & _
        "PROCESS: " & process & vbCrLf & _
        "TOLERANCE: " & toleranceNote

    AddDrawingNote swModel, notes, _
        sheetW * 0.55, sheetH * 0.30

    '---------- REBUILD AND SAVE ----------

    swModel.ClearSelection2 True
    swModel.ForceRebuild3 False

    ok = swModel.Extension.SaveAs( _
        drawingPath, swSaveAsCurrentVersion, _
        swSaveAsOptions_Silent, Nothing, _
        errCode, warnCode)

    If Not ok Or errCode <> 0 Then
        Err.Raise vbObjectError + 103, , _
            "Drawing save failed: " & errCode
    End If

    Dim saveWarnings As String
    If warnCode <> 0 Then saveWarnings = "Drawing save warnings: " & warnCode & vbCrLf

    If EXPORT_PDF Then
        swModel.ClearSelection2 True
        errCode = 0
        warnCode = 0

        ok = swModel.Extension.SaveAs( _
            pdfPath, swSaveAsCurrentVersion, _
            swSaveAsOptions_Silent, Nothing, _
            errCode, warnCode)

        If Not ok Or errCode <> 0 Then
            MsgBox "Drawing saved at " & drawingPath & vbCrLf & _
                "PDF failed. Error: " & errCode & "; warnings: " & warnCode
            Exit Sub
        End If
        If warnCode <> 0 Then saveWarnings = saveWarnings & "PDF save warnings: " & warnCode & vbCrLf
    End If

    MsgBox "Draft drawing generated:" & _
        vbCrLf & drawingPath & vbCrLf & _
        saveWarnings & "Engineering review is required."

    Exit Sub

ErrorHandler:
    MsgBox "Drawing generator error:" & _
        vbCrLf & Err.Description & vbCrLf & _
        "Partial drawing/files may remain. Inspect them before retrying.", vbCritical

End Sub

'=========================================
' CREATE DRAWING VIEW
'=========================================
Private Function AddView( _
    ByVal swDraw As SldWorks.DrawingDoc, _
    ByVal modelPath As String, _
    ByVal viewName As String, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal cfg As String) As SldWorks.View

    Dim v As SldWorks.View

    Set v = swDraw.CreateDrawViewFromModelView3( _
        modelPath, viewName, x, y, 0#)

    If v Is Nothing Then
        Err.Raise vbObjectError + 106, , "Failed to create view: " & viewName
    End If
    v.ReferencedConfiguration = cfg
    If StrComp(v.ReferencedConfiguration, cfg, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 107, , "View configuration mismatch: " & viewName
    End If

    Set AddView = v

End Function

'=========================================
' READ CUSTOM PROPERTY
'=========================================
Private Function GetModelProperty( _
    ByVal model As SldWorks.ModelDoc2, _
    ByVal cfg As String, _
    ByVal propertyName As String) As String

    Dim mgr As SldWorks.CustomPropertyManager
    Dim rawValue As String
    Dim resolvedValue As String
    Dim wasResolved As Boolean
    Dim isLinked As Boolean
    Dim result As Long

    ' First check configuration properties
    Set mgr = model.Extension.CustomPropertyManager(cfg)

    result = mgr.Get6(propertyName, False, _
        rawValue, resolvedValue, _
        wasResolved, isLinked)

    ' Fallback to document-level properties
    If result = swCustomInfoGetResult_NotPresent Then
        Set mgr = model.Extension.CustomPropertyManager("")

        result = mgr.Get6(propertyName, False, _
            rawValue, resolvedValue, _
            wasResolved, isLinked)
    End If

    If result <> swCustomInfoGetResult_NotPresent Then
        If wasResolved Then GetModelProperty = Trim$(resolvedValue)
    End If

End Function

'=========================================
' IMPORT MARKED MODEL DIMENSIONS
'=========================================
Private Sub ImportMarkedDimensions( _
    ByVal model As SldWorks.ModelDoc2, _
    ByVal drawing As SldWorks.DrawingDoc, _
    ByVal viewName As String)

    Dim annotations As Variant

    model.ClearSelection2 True

    If Not drawing.ActivateView(viewName) Then
        Err.Raise vbObjectError + 108, , "Cannot activate view for dimension import."
    End If

    annotations = drawing.InsertModelAnnotations3( _
        swImportModelItemsFromEntireModel, _
        swInsertDimensionsMarkedForDrawing, _
        True, True, False, False)

    model.ClearSelection2 True

End Sub

'=========================================
' CREATE DRAWING NOTE
'=========================================
Private Sub AddDrawingNote( _
    ByVal model As SldWorks.ModelDoc2, _
    ByVal noteText As String, _
    ByVal x As Double, _
    ByVal y As Double)

    Dim nt As SldWorks.Note
    Dim ann As SldWorks.Annotation

    model.ClearSelection2 True

    Set nt = model.InsertNote(noteText)

    If nt Is Nothing Then
        Err.Raise vbObjectError + 110, , "Could not insert review note."
    End If

    Set ann = nt.GetAnnotation

    If ann Is Nothing Then
        Err.Raise vbObjectError + 111, , "Review note has no annotation."
    End If
    If Not ann.SetPosition2(x, y, 0#) Then
        Err.Raise vbObjectError + 112, , "Could not position review note."
    End If

    model.ClearSelection2 True

End Sub

