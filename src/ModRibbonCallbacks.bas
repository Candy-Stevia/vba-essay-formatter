Attribute VB_Name = "ModRibbonCallbacks"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModRibbonCallbacks
' Purpose: Fluent UI Ribbon callbacks & in-memory configuration state management
' Architecture: Academic Document Format Engine (Pure RAM State, No Dirtying Document)
' ==============================================================================

Public Type EngineRibbonConfig
    FontName As String
    FontSize As Single
    LineSpacingMode As String
    IndentPts As Single
    IsBoldTOC As Boolean
    IsFooterPage As Boolean
    TocDepth As Long
    ListLang As String
    CrossRefColor As String
End Type

Public g_Config As EngineRibbonConfig
Public g_RibbonUI As Object
Public g_LastError As String
Private g_Initialized As Boolean

Private Const REG_APP As String = "AcademicFormatter"
Private Const REG_SEC As String = "Config"

Private Function NormalizeLineSpacing(ByVal mode As String) As String
    Select Case LCase$(Trim$(mode))
        Case "ls_115"
            NormalizeLineSpacing = "ls_115"
        Case "ls_130"
            NormalizeLineSpacing = "ls_130"
        Case "ls_150"
            NormalizeLineSpacing = "ls_150"
        Case Else
            NormalizeLineSpacing = "ls_keep"
    End Select
End Function

Private Function NormalizeCrossRefColor(ByVal colorId As String) As String
    Select Case LCase$(Trim$(colorId))
        Case "xref_orange"
            NormalizeCrossRefColor = "xref_orange"
        Case "xref_red"
            NormalizeCrossRefColor = "xref_red"
        Case "xref_black"
            NormalizeCrossRefColor = "xref_black"
        Case "xref_off"
            NormalizeCrossRefColor = "xref_off"
        Case Else
            NormalizeCrossRefColor = "xref_blue"
    End Select
End Function

Private Sub SaveConfigSetting(ByVal keyName As String, ByVal valStr As String)
    On Error Resume Next
    If Application.DisplayAlerts = 0 Then
        Exit Sub
    End If
    SaveSetting REG_APP, REG_SEC, keyName, valStr
    Err.Clear
End Sub

Private Sub LoadConfigFromRegistry()
    On Error Resume Next
    If Application.DisplayAlerts = 0 Then
        Exit Sub
    End If
    Dim sVal As String
    sVal = GetSetting(REG_APP, REG_SEC, "FontName", "")
    Select Case LCase$(Trim$(sVal))
        Case "arial"
            g_Config.FontName = "Arial"
        Case "calibri"
            g_Config.FontName = "Calibri"
        Case "times new roman"
            g_Config.FontName = "Times New Roman"
    End Select
    sVal = GetSetting(REG_APP, REG_SEC, "FontSize", "")
    Select Case Trim$(sVal)
        Case "12"
            g_Config.FontSize = 12#
        Case "14"
            g_Config.FontSize = 14#
        Case "13"
            g_Config.FontSize = 13#
    End Select
    sVal = GetSetting(REG_APP, REG_SEC, "LineSpacingMode", "")
    If Len(sVal) > 0 Then
        g_Config.LineSpacingMode = NormalizeLineSpacing(sVal)
    End If
    sVal = GetSetting(REG_APP, REG_SEC, "IndentId", "")
    Select Case LCase$(Trim$(sVal))
        Case "ind_100"
            g_Config.IndentPts = ModConstants.CmToPoints(1#)
        Case "ind_000"
            g_Config.IndentPts = 0#
        Case "ind_127"
            g_Config.IndentPts = ModConstants.CmToPoints(1.27)
    End Select
    sVal = GetSetting(REG_APP, REG_SEC, "IsBoldTOC", "")
    If sVal = "0" Then
        g_Config.IsBoldTOC = False
    ElseIf sVal = "1" Then
        g_Config.IsBoldTOC = True
    End If
    sVal = GetSetting(REG_APP, REG_SEC, "IsFooterPage", "")
    If sVal = "0" Then
        g_Config.IsFooterPage = False
    ElseIf sVal = "1" Then
        g_Config.IsFooterPage = True
    End If
    sVal = GetSetting(REG_APP, REG_SEC, "TocDepth", "")
    If sVal = "3" Then
        g_Config.TocDepth = 3
    ElseIf sVal = "4" Then
        g_Config.TocDepth = 4
    End If
    sVal = GetSetting(REG_APP, REG_SEC, "ListLang", "")
    If sVal = "lang_en" Then
        g_Config.ListLang = "lang_en"
    ElseIf sVal = "lang_vi" Then
        g_Config.ListLang = "lang_vi"
    End If
    sVal = GetSetting(REG_APP, REG_SEC, "CrossRefColor", "")
    If Len(sVal) > 0 Then
        g_Config.CrossRefColor = NormalizeCrossRefColor(sVal)
    End If
    Err.Clear
End Sub

Public Sub InvalidateRibbonSafe()
    On Error Resume Next
    If Not g_RibbonUI Is Nothing Then
        g_RibbonUI.Invalidate
    End If
    Err.Clear
End Sub

Private Sub InvalidateControlSafe(ByVal ctrlId As String)
    On Error Resume Next
    If Not g_RibbonUI Is Nothing Then
        If Len(ctrlId) > 0 Then
            g_RibbonUI.InvalidateControl ctrlId
        End If
    End If
    Err.Clear
End Sub

Public Sub Ribbon_GetControlsEnabled(ByVal control As Object, ByRef returnedVal As Variant)
    returnedVal = (Not ModMain.g_IsEngineRunning) And (Not ModMathController.g_IsMathBuilding)
End Sub

Public Sub EnsureConfigInitialized()
    If Not g_Initialized Then
        g_Config.FontName = "Times New Roman"
        g_Config.FontSize = 13#
        g_Config.LineSpacingMode = "ls_keep"
        g_Config.IndentPts = ModConstants.CmToPoints(1.27)
        g_Config.IsBoldTOC = True
        g_Config.IsFooterPage = True
        g_Config.TocDepth = 4
        g_Config.ListLang = "lang_vi"
        g_Config.CrossRefColor = "xref_blue"
        LoadConfigFromRegistry
        g_Initialized = True
    End If
End Sub

Public Sub Ribbon_OnLoad(ByVal ribbon As Object)
    Set g_RibbonUI = ribbon
    EnsureConfigInitialized
End Sub

Public Sub Ribbon_GetSelectedFont(ByVal control As Object, ByRef itemID As Variant)
    EnsureConfigInitialized
    Select Case LCase$(g_Config.FontName)
        Case "arial"
            itemID = "font_arial"
        Case "calibri"
            itemID = "font_calibri"
        Case Else
            itemID = "font_tnr"
    End Select
End Sub

Public Sub Ribbon_OnFontChange(ByVal control As Object, ByVal id As String, ByVal index As Long)
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        InvalidateControlSafe "ddFontName"
        Exit Sub
    End If
    EnsureConfigInitialized
    Select Case id
        Case "font_arial"
            g_Config.FontName = "Arial"
        Case "font_calibri"
            g_Config.FontName = "Calibri"
        Case Else
            g_Config.FontName = "Times New Roman"
    End Select
    SaveConfigSetting "FontName", g_Config.FontName
End Sub

Public Sub Ribbon_GetSelectedSize(ByVal control As Object, ByRef itemID As Variant)
    EnsureConfigInitialized
    Select Case CLng(g_Config.FontSize)
        Case 12
            itemID = "size_12"
        Case 14
            itemID = "size_14"
        Case Else
            itemID = "size_13"
    End Select
End Sub

Public Sub Ribbon_OnSizeChange(ByVal control As Object, ByVal id As String, ByVal index As Long)
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        InvalidateControlSafe "ddFontSize"
        Exit Sub
    End If
    EnsureConfigInitialized
    Select Case id
        Case "size_12"
            g_Config.FontSize = 12#
        Case "size_14"
            g_Config.FontSize = 14#
        Case Else
            g_Config.FontSize = 13#
    End Select
    SaveConfigSetting "FontSize", CStr(CLng(g_Config.FontSize))
End Sub

Public Sub Ribbon_GetSelectedLineSpacing(ByVal control As Object, ByRef itemID As Variant)
    EnsureConfigInitialized
    g_Config.LineSpacingMode = NormalizeLineSpacing(g_Config.LineSpacingMode)
    itemID = g_Config.LineSpacingMode
End Sub

Public Sub Ribbon_OnLineSpacingChange(ByVal control As Object, ByVal id As String, ByVal index As Long)
    SetLineSpacing id
End Sub

Public Sub SetLineSpacing(ByVal mode As String)
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        InvalidateControlSafe "ddLineSpacing"
        Exit Sub
    End If
    EnsureConfigInitialized
    g_Config.LineSpacingMode = NormalizeLineSpacing(mode)
    SaveConfigSetting "LineSpacingMode", g_Config.LineSpacingMode
    InvalidateControlSafe "ddLineSpacing"
End Sub

Public Sub Ribbon_GetSelectedIndent(ByVal control As Object, ByRef itemID As Variant)
    EnsureConfigInitialized
    If Abs(g_Config.IndentPts - ModConstants.CmToPoints(1#)) < 1# Then
        itemID = "ind_100"
    ElseIf g_Config.IndentPts < 1# Then
        itemID = "ind_000"
    Else
        itemID = "ind_127"
    End If
End Sub

Public Sub Ribbon_OnIndentChange(ByVal control As Object, ByVal id As String, ByVal index As Long)
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        InvalidateControlSafe "ddIndent"
        Exit Sub
    End If
    EnsureConfigInitialized
    Select Case id
        Case "ind_100"
            g_Config.IndentPts = ModConstants.CmToPoints(1#)
            SaveConfigSetting "IndentId", "ind_100"
        Case "ind_000"
            g_Config.IndentPts = 0#
            SaveConfigSetting "IndentId", "ind_000"
        Case Else
            g_Config.IndentPts = ModConstants.CmToPoints(1.27)
            SaveConfigSetting "IndentId", "ind_127"
    End Select
End Sub

Public Sub Ribbon_GetBoldTOC(ByVal control As Object, ByRef returnedVal As Variant)
    EnsureConfigInitialized
    returnedVal = g_Config.IsBoldTOC
End Sub

Public Sub Ribbon_OnBoldTOCToggle(ByVal control As Object, ByVal pressed As Boolean)
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        InvalidateControlSafe "chkBoldTOC"
        Exit Sub
    End If
    EnsureConfigInitialized
    g_Config.IsBoldTOC = pressed
    If pressed Then
        SaveConfigSetting "IsBoldTOC", "1"
    Else
        SaveConfigSetting "IsBoldTOC", "0"
    End If
End Sub

Public Sub Ribbon_GetFooterPage(ByVal control As Object, ByRef returnedVal As Variant)
    EnsureConfigInitialized
    returnedVal = g_Config.IsFooterPage
End Sub

Public Sub Ribbon_OnFooterPageToggle(ByVal control As Object, ByVal pressed As Boolean)
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        InvalidateControlSafe "chkFooterPage"
        Exit Sub
    End If
    EnsureConfigInitialized
    g_Config.IsFooterPage = pressed
    If pressed Then
        SaveConfigSetting "IsFooterPage", "1"
    Else
        SaveConfigSetting "IsFooterPage", "0"
    End If
End Sub

Public Sub Ribbon_GetTocDepth(ByVal control As Object, ByRef itemID As Variant)
    EnsureConfigInitialized
    If g_Config.TocDepth = 3 Then itemID = "depth_3" Else itemID = "depth_4"
End Sub

Public Sub SetTocDepth(ByVal depthId As String)
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        InvalidateControlSafe "ddTocDepth"
        Exit Sub
    End If
    EnsureConfigInitialized
    If depthId = "depth_3" Then g_Config.TocDepth = 3 Else g_Config.TocDepth = 4
    SaveConfigSetting "TocDepth", CStr(g_Config.TocDepth)
    InvalidateControlSafe "ddTocDepth"
End Sub

Public Sub Ribbon_OnTocDepthChange(ByVal control As Object, ByVal id As String, ByVal index As Long)
    SetTocDepth id
End Sub

Public Sub Ribbon_GetSelectedLang(ByVal control As Object, ByRef itemID As Variant)
    EnsureConfigInitialized
    If g_Config.ListLang = "lang_en" Then itemID = "lang_en" Else itemID = "lang_vi"
End Sub

Public Sub SetListLanguage(ByVal lang As String)
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        InvalidateControlSafe "ddListLang"
        Exit Sub
    End If
    EnsureConfigInitialized
    If lang = "lang_en" Then g_Config.ListLang = "lang_en" Else g_Config.ListLang = "lang_vi"
    SaveConfigSetting "ListLang", g_Config.ListLang
    InvalidateControlSafe "ddListLang"
End Sub

Public Function GetCurrentListLang() As String
    EnsureConfigInitialized
    GetCurrentListLang = g_Config.ListLang
End Function

Public Sub Ribbon_OnLangChange(ByVal control As Object, ByVal id As String, ByVal index As Long)
    SetListLanguage id
End Sub

Public Sub Ribbon_GetCrossRefColor(ByVal control As Object, ByRef itemID As Variant)
    EnsureConfigInitialized
    g_Config.CrossRefColor = NormalizeCrossRefColor(g_Config.CrossRefColor)
    itemID = g_Config.CrossRefColor
End Sub

Public Sub Ribbon_OnCrossRefColorChange(ByVal control As Object, ByVal id As String, ByVal index As Long)
    SetCrossRefColor id
End Sub

Public Sub SetCrossRefColor(ByVal colorId As String)
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        InvalidateControlSafe "ddCrossRefColor"
        Exit Sub
    End If
    EnsureConfigInitialized
    g_Config.CrossRefColor = NormalizeCrossRefColor(colorId)
    SaveConfigSetting "CrossRefColor", g_Config.CrossRefColor
    InvalidateControlSafe "ddCrossRefColor"
End Sub

Private Function CanRunDocumentAction() As Boolean
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        CanRunDocumentAction = False
        Exit Function
    End If
    If Application.Documents.Count = 0 Then
        ModUI.MsgBoxW ModConstants.Uni("No open Word document found. (Kh\u00f4ng t\u00ecm th\u1ea5y t\u00e0i li\u1ec7u Word \u0111ang m\u1edf.)"), _
                      vbExclamation, ModConstants.Uni("Error (L\u1ed7i)")
        CanRunDocumentAction = False
        Exit Function
    End If
    CanRunDocumentAction = True
End Function

Public Sub RunEngineFull()
    If Not CanRunDocumentAction() Then
        Exit Sub
    End If
    EnsureConfigInitialized
    ModMain.RunEngine ActiveDocument, EngineModeFull
End Sub

Public Sub RunEngineListsOnly()
    If Not CanRunDocumentAction() Then
        Exit Sub
    End If
    EnsureConfigInitialized
    ModMain.RunEngine ActiveDocument, EngineModeListsOnly
End Sub

Public Sub RunCleanMarkers()
    If Not CanRunDocumentAction() Then
        Exit Sub
    End If
    ModZoneManager.CleanMarkersForSubmission ActiveDocument
End Sub

Public Sub Ribbon_OnFullRun(Optional ByVal control As Object = Nothing)
    RunEngineFull
End Sub

Public Sub Ribbon_OnUpdateListsOnly(Optional ByVal control As Object = Nothing)
    RunEngineListsOnly
End Sub

Public Sub RunLinkCrossRefsOnly()
    If Not CanRunDocumentAction() Then
        Exit Sub
    End If
    ModCrossRefOnly.RunLinkCrossRefsOnly ActiveDocument
End Sub

Public Sub Ribbon_OnLinkCrossRefsOnly(Optional ByVal control As Object = Nothing)
    RunLinkCrossRefsOnly
End Sub

Public Sub Ribbon_OnCleanMarkers(Optional ByVal control As Object = Nothing)
    RunCleanMarkers
End Sub

Public Sub Ribbon_OnAiPatchClipboard(Optional ByVal control As Object = Nothing)
    ModPatchEngine.RunAiPatchClipboard
End Sub

Public Sub Ribbon_OnAiPatchFile(Optional ByVal control As Object = Nothing)
    ModPatchEngine.RunAiPatchFile
End Sub

Public Sub Ribbon_OnPreAuditScan(Optional ByVal control As Object = Nothing)
    ModCitationScan.RunPreAuditScan
End Sub

Public Sub RunCaptureStyle()
    If Not CanRunDocumentAction() Then
        Exit Sub
    End If
    ModStyleCapture.CaptureStyleFromDocument ActiveDocument, False
End Sub

Public Sub Ribbon_OnCaptureStyle(Optional ByVal control As Object = Nothing)
    RunCaptureStyle
End Sub

Public Sub RunResetCustomStyle()
    If Not CanRunDocumentAction() Then
        Exit Sub
    End If
    ModCustomProfile.ResetCustomProfileInteractive ActiveDocument
End Sub

Public Sub Ribbon_OnResetCustomStyle(Optional ByVal control As Object = Nothing)
    RunResetCustomStyle
End Sub

Public Sub RunExportPdf()
    If Not CanRunDocumentAction() Then
        Exit Sub
    End If
    ModUI.ExportDocumentToPdf ActiveDocument
End Sub

Public Sub Ribbon_OnExportPdf(Optional ByVal control As Object = Nothing)
    RunExportPdf
End Sub

Public Sub Ribbon_OnMathNextPink(Optional ByVal control As Object = Nothing)
    If Not CanRunDocumentAction() Then
        Exit Sub
    End If
    ModMathController.NavigateNextPinkMath ActiveDocument
End Sub

Public Sub Ribbon_OnInsertMarker(ByVal control As Object)
    If control Is Nothing Then
        Exit Sub
    End If
    If Not CanRunDocumentAction() Then
        Exit Sub
    End If
    Select Case control.Id
        Case "btnInsContentStart"
            ModMarkerInserter.InsertMarkerAtSelection ActiveDocument, MARKER_CONTENT_START
        Case "btnInsContentEnd"
            ModMarkerInserter.InsertMarkerAtSelection ActiveDocument, MARKER_CONTENT_END
        Case "btnInsListContent"
            ModMarkerInserter.InsertMarkerAtSelection ActiveDocument, MARKER_LIST_CONTENT
        Case "btnInsListTable"
            ModMarkerInserter.InsertMarkerAtSelection ActiveDocument, MARKER_LIST_TABLE
        Case "btnInsListFigure"
            ModMarkerInserter.InsertMarkerAtSelection ActiveDocument, MARKER_LIST_FIGURE
        Case "btnInsListAcronyms"
            ModMarkerInserter.InsertMarkerAtSelection ActiveDocument, MARKER_LIST_ACRONYMS
    End Select
End Sub

Public Function GetLastError() As String
    GetLastError = g_LastError
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
