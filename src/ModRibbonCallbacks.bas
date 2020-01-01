Attribute VB_Name = "ModRibbonCallbacks"
Option Explicit

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
    EnsureConfigInitialized
    Select Case id
        Case "font_arial"
            g_Config.FontName = "Arial"
        Case "font_calibri"
            g_Config.FontName = "Calibri"
        Case Else
            g_Config.FontName = "Times New Roman"
    End Select
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
    EnsureConfigInitialized
    Select Case id
        Case "size_12"
            g_Config.FontSize = 12#
        Case "size_14"
            g_Config.FontSize = 14#
        Case Else
            g_Config.FontSize = 13#
    End Select
End Sub

Public Sub Ribbon_GetSelectedLineSpacing(ByVal control As Object, ByRef itemID As Variant)
    EnsureConfigInitialized
    itemID = g_Config.LineSpacingMode
End Sub

Public Sub Ribbon_OnLineSpacingChange(ByVal control As Object, ByVal id As String, ByVal index As Long)
    EnsureConfigInitialized
    g_Config.LineSpacingMode = id
End Sub

Public Sub SetLineSpacing(ByVal mode As String)
    EnsureConfigInitialized
    g_Config.LineSpacingMode = mode
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
    EnsureConfigInitialized
    Select Case id
        Case "ind_100"
            g_Config.IndentPts = ModConstants.CmToPoints(1#)
        Case "ind_000"
            g_Config.IndentPts = 0#
        Case Else
            g_Config.IndentPts = ModConstants.CmToPoints(1.27)
    End Select
End Sub

Public Sub Ribbon_GetBoldTOC(ByVal control As Object, ByRef returnedVal As Variant)
    EnsureConfigInitialized
    returnedVal = g_Config.IsBoldTOC
End Sub

Public Sub Ribbon_OnBoldTOCToggle(ByVal control As Object, ByVal pressed As Boolean)
    EnsureConfigInitialized
    g_Config.IsBoldTOC = pressed
End Sub

Public Sub Ribbon_GetFooterPage(ByVal control As Object, ByRef returnedVal As Variant)
    EnsureConfigInitialized
    returnedVal = g_Config.IsFooterPage
End Sub

Public Sub Ribbon_OnFooterPageToggle(ByVal control As Object, ByVal pressed As Boolean)
    EnsureConfigInitialized
    g_Config.IsFooterPage = pressed
End Sub

Public Sub Ribbon_GetTocDepth(ByVal control As Object, ByRef itemID As Variant)
    EnsureConfigInitialized
    If g_Config.TocDepth = 3 Then itemID = "depth_3" Else itemID = "depth_4"
End Sub

Public Sub Ribbon_OnTocDepthChange(ByVal control As Object, ByVal id As String, ByVal index As Long)
    EnsureConfigInitialized
    If id = "depth_3" Then g_Config.TocDepth = 3 Else g_Config.TocDepth = 4
End Sub

Public Sub Ribbon_GetSelectedLang(ByVal control As Object, ByRef itemID As Variant)
    EnsureConfigInitialized
    If g_Config.ListLang = "lang_en" Then itemID = "lang_en" Else itemID = "lang_vi"
End Sub

Public Sub SetListLanguage(ByVal lang As String)
    EnsureConfigInitialized
    If lang = "lang_en" Then g_Config.ListLang = "lang_en" Else g_Config.ListLang = "lang_vi"
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
    If Len(g_Config.CrossRefColor) = 0 Then g_Config.CrossRefColor = "xref_blue"
    itemID = g_Config.CrossRefColor
End Sub

Public Sub Ribbon_OnCrossRefColorChange(ByVal control As Object, ByVal id As String, ByVal index As Long)
    EnsureConfigInitialized
    g_Config.CrossRefColor = id
End Sub

Public Sub SetCrossRefColor(ByVal colorId As String)
    EnsureConfigInitialized
    g_Config.CrossRefColor = colorId
End Sub

Public Sub RunEngineFull()
    EnsureConfigInitialized
    ModMain.RunEngine ActiveDocument, EngineModeFull
End Sub

Public Sub RunEngineListsOnly()
    EnsureConfigInitialized
    ModMain.RunEngine ActiveDocument, EngineModeListsOnly
End Sub

Public Sub RunCleanMarkers()
    ModZoneManager.CleanMarkersForSubmission ActiveDocument
End Sub

Public Sub Ribbon_OnFullRun(Optional ByVal control As Object = Nothing)
    RunEngineFull
End Sub

Public Sub Ribbon_OnUpdateListsOnly(Optional ByVal control As Object = Nothing)
    RunEngineListsOnly
End Sub

Public Sub Ribbon_OnCleanMarkers(Optional ByVal control As Object = Nothing)
    RunCleanMarkers
End Sub

Public Sub Ribbon_OnInsertMarker(ByVal control As Object)
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
    End Select
End Sub

Public Function GetLastError() As String
    GetLastError = g_LastError
End Function