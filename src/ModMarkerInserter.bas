Attribute VB_Name = "ModMarkerInserter"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModMarkerInserter
' Purpose: Interactive standalone marker insertion with duplicate guard and zero bleed
' Architecture: Academic Document Format Engine (Win32 Unicode Safe)
' ==============================================================================

Public Sub InsertMarkerAtSelection(ByVal doc As Document, ByVal markerText As String)
    On Error GoTo ErrorHandler
    
    If doc Is Nothing Then
        Set doc = ActiveDocument
    End If
    If doc Is Nothing Then Exit Sub
    
    ' --- Duplicate Guard ---
    Dim existingPara As Long
    existingPara = ModMarkerLexer.FindMarkerParagraph(doc, LCase$(markerText))
    If existingPara > 0 Then
        ModUI.MsgBoxW ModConstants.Uni("Marker ") & markerText & _
                      ModConstants.Uni(" already exists at paragraph ") & existingPara & _
                      ModConstants.Uni(". (Th\u1ebb ") & markerText & _
                      ModConstants.Uni(" \u0111\u00e3 t\u1ed3n t\u1ea1i \u1edf \u0111o\u1ea1n ") & existingPara & _
                      ModConstants.Uni(", kh\u00f4ng th\u1ec3 ch\u00e8n th\u00eam.)"), _
                      vbExclamation, ModConstants.Uni("Duplicate Marker (Tr\u00f9ng Th\u1ebb)")
        Exit Sub
    End If
    
    ' --- Selection Range Resolution ---
    Dim selRng As Range
    Set selRng = doc.Application.Selection.Range
    If selRng Is Nothing Then Exit Sub
    
    ' Guard against insertion outside main document story (Header/Footer/TextBox/Footnote)
    If selRng.StoryType <> 1 Then ' wdMainTextStory = 1
        ModUI.MsgBoxW ModConstants.Uni("Cannot insert marker outside main document body. Place cursor in the main text. (Kh\u00f4ng th\u1ec3 ch\u00e8n th\u1ebb ngo\u00e0i n\u1ed9i dung ch\u00ednh.)"), _
                      vbExclamation, ModConstants.Uni("Invalid Position (V\u1ecb tr\u00ed kh\u00f4ng h\u1ee3p l\u1ec7)")
        Exit Sub
    End If
    
    ' Guard against insertion inside Table
    If selRng.Information(12) Then ' wdWithInTable = 12
        ModUI.MsgBoxW ModConstants.Uni("Cannot insert marker inside a table. Place cursor outside the table. (Kh\u00f4ng th\u1ec3 ch\u00e8n th\u1ebb trong b\u1ea3ng, h\u00e3y \u0111\u1eb7t con tr\u1ecf ra ngo\u00e0i.)"), _
                      vbExclamation, ModConstants.Uni("Invalid Position (V\u1ecb tr\u00ed kh\u00f4ng h\u1ee3p l\u1ec7)")
        Exit Sub
    End If
    
    ' Start Undo Record safely
    Dim objUndo As Object, weStartedUndo As Boolean
    weStartedUndo = False
    On Error Resume Next
    Set objUndo = Application.UndoRecord
    On Error GoTo ErrorHandler
    If Not objUndo Is Nothing Then
        If Not objUndo.IsRecordingCustomRecord Then
            objUndo.StartCustomRecord "Insert Marker " & markerText
            weStartedUndo = True
        End If
    End If
    
    ' Collapse multi-character selection to avoid deleting content
    If selRng.Start <> selRng.End Then selRng.Collapse 1 ' wdCollapseStart = 1
    
    ' --- Standalone Paragraph Insertion Logic ---
    Dim curPara As Paragraph
    Set curPara = selRng.Paragraphs(1)
    
    Dim pText As String
    pText = Replace(Replace(Replace(curPara.Range.Text, vbCr, ""), Chr(7), ""), Chr(11), "")
    pText = Trim$(pText)
    
    Dim pStart As Long, pEnd As Long
    pStart = curPara.Range.Start
    pEnd = curPara.Range.End
    
    If Len(pText) = 0 Then
        ' Case 1: Empty paragraph (preserve anchored Shapes or InlineShapes)
        If curPara.Range.InlineShapes.Count > 0 Or curPara.Range.ShapeRange.Count > 0 Then
            doc.Range(pStart, pStart).InsertBefore markerText & vbCr
        Else
            curPara.Range.Text = markerText & vbCr
        End If
    ElseIf selRng.Start = pStart Then
        ' Case 2: At start of paragraph
        selRng.InsertBefore markerText & vbCr
    ElseIf selRng.Start >= pEnd - 1 Then
        ' Case 3: At end of paragraph (before \r)
        selRng.InsertBefore vbCr & markerText
    Else
        ' Case 4: In middle of text
        selRng.InsertBefore vbCr & markerText & vbCr
    End If
    
    ' --- Locate Marker Paragraph & Apply Highlighting (Zero Bleed) ---
    Dim markerIdx As Long
    markerIdx = ModMarkerLexer.FindMarkerParagraph(doc, LCase$(markerText))
    If markerIdx > 0 Then
        Dim markerPara As Paragraph
        Set markerPara = doc.Paragraphs(markerIdx)
        If markerPara.Range.ListFormat.ListType <> 0 Then
            markerPara.Range.ListFormat.RemoveNumbers
        End If
        markerPara.Range.HighlightColorIndex = 7 ' wdYellow = 7
        ModFormatSingle.SafeSetOutlineLevel markerPara.Range, 10
        
        ' Collapse Selection to trailing range after marker and force no highlight
        Dim trailingRng As Range
        If markerPara.Range.End >= doc.Range.End Then
            doc.Range(doc.Range.End - 1, doc.Range.End - 1).InsertAfter vbCr
            Set trailingRng = doc.Paragraphs(markerIdx + 1).Range
            trailingRng.HighlightColorIndex = 0
            trailingRng.Collapse 1
            trailingRng.Select
        Else
            Set trailingRng = doc.Range(markerPara.Range.End, markerPara.Range.End)
            trailingRng.HighlightColorIndex = 0 ' wdNoHighlight = 0
            trailingRng.Select
        End If
        doc.Application.Selection.Range.HighlightColorIndex = 0
    End If

CleanExit:
    If weStartedUndo Then
        If Not objUndo Is Nothing Then
            If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
        End If
    End If
    Exit Sub

ErrorHandler:
    If weStartedUndo Then
        If Not objUndo Is Nothing Then
            If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
        End If
    End If
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
