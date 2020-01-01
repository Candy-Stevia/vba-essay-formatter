Attribute VB_Name = "ModFormatSingle"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModFormatSingle
' Purpose: Formats individual non-batchable elements (Headings, Captions, Cells)
' Architecture: Academic Document Format Engine (Modular Single Item Dispatcher)
' ==============================================================================

Public Sub SafeSetOutlineLevel(ByVal rng As Range, ByVal lvl As Long)
    On Error Resume Next
    rng.ParagraphFormat.OutlineLevel = lvl
    Err.Clear
End Sub

Private Function HasHeadingStyleInRange(ByVal rng As Range) As Boolean
    Dim stLvl As Long
    On Error Resume Next
    stLvl = rng.Style.ParagraphFormat.OutlineLevel
    If Err.Number <> 0 Then
        Err.Clear
        HasHeadingStyleInRange = True
        Exit Function
    End If
    HasHeadingStyleInRange = (stLvl < 10)
End Function

Public Sub EnsureOutlineLevelDetachHeading(ByVal doc As Document, ByVal rng As Range, ByVal targetLvl As Long)
    Dim curLvl As Long, p As Paragraph, pLvl As Long, needDetach As Boolean, pStyleLvl As Long
    On Error Resume Next
    curLvl = rng.ParagraphFormat.OutlineLevel
    needDetach = (Err.Number <> 0 Or curLvl <> targetLvl)
    Err.Clear
    If Not needDetach And targetLvl = 10 Then
        If HasHeadingStyleInRange(rng) Then needDetach = True
    End If
    If Not needDetach Then Exit Sub
    rng.ParagraphFormat.OutlineLevel = targetLvl
    needDetach = (Err.Number <> 0 Or rng.ParagraphFormat.OutlineLevel <> targetLvl)
    Err.Clear
    If Not needDetach And targetLvl = 10 Then
        If HasHeadingStyleInRange(rng) Then needDetach = True
    End If
    If needDetach Then
        For Each p In rng.Paragraphs
            pLvl = p.Range.ParagraphFormat.OutlineLevel
            If Err.Number <> 0 Then
                Err.Clear
                pLvl = -1
            End If
            pStyleLvl = 10
            If targetLvl = 10 Then
                pStyleLvl = p.Range.Style.ParagraphFormat.OutlineLevel
                If Err.Number <> 0 Then
                    Err.Clear
                    pStyleLvl = 1
                End If
            End If
            If pLvl <> targetLvl Or (targetLvl = 10 And pStyleLvl < 10) Then
                p.Range.Style = doc.Styles(-1)
                p.Range.ParagraphFormat.OutlineLevel = targetLvl
                Err.Clear
            End If
        Next p
    End If
    Err.Clear
End Sub

Private Sub ApplyFullFontName(ByVal r As Range, ByVal fontName As String)
    On Error Resume Next
    r.Font.Name = fontName
    r.Font.NameAscii = fontName
    r.Font.NameOther = fontName
    r.Font.NameFarEast = fontName
    r.Font.NameBi = fontName
    Err.Clear
End Sub

Public Sub ApplyFontNameSkippingMath(ByVal doc As Document, ByVal rng As Range, _
                                     ByVal fontName As String, ByVal hasOMaths As Boolean)
    If Not hasOMaths Then
        ApplyFullFontName rng, fontName
        Exit Sub
    End If
    Dim mCnt As Long
    mCnt = rng.OMaths.Count
    If mCnt = 0 Then
        ApplyFullFontName rng, fontName
        Exit Sub
    End If
    Dim segPos As Long, mIdx As Long, omRng As Range
    segPos = rng.Start
    For mIdx = 1 To mCnt
        Set omRng = rng.OMaths(mIdx).Range
        If omRng.Start > segPos Then
            ApplyFullFontName doc.Range(segPos, omRng.Start), fontName
        End If
        If omRng.End > segPos Then segPos = omRng.End
    Next mIdx
    If segPos < rng.End Then
        ApplyFullFontName doc.Range(segPos, rng.End), fontName
    End If
End Sub

Public Sub ApplyUpperSkippingMath(ByVal doc As Document, ByVal rng As Range, ByVal hasOMaths As Boolean)
    If Not hasOMaths Then
        rng.Case = 1
        Exit Sub
    End If
    If rng.OMaths.Count = 0 Then
        rng.Case = 1
        Exit Sub
    End If
    Dim segPos As Long, mIdx As Long, omRng As Range
    segPos = rng.Start
    For mIdx = 1 To rng.OMaths.Count
        Set omRng = rng.OMaths(mIdx).Range
        If omRng.Start > segPos Then doc.Range(segPos, omRng.Start).Case = 1
        If omRng.End > segPos Then segPos = omRng.End
    Next mIdx
    If segPos < rng.End Then doc.Range(segPos, rng.End).Case = 1
End Sub

Public Sub ApplyAllCapsSkippingMath(ByVal doc As Document, ByVal rng As Range, _
                                    ByVal enableCaps As Boolean, ByVal hasOMaths As Boolean)
    If Not hasOMaths Then
        rng.Font.AllCaps = enableCaps
        Exit Sub
    End If
    If rng.OMaths.Count = 0 Then
        rng.Font.AllCaps = enableCaps
        Exit Sub
    End If
    Dim segPos As Long, mIdx As Long, omRng As Range
    segPos = rng.Start
    For mIdx = 1 To rng.OMaths.Count
        Set omRng = rng.OMaths(mIdx).Range
        If omRng.Start > segPos Then doc.Range(segPos, omRng.Start).Font.AllCaps = enableCaps
        If omRng.End > segPos Then segPos = omRng.End
    Next mIdx
    If segPos < rng.End Then doc.Range(segPos, rng.End).Font.AllCaps = enableCaps
End Sub

Public Sub FormatSingleItem(ByVal doc As Document, ByVal rng As Range, _
                           ByVal tType As Long, ByVal bmName As String, _
                           ByVal fontName As String, ByVal fontSize As Single, _
                           ByVal indentPts As Single, ByVal hasOMaths As Boolean)
    Dim effFont As String, targetLvl As Long
    effFont = ModCustomProfile.ResolveEffectiveFont(tType, fontName)
    If tType >= ElementTypeLevel1 And tType <= ElementTypeLevel3 Then
        targetLvl = tType
    ElseIf tType = ElementTypeLevel4 Then
        If ModRibbonCallbacks.g_Config.TocDepth = 3 Then targetLvl = 10 Else targetLvl = 4
    Else
        targetLvl = 10
    End If
    If tType >= ElementTypeLevel1 And tType <= ElementTypeIgnoredManualList Then
        EnsureOutlineLevelDetachHeading doc, rng, targetLvl
    End If
    Select Case tType
        Case ElementTypeLevel1
            ModFormatCore.ClearYellowAndRedBleed rng
            ApplyFontNameSkippingMath doc, rng, effFont, hasOMaths
            rng.Font.Size = fontSize
            rng.Font.Bold = True
            rng.Font.Italic = False
            If Not ModCustomProfile.IsL1UpperDisabled() Then
                ApplyUpperSkippingMath doc, rng, hasOMaths
            End If
            rng.ParagraphFormat.Alignment = 1
            rng.ParagraphFormat.FirstLineIndent = 0#
            rng.ParagraphFormat.LeftIndent = 0#
            rng.ParagraphFormat.KeepWithNext = True
            If Len(bmName) > 0 Then
                ModBookmark.AddBookmark doc, bmName, rng
            End If
        Case ElementTypeLevel2
            ModFormatHeading.CapPrefixAndContent doc, rng, ""
            ModFormatCore.ClearYellowAndRedBleed rng
            ApplyFontNameSkippingMath doc, rng, effFont, hasOMaths
            rng.Font.Size = fontSize
            rng.Font.Bold = True
            rng.Bold = True
            rng.Font.Italic = False
            rng.ParagraphFormat.Alignment = 0
            rng.ParagraphFormat.FirstLineIndent = indentPts
            rng.ParagraphFormat.LeftIndent = ModConstants.CmToPoints(0.5)
            rng.ParagraphFormat.KeepWithNext = True
            If Len(bmName) > 0 Then
                ModBookmark.AddBookmark doc, bmName, rng
            End If
        Case ElementTypeLevel3
            ModFormatHeading.CapPrefixAndContent doc, rng, ""
            ModFormatCore.ClearYellowAndRedBleed rng
            ApplyFontNameSkippingMath doc, rng, effFont, hasOMaths
            rng.Font.Size = fontSize
            rng.Font.Bold = True
            rng.Bold = True
            rng.Font.Italic = True
            rng.ParagraphFormat.Alignment = 0
            rng.ParagraphFormat.FirstLineIndent = indentPts
            rng.ParagraphFormat.LeftIndent = ModConstants.CmToPoints(0.75)
            rng.ParagraphFormat.KeepWithNext = True
            If Len(bmName) > 0 Then
                ModBookmark.AddBookmark doc, bmName, rng
            End If
        Case ElementTypeLevel4
            ModFormatHeading.CapPrefixAndContent doc, rng, ""
            ModFormatCore.ClearYellowAndRedBleed rng
            ApplyFontNameSkippingMath doc, rng, effFont, hasOMaths
            rng.Font.Size = fontSize
            rng.Font.Bold = False
            rng.Font.Italic = True
            rng.ParagraphFormat.Alignment = 0
            rng.ParagraphFormat.FirstLineIndent = indentPts
            rng.ParagraphFormat.LeftIndent = ModConstants.CmToPoints(1#)
            rng.ParagraphFormat.KeepWithNext = True
            If Len(bmName) > 0 Then
                ModBookmark.AddBookmark doc, bmName, rng
            End If
        Case ElementTypeTableTitle
            ModFormatHeading.CapPrefixAndContent doc, rng, ". "
            ModFormatCore.ClearYellowAndRedBleed rng
            ApplyFontNameSkippingMath doc, rng, effFont, hasOMaths
            rng.Font.Size = fontSize
            rng.Font.Bold = True
            rng.Bold = True
            rng.Font.Italic = False
            rng.ParagraphFormat.Alignment = 0
            rng.ParagraphFormat.FirstLineIndent = indentPts
            rng.ParagraphFormat.LeftIndent = 0#
            rng.ParagraphFormat.KeepWithNext = True
            If Len(bmName) > 0 Then
                ModBookmark.AddBookmark doc, bmName, rng
            End If
        Case ElementTypeFigureTitle
            If InStr(rng.Text, ":") > 0 Then
                ModFormatHeading.CapPrefixAndContent doc, rng, ": "
            Else
                ModFormatHeading.CapPrefixAndContent doc, rng, ". "
            End If
            ModFormatCore.ClearYellowAndRedBleed rng
            ApplyFontNameSkippingMath doc, rng, effFont, hasOMaths
            rng.Font.Size = fontSize
            rng.Font.Bold = True
            rng.Bold = True
            rng.Font.Italic = False
            rng.ParagraphFormat.Alignment = 1
            rng.ParagraphFormat.FirstLineIndent = 0#
            rng.ParagraphFormat.LeftIndent = 0#
            rng.ParagraphFormat.KeepWithNext = False
            If Len(bmName) > 0 Then
                ModBookmark.AddBookmark doc, bmName, rng
            End If
        Case ElementTypeInlinePicture
            ModFormatCore.ClearYellowAndRedBleed rng
            rng.ParagraphFormat.Alignment = 1
            rng.ParagraphFormat.FirstLineIndent = 0#
            rng.ParagraphFormat.LeftIndent = 0#
            rng.ParagraphFormat.KeepWithNext = True
        Case ElementTypeSourceNote
            ModFormatHeading.FormatSourceNoteRange doc, rng, effFont, fontSize, hasOMaths
        Case ElementTypeTableCell
            ModFormatCore.ClearYellowAndRedBleed rng
            ApplyFontNameSkippingMath doc, rng, fontName, hasOMaths
            rng.Font.Size = fontSize
            If rng.Cells.Count > 0 Then
                rng.Cells(1).VerticalAlignment = 1
            End If
            rng.ParagraphFormat.Alignment = 1
            rng.ParagraphFormat.FirstLineIndent = 0#
            rng.ParagraphFormat.LeftIndent = 0#
        Case ElementTypeNormalParagraph
            ModFormatCore.ClearYellowAndRedBleed rng
            ApplyFontNameSkippingMath doc, rng, effFont, hasOMaths
            rng.Font.Size = fontSize
            rng.ParagraphFormat.Alignment = 3
            rng.ParagraphFormat.FirstLineIndent = indentPts
            rng.ParagraphFormat.LeftIndent = 0#
            rng.ParagraphFormat.KeepWithNext = False
            ModFormatBatch.g_SingleParaCount = ModFormatBatch.g_SingleParaCount + 1
        Case ElementTypeIgnoredManualList
            rng.HighlightColorIndex = 6
    End Select
    If tType >= ElementTypeLevel1 And tType <= ElementTypeNormalParagraph Then
        If ModCustomProfile.g_ForceVisualReset Then
            rng.Font.Color = -16777216
            rng.Font.Underline = 0
            rng.Font.AllCaps = False
            rng.ParagraphFormat.SpaceBefore = 0#
            rng.ParagraphFormat.SpaceAfter = 0#
            If tType = ElementTypeNormalParagraph Then
                rng.Font.Bold = False
                rng.Font.Italic = False
            End If
        End If
        ModCustomProfile.ApplyCustomStyleOverrides doc, rng, tType, hasOMaths, True
    End If
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
