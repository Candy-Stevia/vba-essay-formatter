Attribute VB_Name = "ModFormatCore"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModFormatCore
' Purpose: Body formatting engine with Contiguous Range Batching support
' Architecture: Academic Document Format Engine (Range Batching Candidate)
' ==============================================================================

Public Type FormatStats
    CountLevel1 As Long
    CountSubLevels As Long
    CountTables As Long
    CountFigures As Long
End Type

Private Sub ClearMixedRange(ByVal docRef As Document, ByVal sPos As Long, ByVal ePos As Long)
    If ePos <= sPos Then Exit Sub
    Dim subRng As Range, h As Long, midPos As Long
    Set subRng = docRef.Range(sPos, ePos)
    h = subRng.HighlightColorIndex
    If h = 7 Or h = 6 Then
        subRng.HighlightColorIndex = 0
    ElseIf h = 9999999 And (ePos - sPos > 1) Then
        midPos = sPos + ((ePos - sPos) \ 2)
        ClearMixedRange docRef, sPos, midPos
        ClearMixedRange docRef, midPos, ePos
    End If
End Sub

Public Sub ClearYellowAndRedBleed(ByVal rng As Range)
    On Error Resume Next
    Dim h As Long, docRef As Document
    Dim origStart As Long, origEnd As Long, prevEnd As Long, nextPos As Long, fh As Long
    Dim findErr As Long, hitCount As Long
    Dim fRng As Range
    h = rng.HighlightColorIndex
    If h = 7 Or h = 6 Then
        rng.HighlightColorIndex = 0
    ElseIf h = 9999999 Then
        Set docRef = rng.Parent
        origStart = rng.Start
        origEnd = rng.End
        If origEnd > origStart And Not (docRef Is Nothing) Then
            Err.Clear
            hitCount = 0
            Set fRng = docRef.Range(origStart, origEnd)
            ModMathHelpers.InitDomFind fRng.Find, ""
            fRng.Find.Format = True
            fRng.Find.Highlight = True
            prevEnd = origStart - 1
            Do While fRng.Find.Execute(FindText:="", Forward:=True, Wrap:=0, Format:=True)
                If fRng.Start >= origEnd Then Exit Do
                If fRng.End > origEnd Then fRng.End = origEnd
                hitCount = hitCount + 1
                fh = fRng.HighlightColorIndex
                If fh = 7 Or fh = 6 Then
                    fRng.HighlightColorIndex = 0
                ElseIf fh = 9999999 Then
                    ClearMixedRange docRef, fRng.Start, fRng.End
                End If
                nextPos = fRng.End
                If nextPos <= prevEnd Then nextPos = prevEnd + 1
                prevEnd = nextPos
                If nextPos >= origEnd Then Exit Do
                fRng.SetRange nextPos, origEnd
            Loop
            findErr = Err.Number
            fRng.Find.ClearFormatting
            If hitCount > 0 And findErr = 0 Then Exit Sub
            Err.Clear
            h = rng.HighlightColorIndex
            If h = 7 Or h = 6 Then
                rng.HighlightColorIndex = 0
            ElseIf h = 9999999 Then
                Dim subP As Paragraph, ph As Long
                For Each subP In rng.Paragraphs
                    ph = subP.Range.HighlightColorIndex
                    If ph = 7 Or ph = 6 Then
                        subP.Range.HighlightColorIndex = 0
                    ElseIf ph = 9999999 Then
                        ClearMixedRange docRef, subP.Range.Start, subP.Range.End
                    End If
                Next subP
            End If
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub ApplyLineSpacing(ByVal rng As Range, ByVal lsMode As String)
    Select Case lsMode
        Case "ls_115"
            rng.ParagraphFormat.LineSpacingRule = 5
            rng.ParagraphFormat.LineSpacing = 13.8
        Case "ls_130"
            rng.ParagraphFormat.LineSpacingRule = 5
            rng.ParagraphFormat.LineSpacing = 15.6
        Case "ls_150"
            rng.ParagraphFormat.LineSpacingRule = 1
            rng.ParagraphFormat.LineSpacing = 18#
    End Select
End Sub

Public Function ExecuteRecord1(ByVal doc As Document, ByRef snapshots() As ParagraphSnapshot, _
                              ByVal totalCount As Long, ByVal fontName As String, _
                              ByVal fontSize As Single, ByVal lineSpacingMode As String, _
                              ByVal indentPts As Single, ByRef outStats As FormatStats, _
                              Optional ByVal contentStartIdx As Long = 0, _
                              Optional ByRef outFrontSplits As Long = 0) As Boolean
    Dim objUndo As Object, weStartedUndo As Boolean
    weStartedUndo = False
    On Error Resume Next
    Set objUndo = Application.UndoRecord
    On Error GoTo ErrorHandler
    If Not objUndo Is Nothing Then
        If Not objUndo.IsRecordingCustomRecord Then
            objUndo.StartCustomRecord "Format Phase 1: Content & Bookmarks"
            weStartedUndo = True
        End If
    End If
    
    Dim hasOMaths As Boolean, hasLockedSdt As Boolean, docHasNestedTables As Boolean
    On Error Resume Next
    hasOMaths = (doc.OMaths.Count > 0)
    Err.Clear
    On Error GoTo ErrorHandler
    ModTableIndex.EnsureLockedSdtIndex doc
    hasLockedSdt = ModTableIndex.HasLockedSdts(doc)
    docHasNestedTables = ModTableIndex.HasNestedTables(doc)
    
    outStats.CountLevel1 = 0
    outStats.CountSubLevels = 0
    outStats.CountTables = 0
    outStats.CountFigures = 0
    outFrontSplits = 0
    If totalCount <= 0 Then GoTo CleanExit
    
    Dim bmNames() As String
    ReDim bmNames(1 To totalCount)
    Dim i As Long, tType As Long
    
    For i = 1 To totalCount
        tType = snapshots(i).ElementType
        Select Case tType
            Case ElementTypeLevel1
                outStats.CountLevel1 = outStats.CountLevel1 + 1
                bmNames(i) = ModBookmark.FormatBookmarkId(BM_PREFIX_TOC, outStats.CountLevel1 + outStats.CountSubLevels)
            Case ElementTypeLevel2, ElementTypeLevel3, ElementTypeLevel4
                outStats.CountSubLevels = outStats.CountSubLevels + 1
                bmNames(i) = ModBookmark.FormatBookmarkId(BM_PREFIX_TOC, outStats.CountLevel1 + outStats.CountSubLevels)
            Case ElementTypeTableTitle
                outStats.CountTables = outStats.CountTables + 1
                bmNames(i) = ModBookmark.FormatBookmarkId(BM_PREFIX_TBL, outStats.CountTables)
            Case ElementTypeFigureTitle
                outStats.CountFigures = outStats.CountFigures + 1
                bmNames(i) = ModBookmark.FormatBookmarkId(BM_PREFIX_FIG, outStats.CountFigures)
        End Select
    Next i
    
    Dim rng As Range, splitTailRng As Range, pIdx As Long, j As Long
    Dim bStart As Long, bEnd As Long, bCount As Long, baseNest As Long, prevNest As Long
    Dim hasMathInThisPara As Boolean, itemHasOMath As Boolean, tailHasOMath As Boolean
    Dim docEnd As Long, useMathIdx As Boolean
    docEnd = doc.Range.End
    useMathIdx = ModTableIndex.IsOMathIndexValidForDoc(doc)
    
    i = totalCount
    Do While i >= 1
        pIdx = snapshots(i).ParaIndex
        If snapshots(i).RangeStart < docEnd Then
            ModProgress.ProgressUpdate (totalCount - i + 1), totalCount
            If ModProgress.g_CancelRequested Then
                If weStartedUndo Then
                    If Not objUndo Is Nothing Then
                        If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
                    End If
                End If
                ExecuteRecord1 = False
                Exit Function
            End If
            If hasLockedSdt Then
                If ModTableIndex.IsSpanInLockedSdt(snapshots(i).RangeStart, snapshots(i).RangeEnd) Then
                    i = i - 1
                    GoTo NextLoopIter
                End If
            End If
            tType = snapshots(i).ElementType
            
            If tType = ElementTypeNormalParagraph And ModFormatBatch.IsBatchableNormal(snapshots(i)) Then
                hasMathInThisPara = False
                If hasOMaths Then
                    If useMathIdx Then
                        hasMathInThisPara = ModTableIndex.HasOMathInSpan(snapshots(i).RangeStart, snapshots(i).RangeEnd)
                    Else
                        hasMathInThisPara = True
                    End If
                End If
                
                If Not hasMathInThisPara Then
                    j = i
                    Do While j > 1
                        If ModFormatBatch.IsBatchableNormal(snapshots(j - 1)) Then
                            If snapshots(j - 1).RangeEnd = snapshots(j).RangeStart Then
                                If hasOMaths Then
                                    If (Not useMathIdx) Or ModTableIndex.HasOMathInSpan(snapshots(j - 1).RangeStart, snapshots(j - 1).RangeEnd) Then Exit Do
                                End If
                                If hasLockedSdt Then
                                    If ModTableIndex.IsSpanInLockedSdt(snapshots(j - 1).RangeStart, snapshots(j - 1).RangeEnd) Then Exit Do
                                End If
                                j = j - 1
                            Else
                                Exit Do
                            End If
                        Else
                            Exit Do
                        End If
                    Loop
                    
                    bCount = i - j + 1
                    bStart = snapshots(j).RangeStart
                    bEnd = snapshots(i).RangeEnd
                    ModFormatBatch.FormatNormalBatch doc, bStart, bEnd, fontName, fontSize, indentPts, lineSpacingMode, bCount
                    i = j - 1
                    GoTo NextLoopIter
                End If
            ElseIf tType = ElementTypeTableCell And Not snapshots(i).IsExcluded Then
                hasMathInThisPara = False
                If hasOMaths Then
                    If useMathIdx Then
                        hasMathInThisPara = ModTableIndex.HasOMathInSpan(snapshots(i).RangeStart, snapshots(i).RangeEnd)
                    Else
                        hasMathInThisPara = True
                    End If
                End If
                If Not hasMathInThisPara Then
                    If docHasNestedTables Then
                        baseNest = doc.Range(snapshots(i).RangeStart, snapshots(i).RangeEnd).Information(16)
                    End If
                    j = i
                    Do While j > 1
                        If snapshots(j - 1).ElementType = ElementTypeTableCell And Not snapshots(j - 1).IsExcluded And _
                           snapshots(j - 1).RangeEnd = snapshots(j).RangeStart Then
                            If hasOMaths Then
                                If (Not useMathIdx) Or ModTableIndex.HasOMathInSpan(snapshots(j - 1).RangeStart, snapshots(j - 1).RangeEnd) Then Exit Do
                            End If
                            If hasLockedSdt Then
                                If ModTableIndex.IsSpanInLockedSdt(snapshots(j - 1).RangeStart, snapshots(j - 1).RangeEnd) Then Exit Do
                            End If
                            If docHasNestedTables Then
                                prevNest = doc.Range(snapshots(j - 1).RangeStart, snapshots(j - 1).RangeEnd).Information(16)
                                If prevNest <> baseNest Then Exit Do
                            End If
                            j = j - 1
                        Else
                            Exit Do
                        End If
                    Loop
                    ModFormatBatch.FormatTableCellBatch doc, snapshots, j, i, fontName, fontSize
                    i = j - 1
                    GoTo NextLoopIter
                End If
            End If
            
            Set rng = doc.Range(snapshots(i).RangeStart, snapshots(i).RangeEnd)
            
            If snapshots(i).HasSoftBreak And (tType >= ElementTypeLevel1 And tType <= ElementTypeSourceNote) Then
                If Not ModLocale.IsPrefixOnlyHeading(snapshots(i).FirstLineText) Then
                    ModMathHelpers.InitDomFind rng.Find, Chr(11)
                    rng.Find.Replacement.ClearFormatting
                    rng.Find.Replacement.Text = vbCr
                    If rng.Find.Execute(Replace:=1) Then
                        docEnd = doc.Range.End
                        If contentStartIdx > 0 And snapshots(i).ParaIndex < contentStartIdx Then
                            outFrontSplits = outFrontSplits + 1
                        End If
                        Set rng = rng.Paragraphs(1).Range
                        snapshots(i).RangeEnd = rng.End
                        snapshots(i).RawText = rng.Text
                        snapshots(i).CleanText = snapshots(i).FirstLineText
                        If Not rng.Paragraphs(1).Next Is Nothing Then
                            Set splitTailRng = rng.Paragraphs(1).Next.Range
                            splitTailRng.Font.Bold = False
                            splitTailRng.Font.Italic = False
                            tailHasOMath = False
                            If hasOMaths Then tailHasOMath = (splitTailRng.OMaths.Count > 0)
                            ModFormatSingle.FormatSingleItem doc, splitTailRng, ElementTypeNormalParagraph, "", fontName, fontSize, indentPts, tailHasOMath
                            If lineSpacingMode <> "ls_keep" Then
                                ApplyLineSpacing splitTailRng, lineSpacingMode
                            End If
                        End If
                    Else
                        Set rng = doc.Range(snapshots(i).RangeStart, snapshots(i).RangeEnd)
                    End If
                End If
            End If
            
            itemHasOMath = False
            If hasOMaths Then
                If useMathIdx Then
                    itemHasOMath = ModTableIndex.HasOMathInSpan(rng.Start, rng.End)
                Else
                    itemHasOMath = True
                End If
            End If
            ModFormatSingle.FormatSingleItem doc, rng, tType, bmNames(i), fontName, fontSize, indentPts, itemHasOMath
            
            If lineSpacingMode <> "ls_keep" And (Not snapshots(i).IsExcluded) And _
               (tType >= ElementTypeLevel1 And tType <= ElementTypeNormalParagraph) Then
                ApplyLineSpacing rng, lineSpacingMode
            End If
            
            i = i - 1
        Else
            i = i - 1
        End If
NextLoopIter:
    Loop
    
    ModBookmark.CleanExcessBookmarks doc, outStats.CountLevel1 + outStats.CountSubLevels, outStats.CountTables, outStats.CountFigures
    
CleanExit:
    If weStartedUndo Then
        If Not objUndo Is Nothing Then
            If objUndo.IsRecordingCustomRecord Then
                objUndo.EndCustomRecord
            End If
        End If
    End If
    ExecuteRecord1 = True
    Exit Function
    
ErrorHandler:
    Application.StatusBar = ""
    ModRibbonCallbacks.g_LastError = "FormatCore Error " & Err.Number & ": " & Err.Description
    If weStartedUndo Then
        If Not objUndo Is Nothing Then
            If objUndo.IsRecordingCustomRecord Then
                objUndo.EndCustomRecord
            End If
        End If
    End If
    ExecuteRecord1 = False
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
