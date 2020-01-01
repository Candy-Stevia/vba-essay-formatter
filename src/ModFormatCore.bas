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

Public Sub ClearYellowAndRedBleed(ByVal rng As Range)
    On Error Resume Next
    Dim h As Long
    h = rng.HighlightColorIndex
    If h = 7 Or h = 6 Then
        rng.HighlightColorIndex = 0
    End If
    If Err.Number <> 0 Then
        Err.Clear
    End If
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
                              ByVal indentPts As Single, ByRef outStats As FormatStats) As Boolean
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
    
    Dim hasOMaths As Boolean
    On Error Resume Next
    hasOMaths = (doc.OMaths.Count > 0)
    On Error GoTo ErrorHandler
    
    Dim bmNames() As String
    ReDim bmNames(1 To totalCount)
    Dim i As Long, tType As Long
    outStats.CountLevel1 = 0
    outStats.CountSubLevels = 0
    outStats.CountTables = 0
    outStats.CountFigures = 0
    
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
    
    Dim rng As Range, pIdx As Long, j As Long
    Dim bStart As Long, bEnd As Long, bCount As Long
    Dim hasMathInThisPara As Boolean
    
    i = totalCount
    Do While i >= 1
        pIdx = snapshots(i).ParaIndex
        If snapshots(i).RangeStart < doc.Range.End Then
            ModProgress.ProgressUpdate (totalCount - i + 1), totalCount
            tType = snapshots(i).ElementType
            
            If tType = ElementTypeNormalParagraph And ModFormatBatch.IsBatchableNormal(snapshots(i)) Then
                hasMathInThisPara = False
                If hasOMaths Then
                    Set rng = doc.Range(snapshots(i).RangeStart, snapshots(i).RangeEnd)
                    If rng.OMaths.Count > 0 Then
                        hasMathInThisPara = True
                    End If
                End If
                
                If Not hasMathInThisPara Then
                    j = i
                    Do While j > 1
                        If ModFormatBatch.IsBatchableNormal(snapshots(j - 1)) Then
                            If snapshots(j - 1).RangeEnd = snapshots(j).RangeStart Then
                                If hasOMaths Then
                                    Set rng = doc.Range(snapshots(j - 1).RangeStart, snapshots(j - 1).RangeEnd)
                                    If rng.OMaths.Count > 0 Then
                                        Exit Do
                                    End If
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
            End If
            
            Set rng = doc.Range(snapshots(i).RangeStart, snapshots(i).RangeEnd)
            
            If snapshots(i).HasSoftBreak And (tType >= ElementTypeLevel1 And tType <= ElementTypeSourceNote) Then
                If Not ModLocale.IsPrefixOnlyHeading(snapshots(i).FirstLineText) Then
                    rng.Find.Execute FindText:=Chr(11), ReplaceWith:=vbCr, Replace:=1
                    Set rng = rng.Paragraphs(1).Range
                End If
            End If
            
            ModFormatSingle.FormatSingleItem doc, rng, tType, bmNames(i), fontName, fontSize, indentPts, hasOMaths
            
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
    Application.StatusBar = False
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
