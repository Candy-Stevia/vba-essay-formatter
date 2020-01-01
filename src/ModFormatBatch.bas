Attribute VB_Name = "ModFormatBatch"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModFormatBatch
' Purpose: Contiguous Range Batching Engine for Normal Paragraphs
' Architecture: Formats contiguous runs of homogeneous normal paragraphs in a
'               single COM Range call, eliminating per-paragraph COM overhead.
' Invariants: Never batch across headings, tables, Word Math, lists, or excluded.
' Highlight Parity: Handles mixed highlight (wdUndefined = 9999999) per paragraph.
' ==============================================================================

Public Type BatchListEntry
    TitleText As String
    TargetBm As String
    LeftIndentPts As Single
    IsBold As Boolean
End Type

Public g_BatchCount As Long
Public g_BatchedParaCount As Long
Public g_SingleParaCount As Long

Public Sub ResetBatchCounters()
    g_BatchCount = 0
    g_BatchedParaCount = 0
    g_SingleParaCount = 0
End Sub

Public Function GetBatchTelemetry() As String
    Dim s As String
    s = "{"
    s = s & """batch_count"":" & g_BatchCount & ","
    s = s & """batched_para_count"":" & g_BatchedParaCount & ","
    s = s & """single_para_count"":" & g_SingleParaCount
    s = s & "}"
    GetBatchTelemetry = s
End Function

Public Function IsBatchableNormal(ByRef snap As ParagraphSnapshot) As Boolean
    IsBatchableNormal = False
    If snap.ElementType <> ElementTypeNormalParagraph Then
        Exit Function
    End If
    If snap.IsExcluded Then
        Exit Function
    End If
    If snap.HasSoftBreak Then
        Exit Function
    End If
    IsBatchableNormal = True
End Function

Public Sub FormatNormalBatch(ByVal doc As Document, ByVal rStart As Long, _
                            ByVal rEnd As Long, ByVal fontName As String, _
                            ByVal fontSize As Single, ByVal indentPts As Single, _
                            ByVal lineSpacingMode As String, ByVal paraCount As Long)
    Dim rng As Range, effFont As String
    Set rng = doc.Range(rStart, rEnd)
    ModFormatSingle.EnsureOutlineLevelDetachHeading doc, rng, 10
    ModFormatCore.ClearYellowAndRedBleed rng
    effFont = ModCustomProfile.ResolveEffectiveFont(ElementTypeNormalParagraph, fontName)
    ModFormatSingle.ApplyFontNameSkippingMath doc, rng, effFont, False
    rng.Font.Size = fontSize
    rng.ParagraphFormat.Alignment = 3
    rng.ParagraphFormat.FirstLineIndent = indentPts
    rng.ParagraphFormat.LeftIndent = 0#
    rng.ParagraphFormat.KeepWithNext = False
    If lineSpacingMode <> "ls_keep" Then
        ModFormatCore.ApplyLineSpacing rng, lineSpacingMode
    End If
    If ModCustomProfile.g_ForceVisualReset Then
        rng.Font.Color = -16777216
        rng.Font.Underline = 0
        rng.Font.AllCaps = False
        rng.Font.Bold = False
        rng.Font.Italic = False
        rng.ParagraphFormat.SpaceBefore = 0#
        rng.ParagraphFormat.SpaceAfter = 0#
    End If
    ModCustomProfile.ApplyCustomStyleOverrides doc, rng, ElementTypeNormalParagraph, False, True
    If paraCount > 1 Then
        g_BatchCount = g_BatchCount + 1
        g_BatchedParaCount = g_BatchedParaCount + paraCount
    Else
        g_SingleParaCount = g_SingleParaCount + 1
    End If
End Sub

Public Sub FormatTableCellBatch(ByVal doc As Document, ByRef snapshots() As ParagraphSnapshot, _
                                ByVal idxStart As Long, ByVal idxEnd As Long, _
                                ByVal fontName As String, ByVal fontSize As Single)
    Dim rng As Range, k As Long, cellRng As Range
    Set rng = doc.Range(snapshots(idxStart).RangeStart, snapshots(idxEnd).RangeEnd)
    ModFormatSingle.EnsureOutlineLevelDetachHeading doc, rng, 10
    ModFormatCore.ClearYellowAndRedBleed rng
    ModFormatSingle.ApplyFontNameSkippingMath doc, rng, fontName, False
    rng.Font.Size = fontSize
    rng.ParagraphFormat.Alignment = 1
    rng.ParagraphFormat.FirstLineIndent = 0#
    rng.ParagraphFormat.LeftIndent = 0#
    On Error Resume Next
    rng.Cells.VerticalAlignment = 1
    If Err.Number <> 0 Then
        Err.Clear
        For k = idxStart To idxEnd
            Set cellRng = doc.Range(snapshots(k).RangeStart, snapshots(k).RangeEnd)
            If cellRng.Cells.Count > 0 Then cellRng.Cells(1).VerticalAlignment = 1
        Next k
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Sub AppendBatchEntry(ByRef entries() As BatchListEntry, ByRef eCnt As Long, ByRef eCap As Long, _
                            ByVal titleTxt As String, ByVal targetBm As String, _
                            ByVal leftIndentPts As Single, ByVal isBold As Boolean)
    eCnt = eCnt + 1
    If eCnt > eCap Then
        If eCap < 16 Then eCap = 32 Else eCap = eCap * 2
        ReDim Preserve entries(1 To eCap)
    End If
    entries(eCnt).TitleText = titleTxt
    entries(eCnt).TargetBm = targetBm
    entries(eCnt).LeftIndentPts = leftIndentPts
    entries(eCnt).IsBold = isBold
End Sub

Public Sub AppendQuickSnap(ByRef qSnaps() As ParagraphSnapshot, ByRef qCount As Long, ByRef qCap As Long, _
                           ByVal pIdx As Long, ByVal rStart As Long, ByVal rEnd As Long, _
                           ByVal elemType As Long, ByVal rawTxt As String, ByVal cleanTxt As String)
    qCount = qCount + 1
    If qCount > qCap Then
        If qCap < 32 Then qCap = 64 Else qCap = qCap * 2
        ReDim Preserve qSnaps(1 To qCap)
    End If
    qSnaps(qCount).ParaIndex = pIdx
    qSnaps(qCount).RangeStart = rStart
    qSnaps(qCount).RangeEnd = rEnd
    qSnaps(qCount).ElementType = elemType
    qSnaps(qCount).RawText = rawTxt
    qSnaps(qCount).CleanText = cleanTxt
End Sub

Public Function InsertListEntriesBatch(ByVal doc As Document, ByVal atRange As Range, _
                                       ByRef entries() As BatchListEntry, ByVal entryCount As Long, _
                                       ByVal fontName As String, ByVal fontSize As Single, _
                                       ByVal tabPos As Single, ByVal defaultLeftIndent As Single, _
                                       ByVal defaultBold As Boolean) As Range
    If entryCount <= 0 Then
        Set InsertListEntriesBatch = atRange
        Exit Function
    End If
    atRange.HighlightColorIndex = 0
    Dim linesArr() As String, lineStarts() As Long, lineEnds() As Long
    ReDim linesArr(1 To entryCount)
    ReDim lineStarts(1 To entryCount)
    ReDim lineEnds(1 To entryCount)
    Dim i As Long, actualTitle As String
    For i = 1 To entryCount
        actualTitle = ModListBuilder.GetFreshTitleText(doc, entries(i).TargetBm, entries(i).TitleText)
        linesArr(i) = actualTitle & vbTab & DUMMY_PAGE_STRING
    Next i
    Dim fullText As String, startPos As Long, pos As Long
    fullText = Join(linesArr, vbCr) & vbCr
    startPos = atRange.End
    atRange.InsertAfter fullText
    Dim blockEnd As Long, blockRng As Range, endRng As Range, exactLen As Boolean
    blockEnd = atRange.End
    Set blockRng = doc.Range(startPos, blockEnd - 1)
    Set endRng = doc.Range(blockEnd, blockEnd)
    ModPageNumberSync.ApplyListLineFormat blockRng, fontName, fontSize, tabPos, defaultLeftIndent, defaultBold
    exactLen = ((blockEnd - startPos) = Len(fullText))
    If exactLen Then
        pos = startPos
        For i = 1 To entryCount
            lineStarts(i) = pos
            lineEnds(i) = pos + Len(linesArr(i))
            pos = lineEnds(i) + 1
        Next i
    Else
        Dim bParaCnt As Long, pR As Range
        bParaCnt = blockRng.Paragraphs.Count
        For i = 1 To entryCount
            If i <= bParaCnt Then
                Set pR = blockRng.Paragraphs(i).Range
                lineStarts(i) = pR.Start
                lineEnds(i) = pR.End - 1
                If lineEnds(i) < lineStarts(i) Then lineEnds(i) = lineStarts(i)
            End If
        Next i
    End If
    Dim lineRng As Range, hlAdded As Boolean
    hlAdded = False
    For i = entryCount To 1 Step -1
        If lineEnds(i) > lineStarts(i) Then
            Set lineRng = doc.Range(lineStarts(i), lineEnds(i))
            If Abs(entries(i).LeftIndentPts - defaultLeftIndent) > 0.01 Then
                lineRng.ParagraphFormat.LeftIndent = entries(i).LeftIndentPts
            End If
            If entries(i).IsBold <> defaultBold Then
                lineRng.Font.Bold = entries(i).IsBold
            End If
            If Len(entries(i).TargetBm) > 0 Then
                If doc.Bookmarks.Exists(entries(i).TargetBm) Then
                    doc.Hyperlinks.Add Anchor:=lineRng, Address:="", SubAddress:=entries(i).TargetBm
                    hlAdded = True
                End If
            End If
        End If
    Next i
    endRng.Collapse 0
    If hlAdded And endRng.Start - 1 > startPos Then
        Dim allLinksRng As Range
        Set allLinksRng = doc.Range(startPos, endRng.Start - 1)
        allLinksRng.Font.Underline = 0
        allLinksRng.Font.ColorIndex = 0
    End If
    endRng.HighlightColorIndex = 0
    Set InsertListEntriesBatch = endRng
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
