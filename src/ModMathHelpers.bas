Attribute VB_Name = "ModMathHelpers"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ModMathHelpers: Math ID cache, Variables, Shading, Bookmark lookup & DOM fallback
Private g_MathSeqCached As Long
Private g_HasSeqCache As Boolean
Private g_SeqDocPtr As LongPtr

Private Sub UpdateMaxSeq(ByVal itemName As String, ByRef curSeq As Long)
    Dim suffix As String, bmNum As Long
    If Left$(itemName, 8) = "bm_math_" Then
        suffix = Mid$(itemName, InStrRev(itemName, "_") + 1)
        If Len(suffix) > 0 And Len(suffix) <= 9 And IsNumeric(suffix) Then
            bmNum = CLng(suffix)
            If bmNum > curSeq Then curSeq = bmNum
        End If
    End If
End Sub

Public Sub InitMathSeqCache(ByVal doc As Document)
    Dim curSeq As Long, bm As Bookmark, v As Variable, vStr As String
    ResetMathSeqCache
    If doc Is Nothing Then Exit Sub
    g_SeqDocPtr = ObjPtr(doc)
    On Error Resume Next
    vStr = doc.Variables("mathcfg_seq").Value
    If Err.Number = 0 And Len(vStr) > 0 And Len(vStr) <= 9 Then curSeq = CLng(vStr)
    Err.Clear
    For Each bm In doc.Bookmarks
        UpdateMaxSeq bm.Name, curSeq
    Next bm
    For Each v In doc.Variables
        UpdateMaxSeq v.Name, curSeq
    Next v
    On Error GoTo 0
    g_MathSeqCached = curSeq
    g_HasSeqCache = True
End Sub

Public Sub ResetMathSeqCache()
    g_MathSeqCached = 0
    g_HasSeqCache = False
    g_SeqDocPtr = 0
End Sub
Public Sub FlushMathSeqCache(ByVal doc As Document)
    If Not doc Is Nothing Then
        If g_HasSeqCache And g_SeqDocPtr = ObjPtr(doc) And g_MathSeqCached > 0 Then
            SetDocVariable doc, "mathcfg_seq", CStr(g_MathSeqCached)
        End If
    End If
End Sub
Public Function NextMathSeq(ByVal doc As Document) As Long
    If Not g_HasSeqCache Or g_SeqDocPtr <> ObjPtr(doc) Then
        InitMathSeqCache doc
        g_HasSeqCache = False
        g_SeqDocPtr = 0
        g_MathSeqCached = g_MathSeqCached + 1
        SetDocVariable doc, "mathcfg_seq", CStr(g_MathSeqCached)
        NextMathSeq = g_MathSeqCached
        Exit Function
    End If
    g_MathSeqCached = g_MathSeqCached + 1
    NextMathSeq = g_MathSeqCached
End Function
Public Sub SetDocVariable(ByVal doc As Document, ByVal varName As String, ByVal varVal As String)
    If Len(varVal) > 65000 Then
        Exit Sub
    End If
    On Error Resume Next
    doc.Variables(varName).Value = varVal
    If Err.Number <> 0 Then
        Err.Clear
        doc.Variables.Add varName, varVal
    End If
    Err.Clear
    On Error GoTo 0
End Sub
Public Function GetDocVariable(ByVal doc As Document, ByVal varName As String) As String
    On Error Resume Next
    GetDocVariable = doc.Variables(varName).Value
    On Error GoTo 0
End Function
Public Sub ApplyPastelShading(ByVal rng As Range, ByVal rgbColor As Long)
    On Error Resume Next
    rng.Shading.Texture = 0
    rng.Shading.BackgroundPatternColor = rgbColor
    On Error GoTo 0
End Sub
Public Sub NotifyUser(ByVal msg As String)
    If Application.DisplayAlerts <> 0 Then
        ModUI.MsgBoxW msg, vbExclamation, ModConstants.Uni("Notice (Th\u00f4ng b\u00e1o)")
    End If
End Sub

Public Function FindEnclosingMathBookmark(ByVal doc As Document, ByVal selStart As Long, ByVal selEnd As Long, ByRef outBmName As String, ByRef outBmKind As Long) As Boolean
    Dim bm As Bookmark, bStart As Long, bEnd As Long, ovLen As Long
    Dim bestOverlap As Long, bestDist As Long, kind As Long, bestKind As Long, bestBm As String, isOverlapping As Boolean
    outBmName = ""
    outBmKind = ModMathScanner.BM_KIND_NONE
    bestOverlap = -1
    bestDist = 999999
    For Each bm In doc.Bookmarks
        If Left$(bm.Name, 8) = "bm_math_" Then
            bStart = bm.Range.Start
            bEnd = bm.Range.End
            If selEnd > selStart Then isOverlapping = (selStart < bEnd And selEnd > bStart) Else isOverlapping = (selStart >= bStart And selEnd <= bEnd)
            If isOverlapping Then
                kind = ModMathScanner.BM_KIND_OMATH
                If Left$(bm.Name, 13) = "bm_math_pink_" Then
                    kind = ModMathScanner.BM_KIND_PINK
                ElseIf Left$(bm.Name, 12) = "bm_math_src_" Then
                    kind = ModMathScanner.BM_KIND_SRC
                End If
                If selStart >= bStart And selEnd <= bEnd Then
                    outBmName = bm.Name
                    outBmKind = kind
                    FindEnclosingMathBookmark = True
                    Exit Function
                End If
                ovLen = IIf(bEnd < selEnd, bEnd, selEnd) - IIf(bStart > selStart, bStart, selStart)
                If ovLen > bestOverlap Or (ovLen = bestOverlap And (bEnd - bStart) < bestDist) Then
                    bestOverlap = ovLen
                    bestDist = bEnd - bStart
                    bestBm = bm.Name
                    bestKind = kind
                End If
            End If
        End If
    Next bm
    If Len(bestBm) > 0 Then
        outBmName = bestBm
        outBmKind = bestKind
        FindEnclosingMathBookmark = True
    End If
End Function
Public Sub AppendPairUnique(ByRef starts() As Long, ByRef ends() As Long, ByRef pairCount As Long, ByVal firstPos As Long, ByVal lastPos As Long)
    Dim i As Long
    For i = 1 To pairCount
        If firstPos < ends(i) And lastPos > starts(i) Then
            Exit Sub
        End If
    Next i
    pairCount = pairCount + 1
    If pairCount = 1 Then
        ReDim starts(1 To 1)
        ReDim ends(1 To 1)
    Else
        ReDim Preserve starts(1 To pairCount)
        ReDim Preserve ends(1 To pairCount)
    End If
    starts(pairCount) = firstPos
    ends(pairCount) = lastPos
End Sub
Private Function IsDomPosEscaped(ByVal doc As Document, ByVal pStart As Long, ByVal pos As Long) As Boolean
    Dim slashCount As Long
    Do While pos > pStart
        If doc.Range(pos - 1, pos).Text <> "\" Then
            Exit Do
        End If
        slashCount = slashCount + 1
        pos = pos - 1
    Loop
    IsDomPosEscaped = (slashCount Mod 2 = 1)
End Function
Public Sub MarkSpanPink(ByVal doc As Document, ByVal rng As Range, ByVal seq As Long, ByVal varText As String)
    Dim idStr As String
    If Right$(rng.Text, 1) = Chr(7) Then
        ModMathInline.TrimTrailingDomMarks rng
    End If
    idStr = Format$(seq, "00000")
    ApplyPastelShading rng, ModMathScanner.COLOR_PASTEL_PINK
    doc.Bookmarks.Add "bm_math_pink_" & idStr, rng
    SetDocVariable doc, "bm_math_" & idStr, varText
End Sub
Public Function InferIsBlock(ByVal rawLatex As String, ByVal doc As Document, ByVal idStr As String) As Boolean
    If Left$(rawLatex, 2) = "$$" Or Left$(rawLatex, 2) = "\[" Or (Left$(rawLatex, 6) = "\begin" And Left$(rawLatex, 2) <> "\(") Then
        InferIsBlock = True
    ElseIf Left$(rawLatex, 1) = "$" Or Left$(rawLatex, 2) = "\(" Then
        InferIsBlock = False
    ElseIf Len(idStr) > 0 Then
        InferIsBlock = (GetDocVariable(doc, "bm_math_" & idStr & "_isblock") = "1")
    End If
End Function
Private Function CheckOpeningDollarDom(ByVal doc As Document, ByVal posEnd As Long, ByVal limit As Long, ByVal cAfter As String) As Boolean
    Dim aheadText As String

    If cAfter = " " Or cAfter = vbTab Or cAfter = vbCr Or cAfter = Chr(11) Or cAfter = Chr(7) Then
        Exit Function
    End If
    If cAfter >= "0" And cAfter <= "9" Then
        aheadText = doc.Range(posEnd, limit).Text
        CheckOpeningDollarDom = ModMathInline.IsMathOpeningAheadText(aheadText, 1, Len(aheadText))
    Else
        CheckOpeningDollarDom = True
    End If
End Function

Public Sub InitDomFind(ByVal f As Find, ByVal findTxt As String)
    f.ClearFormatting
    f.Text = findTxt
    f.Forward = True
    f.Wrap = 0
    f.Format = False
    f.MatchWildcards = False
    f.MatchWholeWord = False
    f.MatchAllWordForms = False
    f.MatchSoundsLike = False
    f.IgnorePunct = False
End Sub

Public Sub FindDollarPairsDom(ByVal doc As Document, ByVal p As Paragraph, ByRef outStarts() As Long, ByRef outEnds() As Long, ByRef outPairCount As Long)
    Dim fRng As Range
    Dim candRng As Range
    Dim limit As Long
    Dim pStart As Long
    Dim openStart As Long
    Dim isEscaped As Boolean
    Dim isDbl As Boolean
    Dim hasOpen As Boolean
    Dim isValidClose As Boolean
    Dim cBefore As String
    Dim cAfter As String

    outPairCount = 0
    pStart = p.Range.Start
    limit = p.Range.End
    If Right$(p.Range.Text, 1) = Chr(7) Or Right$(p.Range.Text, 1) = vbCr Then
        limit = limit - 1
    End If
    If limit <= pStart Then
        Exit Sub
    End If
    Set fRng = doc.Range(pStart, limit)
    InitDomFind fRng.Find, "$"
    Dim prevEnd As Long
    prevEnd = pStart
    Do While fRng.Find.Execute
        If fRng.Start >= limit Or fRng.Start < prevEnd Or fRng.End <= prevEnd Then
            Exit Do
        End If
        prevEnd = fRng.End
        If fRng.OMaths.Count = 0 And fRng.Fields.Count = 0 Then
            isEscaped = IsDomPosEscaped(doc, pStart, fRng.Start)
            isDbl = False
            If Not isEscaped Then
                If fRng.End < limit Then
                    If doc.Range(fRng.Start, fRng.Start + 2).Text = "$$" Then
                        isDbl = True
                    End If
                End If
                If fRng.Start > pStart Then
                    If doc.Range(fRng.Start - 1, fRng.Start + 1).Text = "$$" Then
                        If Not IsDomPosEscaped(doc, pStart, fRng.Start - 1) Then
                            isDbl = True
                        End If
                    End If
                End If
            End If
            If isDbl Then
                hasOpen = False
                openStart = 0
            ElseIf Not isEscaped Then
                cBefore = ""
                If fRng.Start > pStart Then
                    cBefore = doc.Range(fRng.Start - 1, fRng.Start).Text
                End If
                cAfter = ""
                If fRng.End < limit Then
                    cAfter = doc.Range(fRng.End, fRng.End + 1).Text
                End If
                If hasOpen Then
                    isValidClose = False
                    If cBefore <> " " And cBefore <> vbTab And cBefore <> vbCr And cBefore <> Chr(11) And cBefore <> Chr(7) Then
                        If Not (cAfter >= "0" And cAfter <= "9") Then
                            If fRng.End - openStart >= 2 And fRng.End - openStart <= 2000 Then
                                isValidClose = True
                            End If
                        End If
                    End If
                    If isValidClose Then
                        Set candRng = doc.Range(openStart, fRng.End)
                        If candRng.OMaths.Count = 0 And candRng.Fields.Count = 0 And Len(Trim$(doc.Range(openStart + 1, fRng.Start).Text)) > 0 Then
                            AppendPairUnique outStarts, outEnds, outPairCount, openStart, fRng.End
                        End If
                        hasOpen = False
                        openStart = 0
                    ElseIf CheckOpeningDollarDom(doc, fRng.End, limit, cAfter) Then
                        openStart = fRng.Start
                        hasOpen = True
                    End If
                ElseIf CheckOpeningDollarDom(doc, fRng.End, limit, cAfter) Then
                    openStart = fRng.Start
                    hasOpen = True
                End If
            End If
        End If
        If fRng.End >= limit Then
            Exit Do
        End If
        fRng.SetRange fRng.End, limit
    Loop
End Sub

Public Sub FindParenPairsDom(ByVal doc As Document, ByVal p As Paragraph, ByRef starts() As Long, ByRef ends() As Long, ByRef pairCount As Long)
    Dim fOpen As Range
    Dim fClose As Range
    Dim candRng As Range
    Dim pStart As Long
    Dim limit As Long
    Dim openPos As Long
    Dim closePos As Long
    Dim foundClose As Boolean

    pStart = p.Range.Start
    limit = p.Range.End
    If Right$(p.Range.Text, 1) = Chr(7) Or Right$(p.Range.Text, 1) = vbCr Then
        limit = limit - 1
    End If
    If limit <= pStart Then
        Exit Sub
    End If
    Set fOpen = doc.Range(pStart, limit)
    InitDomFind fOpen.Find, "\("
    Dim prevOpenEnd As Long
    Dim prevCloseEnd As Long
    Dim prevInEnd As Long
    prevOpenEnd = pStart
    Do While fOpen.Find.Execute
        If fOpen.Start >= limit Or fOpen.Start < prevOpenEnd Or fOpen.End <= prevOpenEnd Then
            Exit Do
        End If
        prevOpenEnd = fOpen.End
        If fOpen.OMaths.Count = 0 And fOpen.Fields.Count = 0 And Not IsDomPosEscaped(doc, pStart, fOpen.Start) Then
            openPos = fOpen.Start
            If fOpen.End >= limit Then
                Exit Do
            End If
            Set fClose = doc.Range(fOpen.End, limit)
            InitDomFind fClose.Find, "\)"
            foundClose = False
            prevCloseEnd = fOpen.End
            Do While fClose.Find.Execute
                If fClose.Start >= limit Or fClose.Start < prevCloseEnd Or fClose.End <= prevCloseEnd Then
                    Exit Do
                End If
                prevCloseEnd = fClose.End
                If fClose.OMaths.Count = 0 And fClose.Fields.Count = 0 And Not IsDomPosEscaped(doc, pStart, fClose.Start) Then
                    closePos = fClose.End
                    foundClose = True
                    Exit Do
                End If
                If fClose.End >= limit Then
                    Exit Do
                End If
                fClose.SetRange fClose.End, limit
            Loop
            If Not foundClose Then
                Exit Do
            End If
            If fOpen.End < fClose.Start Then
                fOpen.SetRange fOpen.End, fClose.Start
                prevInEnd = fOpen.Start
                Do While fOpen.Find.Execute
                    If fOpen.Start >= fClose.Start Or fOpen.Start >= limit Or fOpen.Start < prevInEnd Or fOpen.End <= prevInEnd Then
                        Exit Do
                    End If
                    prevInEnd = fOpen.End
                    If fOpen.OMaths.Count = 0 And fOpen.Fields.Count = 0 And Not IsDomPosEscaped(doc, pStart, fOpen.Start) Then
                        openPos = fOpen.Start
                    End If
                    If fOpen.End >= fClose.Start Then
                        Exit Do
                    End If
                    fOpen.SetRange fOpen.End, fClose.Start
                Loop
            End If
            Set candRng = doc.Range(openPos, closePos)
            If candRng.OMaths.Count = 0 And candRng.Fields.Count = 0 And closePos - openPos > 4 Then
                If Len(Trim$(doc.Range(openPos + 2, closePos - 2).Text)) > 0 Then
                    AppendPairUnique starts, ends, pairCount, openPos, closePos
                End If
            End If
            If closePos >= limit Then
                Exit Do
            End If
            prevOpenEnd = closePos
            fOpen.SetRange closePos, limit
        Else
            If fOpen.End >= limit Then
                Exit Do
            End If
            fOpen.SetRange fOpen.End, limit
        End If
    Loop
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
