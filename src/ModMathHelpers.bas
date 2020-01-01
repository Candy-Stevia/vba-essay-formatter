Attribute VB_Name = "ModMathHelpers"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModMathHelpers
' Purpose: Math ID cache, Variables, Shading, Bookmark lookup & DOM fallback
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Private g_MathSeqCached As Long
Private g_HasSeqCache As Boolean

Public Sub InitMathSeqCache(ByVal doc As Document)
    Dim curSeq As Long
    Dim bmNum As Long
    Dim bm As Bookmark
    Dim v As Variable
    Dim vStr As String
    Dim suffix As String

    curSeq = 0
    On Error Resume Next
    vStr = doc.Variables("mathcfg_seq").Value
    If Err.Number = 0 And Len(vStr) > 0 Then
        curSeq = CLng(vStr)
    End If
    Err.Clear
    For Each bm In doc.Bookmarks
        If Left$(bm.Name, 8) = "bm_math_" Then
            suffix = Mid$(bm.Name, InStrRev(bm.Name, "_") + 1)
            If IsNumeric(suffix) Then
                bmNum = CLng(suffix)
                If bmNum > curSeq Then
                    curSeq = bmNum
                End If
            End If
        End If
    Next bm
    For Each v In doc.Variables
        If Left$(v.Name, 8) = "bm_math_" Then
            suffix = Mid$(v.Name, InStrRev(v.Name, "_") + 1)
            If IsNumeric(suffix) Then
                bmNum = CLng(suffix)
                If bmNum > curSeq Then
                    curSeq = bmNum
                End If
            End If
        End If
    Next v
    On Error GoTo 0
    g_MathSeqCached = curSeq
    g_HasSeqCache = True
End Sub

Public Sub ResetMathSeqCache()
    g_MathSeqCached = 0
    g_HasSeqCache = False
End Sub
Public Sub FlushMathSeqCache(ByVal doc As Document)
    If g_HasSeqCache And g_MathSeqCached > 0 Then
        SetDocVariable doc, "mathcfg_seq", CStr(g_MathSeqCached)
    End If
End Sub
Public Function NextMathSeq(ByVal doc As Document) As Long
    If Not g_HasSeqCache Then
        InitMathSeqCache doc
        g_HasSeqCache = False
        g_MathSeqCached = g_MathSeqCached + 1
        SetDocVariable doc, "mathcfg_seq", CStr(g_MathSeqCached)
        NextMathSeq = g_MathSeqCached
        Exit Function
    End If
    g_MathSeqCached = g_MathSeqCached + 1
    NextMathSeq = g_MathSeqCached
End Function
Public Sub SetDocVariable(ByVal doc As Document, ByVal varName As String, ByVal varVal As String)
    On Error Resume Next
    doc.Variables(varName).Value = varVal
    If Err.Number <> 0 Then
        Err.Clear
        doc.Variables.Add varName, varVal
    End If
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
    Dim bm As Bookmark
    Dim bStart As Long
    Dim bEnd As Long
    Dim ovStart As Long
    Dim ovEnd As Long
    Dim ovLen As Long
    Dim bestOverlap As Long
    Dim bestDist As Long
    Dim kind As Long
    Dim bestKind As Long
    Dim bestBm As String
    Dim isOverlapping As Boolean

    outBmName = ""
    outBmKind = ModMathScanner.BM_KIND_NONE
    bestBm = ""
    bestOverlap = -1
    bestDist = 999999
    bestKind = ModMathScanner.BM_KIND_NONE
    For Each bm In doc.Bookmarks
        If Left$(bm.Name, 8) = "bm_math_" Then
            bStart = bm.Range.Start
            bEnd = bm.Range.End
            If selEnd > selStart Then
                isOverlapping = (selStart < bEnd And selEnd > bStart)
            Else
                isOverlapping = (selStart >= bStart And selEnd <= bEnd)
            End If
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
                ovStart = selStart
                If bStart > ovStart Then
                    ovStart = bStart
                End If
                ovEnd = selEnd
                If bEnd < ovEnd Then
                    ovEnd = bEnd
                End If
                ovLen = ovEnd - ovStart
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
    Dim overlaps As Boolean

    overlaps = False
    For i = 1 To pairCount
        If firstPos < ends(i) And lastPos > starts(i) Then
            overlaps = True
            Exit For
        End If
    Next i
    If Not overlaps Then
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
    End If
End Sub
Private Function IsDomPosEscaped(ByVal doc As Document, ByVal pStart As Long, ByVal pos As Long) As Boolean
    Dim curPos As Long
    Dim slashCount As Long
    curPos = pos
    slashCount = 0
    Do While curPos > pStart
        If doc.Range(curPos - 1, curPos).Text <> "\" Then
            Exit Do
        End If
        slashCount = slashCount + 1
        curPos = curPos - 1
    Loop
    IsDomPosEscaped = (slashCount Mod 2 = 1)
End Function
Private Function IsMathOpeningAheadDom(ByVal doc As Document, ByVal posAfter As Long, ByVal limit As Long) As Boolean
    Dim aheadText As String
    Dim spanText As String
    Dim ch As String
    Dim chAfterNext As String
    Dim nextDollarPos As Long
    aheadText = doc.Range(posAfter, limit).Text
    nextDollarPos = InStr(1, aheadText, "$")
    If nextDollarPos <= 0 Then
        IsMathOpeningAheadDom = False
        Exit Function
    End If
    spanText = Left$(aheadText, nextDollarPos - 1)
    If Len(spanText) = 0 Or Len(spanText) > 250 Then
        IsMathOpeningAheadDom = False
        Exit Function
    End If
    If InStr(1, spanText, vbCr) > 0 Or InStr(1, spanText, Chr(11)) > 0 Then
        IsMathOpeningAheadDom = False
        Exit Function
    End If
    ch = Right$(spanText, 1)
    If ch = " " Or ch = vbTab Then
        IsMathOpeningAheadDom = False
        Exit Function
    End If
    If nextDollarPos < Len(aheadText) Then
        chAfterNext = Mid$(aheadText, nextDollarPos + 1, 1)
        If chAfterNext >= "0" And chAfterNext <= "9" Then
            IsMathOpeningAheadDom = False
            Exit Function
        End If
    End If
    IsMathOpeningAheadDom = True
End Function

Private Function CheckOpeningDollarDom(ByVal doc As Document, ByVal posEnd As Long, ByVal limit As Long, ByVal cAfter As String) As Boolean
    If cAfter = " " Or cAfter = vbTab Or cAfter = vbCr Or cAfter = Chr(11) Or cAfter = Chr(7) Then
        CheckOpeningDollarDom = False
        Exit Function
    End If
    If cAfter >= "0" And cAfter <= "9" Then
        CheckOpeningDollarDom = IsMathOpeningAheadDom(doc, posEnd, limit)
    Else
        CheckOpeningDollarDom = True
    End If
End Function
Public Sub FindDollarPairsDom(ByVal doc As Document, ByVal p As Paragraph, ByRef outStarts() As Long, ByRef outEnds() As Long, ByRef outPairCount As Long)
    Dim fRng As Range
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
    Set fRng = p.Range.Duplicate
    fRng.Find.ClearFormatting
    fRng.Find.MatchWildcards = False
    fRng.Find.Text = "$"
    fRng.Find.Forward = True
    fRng.Find.Wrap = 0
    hasOpen = False
    openStart = 0
    Do While fRng.Find.Execute
        If fRng.Start >= limit Then
            Exit Do
        End If
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
                    isDbl = True
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
                    AppendPairUnique outStarts, outEnds, outPairCount, openStart, fRng.End
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
        fRng.Collapse 0
    Loop
End Sub

Public Sub FindParenPairsDom(ByVal doc As Document, ByVal p As Paragraph, ByRef starts() As Long, ByRef ends() As Long, ByRef pairCount As Long)
    Dim fOpen As Range
    Dim fClose As Range
    Dim pStart As Long
    Dim limit As Long
    Dim openPos As Long
    Dim closePos As Long
    Dim foundClose As Boolean

    pStart = p.Range.Start
    limit = p.Range.End
    Set fOpen = p.Range.Duplicate
    fOpen.Find.ClearFormatting
    fOpen.Find.MatchWildcards = False
    fOpen.Find.Text = "\("
    fOpen.Find.Forward = True
    fOpen.Find.Wrap = 0

    Do While fOpen.Find.Execute
        If fOpen.Start >= limit Then
            Exit Do
        End If
        If Not IsDomPosEscaped(doc, pStart, fOpen.Start) Then
            openPos = fOpen.Start
            Set fClose = doc.Range(fOpen.End, limit)
            fClose.Find.ClearFormatting
            fClose.Find.MatchWildcards = False
            fClose.Find.Text = "\)"
            fClose.Find.Forward = True
            fClose.Find.Wrap = 0
            foundClose = False
            Do While fClose.Find.Execute
                If fClose.Start >= limit Then
                    Exit Do
                End If
                If Not IsDomPosEscaped(doc, pStart, fClose.Start) Then
                    closePos = fClose.End
                    foundClose = True
                    Exit Do
                End If
                fClose.Collapse 0
            Loop
            If Not foundClose Then
                Exit Do
            End If
            fOpen.Collapse 0
            Do While fOpen.Find.Execute
                If fOpen.Start >= fClose.Start Or fOpen.Start >= limit Then
                    Exit Do
                End If
                If Not IsDomPosEscaped(doc, pStart, fOpen.Start) Then
                    openPos = fOpen.Start
                End If
                fOpen.Collapse 0
            Loop
            AppendPairUnique starts, ends, pairCount, openPos, closePos
            fOpen.SetRange closePos, limit
        Else
            fOpen.Collapse 0
        End If
    Loop
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
