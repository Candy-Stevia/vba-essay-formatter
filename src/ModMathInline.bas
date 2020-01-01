Attribute VB_Name = "ModMathInline"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModMathInline
' Purpose: Hybrid RAM/DOM inline delimiter pair scanner ($ and \()
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Private Function IsMathOpeningAheadText(ByVal txt As String, ByVal idxAfter As Long, ByVal txtLen As Long) As Boolean
    Dim nextDollarPos As Long
    Dim spanText As String
    Dim ch As String
    Dim chAfterNext As String

    nextDollarPos = InStr(idxAfter, txt, "$")
    If nextDollarPos <= 0 Then
        IsMathOpeningAheadText = False
        Exit Function
    End If
    spanText = Mid$(txt, idxAfter, nextDollarPos - idxAfter)
    If Len(spanText) = 0 Or Len(spanText) > 250 Then
        IsMathOpeningAheadText = False
        Exit Function
    End If
    If InStr(1, spanText, vbCr) > 0 Or InStr(1, spanText, Chr(11)) > 0 Then
        IsMathOpeningAheadText = False
        Exit Function
    End If
    ch = Right$(spanText, 1)
    If ch = " " Or ch = vbTab Then
        IsMathOpeningAheadText = False
        Exit Function
    End If
    If nextDollarPos < txtLen Then
        chAfterNext = Mid$(txt, nextDollarPos + 1, 1)
        If chAfterNext >= "0" And chAfterNext <= "9" Then
            IsMathOpeningAheadText = False
            Exit Function
        End If
    End If
    IsMathOpeningAheadText = True
End Function

Private Function CheckOpeningDollarText(ByVal txt As String, ByVal idxAfter As Long, ByVal txtLen As Long, ByVal cAfter As String) As Boolean
    If cAfter = " " Or cAfter = vbTab Or cAfter = vbCr Or cAfter = Chr(11) Or cAfter = Chr(7) Then
        CheckOpeningDollarText = False
        Exit Function
    End If
    If cAfter >= "0" And cAfter <= "9" Then
        CheckOpeningDollarText = IsMathOpeningAheadText(txt, idxAfter, txtLen)
    Else
        CheckOpeningDollarText = True
    End If
End Function

Public Sub FindDollarPairsInText(ByVal pStart As Long, ByVal txt As String, ByRef outStarts() As Long, ByRef outEnds() As Long, ByRef outPairCount As Long)
    Dim txtLen As Long
    Dim pos As Long
    Dim openIdx As Long
    Dim isEscaped As Boolean
    Dim isDbl As Boolean
    Dim hasOpen As Boolean
    Dim isValidClose As Boolean
    Dim cBefore As String
    Dim cAfter As String

    outPairCount = 0
    txtLen = Len(txt)
    hasOpen = False
    openIdx = 0

    pos = InStr(1, txt, "$")
    Do While pos > 0
        isEscaped = IsTextPosEscaped(txt, pos)
        isDbl = False
        If Not isEscaped Then
            If pos < txtLen Then
                If Mid$(txt, pos + 1, 1) = "$" Then
                    isDbl = True
                End If
            End If
            If pos > 1 Then
                If Mid$(txt, pos - 1, 1) = "$" Then
                    isDbl = True
                End If
            End If
        End If

        If isDbl Then
            hasOpen = False
            openIdx = 0
        ElseIf Not isEscaped Then
            cBefore = ""
            If pos > 1 Then
                cBefore = Mid$(txt, pos - 1, 1)
            End If
            cAfter = ""
            If pos < txtLen Then
                cAfter = Mid$(txt, pos + 1, 1)
            End If

            If hasOpen Then
                isValidClose = False
                If cBefore <> " " And cBefore <> vbTab And cBefore <> vbCr And cBefore <> Chr(11) And cBefore <> Chr(7) Then
                    If Not (cAfter >= "0" And cAfter <= "9") Then
                        If pos - openIdx + 1 >= 2 And pos - openIdx + 1 <= 2000 Then
                            isValidClose = True
                        End If
                    End If
                End If
                If isValidClose Then
                    ModMathHelpers.AppendPairUnique outStarts, outEnds, outPairCount, pStart + openIdx - 1, pStart + pos
                    hasOpen = False
                    openIdx = 0
                ElseIf CheckOpeningDollarText(txt, pos + 1, txtLen, cAfter) Then
                    openIdx = pos
                    hasOpen = True
                End If
            ElseIf CheckOpeningDollarText(txt, pos + 1, txtLen, cAfter) Then
                openIdx = pos
                hasOpen = True
            End If
        End If
        pos = InStr(pos + 1, txt, "$")
    Loop
End Sub

Public Sub FindDollarPairsInPara(ByVal doc As Document, ByVal p As Paragraph, ByRef outStarts() As Long, ByRef outEnds() As Long, ByRef outPairCount As Long)
    ModMathHelpers.FindDollarPairsDom doc, p, outStarts, outEnds, outPairCount
End Sub

Public Sub FindInlinePairs(ByVal doc As Document, ByVal p As Paragraph, ByRef starts() As Long, ByRef ends() As Long, ByRef pairCount As Long, Optional ByVal preReadText As String = "")
    Dim txt As String
    Dim pStart As Long
    Dim pEnd As Long
    Dim openAt As Long
    Dim closeAt As Long
    Dim searchAt As Long
    Dim scanPos As Long
    Dim nextOpen As Long
    Dim i As Long
    Dim j As Long
    Dim temp As Long
    Dim canUseRam As Boolean

    pairCount = 0
    If Len(preReadText) > 0 Then
        txt = preReadText
    Else
        txt = p.Range.Text
    End If
    pStart = p.Range.Start
    pEnd = p.Range.End
    canUseRam = ((pEnd - pStart) = Len(txt) And p.Range.OMaths.Count = 0 And p.Range.Fields.Count = 0)

    If canUseRam Then
        If InStr(1, txt, "$") > 0 Then
            FindDollarPairsInText pStart, txt, starts, ends, pairCount
        End If
        If InStr(1, txt, "\(") > 0 Then
            searchAt = 1
            Do While searchAt <= Len(txt)
                openAt = InStr(searchAt, txt, "\(")
                If openAt = 0 Then
                    Exit Do
                End If
                If IsTextPosEscaped(txt, openAt) Then
                    searchAt = openAt + 2
                Else
                    closeAt = FindClosing(txt, openAt + 2)
                    If closeAt = 0 Then
                        Exit Do
                    End If
                    scanPos = openAt + 2
                    Do While scanPos < closeAt
                        nextOpen = InStr(scanPos, txt, "\(")
                        If nextOpen = 0 Or nextOpen >= closeAt Then
                            Exit Do
                        End If
                        If Not IsTextPosEscaped(txt, nextOpen) Then
                            openAt = nextOpen
                        End If
                        scanPos = nextOpen + 2
                    Loop
                    ModMathHelpers.AppendPairUnique starts, ends, pairCount, pStart + openAt - 1, pStart + closeAt + 1
                    searchAt = closeAt + 2
                End If
            Loop
        End If
    Else
        If InStr(1, txt, "$") > 0 Then
            ModMathHelpers.FindDollarPairsDom doc, p, starts, ends, pairCount
        End If
        If InStr(1, txt, "\(") > 0 Then
            ModMathHelpers.FindParenPairsDom doc, p, starts, ends, pairCount
        End If
    End If

    For i = 1 To pairCount
        For j = i + 1 To pairCount
            If starts(i) > starts(j) Then
                temp = starts(i)
                starts(i) = starts(j)
                starts(j) = temp
                temp = ends(i)
                ends(i) = ends(j)
                ends(j) = temp
            End If
        Next j
    Next i
End Sub

Private Function FindClosing(ByVal txt As String, ByVal searchAt As Long) As Long
    Dim pos As Long

    pos = InStr(searchAt, txt, "\)")
    Do While pos > 0
        If Not IsTextPosEscaped(txt, pos) Then
            FindClosing = pos
            Exit Function
        End If
        pos = InStr(pos + 2, txt, "\)")
    Loop
End Function

Private Function IsTextPosEscaped(ByVal txt As String, ByVal pos As Long) As Boolean
    Dim i As Long
    Dim slashCount As Long

    i = pos - 1
    slashCount = 0
    Do While i > 0
        If Mid$(txt, i, 1) <> "\" Then
            Exit Do
        End If
        slashCount = slashCount + 1
        i = i - 1
    Loop
    IsTextPosEscaped = (slashCount Mod 2 = 1)
End Function

Public Sub TrimMathRange(ByRef rng As Range)
    Dim c As String

    Do While rng.End > rng.Start
        c = Left$(rng.Text, 1)
        If InStr(1, " " & vbTab & vbCr & Chr(11) & ".,;:?", c) > 0 Then
            rng.Start = rng.Start + 1
        Else
            Exit Do
        End If
    Loop
    Do While rng.End > rng.Start
        c = Right$(rng.Text, 1)
        If InStr(1, " " & vbTab & vbCr & Chr(11) & ".,;:?", c) > 0 Then
            rng.End = rng.End - 1
        Else
            Exit Do
        End If
    Loop
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
