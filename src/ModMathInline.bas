Attribute VB_Name = "ModMathInline"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModMathInline
' Purpose: Hybrid RAM/DOM inline delimiter pair scanner ($ and \()
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Public Function IsMathOpeningAheadText(ByVal txt As String, ByVal idxAfter As Long, ByVal txtLen As Long) As Boolean
    Dim searchPos As Long
    Dim nextDollarPos As Long
    Dim spanText As String
    Dim ch As String
    Dim chAfterNext As String

    searchPos = idxAfter
    nextDollarPos = 0
    Do While searchPos <= txtLen
        nextDollarPos = InStr(searchPos, txt, "$")
        If nextDollarPos <= 0 Then
            Exit Function
        End If
        If Not IsTextPosEscaped(txt, nextDollarPos) Then
            Exit Do
        End If
        searchPos = nextDollarPos + 1
        nextDollarPos = 0
    Loop
    If nextDollarPos <= 0 Then
        Exit Function
    End If
    If nextDollarPos < txtLen Then
        If Mid$(txt, nextDollarPos + 1, 1) = "$" Then
            Exit Function
        End If
    End If
    If nextDollarPos > 1 Then
        If Mid$(txt, nextDollarPos - 1, 1) = "$" Then
            If Not IsTextPosEscaped(txt, nextDollarPos - 1) Then
                Exit Function
            End If
        End If
    End If
    spanText = Mid$(txt, idxAfter, nextDollarPos - idxAfter)
    If Len(spanText) = 0 Or Len(spanText) > 250 Then
        Exit Function
    End If
    If InStr(1, spanText, vbCr) > 0 Or InStr(1, spanText, Chr(11)) > 0 Or InStr(1, spanText, Chr(7)) > 0 Then
        Exit Function
    End If
    ch = Right$(spanText, 1)
    If ch = " " Or ch = vbTab Or InStr(1, "/([{\+-*:;,=", ch) > 0 Then
        Exit Function
    End If
    If Not ModMathScanner.HasBalancedBrackets(spanText) Then
        Exit Function
    End If
    If nextDollarPos < txtLen Then
        chAfterNext = Mid$(txt, nextDollarPos + 1, 1)
        If chAfterNext >= "0" And chAfterNext <= "9" Then
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
                    If Not IsTextPosEscaped(txt, pos - 1) Then
                        isDbl = True
                    End If
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
                            If Len(Trim$(Mid$(txt, openIdx + 1, pos - openIdx - 1))) > 0 Then
                                isValidClose = True
                            End If
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
    canUseRam = (((pEnd - pStart) = Len(txt) Or ((pEnd - pStart) = Len(txt) - 1 And Right$(txt, 2) = vbCr & Chr(7))) And p.Range.OMaths.Count = 0 And p.Range.Fields.Count = 0)

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
                    If Len(Trim$(Mid$(txt, openAt + 2, closeAt - openAt - 2))) > 0 Then
                        ModMathHelpers.AppendPairUnique starts, ends, pairCount, pStart + openAt - 1, pStart + closeAt + 1
                    End If
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

Public Function IsTextPosEscaped(ByVal txt As String, ByVal pos As Long) As Boolean
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

Public Function FindUnescapedToken(ByVal txt As String, ByVal startAt As Long, ByVal tok As String) As Long
    Dim p As Long

    p = InStr(startAt, txt, tok)
    Do While p > 0
        If Not IsTextPosEscaped(txt, p) Then
            FindUnescapedToken = p
            Exit Function
        End If
        p = InStr(p + 1, txt, tok)
    Loop
End Function

Public Function CountTokenSubstr(ByVal txt As String, ByVal tok As String) As Long
    Dim cnt As Long
    Dim p As Long

    p = InStr(1, txt, tok)
    Do While p > 0
        cnt = cnt + 1
        p = InStr(p + Len(tok), txt, tok)
    Loop
    CountTokenSubstr = cnt
End Function

Public Function IsSingleLineDisplayBlock(ByVal pText As String) As Boolean
    Dim braceClose As Long
    Dim envName As String
    Dim beginTok As String
    Dim endTok As String

    If Len(pText) <= 4 Then
        Exit Function
    End If
    If Left$(pText, 2) = "$$" And Right$(pText, 2) = "$$" Then
        IsSingleLineDisplayBlock = (FindUnescapedToken(pText, 3, "$$") = Len(pText) - 1)
    ElseIf Left$(pText, 2) = "\[" And Right$(pText, 2) = "\]" Then
        IsSingleLineDisplayBlock = (FindUnescapedToken(pText, 3, "\]") = Len(pText) - 1)
    ElseIf Left$(pText, 7) = "\begin{" And Right$(pText, 1) = "}" Then
        braceClose = InStr(8, pText, "}")
        If braceClose > 8 Then
            envName = Trim$(Mid$(pText, 8, braceClose - 8))
            beginTok = "\begin{" & envName & "}"
            endTok = "\end{" & envName & "}"
            If Len(envName) > 0 And Right$(pText, Len(endTok)) = endTok And Len(pText) > Len(beginTok) + Len(endTok) Then
                IsSingleLineDisplayBlock = (CountTokenSubstr(pText, beginTok) = CountTokenSubstr(pText, endTok))
            End If
        End If
    End If
End Function

Public Sub TrimMathRange(ByRef rng As Range)
    Dim txt As String
    Dim c As String
    Dim prevPos As Long

    Do While rng.End > rng.Start
        txt = rng.Text
        If Len(txt) = 0 Then
            Exit Do
        End If
        c = Left$(txt, 1)
        If c = "." And Len(txt) >= 2 Then
            If Mid$(txt, 2, 1) >= "0" And Mid$(txt, 2, 1) <= "9" Then
                Exit Do
            End If
        End If
        If InStr(1, " " & vbTab & vbCr & Chr(11) & ".,;:?", c) > 0 Then
            prevPos = rng.Start
            rng.Start = rng.Start + 1
            If rng.Start <= prevPos Then
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    Do While rng.End > rng.Start
        txt = rng.Text
        If Len(txt) = 0 Then
            Exit Do
        End If
        c = Right$(txt, 1)
        If c = "." And Len(txt) >= 7 Then
            If Right$(txt, 7) = "\right." Then
                Exit Do
            End If
        End If
        If InStr(1, " " & vbTab & vbCr & Chr(11) & ".,;:?", c) > 0 Then
            prevPos = rng.End
            rng.End = rng.End - 1
            If rng.End >= prevPos Then
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
End Sub
Public Function HasForbiddenDomChar(ByVal txt As String) As Boolean
    Do While Len(txt) > 0 And (Right$(txt, 1) = vbCr Or Right$(txt, 1) = Chr(7))
        txt = Left$(txt, Len(txt) - 1)
    Loop
    If InStr(1, txt, Chr(1)) > 0 Or InStr(1, txt, Chr(2)) > 0 Or InStr(1, txt, Chr(5)) > 0 Or InStr(1, txt, Chr(7)) > 0 Or InStr(1, txt, Chr(12)) > 0 Then
        HasForbiddenDomChar = True
    End If
End Function

Public Sub TrimTrailingDomMarks(ByVal rng As Range)
    Dim prevEnd As Long
    Do While rng.End > rng.Start
        If Right$(rng.Text, 1) = vbCr Or Right$(rng.Text, 1) = Chr(7) Then
            prevEnd = rng.End
            rng.End = rng.End - 1
            If rng.End >= prevEnd Then
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
