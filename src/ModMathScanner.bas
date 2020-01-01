Attribute VB_Name = "ModMathScanner"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ModMathScanner: DOM-level coordinate scanning, Pandoc dollar rules, and No-Fly zone
Public Const COLOR_PASTEL_MINT As Long = 14808540
Public Const COLOR_PASTEL_CYAN As Long = 16119260
Public Const COLOR_PASTEL_PINK As Long = 15132415
Public Const COLOR_PASTEL_AMBER As Long = 13497343
Public Const COLOR_AUTO_CLEAR As Long = -16777216
Public Const MATH_MERGE_MAX_SPAN As Long = 40
Public Const BM_KIND_NONE As Long = 0
Public Const BM_KIND_OMATH As Long = 1
Public Const BM_KIND_SRC As Long = 2
Public Const BM_KIND_PINK As Long = 3
Private g_HasZone(1 To 3) As Boolean
Private g_ZoneBms(1 To 3) As Bookmark
Private g_ZoneDocPtr As LongPtr

Public Sub InitNoFlyZoneCache(ByVal doc As Document)
    Dim names(1 To 3) As String
    Dim i As Long
    ResetNoFlyZoneCache
    If doc Is Nothing Then
        Exit Sub
    End If
    g_ZoneDocPtr = ObjPtr(doc)
    names(1) = "bm_zone_toc"
    names(2) = "bm_zone_lot"
    names(3) = "bm_zone_lof"
    On Error Resume Next
    For i = 1 To 3
        If doc.Bookmarks.Exists(names(i)) Then
            Set g_ZoneBms(i) = doc.Bookmarks(names(i))
            g_HasZone(i) = (Err.Number = 0 And Not g_ZoneBms(i) Is Nothing)
            Err.Clear
        End If
    Next i
    On Error GoTo 0
End Sub

Public Sub ResetNoFlyZoneCache()
    Dim i As Long
    For i = 1 To 3
        g_HasZone(i) = False
        Set g_ZoneBms(i) = Nothing
    Next i
    g_ZoneDocPtr = 0
End Sub

Public Function HasBalancedBrackets(ByVal txt As String) As Boolean
    Dim i As Long
    Dim braceDepth As Long
    Dim pDepth As Long
    Dim bDepth As Long
    Dim exactBroken As Boolean
    Dim iBroken As Boolean
    Dim isEsc As Boolean
    Dim c As String

    For i = 1 To Len(txt)
        c = Mid$(txt, i, 1)
        If isEsc Then
            isEsc = False
        Else
            Select Case c
                Case "\"
                    isEsc = True
                Case "{"
                    braceDepth = braceDepth + 1
                Case "}"
                    braceDepth = braceDepth - 1
                    If braceDepth < 0 Then
                        Exit Function
                    End If
                Case "("
                    pDepth = pDepth + 1
                Case ")"
                    pDepth = pDepth - 1
                Case "["
                    bDepth = bDepth + 1
                Case "]"
                    bDepth = bDepth - 1
            End Select
            If pDepth < 0 Or bDepth < 0 Then
                exactBroken = True
            End If
            If pDepth + bDepth < 0 Then
                iBroken = True
            End If
        End If
    Next i
    If braceDepth <> 0 Then
        Exit Function
    End If
    If Not exactBroken And pDepth = 0 And bDepth = 0 Then
        HasBalancedBrackets = True
    ElseIf Not iBroken And (pDepth + bDepth) = 0 And (InStr(1, txt, ",") > 0 Or InStr(1, txt, ";") > 0) Then
        HasBalancedBrackets = True
    ElseIf InStr(1, txt, "\left") > 0 And InStr(1, txt, "\right") > 0 Then
        If ModMathInline.CountTokenSubstr(txt, "\left") = ModMathInline.CountTokenSubstr(txt, "\right") Then
            HasBalancedBrackets = True
        End If
    End If
End Function

Private Function IsSingleCurrencyUnit(ByVal rawUnit As String, ByVal lowUnit As String, ByVal numPart As String, ByVal currUnits As String) As Boolean
    If InStr(1, currUnits, "|" & lowUnit & "|") > 0 Then
        IsSingleCurrencyUnit = True
    ElseIf rawUnit = "M" Or rawUnit = "B" Or rawUnit = "K" Then
        IsSingleCurrencyUnit = True
    ElseIf lowUnit = "k" Then
        If Val(Replace(Replace(numPart, ",", ""), " ", "")) >= 10# Then
            IsSingleCurrencyUnit = True
        End If
    End If
End Function

Public Function IsCurrencyText(ByVal txt As String) As Boolean
    Dim rawT As String
    Dim t As String
    Dim c As String
    Dim i As Long
    Dim numPart As String
    Dim rawUnitPart As String
    Dim unitPart As String
    Dim slashPos As Long
    Dim leftTok As String
    Dim rightTok As String
    Dim currUnits As String
    Dim periodUnits As String

    currUnits = "|usd|vnd|eur|gbp|jpy|cny|chf|tr|bn|dong|nghin|ngan|trieu|ty|dollar|dollars|cent|cents|thousand|million|billion|" & _
                ModConstants.Uni("tri\u1ec7u|t\u1ef7|ngh\u00ecn|ng\u00e0n|\u0111\u1ed3ng|\u0111|")
    periodUnits = "|month|mo|year|yr|day|ngay|unit|thang|nam|nguoi|don vi|" & ModConstants.Uni("th\u00e1ng|n\u0103m|ng\u00e0y|ng\u01b0\u1eddi|\u0111\u01a1n v\u1ecb|")
    rawT = Trim$(txt)
    t = LCase$(rawT)
    If Len(t) = 0 Or InStr(1, t, "\") > 0 Or InStr(1, t, "=") > 0 Or InStr(1, t, "{") > 0 Or InStr(1, t, "}") > 0 Or InStr(1, t, "^") > 0 Or InStr(1, t, "_") > 0 Then
        Exit Function
    End If
    c = Left$(t, 1)
    If c < "0" Or c > "9" Then
        Exit Function
    End If
    For i = 1 To Len(rawT)
        c = Mid$(rawT, i, 1)
        If InStr(1, "0123456789., ", c) = 0 Then
            numPart = Left$(rawT, i - 1)
            rawUnitPart = Trim$(Mid$(rawT, i))
            Exit For
        End If
    Next i
    unitPart = LCase$(rawUnitPart)
    If Len(unitPart) = 0 Then
        Exit Function
    End If
    slashPos = InStr(1, unitPart, "/")
    If slashPos > 0 Then
        leftTok = Trim$(Left$(unitPart, slashPos - 1))
        rightTok = Trim$(Mid$(unitPart, slashPos + 1))
        If Len(rightTok) > 0 And InStr(1, periodUnits, "|" & rightTok & "|") > 0 Then
            If Len(leftTok) = 0 Then
                IsCurrencyText = True
            Else
                IsCurrencyText = IsSingleCurrencyUnit(Trim$(Left$(rawUnitPart, slashPos - 1)), leftTok, numPart, currUnits)
            End If
        End If
    Else
        IsCurrencyText = IsSingleCurrencyUnit(rawUnitPart, unitPart, numPart, currUnits)
    End If
End Function

Public Function ClassifyDollarSpan(ByVal innerText As String, ByVal charBeforeOpen As String, ByVal charAfterClose As String) As Long
    Dim t As String
    Dim c As String
    Dim c1 As String
    Dim cLast As String
    Dim i As Long
    Dim hasNonDigits As Boolean
    Dim hasMathProof As Boolean

    t = innerText
    If Len(t) > 250 Then
        ClassifyDollarSpan = 2
    ElseIf Not HasBalancedBrackets(t) Then
        ClassifyDollarSpan = 2
    Else
        c1 = Left$(t, 1)
        cLast = Right$(t, 1)
        hasMathProof = (InStr(1, t, "\") > 0 Or InStr(1, t, "=") > 0 Or InStr(1, t, "^") > 0 Or InStr(1, t, "_") > 0)
        If (c1 = " " Or c1 = vbTab Or cLast = " " Or cLast = vbTab) And Not hasMathProof Then
            ClassifyDollarSpan = 0
        ElseIf Len(charAfterClose) > 0 And charAfterClose >= "0" And charAfterClose <= "9" Then
            ClassifyDollarSpan = 0
        ElseIf IsCurrencyText(t) Then
            ClassifyDollarSpan = 0
        Else
            For i = 1 To Len(t)
                c = Mid$(t, i, 1)
                If InStr(1, "0123456789.,% ", c) = 0 Then
                    hasNonDigits = True
                    Exit For
                End If
            Next i
            If Not hasNonDigits And Len(Trim$(t)) > 0 Then
                ClassifyDollarSpan = 2
            ElseIf (c1 >= "0" And c1 <= "9") And InStr(1, t, " ") > 0 And Not hasMathProof Then
                ClassifyDollarSpan = 2
            Else
                ClassifyDollarSpan = 1
            End If
        End If
    End If
End Function

Public Function StripOuterDelims(ByVal txt As String) As String
    Dim s As String
    s = Trim$(txt)
    If Left$(s, 2) = "$$" And Right$(s, 2) = "$$" And Len(s) >= 4 Then
        s = Trim$(Mid$(s, 3, Len(s) - 4))
    ElseIf Left$(s, 2) = "\[" And Right$(s, 2) = "\]" And Len(s) >= 4 Then
        s = Trim$(Mid$(s, 3, Len(s) - 4))
    ElseIf Left$(s, 2) = "\(" And Right$(s, 2) = "\)" And Len(s) >= 4 Then
        s = Trim$(Mid$(s, 3, Len(s) - 4))
    ElseIf Left$(s, 1) = "$" And Right$(s, 1) = "$" And Len(s) >= 2 Then
        s = Trim$(Mid$(s, 2, Len(s) - 2))
    End If
    StripOuterDelims = s
End Function

Private Function CheckNoFlyHeading(ByVal rawText As String, ByRef isScanningActive As Boolean) As Boolean
    Dim cleanHead As String
    cleanHead = LCase$(Trim$(Replace(Replace(Replace(rawText, vbCr, ""), Chr(11), ""), Chr(7), "")))
    If Len(cleanHead) <= 45 Then
        If cleanHead = "references" Or cleanHead = "bibliography" Or cleanHead = "tai lieu tham khao" Or cleanHead = "danh muc tai lieu tham khao" Or cleanHead = ModConstants.Uni("t\u00e0i li\u1ec7u tham kh\u1ea3o") Or cleanHead = ModConstants.Uni("danh m\u1ee5c t\u00e0i li\u1ec7u tham kh\u1ea3o") Then
            isScanningActive = False
            CheckNoFlyHeading = True
            Exit Function
        ElseIf Left$(cleanHead, 8) = "appendix" Or Left$(cleanHead, 10) = "appendices" Or Left$(cleanHead, 7) = "phu luc" Or Left$(cleanHead, 7) = ModConstants.Uni("ph\u1ee5 l\u1ee5c") Then
            isScanningActive = True
        End If
    End If
    CheckNoFlyHeading = (Not isScanningActive)
End Function

Public Function IsInNoFlyZoneCached(ByVal pRng As Range, ByVal rawText As String, ByRef isScanningActive As Boolean) As Boolean
    Dim i As Long
    Dim bmStart As Long
    Dim bmEnd As Long
    Dim rStart As Long
    Dim rEnd As Long
    Dim docPtr As LongPtr

    If g_ZoneDocPtr <> 0 Then
        On Error Resume Next
        docPtr = ObjPtr(pRng.Parent)
        If Err.Number <> 0 Then
            Err.Clear
            docPtr = 0
        End If
        On Error GoTo 0
        If docPtr <> g_ZoneDocPtr Then
            ResetNoFlyZoneCache
        End If
    End If
    On Error Resume Next
    rStart = pRng.Start
    rEnd = pRng.End
    If Err.Number = 0 Then
        For i = 1 To 3
            If g_HasZone(i) And Not g_ZoneBms(i) Is Nothing Then
                bmStart = g_ZoneBms(i).Start
                bmEnd = g_ZoneBms(i).End
                If Err.Number <> 0 Then
                    Err.Clear
                    g_HasZone(i) = False
                    Set g_ZoneBms(i) = Nothing
                ElseIf rStart < bmEnd And rEnd > bmStart Then
                    On Error GoTo 0
                    IsInNoFlyZoneCached = True
                    Exit Function
                End If
            End If
        Next i
    Else
        Err.Clear
    End If
    On Error GoTo 0
    IsInNoFlyZoneCached = CheckNoFlyHeading(rawText, isScanningActive)
End Function

Public Function IsInNoFlyZone(ByVal doc As Document, ByVal pRng As Range, ByRef isScanningActive As Boolean) As Boolean
    InitNoFlyZoneCache doc
    IsInNoFlyZone = IsInNoFlyZoneCached(pRng, pRng.Text, isScanningActive)
    ResetNoFlyZoneCache
End Function

Public Function CheckMultilineBlock(ByVal doc As Document, ByVal startPara As Paragraph, ByRef outEndPara As Paragraph, ByRef outRawLatex As String, ByRef outIsAmbiguous As Boolean) As Boolean
    Dim p1Raw As String
    Dim p1Text As String
    Dim beginToken As String
    Dim endToken As String
    Dim envName As String
    Dim curText As String
    Dim curTrim As String
    Dim combined As String
    Dim curPara As Paragraph
    Dim spanStep As Long
    Dim braceClose As Long
    Dim blockKind As Long
    Dim depth As Long
    Dim startInTable As Boolean
    Dim tmpScan As Boolean
    Dim closedHere As Boolean
    p1Raw = startPara.Range.Text
    If InStr(1, p1Raw, Chr(7)) > 0 Then
        Exit Function
    End If
    p1Text = Trim$(Replace(Replace(p1Raw, vbCr, ""), Chr(11), " "))
    If Left$(p1Text, 2) = "$$" And Left$(p1Text, 3) <> "$$$" Then
        If ModMathInline.FindUnescapedToken(p1Text, 3, "$$") = 0 Then
            blockKind = 1
            endToken = "$$"
        End If
    ElseIf Left$(p1Text, 2) = "\[" Then
        If ModMathInline.FindUnescapedToken(p1Text, 3, "\]") = 0 Then
            blockKind = 2
            endToken = "\]"
        End If
    ElseIf Left$(p1Text, 7) = "\begin{" Then
        braceClose = InStr(8, p1Text, "}")
        If braceClose > 8 Then
            envName = Trim$(Mid$(p1Text, 8, braceClose - 8))
            If Len(envName) > 0 Then
                beginToken = "\begin{" & envName & "}"
                endToken = "\end{" & envName & "}"
                depth = ModMathInline.CountTokenSubstr(p1Text, beginToken) - ModMathInline.CountTokenSubstr(p1Text, endToken)
                If depth > 0 Then
                    blockKind = 3
                End If
            End If
        End If
    End If
    If blockKind = 0 Then
        Exit Function
    End If
    startInTable = startPara.Range.Information(12)
    combined = Replace(Replace(p1Raw, vbCr, " "), Chr(11), " ")
    Set curPara = startPara.Next
    spanStep = 1
    Do While Not curPara Is Nothing And spanStep < MATH_MERGE_MAX_SPAN
        If curPara.Range.Information(12) <> startInTable Then
            Exit Do
        End If
        curText = curPara.Range.Text
        curTrim = Trim$(Replace(Replace(Replace(curText, vbCr, ""), Chr(11), " "), Chr(7), ""))
        tmpScan = True
        If LCase$(curTrim) = "=contentend()" Or IsInNoFlyZoneCached(curPara.Range, curText, tmpScan) Then
            Exit Do
        End If
        combined = combined & " " & Replace(Replace(Replace(curText, vbCr, " "), Chr(11), " "), Chr(7), "")
        If blockKind = 3 Then
            depth = depth + ModMathInline.CountTokenSubstr(curText, beginToken) - ModMathInline.CountTokenSubstr(curText, endToken)
            If depth <= 0 Then
                If InStr(1, curText, endToken) > 0 Then
                    closedHere = True
                Else
                    Exit Function
                End If
            End If
        ElseIf (Left$(curTrim, 2) = "$$" And Len(curTrim) > 2 And ModMathInline.FindUnescapedToken(curTrim, 3, "$$") > 0) Or (Left$(curTrim, 2) = "\[" And Len(curTrim) > 2 And ModMathInline.FindUnescapedToken(curTrim, 3, "\]") > 0) Then
            Exit Do
        ElseIf ModMathInline.FindUnescapedToken(curText, 1, endToken) > 0 Then
            closedHere = True
        End If
        If closedHere Then
            Set outEndPara = curPara
            outRawLatex = combined
            outIsAmbiguous = False
            CheckMultilineBlock = True
            Exit Function
        End If
        If InStr(1, curText, Chr(7)) > 0 Then
            Exit Do
        End If
        Set curPara = curPara.Next
        spanStep = spanStep + 1
    Loop
End Function
Public Sub FindDollarPairsInPara(ByVal doc As Document, ByVal p As Paragraph, ByRef outStarts() As Long, ByRef outEnds() As Long, ByRef outPairCount As Long)
    ModMathHelpers.FindDollarPairsDom doc, p, outStarts, outEnds, outPairCount
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
