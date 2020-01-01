Attribute VB_Name = "ModMathScanner"
Option Explicit

' ==============================================================================
' Module: ModMathScanner
' Purpose: DOM-level coordinate scanning, Pandoc dollar rules, and No-Fly zone
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

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

Private g_HasTocZone As Boolean
Private g_HasLotZone As Boolean
Private g_HasLofZone As Boolean
Private g_BmToc As Bookmark
Private g_BmLot As Bookmark
Private g_BmLof As Bookmark

Public Sub InitNoFlyZoneCache(ByVal doc As Document)
    ResetNoFlyZoneCache
    On Error Resume Next
    If doc.Bookmarks.Exists("bm_zone_toc") Then
        Set g_BmToc = doc.Bookmarks("bm_zone_toc")
        g_HasTocZone = (Err.Number = 0 And Not g_BmToc Is Nothing)
        Err.Clear
    End If
    If doc.Bookmarks.Exists("bm_zone_lot") Then
        Set g_BmLot = doc.Bookmarks("bm_zone_lot")
        g_HasLotZone = (Err.Number = 0 And Not g_BmLot Is Nothing)
        Err.Clear
    End If
    If doc.Bookmarks.Exists("bm_zone_lof") Then
        Set g_BmLof = doc.Bookmarks("bm_zone_lof")
        g_HasLofZone = (Err.Number = 0 And Not g_BmLof Is Nothing)
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Sub ResetNoFlyZoneCache()
    g_HasTocZone = False
    g_HasLotZone = False
    g_HasLofZone = False
    Set g_BmToc = Nothing
    Set g_BmLot = Nothing
    Set g_BmLof = Nothing
End Sub

Public Function HasBalancedBrackets(ByVal txt As String) As Boolean
    Dim i As Long
    Dim braceDepth As Long
    Dim parenDepth As Long
    Dim c As String

    braceDepth = 0
    parenDepth = 0
    For i = 1 To Len(txt)
        c = Mid$(txt, i, 1)
        If c = "{" Then
            braceDepth = braceDepth + 1
        ElseIf c = "}" Then
            braceDepth = braceDepth - 1
            If braceDepth < 0 Then
                HasBalancedBrackets = False
                Exit Function
            End If
        ElseIf c = "(" Then
            parenDepth = parenDepth + 1
        ElseIf c = ")" Then
            parenDepth = parenDepth - 1
            If parenDepth < 0 Then
                HasBalancedBrackets = False
                Exit Function
            End If
        End If
    Next i
    HasBalancedBrackets = (braceDepth = 0 And parenDepth = 0)
End Function

Public Function IsCurrencyText(ByVal txt As String) As Boolean
    Dim t As String
    Dim firstC As String
    Dim c As String
    Dim i As Long
    Dim unitPart As String
    Dim slashPos As Long
    Dim leftTok As String
    Dim rightTok As String
    Dim numDone As Boolean
    Dim currUnits As String
    Dim periodUnits As String

    currUnits = "|usd|vnd|eur|gbp|jpy|cny|chf|k|m|b|tr|bn|dong|d|nghin|ngan|trieu|ty|dollar|dollars|cent|cents|thousand|million|billion|" & _
                ModConstants.Uni("tri\u1ec7u|t\u1ef7|ngh\u00ecn|ng\u00e0n|\u0111\u1ed3ng|\u0111|th\u00e1ng|n\u0103m|ng\u01b0\u1eddi|\u0111\u01a1n v\u1ecb|")
    periodUnits = "|month|mo|year|yr|day|unit|thang|nam|nguoi|don vi|" & ModConstants.Uni("th\u00e1ng|n\u0103m|ng\u01b0\u1eddi|\u0111\u01a1n v\u1ecb|")

    t = LCase$(Trim$(txt))
    If Len(t) = 0 Then
        IsCurrencyText = False
        Exit Function
    End If
    If InStr(1, t, "\") > 0 Or InStr(1, t, "=") > 0 Or InStr(1, t, "{") > 0 Or InStr(1, t, "}") > 0 Or InStr(1, t, "^") > 0 Or InStr(1, t, "_") > 0 Then
        IsCurrencyText = False
        Exit Function
    End If
    firstC = Left$(t, 1)
    If firstC < "0" Or firstC > "9" Then
        IsCurrencyText = False
        Exit Function
    End If
    numDone = False
    unitPart = ""
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If Not numDone Then
            If InStr(1, "0123456789., ", c) = 0 Then
                numDone = True
                unitPart = unitPart & c
            End If
        Else
            unitPart = unitPart & c
        End If
    Next i
    unitPart = Trim$(unitPart)
    If Left$(unitPart, 1) = "/" Then
        unitPart = Trim$(Mid$(unitPart, 2))
    End If
    If Len(unitPart) = 0 Then
        IsCurrencyText = False
        Exit Function
    End If
    If InStr(1, currUnits, "|" & unitPart & "|") > 0 Then
        IsCurrencyText = True
        Exit Function
    End If
    slashPos = InStr(1, unitPart, "/")
    If slashPos > 0 Then
        leftTok = Trim$(Left$(unitPart, slashPos - 1))
        rightTok = Trim$(Mid$(unitPart, slashPos + 1))
        If InStr(1, currUnits, "|" & leftTok & "|") > 0 Or InStr(1, periodUnits, "|" & rightTok & "|") > 0 Then
            IsCurrencyText = True
            Exit Function
        End If
    End If
    IsCurrencyText = False
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
        Exit Function
    End If
    If Not HasBalancedBrackets(t) Then
        ClassifyDollarSpan = 2
        Exit Function
    End If
    c1 = Left$(t, 1)
    cLast = Right$(t, 1)
    If c1 = " " Or c1 = vbTab Or cLast = " " Or cLast = vbTab Then
        If InStr(1, t, "\") = 0 And InStr(1, t, "_") = 0 And InStr(1, t, "^") = 0 And InStr(1, t, "=") = 0 Then
            ClassifyDollarSpan = 0
            Exit Function
        End If
    End If
    If Len(charAfterClose) > 0 Then
        If charAfterClose >= "0" And charAfterClose <= "9" Then
            ClassifyDollarSpan = 0
            Exit Function
        End If
    End If
    If IsCurrencyText(t) Then
        ClassifyDollarSpan = 0
        Exit Function
    End If
    hasMathProof = (InStr(1, t, "\") > 0 Or InStr(1, t, "=") > 0 Or InStr(1, t, "^") > 0 Or InStr(1, t, "_") > 0)
    hasNonDigits = False
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If InStr(1, "0123456789.,% ", c) = 0 Then
            hasNonDigits = True
            Exit For
        End If
    Next i
    If Not hasNonDigits And Len(Trim$(t)) > 0 Then
        ClassifyDollarSpan = 2
        Exit Function
    End If
    If (c1 >= "0" And c1 <= "9") And InStr(1, t, " ") > 0 And Not hasMathProof Then
        ClassifyDollarSpan = 2
        Exit Function
    End If
    ClassifyDollarSpan = 1
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
        If cleanHead = "references" Or cleanHead = "bibliography" Or cleanHead = "tai lieu tham khao" Or cleanHead = "danh muc tai lieu tham khao" Then
            isScanningActive = False
            CheckNoFlyHeading = True
            Exit Function
        ElseIf cleanHead = ModConstants.Uni("t\u00e0i li\u1ec7u tham kh\u1ea3o") Or cleanHead = ModConstants.Uni("danh m\u1ee5c t\u00e0i li\u1ec7u tham kh\u1ea3o") Then
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
    On Error Resume Next
    If g_HasTocZone Then
        If pRng.Start >= g_BmToc.Start And pRng.End <= g_BmToc.End Then
            IsInNoFlyZoneCached = True
            Exit Function
        End If
    End If
    If g_HasLotZone Then
        If pRng.Start >= g_BmLot.Start And pRng.End <= g_BmLot.End Then
            IsInNoFlyZoneCached = True
            Exit Function
        End If
    End If
    If g_HasLofZone Then
        If pRng.Start >= g_BmLof.Start And pRng.End <= g_BmLof.End Then
            IsInNoFlyZoneCached = True
            Exit Function
        End If
    End If
    On Error GoTo 0
    IsInNoFlyZoneCached = CheckNoFlyHeading(rawText, isScanningActive)
End Function

Public Function IsInNoFlyZone(ByVal doc As Document, ByVal pRng As Range, ByRef isScanningActive As Boolean) As Boolean
    InitNoFlyZoneCache doc
    IsInNoFlyZone = IsInNoFlyZoneCached(pRng, pRng.Text, isScanningActive)
End Function

Private Function CountTokenSubstr(ByVal txt As String, ByVal tok As String) As Long
    Dim cnt As Long
    Dim p As Long

    cnt = 0
    p = InStr(1, txt, tok)
    Do While p > 0
        cnt = cnt + 1
        p = InStr(p + Len(tok), txt, tok)
    Loop
    CountTokenSubstr = cnt
End Function

Public Function CheckMultilineBlock(ByVal doc As Document, ByVal startParaIdx As Long, ByRef outEndParaIdx As Long, ByRef outRawLatex As String, ByRef outIsAmbiguous As Boolean) As Boolean
    Dim p1Raw As String
    Dim p1Text As String
    Dim endToken As String
    Dim curText As String
    Dim curTrim As String
    Dim combined As String
    Dim maxSpan As Long
    Dim k As Long
    Dim braceClose As Long
    Dim blockKind As Long
    Dim depth As Long
    Dim foundEnd As Boolean

    p1Raw = doc.Paragraphs(startParaIdx).Range.Text
    p1Text = Trim$(Replace(Replace(Replace(p1Raw, vbCr, ""), Chr(11), ""), Chr(7), ""))
    blockKind = 0
    endToken = ""
    depth = 0
    If Left$(p1Text, 2) = "$$" And (Right$(p1Text, 2) <> "$$" Or Len(p1Text) = 2) Then
        blockKind = 1
        endToken = "$$"
    ElseIf Left$(p1Text, 2) = "\[" And (Right$(p1Text, 2) <> "\]" Or Len(p1Text) = 2) Then
        blockKind = 2
        endToken = "\]"
    ElseIf Left$(p1Text, 7) = "\begin{" Then
        braceClose = InStr(8, p1Text, "}")
        If braceClose > 8 Then
            endToken = "\end{" & Trim$(Mid$(p1Text, 8, braceClose - 8)) & "}"
            If InStr(braceClose + 1, p1Text, endToken) = 0 Then
                blockKind = 3
                depth = CountTokenSubstr(p1Text, "\begin{") - CountTokenSubstr(p1Text, "\end{")
            End If
        End If
    End If
    If blockKind = 0 Then
        CheckMultilineBlock = False
        Exit Function
    End If

    combined = Replace(Replace(p1Raw, vbCr, " "), Chr(11), " ")
    maxSpan = startParaIdx + MATH_MERGE_MAX_SPAN - 1
    If maxSpan > doc.Paragraphs.Count Then
        maxSpan = doc.Paragraphs.Count
    End If
    foundEnd = False
    For k = startParaIdx + 1 To maxSpan
        curText = doc.Paragraphs(k).Range.Text
        curTrim = Trim$(Replace(Replace(Replace(curText, vbCr, ""), Chr(11), ""), Chr(7), ""))
        If LCase$(curTrim) = "=contentend()" Then
            Exit For
        End If
        combined = combined & " " & Replace(Replace(curText, vbCr, " "), Chr(11), " ")
        If blockKind = 1 Or blockKind = 2 Then
            If (Left$(curTrim, 2) = "$$" And Right$(curTrim, 2) = "$$" And Len(curTrim) > 2) Or (Left$(curTrim, 2) = "\[" And Right$(curTrim, 2) = "\]" And Len(curTrim) > 2) Then
                Exit For
            ElseIf InStr(1, curText, endToken) > 0 Then
                foundEnd = True
                outEndParaIdx = k
                Exit For
            End If
        Else
            depth = depth + CountTokenSubstr(curText, "\begin{") - CountTokenSubstr(curText, "\end{")
            If depth <= 0 Then
                If InStr(1, curText, endToken) > 0 Or InStr(1, curText, "\end{") > 0 Then
                    foundEnd = True
                    outEndParaIdx = k
                End If
                Exit For
            End If
        End If
    Next k

    If foundEnd Then
        outRawLatex = combined
        outIsAmbiguous = False
        CheckMultilineBlock = True
    Else
        CheckMultilineBlock = False
    End If
End Function

Public Sub FindDollarPairsInPara(ByVal doc As Document, ByVal p As Paragraph, ByRef outStarts() As Long, ByRef outEnds() As Long, ByRef outPairCount As Long)
    ModMathHelpers.FindDollarPairsDom doc, p, outStarts, outEnds, outPairCount
End Sub
