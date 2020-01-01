Attribute VB_Name = "ModCrossRef"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModCrossRef
' Purpose: Deterministic Cross-Reference Auto-Linking & Normalization
' Architecture: Academic Document Format Engine (v1.0.1)
' ==============================================================================

Public Function ResolveLinkColor(ByVal colorId As String) As Long
    ResolveLinkColor = RGB(43, 108, 176)
    If colorId = "xref_orange" Then ResolveLinkColor = RGB(197, 90, 17)
    If colorId = "xref_red" Then ResolveLinkColor = RGB(185, 28, 28)
    If colorId = "xref_black" Then ResolveLinkColor = -16777216
End Function

Private Function CanonicalKeyword(ByVal rawKw As String, ByRef outKind As String) As String
    Static d As Object
    If d Is Nothing Then
        Set d = CreateObject("Scripting.Dictionary")
        Dim pairs As Variant, j As Long, p() As String
        pairs = Array("table|TBL|Table", ModConstants.Uni("b\u1ea3ng") & "|TBL|" & ModConstants.Uni("B\u1ea3ng"), _
                      "bang|TBL|" & ModConstants.Uni("B\u1ea3ng"), "figure|FIG|Figure", "fig.|FIG|Fig.", "fig|FIG|Fig", _
                      "chart|FIG|Chart", "diagram|FIG|Diagram", ModConstants.Uni("h\u00ecnh") & "|FIG|" & ModConstants.Uni("H\u00ecnh"), _
                      "hinh|FIG|" & ModConstants.Uni("H\u00ecnh"), ModConstants.Uni("s\u01a1 \u0111\u1ed3") & "|FIG|" & ModConstants.Uni("S\u01a1 \u0111\u1ed3"), _
                      ModConstants.Uni("bi\u1ec3u \u0111\u1ed3") & "|FIG|" & ModConstants.Uni("Bi\u1ec3u \u0111\u1ed3"), _
                      ModConstants.Uni("\u0111\u1ed3 th\u1ecb") & "|FIG|" & ModConstants.Uni("\u0110\u1ed3 th\u1ecb"), _
                      "chapter|CHAP|Chapter", ModConstants.Uni("ch\u01b0\u01a1ng") & "|CHAP|" & ModConstants.Uni("Ch\u01b0\u01a1ng"), _
                      "appendix|APP|Appendix", ModConstants.Uni("ph\u1ee5 l\u1ee5c") & "|APP|" & ModConstants.Uni("Ph\u1ee5 l\u1ee5c"), _
                      "section|SEC|Section", ModConstants.Uni("m\u1ee5c") & "|SEC|" & ModConstants.Uni("M\u1ee5c"), _
                      ModConstants.Uni("ph\u1ea7n") & "|SEC|" & ModConstants.Uni("Ph\u1ea7n"))
        For j = LBound(pairs) To UBound(pairs)
            p = Split(CStr(pairs(j)), "|")
            d.Add p(0), p(1) & "|" & p(2)
        Next j
    End If
    Dim k As String, v() As String
    k = LCase$(Trim$(rawKw))
    If d.Exists(k) Then
        v = Split(CStr(d(k)), "|")
        outKind = v(0)
        CanonicalKeyword = v(1)
    End If
End Function

Private Sub RegisterTarget(ByVal dictBm As Object, ByVal dictTip As Object, ByVal key As String, ByVal bmName As String, ByVal tipText As String)
    If Len(key) = 0 Or dictBm.Exists(key) Then Exit Sub
    dictBm.Add key, bmName
    dictTip.Add key, Left$(tipText, 120)
End Sub

Private Function TryRegMatch(ByVal re As Object, ByVal pat As String, ByVal keyPfx As String, _
                             ByVal txt As String, ByVal bm As String, ByVal dictBm As Object, ByVal dictTip As Object) As Boolean
    re.Pattern = pat
    If re.Test(txt) Then
        RegisterTarget dictBm, dictTip, keyPfx & UCase$(Trim$(re.Execute(txt)(0).SubMatches(0))), bm, txt
        TryRegMatch = True
    End If
End Function

Private Sub BuildTargetIndex(ByVal doc As Document, ByRef snapshots() As ParagraphSnapshot, _
                             ByVal totalCount As Long, ByVal dictBm As Object, ByVal dictTip As Object)
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = False
    Dim patChap As String, patPart As String, patApp As String, patSec As String, patTbl As String, patFig As String
    patChap = "^(?:Chapter|" & ModConstants.Uni("Ch\u01b0\u01a1ng") & ")\s+([A-Za-z0-9]+)\s*[:\.]"
    patPart = "^" & ModConstants.Uni("Ph\u1ea7n") & "\s+([A-Za-z0-9]+)\s*[:\.]"
    patApp = "^(?:Appendix|" & ModConstants.Uni("Ph\u1ee5 l\u1ee5c") & ")(?:\s+([A-Za-z0-9]+))?(?:\s*[:\.]|$)"
    patSec = "^(\d+(?:\.\d+)+)\.\s+"
    patTbl = "^(?:Table|" & ModConstants.Uni("B(?:\u1ea3|a)ng") & ")\s+(" & ModLocale.NumToken() & ")" & ModLocale.CaptionSep()
    patFig = "^(?:" & ModLocale.GetFigureKeywords() & ")\s+(" & ModLocale.NumToken() & ")" & ModLocale.CaptionSep()
    Dim cntToc As Long, cntTbl As Long, cntFig As Long, i As Long, tType As Long, bm As String, txt As String, sepPos As Long
    For i = 1 To totalCount
        tType = snapshots(i).ElementType
        txt = snapshots(i).CleanText
        If tType >= ElementTypeLevel1 And tType <= ElementTypeLevel4 Then
            cntToc = cntToc + 1
            bm = ModBookmark.FormatBookmarkId(BM_PREFIX_TOC, cntToc)
            If doc.Bookmarks.Exists(bm) Then
                If tType > ElementTypeLevel1 Then
                    TryRegMatch re, patSec, "SEC|", txt, bm, dictBm, dictTip
                ElseIf Not TryRegMatch(re, patChap, "CHAP|", txt, bm, dictBm, dictTip) Then
                    If Not TryRegMatch(re, patPart, "SEC|", txt, bm, dictBm, dictTip) Then
                        If ModLocale.IsRomanLevel1(txt) Then
                            sepPos = InStr(txt, ".")
                            If sepPos = 0 Or (InStr(txt, ":") > 0 And InStr(txt, ":") < sepPos) Then sepPos = InStr(txt, ":")
                            If sepPos > 1 Then RegisterTarget dictBm, dictTip, "SEC|" & UCase$(Trim$(Left$(txt, sepPos - 1))), bm, txt
                        Else
                            TryRegMatch re, patApp, "APP|", txt, bm, dictBm, dictTip
                        End If
                    End If
                End If
            End If
        ElseIf tType = ElementTypeTableTitle Then
            cntTbl = cntTbl + 1
            bm = ModBookmark.FormatBookmarkId(BM_PREFIX_TBL, cntTbl)
            If doc.Bookmarks.Exists(bm) Then TryRegMatch re, patTbl, "TBL|", txt, bm, dictBm, dictTip
        ElseIf tType = ElementTypeFigureTitle Then
            cntFig = cntFig + 1
            bm = ModBookmark.FormatBookmarkId(BM_PREFIX_FIG, cntFig)
            If doc.Bookmarks.Exists(bm) Then TryRegMatch re, patFig, "FIG|", txt, bm, dictBm, dictTip
        End If
    Next i
End Sub

Private Function ClearOldCrossRefHyperlinks(ByVal doc As Document, ByVal bodyStartPos As Long) As Boolean
    If doc.Hyperlinks.Count = 0 Then Exit Function
    Dim toDel() As Hyperlink, delCount As Long, delCap As Long, hl As Hyperlink, pfx As String, i As Long
    delCap = 32
    ReDim toDel(1 To delCap)
    On Error Resume Next
    For Each hl In doc.Hyperlinks
        Err.Clear
        If hl.Type = 0 And hl.Address = "" Then
            pfx = Left$(hl.SubAddress, 7)
            If (pfx = BM_PREFIX_TOC Or pfx = BM_PREFIX_TBL Or pfx = BM_PREFIX_FIG) And Left$(hl.ScreenTip, 14) = "Ctrl+Click -> " Then
                If hl.Range.StoryType = 1 And hl.Range.Start >= bodyStartPos And Err.Number = 0 Then
                    delCount = delCount + 1
                    If delCount > delCap Then
                        delCap = delCap * 2
                        ReDim Preserve toDel(1 To delCap)
                    End If
                    Set toDel(delCount) = hl
                End If
            End If
        End If
    Next hl
    For i = delCount To 1 Step -1
        Err.Clear
        toDel(i).Range.Font.Color = -16777216
        toDel(i).Range.Font.Underline = 0
        toDel(i).Delete
        If Err.Number = 0 Then ClearOldCrossRefHyperlinks = True
    Next i
    Err.Clear
End Function

Private Function IsErroneousTrailingDot(ByVal fullTxt As String, ByVal dotIdx1 As Long) As Boolean
    If dotIdx1 > Len(fullTxt) Then Exit Function
    Dim ch As String, nextCh As String
    ch = Mid$(fullTxt, dotIdx1, 1)
    If InStr(1, ",;:)]", ch) > 0 Then
        IsErroneousTrailingDot = True
    ElseIf ch = " " Then
        nextCh = Mid$(fullTxt, dotIdx1 + 1, 1)
        If Len(nextCh) > 0 And nextCh <> vbCr And nextCh <> vbLf And nextCh = LCase$(nextCh) And nextCh <> UCase$(nextCh) Then IsErroneousTrailingDot = True
    End If
End Function

Private Sub ApplySingleCrossRef(ByVal doc As Document, ByVal matchRng As Range, _
                                ByVal normText As String, ByVal rawSlice As String, _
                                ByVal bmName As String, ByVal tipText As String, _
                                ByVal linkColor As Long, ByVal skipGuards As Boolean)
    On Error Resume Next
    If Not skipGuards Then
        If matchRng.Hyperlinks.Count > 0 Or matchRng.OMaths.Count > 0 Or matchRng.Fields.Count > 0 Then Exit Sub
        If matchRng.Text <> normText Then matchRng.Text = normText
    Else
        If StrComp(rawSlice, normText, vbBinaryCompare) <> 0 Then matchRng.Text = normText
    End If
    Dim hl As Hyperlink
    Set hl = doc.Hyperlinks.Add(Anchor:=matchRng, Address:="", SubAddress:=bmName, ScreenTip:="Ctrl+Click -> " & tipText)
    If Not hl Is Nothing Then
        hl.Range.Font.Underline = 0
        hl.Range.Font.Color = linkColor
    End If
    If Err.Number <> 0 Then ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
    Err.Clear
End Sub

Private Function IsUpperRomanToken(ByVal tok As String) As Boolean
    If Len(tok) = 0 Or tok <> UCase$(tok) Then Exit Function
    Dim k As Long
    For k = 1 To Len(tok)
        If InStr("IVX", Mid$(tok, k, 1)) = 0 Then Exit Function
    Next k
    IsUpperRomanToken = True
End Function

Private Function BuildCrossRefPattern() As String
    Dim kwNum As String, kwId As String, numTok As String
    kwNum = "Table|" & ModConstants.Uni("B(?:\u1ea3|a)ng") & "|Figure|Fig\.?|Chart|Diagram|" & ModConstants.Uni("H\u00ecnh|S\u01a1 \u0111\u1ed3|Bi\u1ec3u \u0111\u1ed3|\u0110\u1ed3 th\u1ecb")
    kwId = "Chapter|" & ModConstants.Uni("Ch\u01b0\u01a1ng") & "|Section|" & ModConstants.Uni("M\u1ee5c|Ph\u1ea7n")
    numTok = "(?=[A-Za-z0-9.\-_]*\d)[A-Za-z0-9]+(?:[\.\-_][A-Za-z0-9]+)*"
    BuildCrossRefPattern = ModConstants.Uni("(^|[^A-Za-z0-9\u00C0-\u1EF9_])") & "(?:(" & kwNum & ")\s+(" & numTok & ")(\.?)|(" & _
                           kwId & ")\s+([A-Za-z0-9]+(?:[\.\-_][A-Za-z0-9]+)*)(\.?)|(Appendix|" & ModConstants.Uni("Ph\u1ee5 l\u1ee5c") & _
                           ")(?:\s+([A-Za-z0-9]+))?(\.?))" & ModConstants.Uni("(?![A-Za-z0-9\u00C0-\u1EF9_])")
End Function

Private Function IsPrecededByDeterminer(ByVal txt As String, ByVal kwPos1 As Long) As Boolean
    If kwPos1 <= 2 Then Exit Function
    Dim prefix As String, trimmed As String, i As Long, ch As String
    prefix = Left$(txt, kwPos1 - 1)
    trimmed = RTrim$(prefix)
    If Len(trimmed) = Len(prefix) Or Len(trimmed) = 0 Then Exit Function
    For i = Len(trimmed) To 1 Step -1
        ch = Mid$(trimmed, i, 1)
        If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z")) Then Exit For
    Next i
    Select Case LCase$(Mid$(trimmed, i + 1))
        Case "this", "that", "the", "each", "every", "my", "our", "any", "one", "which"
            IsPrecededByDeterminer = True
    End Select
End Function

Private Function IsIdentChar(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    Dim c As Long
    c = AscW(ch)
    IsIdentChar = (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c = 95 Or (c >= &HC0 And c <= &H1EF9)
End Function

Private Function IsCrossRefBoundaryOk(ByVal doc As Document, ByVal mRng As Range, ByVal pRng As Range) As Boolean
    Dim chR As String
    If mRng.Start > pRng.Start Then If IsIdentChar(doc.Range(mRng.Start - 1, mRng.Start).Text) Then Exit Function
    If mRng.End < pRng.End Then
        chR = doc.Range(mRng.End, mRng.End + 1).Text
        If IsIdentChar(chR) Then Exit Function
        If (chR = "." Or chR = "-" Or chR = "_") And mRng.End + 1 < pRng.End Then
            If IsIdentChar(doc.Range(mRng.End + 1, mRng.End + 2).Text) Then Exit Function
        End If
    End If
    IsCrossRefBoundaryOk = True
End Function

Private Function ResolveMatchTarget(ByVal m As Object, ByVal txt As String, _
                                    ByVal dictBm As Object, ByVal dictTip As Object, _
                                    ByRef outBm As String, ByRef outTip As String, _
                                    ByRef outNorm As String, ByRef outLen As Long) As Boolean
    Dim off As Long, rawKw As String, rawId As String, rawDot As String
    off = 1
    If Len(m.SubMatches(4)) > 0 Then off = 4
    If Len(m.SubMatches(7)) > 0 Then off = 7
    rawKw = m.SubMatches(off)
    rawId = m.SubMatches(off + 1)
    rawDot = m.SubMatches(off + 2)
    Dim kind As String, canKw As String, leadLen As Long, firstCh As String
    canKw = CanonicalKeyword(rawKw, kind)
    If Len(canKw) = 0 Then Exit Function
    If kind = "SEC" And Not (rawId Like "*[0-9]*") Then
        If Not IsUpperRomanToken(rawId) Then Exit Function
    End If
    leadLen = Len(m.SubMatches(0))
    If (kind = "APP" Or kind = "CHAP") And Len(rawId) = 1 And Not IsNumeric(rawId) Then
        If rawId <> UCase$(rawId) And (kind = "CHAP" Or IsPrecededByDeterminer(txt, m.FirstIndex + 1 + leadLen)) Then Exit Function
    End If
    If (kind = "CHAP" Or kind = "SEC") And rawId = "I" Then
        firstCh = Left$(rawKw, 1)
        If firstCh <> UCase$(firstCh) Or firstCh = LCase$(firstCh) Then Exit Function
        If IsPrecededByDeterminer(txt, m.FirstIndex + 1 + leadLen) Then Exit Function
    End If
    Dim key As String
    key = kind & "|" & UCase$(rawId)
    If kind = "APP" And Len(rawId) > 1 And Not dictBm.Exists(key) Then
        If dictBm.Exists("APP|") And rawId = LCase$(rawId) And Not IsNumeric(Left$(rawId, 1)) Then
            key = "APP|"
            rawId = ""
            rawDot = ""
        End If
    End If
    If Not dictBm.Exists(key) Then Exit Function
    outBm = dictBm(key)
    outTip = dictTip(key)
    Dim stripDot As Boolean
    If rawDot = "." Then stripDot = IsErroneousTrailingDot(txt, m.FirstIndex + m.Length + 1)
    If Len(rawId) > 0 Then
        Dim idOut As String
        If kind = "CHAP" Or kind = "APP" Then idOut = UCase$(rawId) Else idOut = rawId
        outNorm = canKw & " " & idOut
        outLen = m.Length - leadLen
        If rawDot = "." And Not stripDot Then outLen = outLen - 1
    Else
        outNorm = canKw
        outLen = Len(rawKw)
    End If
    ResolveMatchTarget = True
End Function

Public Sub ApplyCrossReferences(ByVal doc As Document, ByRef markers As MarkerLocations, _
                                ByRef snapshots() As ParagraphSnapshot, ByVal totalCount As Long, _
                                ByVal colorId As String)
    If Not markers.HasContentStart Or totalCount < 1 Then Exit Sub
    Dim startP As Long, bodyStartPos As Long, didClear As Boolean
    startP = markers.ContentStartParaIndex + 1
    If startP > doc.Paragraphs.Count Then Exit Sub
    If startP <= totalCount Then bodyStartPos = snapshots(startP).RangeStart Else bodyStartPos = doc.Paragraphs(startP).Range.Start
    If colorId = "xref_off" Then GoTo ClearOnly
    Dim dictBm As Object, dictTip As Object
    Set dictBm = CreateObject("Scripting.Dictionary")
    Set dictTip = CreateObject("Scripting.Dictionary")
    BuildTargetIndex doc, snapshots, totalCount, dictBm, dictTip
    If dictBm.Count = 0 Then
ClearOnly:
        ClearOldCrossRefHyperlinks doc, bodyStartPos
        If colorId <> "xref_off" Then ModPatchMatcher.SyncCitationLinkColors doc, ResolveLinkColor(colorId)
        Exit Sub
    End If
    Dim refFlags() As Boolean, inRefZone As Boolean, sIdx As Long, tType As Long, refCls As Long
    ReDim refFlags(1 To totalCount)
    For sIdx = 1 To totalCount
        If snapshots(sIdx).ElementType = ElementTypeLevel1 Then
            refCls = ModLocale.ClassifyRefZoneHeading(snapshots(sIdx).CleanText)
            If refCls = 1 Then
                inRefZone = True
            ElseIf refCls = -1 Then
                inRefZone = False
            End If
        End If
        refFlags(sIdx) = inRefZone
    Next sIdx
    Dim linkColor As Long, docHasMath As Boolean, reMain As Object
    linkColor = ResolveLinkColor(colorId)
    docHasMath = (doc.OMaths.Count > 0)
    Set reMain = CreateObject("VBScript.RegExp")
    reMain.IgnoreCase = True
    reMain.Global = True
    reMain.Pattern = BuildCrossRefPattern()
    Dim pRngs() As Range
    ReDim pRngs(1 To totalCount)
    For sIdx = 1 To totalCount
        tType = snapshots(sIdx).ElementType
        If snapshots(sIdx).ParaIndex >= startP And Not refFlags(sIdx) And Not snapshots(sIdx).IsExcluded And _
           (tType = ElementTypeNormalParagraph Or tType = ElementTypeSourceNote Or tType = ElementTypeTableCell Or tType = ElementTypeUnknown) And _
           Len(snapshots(sIdx).CleanText) > 5 Then
            If reMain.Test(snapshots(sIdx).CleanText) Then Set pRngs(sIdx) = doc.Range(snapshots(sIdx).RangeStart, snapshots(sIdx).RangeEnd)
        End If
    Next sIdx
    didClear = ClearOldCrossRefHyperlinks(doc, bodyStartPos)
    Dim pRng As Range, txt As String, pStart As Long, pEnd As Long, matches As Object, mIdx As Long, m As Object
    Dim bmName As String, tipText As String, normStr As String, spanLen As Long
    Dim canDirectOffset As Boolean, mRng As Range, rawSub As String, searchPos As Long, prevPos As Long
    For sIdx = totalCount To 1 Step -1
        If Not pRngs(sIdx) Is Nothing Then
            If didClear Then
                Set pRng = pRngs(sIdx)
                txt = pRng.Text
                pStart = pRng.Start
                pEnd = pRng.End
            Else
                txt = snapshots(sIdx).RawText
                pStart = snapshots(sIdx).RangeStart
                pEnd = snapshots(sIdx).RangeEnd
            End If
            canDirectOffset = ((pEnd - pStart) = Len(txt))
            If canDirectOffset And docHasMath Then If doc.Range(pStart, pEnd).OMaths.Count > 0 Then canDirectOffset = False
            Set matches = reMain.Execute(txt)
            If Not canDirectOffset And Not didClear Then Set pRng = doc.Range(pStart, pEnd)
            For mIdx = matches.Count - 1 To 0 Step -1
                Set m = matches(mIdx)
                If ResolveMatchTarget(m, txt, dictBm, dictTip, bmName, tipText, normStr, spanLen) Then
                    rawSub = Mid$(txt, m.FirstIndex + 1 + Len(m.SubMatches(0)), spanLen)
                    If canDirectOffset Then
                        Set mRng = doc.Range(pStart + m.FirstIndex + Len(m.SubMatches(0)), pStart + m.FirstIndex + Len(m.SubMatches(0)) + spanLen)
                        ApplySingleCrossRef doc, mRng, normStr, rawSub, bmName, tipText, linkColor, True
                    Else
                        searchPos = pRng.Start
                        Set mRng = pRng.Duplicate
                        Do While searchPos < pRng.End
                            mRng.SetRange searchPos, pRng.End
                            mRng.Find.ClearFormatting
                            If Not mRng.Find.Execute(FindText:=rawSub, MatchCase:=False, MatchWholeWord:=False, _
                                                     MatchWildcards:=False, Forward:=True, Wrap:=0, Format:=False) Then Exit Do
                            If mRng.End > pRng.End Then Exit Do
                            prevPos = searchPos
                            searchPos = mRng.End
                            If searchPos <= prevPos Then searchPos = prevPos + 1
                            If IsCrossRefBoundaryOk(doc, mRng, pRng) And mRng.OMaths.Count = 0 And mRng.Fields.Count = 0 And mRng.Hyperlinks.Count = 0 Then
                                ApplySingleCrossRef doc, mRng, normStr, rawSub, bmName, tipText, linkColor, False
                                If mRng.End > prevPos Then searchPos = mRng.End
                            End If
                        Loop
                    End If
                End If
            Next mIdx
        End If
    Next sIdx
    ModPatchMatcher.SyncCitationLinkColors doc, ResolveLinkColor(colorId)
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
