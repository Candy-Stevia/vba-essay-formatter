Attribute VB_Name = "ModCrossRef"
Option Explicit

' ==============================================================================
' Module: ModCrossRef
' Purpose: Deterministic Cross-Reference Auto-Linking & Normalization
' Architecture: Academic Document Format Engine (v1.0.1)
' ==============================================================================

Private Function ResolveLinkColor(ByVal colorId As String) As Long
    Select Case colorId
        Case "xref_orange"
            ResolveLinkColor = RGB(197, 90, 17)
        Case "xref_red"
            ResolveLinkColor = RGB(185, 28, 28)
        Case "xref_black"
            ResolveLinkColor = -16777216 ' wdColorAutomatic
        Case Else
            ResolveLinkColor = RGB(43, 108, 176) ' Soft Academic Blue
    End Select
End Function

Private Function CanonicalKeyword(ByVal rawKw As String, ByRef outKind As String) As String
    Dim k As String
    k = LCase$(Trim$(rawKw))
    Select Case k
        Case "table"
            outKind = "TBL"
            CanonicalKeyword = "Table"
        Case ModConstants.Uni("b\u1ea3ng"), "bang"
            outKind = "TBL"
            CanonicalKeyword = ModConstants.Uni("B\u1ea3ng")
        Case "figure"
            outKind = "FIG"
            CanonicalKeyword = "Figure"
        Case "fig.", "fig"
            outKind = "FIG"
            If Right$(k, 1) = "." Then CanonicalKeyword = "Fig." Else CanonicalKeyword = "Fig"
        Case "chart"
            outKind = "FIG"
            CanonicalKeyword = "Chart"
        Case "diagram"
            outKind = "FIG"
            CanonicalKeyword = "Diagram"
        Case ModConstants.Uni("h\u00ecnh"), "hinh"
            outKind = "FIG"
            CanonicalKeyword = ModConstants.Uni("H\u00ecnh")
        Case ModConstants.Uni("s\u01a1 \u0111\u1ed3")
            outKind = "FIG"
            CanonicalKeyword = ModConstants.Uni("S\u01a1 \u0111\u1ed3")
        Case ModConstants.Uni("bi\u1ec3u \u0111\u1ed3")
            outKind = "FIG"
            CanonicalKeyword = ModConstants.Uni("Bi\u1ec3u \u0111\u1ed3")
        Case ModConstants.Uni("\u0111\u1ed3 th\u1ecb")
            outKind = "FIG"
            CanonicalKeyword = ModConstants.Uni("\u0110\u1ed3 th\u1ecb")
        Case "chapter"
            outKind = "CHAP"
            CanonicalKeyword = "Chapter"
        Case ModConstants.Uni("ch\u01b0\u01a1ng")
            outKind = "CHAP"
            CanonicalKeyword = ModConstants.Uni("Ch\u01b0\u01a1ng")
        Case "appendix"
            outKind = "APP"
            CanonicalKeyword = "Appendix"
        Case ModConstants.Uni("ph\u1ee5 l\u1ee5c")
            outKind = "APP"
            CanonicalKeyword = ModConstants.Uni("Ph\u1ee5 l\u1ee5c")
        Case "section"
            outKind = "SEC"
            CanonicalKeyword = "Section"
        Case ModConstants.Uni("m\u1ee5c")
            outKind = "SEC"
            CanonicalKeyword = ModConstants.Uni("M\u1ee5c")
        Case ModConstants.Uni("ph\u1ea7n")
            outKind = "SEC"
            CanonicalKeyword = ModConstants.Uni("Ph\u1ea7n")
    End Select
End Function

Private Sub RegisterTarget(ByVal dictBm As Object, ByVal dictTip As Object, _
                           ByVal key As String, ByVal bmName As String, ByVal tipText As String)
    If Len(key) = 0 Then Exit Sub
    If Not dictBm.Exists(key) Then
        dictBm.Add key, bmName
        dictTip.Add key, Left$(tipText, 120)
    End If
End Sub

Private Sub BuildTargetIndex(ByVal doc As Document, ByRef snapshots() As ParagraphSnapshot, _
                             ByVal totalCount As Long, ByVal dictBm As Object, ByVal dictTip As Object)
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = False

    Dim patChap As String, patApp As String, patSec As String, patTbl As String, patFig As String
    patChap = "^(?:Chapter|" & ModConstants.Uni("Ch\u01b0\u01a1ng") & ")\s+([A-Za-z0-9]+)\s*[:\.]"
    patApp = "^(?:Appendix|" & ModConstants.Uni("Ph\u1ee5 l\u1ee5c") & ")(?:\s+([A-Za-z0-9]+))?(?:\s*[:\.]|$)"
    patSec = "^(\d+(?:\.\d+)+)\.\s+"
    patTbl = "^(?:Table|" & ModConstants.Uni("B(?:\u1ea3|a)ng") & ")\s+(" & ModLocale.NumToken() & ")" & ModLocale.CaptionSep()
    patFig = "^(?:" & ModLocale.GetFigureKeywords() & ")\s+(" & ModLocale.NumToken() & ")" & ModLocale.CaptionSep()

    Dim cntToc As Long, cntTbl As Long, cntFig As Long
    Dim i As Long, tType As Long, bm As String, txt As String, m As Object
    For i = 1 To totalCount
        tType = snapshots(i).ElementType
        txt = snapshots(i).CleanText
        If tType = ElementTypeLevel1 Then
            cntToc = cntToc + 1
            bm = ModBookmark.FormatBookmarkId(BM_PREFIX_TOC, cntToc)
            If doc.Bookmarks.Exists(bm) Then
                re.Pattern = patChap
                If re.Test(txt) Then
                    Set m = re.Execute(txt)(0)
                    RegisterTarget dictBm, dictTip, "CHAP|" & UCase$(m.SubMatches(0)), bm, txt
                Else
                    re.Pattern = patApp
                    If re.Test(txt) Then
                        Set m = re.Execute(txt)(0)
                        RegisterTarget dictBm, dictTip, "APP|" & UCase$(Trim$(m.SubMatches(0))), bm, txt
                    End If
                End If
            End If
        ElseIf tType >= ElementTypeLevel2 And tType <= ElementTypeLevel4 Then
            cntToc = cntToc + 1
            bm = ModBookmark.FormatBookmarkId(BM_PREFIX_TOC, cntToc)
            If doc.Bookmarks.Exists(bm) Then
                re.Pattern = patSec
                If re.Test(txt) Then
                    Set m = re.Execute(txt)(0)
                    RegisterTarget dictBm, dictTip, "SEC|" & UCase$(m.SubMatches(0)), bm, txt
                End If
            End If
        ElseIf tType = ElementTypeTableTitle Then
            cntTbl = cntTbl + 1
            bm = ModBookmark.FormatBookmarkId(BM_PREFIX_TBL, cntTbl)
            If doc.Bookmarks.Exists(bm) Then
                re.Pattern = patTbl
                If re.Test(txt) Then
                    Set m = re.Execute(txt)(0)
                    RegisterTarget dictBm, dictTip, "TBL|" & UCase$(m.SubMatches(0)), bm, txt
                End If
            End If
        ElseIf tType = ElementTypeFigureTitle Then
            cntFig = cntFig + 1
            bm = ModBookmark.FormatBookmarkId(BM_PREFIX_FIG, cntFig)
            If doc.Bookmarks.Exists(bm) Then
                re.Pattern = patFig
                If re.Test(txt) Then
                    Set m = re.Execute(txt)(0)
                    RegisterTarget dictBm, dictTip, "FIG|" & UCase$(m.SubMatches(0)), bm, txt
                End If
            End If
        End If
    Next i
End Sub

Private Sub ClearOldCrossRefHyperlinks(ByVal doc As Document, ByVal bodyStartPos As Long)
    If doc.Hyperlinks.Count = 0 Then Exit Sub
    Dim i As Long, hl As Hyperlink, subA As String, pfx As String
    On Error Resume Next
    For i = doc.Hyperlinks.Count To 1 Step -1
        Set hl = doc.Hyperlinks(i)
        If hl.Address = "" Then
            subA = hl.SubAddress
            pfx = Left$(subA, 7)
            If pfx = BM_PREFIX_TOC Or pfx = BM_PREFIX_TBL Or pfx = BM_PREFIX_FIG Then
                If hl.Range.Start >= bodyStartPos Then
                    hl.Range.Font.Color = -16777216
                    hl.Range.Font.Underline = 0
                    hl.Delete
                End If
            End If
        End If
    Next i
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function IsErroneousTrailingDot(ByVal fullTxt As String, ByVal dotIdx1 As Long) As Boolean
    ' Returns True only when trailing '.' is followed by space+lowercase or punctuation (, ; : ) ])
    If dotIdx1 > Len(fullTxt) Then Exit Function
    Dim ch As String
    ch = Mid$(fullTxt, dotIdx1, 1)
    If ch = "," Or ch = ";" Or ch = ":" Or ch = ")" Or ch = "]" Then
        IsErroneousTrailingDot = True
        Exit Function
    End If
    If ch = " " Then
        Dim nextCh As String
        nextCh = Mid$(fullTxt, dotIdx1 + 1, 1)
        If Len(nextCh) > 0 And nextCh <> vbCr And nextCh <> vbLf Then
            If nextCh = LCase$(nextCh) And nextCh <> UCase$(nextCh) Then
                IsErroneousTrailingDot = True
            End If
        End If
    End If
End Function

Private Sub UpdateRefZoneState(ByVal cleanTxt As String, ByRef inRefZone As Boolean)
    Dim low As String
    low = LCase$(cleanTxt)
    If Left$(low, 10) = "references" Or Left$(low, 12) = "bibliography" Or _
       Left$(low, 18) = ModConstants.Uni("t\u00e0i li\u1ec7u tham kh\u1ea3o") Or _
       Left$(low, 27) = ModConstants.Uni("danh m\u1ee5c t\u00e0i li\u1ec7u tham kh\u1ea3o") Then
        inRefZone = True
    ElseIf Left$(low, 8) = "appendix" Or Left$(low, 8) = ModConstants.Uni("ph\u1ee5 l\u1ee5c") Then
        inRefZone = False
    End If
End Sub

Private Sub ApplySingleCrossRef(ByVal doc As Document, ByVal matchRng As Range, _
                                ByVal normText As String, ByVal bmName As String, _
                                ByVal tipText As String, ByVal linkColor As Long)
    On Error Resume Next
    If matchRng.Hyperlinks.Count > 0 Or matchRng.OMaths.Count > 0 Or matchRng.Fields.Count > 0 Then Exit Sub
    If matchRng.Text <> normText Then
        matchRng.Text = normText
    End If
    Dim hl As Hyperlink
    Set hl = doc.Hyperlinks.Add(Anchor:=matchRng, Address:="", SubAddress:=bmName, _
                                ScreenTip:="Ctrl+Click -> " & tipText)
    If Not hl Is Nothing Then
        hl.Range.Font.Underline = 0
        hl.Range.Font.Color = linkColor
    End If
    If Err.Number <> 0 Then
        ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
        Err.Clear
    End If
End Sub

Private Function BuildCrossRefPattern() As String
    Dim kwNum As String, kwId As String, kwApp As String
    kwNum = "Table|" & ModConstants.Uni("B(?:\u1ea3|a)ng") & "|Figure|Fig\.?|Chart|Diagram|" & _
            ModConstants.Uni("H\u00ecnh|S\u01a1 \u0111\u1ed3|Bi\u1ec3u \u0111\u1ed3|\u0110\u1ed3 th\u1ecb") & _
            "|Section|" & ModConstants.Uni("M\u1ee5c|Ph\u1ea7n")
    kwId = "Chapter|" & ModConstants.Uni("Ch\u01b0\u01a1ng")
    kwApp = "Appendix|" & ModConstants.Uni("Ph\u1ee5 l\u1ee5c")
    Dim numTok As String
    numTok = "(?=[A-Za-z0-9.\-_]*\d)[A-Za-z0-9]+(?:[\.\-_][A-Za-z0-9]+)*"
    Dim pfxBound As String
    pfxBound = ModConstants.Uni("(^|[^A-Za-z0-9\u00C0-\u1EF9_])")
    Dim altNum As String, altId As String, altApp As String
    altNum = "(" & kwNum & ")\s+(" & numTok & ")(\.?)"
    altId = "(" & kwId & ")\s+([A-Za-z0-9]+)(\.?)"
    altApp = "(" & kwApp & ")(?:\s+([A-Za-z0-9]+))?(\.?)"
    BuildCrossRefPattern = pfxBound & "(?:" & altNum & "|" & altId & "|" & altApp & ")" & ModConstants.Uni("(?![A-Za-z0-9\u00C0-\u1EF9_])")
End Function

Private Function ResolveMatchTarget(ByVal m As Object, ByVal txt As String, _
                                    ByVal dictBm As Object, ByVal dictTip As Object, _
                                    ByRef outBm As String, ByRef outTip As String, _
                                    ByRef outNorm As String, ByRef outLen As Long) As Boolean
    Dim rawKw As String, rawId As String, rawDot As String
    If Len(m.SubMatches(1)) > 0 Then
        rawKw = m.SubMatches(1)
        rawId = m.SubMatches(2)
        rawDot = m.SubMatches(3)
    ElseIf Len(m.SubMatches(4)) > 0 Then
        rawKw = m.SubMatches(4)
        rawId = m.SubMatches(5)
        rawDot = m.SubMatches(6)
    Else
        rawKw = m.SubMatches(7)
        rawId = m.SubMatches(8)
        rawDot = m.SubMatches(9)
    End If

    Dim kind As String, canKw As String
    canKw = CanonicalKeyword(rawKw, kind)
    If Len(canKw) = 0 Then Exit Function

    Dim key As String
    key = kind & "|" & UCase$(rawId)
    If kind = "APP" And Len(rawId) > 0 And Not dictBm.Exists(key) Then
        If dictBm.Exists("APP|") Then
            ' Standalone Appendix/Phu luc followed by lowercase prose word (>1 chars, non-numeric)
            If (rawId = LCase$(rawId)) And (Not IsNumeric(Left$(rawId, 1))) And (Len(rawId) > 1) Then
                key = "APP|"
                rawId = ""
                rawDot = ""
            End If
        End If
    End If
    If Not dictBm.Exists(key) Then Exit Function

    outBm = dictBm(key)
    outTip = dictTip(key)
    Dim leadLen As Long
    leadLen = Len(m.SubMatches(0))
    Dim matchEnd1 As Long
    matchEnd1 = m.FirstIndex + m.Length + 1
    Dim stripDot As Boolean
    If rawDot = "." Then
        stripDot = IsErroneousTrailingDot(txt, matchEnd1)
    End If

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
    If Not markers.HasContentStart Then Exit Sub
    Dim startP As Long
    startP = markers.ContentStartParaIndex + 1
    If startP > doc.Paragraphs.Count Then Exit Sub

    Dim bodyStartPos As Long
    bodyStartPos = doc.Paragraphs(startP).Range.Start
    ClearOldCrossRefHyperlinks doc, bodyStartPos
    If colorId = "xref_off" Then Exit Sub

    Dim dictBm As Object, dictTip As Object
    Set dictBm = CreateObject("Scripting.Dictionary")
    Set dictTip = CreateObject("Scripting.Dictionary")
    BuildTargetIndex doc, snapshots, totalCount, dictBm, dictTip
    If dictBm.Count = 0 Then Exit Sub

    Dim linkColor As Long
    linkColor = ResolveLinkColor(colorId)

    Dim reMain As Object
    Set reMain = CreateObject("VBScript.RegExp")
    reMain.IgnoreCase = True
    reMain.Global = True
    reMain.Pattern = BuildCrossRefPattern()

    Dim sIdx As Long, tType As Long, pRng As Range, txt As String, pStart As Long, pEnd As Long
    Dim matches As Object, mIdx As Long, m As Object
    Dim bmName As String, tipText As String, normStr As String, spanLen As Long
    Dim inRefZone As Boolean, canDirectOffset As Boolean, mRng As Range, rawSub As String
    inRefZone = False

    For sIdx = 1 To totalCount
        If snapshots(sIdx).ParaIndex >= startP Then
            tType = snapshots(sIdx).ElementType
            If tType = ElementTypeLevel1 Then
                UpdateRefZoneState snapshots(sIdx).CleanText, inRefZone
            ElseIf (Not inRefZone) And (Not snapshots(sIdx).IsExcluded) Then
                If tType = ElementTypeNormalParagraph Or tType = ElementTypeSourceNote Or _
                   tType = ElementTypeTableCell Or tType = ElementTypeUnknown Then
                    If Len(snapshots(sIdx).CleanText) > 5 Then
                        If reMain.Test(snapshots(sIdx).CleanText) Then
                            Set pRng = doc.Paragraphs(snapshots(sIdx).ParaIndex).Range
                            txt = pRng.Text
                            pStart = pRng.Start
                            pEnd = pRng.End
                            canDirectOffset = ((pEnd - pStart) = Len(txt)) And _
                                              (pRng.OMaths.Count = 0) And (pRng.Fields.Count = 0)
                            Set matches = reMain.Execute(txt)
                            For mIdx = matches.Count - 1 To 0 Step -1
                                Set m = matches(mIdx)
                                If ResolveMatchTarget(m, txt, dictBm, dictTip, bmName, tipText, normStr, spanLen) Then
                                    If canDirectOffset Then
                                        Set mRng = doc.Range(pStart + m.FirstIndex + Len(m.SubMatches(0)), _
                                                             pStart + m.FirstIndex + Len(m.SubMatches(0)) + spanLen)
                                        ApplySingleCrossRef doc, mRng, normStr, bmName, tipText, linkColor
                                    Else
                                        rawSub = Mid$(txt, m.FirstIndex + 1 + Len(m.SubMatches(0)), spanLen)
                                        Set mRng = pRng.Duplicate
                                        With mRng.Find
                                            .ClearFormatting
                                            .Text = rawSub
                                            .MatchWildcards = False
                                            .MatchCase = False
                                            .Forward = True
                                            .Wrap = 0
                                            Do While .Execute
                                                If mRng.OMaths.Count = 0 And mRng.Fields.Count = 0 And mRng.Hyperlinks.Count = 0 Then
                                                    ApplySingleCrossRef doc, mRng, normStr, bmName, tipText, linkColor
                                                End If
                                                mRng.Collapse 0
                                            Loop
                                        End With
                                    End If
                                End If
                            Next mIdx
                        End If
                    End If
                End If
            End If
        End If
    Next sIdx
End Sub
