Attribute VB_Name = "ModPatchMatcher"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModPatchMatcher
' Purpose: Normalization with inverse posMap, body/reference scanner, and Ref matching
' Architecture: Academic Document Format Engine (100% 7-bit ASCII, < 400 lines)
' ==============================================================================

Public Const OP_KIND_FIX As Long = 1
Public Const OP_KIND_CITE As Long = 2
Private Const MAX_HEADING_LEN As Long = 350

Public Type RefTargetInfo
    KeyId As String
    PrefixText As String
    BmName As String
    ParaIndex As Long
    RefArrIdx As Long
    ShortTip As String
    Found As Boolean
    IsAmbiguous As Boolean
    ErrReason As String
End Type

Public Type ScannedPara
    ParaIndex As Long
    RangeStart As Long
    RangeEnd As Long
    RawText As String
    NormText As String
    PosStart() As Long
    PosEnd() As Long
    NormLen As Long
    InBodyScope As Boolean
    IsLevel1 As Boolean
    IsHeading As Boolean
    IsInTable As Boolean
    IsInRefZone As Boolean
    IsExcluded As Boolean
End Type

Public Type PatchOp
    Kind As Long
    ParaIdxInArr As Long
    ParaIndex As Long
    StartPos As Long
    EndPos As Long
    RawTarget As String
    ReplaceText As String
    HasReplace As Boolean
    BmName As String
    ShortTip As String
    SrcLine As Long
    KeepExistingLink As Boolean
    IsAlreadySynced As Boolean
End Type

Public Sub SyncCitationLinkColors(ByVal doc As Document, ByVal linkColor As Long)
    If doc Is Nothing Then Exit Sub
    If doc.Hyperlinks.Count = 0 Then Exit Sub
    Dim hl As Hyperlink
    On Error Resume Next
    For Each hl In doc.Hyperlinks
        Err.Clear
        If hl.Type = 0 And hl.Address = "" And Left$(hl.SubAddress, 7) = BM_PREFIX_REF Then
            If hl.Range.StoryType = 1 And Err.Number = 0 Then
                hl.Range.Font.Color = linkColor
                hl.Range.Font.Underline = 0
            End If
        End If
    Next hl
    Err.Clear
End Sub

Public Function IsIdentChar(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    Dim c As Long
    c = AscW(ch) And &HFFFF&
    IsIdentChar = (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c = 95 Or (c >= &HC0 And c <= &H1EF9)
End Function

Public Sub UpdateRefZoneStateText(ByVal cleanTxt As String, ByRef inRefZone As Boolean)
    Static reStrip As Object
    If reStrip Is Nothing Then
        Set reStrip = CreateObject("VBScript.RegExp")
        reStrip.IgnoreCase = True
        reStrip.Global = False
        reStrip.Pattern = "^(?:(?:chapter|" & ModConstants.Uni("ch\u01b0\u01a1ng") & "|part|" & _
                          ModConstants.Uni("ph\u1ea7n") & ")\s+[a-z0-9]+\s*[:\.\-]?\s*|(?:\d+(?:\.\d+)*|[ivxlcdm]+)\s*[:\.\-]?\s+)"
    End If
    Dim low As String
    low = Trim$(reStrip.Replace(LCase$(Trim$(cleanTxt)), ""))
    If Left$(low, 9) = "reference" Or Left$(low, 12) = "bibliography" Or _
       Left$(low, 18) = ModConstants.Uni("t\u00e0i li\u1ec7u tham kh\u1ea3o") Or _
       Left$(low, 27) = ModConstants.Uni("danh m\u1ee5c t\u00e0i li\u1ec7u tham kh\u1ea3o") Then
        inRefZone = True
    ElseIf Left$(low, 8) = "appendix" Or Left$(low, 10) = "appendices" Or Left$(low, 8) = "glossary" Or _
           Left$(low, 8) = ModConstants.Uni("ph\u1ee5 l\u1ee5c") Then
        inRefZone = False
    End If
End Sub

Public Sub NormalizeWithPosMap(ByVal rawTxt As String, ByRef outNorm As String, _
                                ByRef outPosStart() As Long, ByRef outPosEnd() As Long, _
                                ByRef outLen As Long)
    Dim rawLen As Long, i As Long, code As Long, ch As String, mapped As String
    Dim buf As String, inSpace As Boolean
    rawLen = Len(rawTxt)
    outLen = 0
    outNorm = ""
    If rawLen = 0 Then Exit Sub
    ReDim outPosStart(1 To rawLen)
    ReDim outPosEnd(1 To rawLen)
    buf = Space$(rawLen)
    inSpace = False
    For i = 1 To rawLen
        ch = Mid$(rawTxt, i, 1)
        code = AscW(ch) And &HFFFF&
        If code = 13 Or code = 7 Or code = &HAD Then
            ' Strip paragraph mark, cell bell marker, and soft hyphen
        ElseIf code = 32 Or code = 9 Or code = 10 Or code = 11 Or code = 12 Or code = &HA0 Or code = &H202F Then
            If outLen > 0 Then
                If Not inSpace Then
                    outLen = outLen + 1
                    Mid$(buf, outLen, 1) = " "
                    outPosStart(outLen) = i
                    outPosEnd(outLen) = i
                    inSpace = True
                Else
                    outPosEnd(outLen) = i
                End If
            End If
        Else
            inSpace = False
            If code = &H2018 Or code = &H2019 Then
                mapped = "'"
            ElseIf code = &H201C Or code = &H201D Then
                mapped = """"
            ElseIf code = &H2013 Or code = &H2014 Or code = &H2212 Or code = &H2011 Then
                mapped = "-"
            Else
                mapped = ch
            End If
            outLen = outLen + 1
            Mid$(buf, outLen, 1) = mapped
            outPosStart(outLen) = i
            outPosEnd(outLen) = i
        End If
    Next i
    If outLen > 0 Then
        If Mid$(buf, outLen, 1) = " " Then outLen = outLen - 1
    End If
    If outLen > 0 Then outNorm = Left$(buf, outLen)
End Sub

Public Function NormalizeStr(ByVal s As String) As String
    Dim ps() As Long, pe() As Long, nLen As Long, norm As String
    NormalizeWithPosMap s, norm, ps, pe, nLen
    NormalizeStr = norm
End Function

Private Function IsMathBlockOnly(ByVal rng As Range, ByVal rawTxt As String) As Boolean
    On Error GoTo ErrDone
    If rng.OMaths.Count > 0 Then
        If rng.OMaths(1).Type = 0 Or rng.OMaths(1).Range.Text = rawTxt Then IsMathBlockOnly = True
    End If
ErrDone:
End Function

Private Function ExtractRefBmNum(ByVal bmName As String) As Long
    If Left$(bmName, 7) <> BM_PREFIX_REF Then Exit Function
    Dim tail As String, k As Long, ch As Long
    tail = Mid$(bmName, 8)
    If Len(tail) = 0 Then Exit Function
    For k = 1 To Len(tail)
        ch = AscW(Mid$(tail, k, 1))
        If ch < 48 Or ch > 57 Then Exit Function
    Next k
    ExtractRefBmNum = CLng(tail)
End Function

Public Function ScanDocForPatch(ByVal doc As Document, ByRef outMarkers As MarkerLocations, _
                                ByRef outParas() As ScannedPara, ByRef outCount As Long, _
                                ByRef outMaxRefNum As Long, ByRef outErr As String) As Boolean
    outCount = 0
    outMaxRefNum = 0
    outErr = ""
    If doc Is Nothing Then
        outErr = "No active document."
        Exit Function
    End If
    If Not ModMarkerLexer.ScanMarkers(doc, outMarkers, outErr) Then Exit Function
    If Not outMarkers.HasContentStart Then
        outErr = ModConstants.Uni("Missing =contentStart() marker. Insert =contentStart() before running AI Patch. (Thi\u1ebfu th\u1ebb =contentStart().)")
        Exit Function
    End If
    Dim bm As Bookmark, bmNum As Long
    For Each bm In doc.Bookmarks
        bmNum = ExtractRefBmNum(bm.Name)
        If bmNum > outMaxRefNum Then outMaxRefNum = bmNum
    Next bm
    ModTableIndex.BuildTableIndex doc
    Dim reL1 As Object, reL2 As Object, reL3 As Object, reL4 As Object
    Dim reTbl As Object, reFig As Object, reMark As Object
    Set reL1 = CreateObject("VBScript.RegExp")
    reL1.IgnoreCase = True
    reL1.Pattern = ModLocale.GetPatternLevel1()
    Set reL2 = CreateObject("VBScript.RegExp")
    reL2.IgnoreCase = True
    reL2.Pattern = ModConstants.GetPatternLevel2()
    Set reL3 = CreateObject("VBScript.RegExp")
    reL3.IgnoreCase = True
    reL3.Pattern = ModConstants.GetPatternLevel3()
    Set reL4 = CreateObject("VBScript.RegExp")
    reL4.IgnoreCase = True
    reL4.Pattern = ModConstants.GetPatternLevel4()
    Set reTbl = CreateObject("VBScript.RegExp")
    reTbl.IgnoreCase = True
    reTbl.Pattern = ModLocale.GetPatternTable()
    Set reFig = CreateObject("VBScript.RegExp")
    reFig.IgnoreCase = True
    reFig.Pattern = ModLocale.GetPatternFigure()
    Set reMark = CreateObject("VBScript.RegExp")
    reMark.IgnoreCase = True
    reMark.Pattern = ModConstants.GetPatternMarker()

    Dim startP As Long, endBodyP As Long, totalParas As Long
    startP = outMarkers.ContentStartParaIndex + 1
    totalParas = doc.Paragraphs.Count
    If outMarkers.HasContentEnd Then endBodyP = outMarkers.ContentEndParaIndex - 1 Else endBodyP = totalParas
    If startP > totalParas Then Exit Function
    Dim cap As Long
    cap = totalParas - startP + 1
    If cap < 8 Then cap = 8
    ReDim outParas(1 To cap)

    Dim docHasMath As Boolean, inRefZone As Boolean, p As Paragraph, pIdx As Long
    Dim rng As Range, rawTxt As String, normTxt As String, ps() As Long, pe() As Long, nLen As Long
    Dim inTbl As Boolean, isL1 As Boolean, isHead As Boolean, isExcl As Boolean, stName As String
    docHasMath = (doc.OMaths.Count > 0)
    inRefZone = False
    pIdx = 0
    For Each p In doc.Paragraphs
        pIdx = pIdx + 1
        If pIdx >= startP Then
            Set rng = p.Range
            rawTxt = rng.Text
            NormalizeWithPosMap rawTxt, normTxt, ps, pe, nLen
            If nLen > 0 And Not reMark.Test(normTxt) Then
                inTbl = ModTableIndex.IsRangeInTable(rng)
                isL1 = False
                isHead = False
                isExcl = False
                If docHasMath Then isExcl = IsMathBlockOnly(rng, rawTxt)
                If Not isExcl And Not inTbl And nLen <= MAX_HEADING_LEN Then
                    stName = ""
                    On Error Resume Next
                    stName = LCase$(rng.Style.NameLocal)
                    On Error GoTo 0
                    If reL1.Test(normTxt) Or ModLocale.IsRomanLevel1(normTxt) Or Left$(stName, 9) = "heading 1" Then isL1 = True
                    If isL1 Or reL2.Test(normTxt) Or reL3.Test(normTxt) Or reL4.Test(normTxt) Or reTbl.Test(normTxt) Or reFig.Test(normTxt) Then
                        isHead = True
                    End If
                    If isL1 Then UpdateRefZoneStateText normTxt, inRefZone
                End If
                outCount = outCount + 1
                outParas(outCount).ParaIndex = pIdx
                outParas(outCount).RangeStart = rng.Start
                outParas(outCount).RangeEnd = rng.End
                outParas(outCount).RawText = rawTxt
                outParas(outCount).NormText = normTxt
                outParas(outCount).PosStart = ps
                outParas(outCount).PosEnd = pe
                outParas(outCount).NormLen = nLen
                outParas(outCount).InBodyScope = (pIdx <= endBodyP)
                outParas(outCount).IsLevel1 = isL1
                outParas(outCount).IsHeading = isHead
                outParas(outCount).IsInTable = inTbl
                outParas(outCount).IsInRefZone = inRefZone
                outParas(outCount).IsExcluded = isExcl
            End If
        End If
    Next p
    ScanDocForPatch = True
End Function

Private Function FindExistingRefBmInPara(ByVal doc As Document, ByVal pStart As Long, ByVal pEnd As Long) As String
    Dim bm As Bookmark
    For Each bm In doc.Bookmarks
        If Left$(bm.Name, 7) = BM_PREFIX_REF Then
            If bm.Range.Start >= pStart And bm.Range.Start < pEnd Then
                FindExistingRefBmInPara = bm.Name
                Exit Function
            End If
        End If
    Next bm
End Function

Private Function BuildShortRefTip(ByVal normTxt As String) As String
    Dim s As String
    s = Trim$(normTxt)
    If Len(s) > 80 Then s = Left$(s, 77) & "..."
    BuildShortRefTip = s
End Function

Public Sub ResolveRefsTargets(ByVal doc As Document, ByRef rawRefs() As PatchRefItem, ByVal rawCount As Long, _
                              ByRef paras() As ScannedPara, ByVal pCount As Long, ByVal maxRefNum As Long, _
                              ByRef outTargets() As RefTargetInfo, ByVal outDict As Object, _
                              ByRef outMatched As Long, ByRef outFailed As Long)
    outMatched = 0
    outFailed = 0
    If rawCount <= 0 Then Exit Sub
    ReDim outTargets(1 To rawCount)
    Dim i As Long, j As Long, pfxNorm As String, pfxLen As Long
    Dim hitCount As Long, hitParaArrIdx As Long, nextNum As Long, existBm As String
    nextNum = maxRefNum
    For i = 1 To rawCount
        outTargets(i).KeyId = rawRefs(i).KeyId
        outTargets(i).PrefixText = rawRefs(i).PrefixText
        pfxNorm = NormalizeStr(rawRefs(i).PrefixText)
        pfxLen = Len(pfxNorm)
        hitCount = 0
        hitParaArrIdx = 0
        If pfxLen > 0 Then
            For j = 1 To pCount
                If paras(j).IsInRefZone And Not paras(j).IsHeading And Not paras(j).IsInTable And Not paras(j).IsExcluded Then
                    If paras(j).NormLen >= pfxLen Then
                        If StrComp(Left$(paras(j).NormText, pfxLen), pfxNorm, vbTextCompare) = 0 Then
                            hitCount = hitCount + 1
                            hitParaArrIdx = j
                        End If
                    End If
                End If
            Next j
            If hitCount = 0 And pfxLen >= 12 Then
                For j = 1 To pCount
                    If paras(j).IsInRefZone And Not paras(j).IsHeading And Not paras(j).IsInTable And Not paras(j).IsExcluded Then
                        If InStr(1, paras(j).NormText, pfxNorm, vbTextCompare) > 0 Then
                            hitCount = hitCount + 1
                            hitParaArrIdx = j
                        End If
                    End If
                Next j
            End If
        End If
        If outDict.Exists(rawRefs(i).KeyId) Then
            outTargets(i).Found = False
            outTargets(i).IsAmbiguous = True
            outTargets(i).ErrReason = "Duplicate @REFS key"
            outFailed = outFailed + 1
        ElseIf hitCount = 1 Then
            outTargets(i).Found = True
            outTargets(i).ParaIndex = paras(hitParaArrIdx).ParaIndex
            outTargets(i).RefArrIdx = hitParaArrIdx
            existBm = FindExistingRefBmInPara(doc, paras(hitParaArrIdx).RangeStart, paras(hitParaArrIdx).RangeEnd)
            If Len(existBm) > 0 Then
                outTargets(i).BmName = existBm
            Else
                nextNum = nextNum + 1
                outTargets(i).BmName = ModBookmark.FormatBookmarkId(BM_PREFIX_REF, nextNum)
            End If
            outTargets(i).ShortTip = BuildShortRefTip(paras(hitParaArrIdx).NormText)
            outDict.Add rawRefs(i).KeyId, i
            outMatched = outMatched + 1
        Else
            outTargets(i).Found = False
            outTargets(i).IsAmbiguous = (hitCount > 1)
            If hitCount = 0 Then outTargets(i).ErrReason = "Ref prefix not found" Else outTargets(i).ErrReason = "Ambiguous ref (" & hitCount & ")"
            outFailed = outFailed + 1
        End If
    Next i
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
