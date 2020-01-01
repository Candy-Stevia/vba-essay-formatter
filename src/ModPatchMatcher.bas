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
Public Const OP_KIND_STRUCT As Long = 3
Private Const MAX_HEADING_LEN As Long = 350
Private mRefBmDocPtr As LongPtr
Private mRefBmCount As Long
Private mRefBmNames() As String
Private mRefBmStarts() As Long
Public g_LastInferredBodyStart As Long
Public g_LastHasExplicitMarkers As Boolean

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

Public Sub ResetRefBmCache()
    mRefBmDocPtr = 0
    mRefBmCount = 0
    Erase mRefBmNames
    Erase mRefBmStarts
End Sub

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
        If code = 13 Or code = 7 Or code = 2 Or code = 5 Or code = &HAD Or code = &H200B Or code = &H200C Or code = &H200D Or code = &HFEFF& Then
            ' Strip paragraph mark, cell bell, footnote/comment markers, soft hyphen, and zero-width chars
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
    If Len(tail) = 0 Or Len(tail) > 9 Then Exit Function
    For k = 1 To Len(tail)
        ch = AscW(Mid$(tail, k, 1))
        If ch < 48 Or ch > 57 Then Exit Function
    Next k
    ExtractRefBmNum = CLng(tail)
End Function

Public Function InferBodyStartPara(ByVal doc As Document, ByRef outHasExplicitMarkers As Boolean) As Long
    outHasExplicitMarkers = False
    InferBodyStartPara = 1
    If doc Is Nothing Then Exit Function
    Dim totalParas As Long, mk As MarkerLocations, errTmp As String
    totalParas = doc.Paragraphs.Count
    If totalParas < 1 Then Exit Function
    If ModMarkerLexer.ScanMarkers(doc, mk, errTmp) Then
        If mk.HasContentStart Then
            outHasExplicitMarkers = True
            InferBodyStartPara = mk.ContentStartParaIndex + 1
            Exit Function
        End If
    End If
    Dim maxFrontPos As Long, toc As TableOfContents, tof As TableOfFigures
    On Error Resume Next
    For Each toc In doc.TablesOfContents
        If toc.Range.End > maxFrontPos Then maxFrontPos = toc.Range.End
    Next toc
    For Each tof In doc.TablesOfFigures
        If tof.Range.End > maxFrontPos Then maxFrontPos = tof.Range.End
    Next tof
    Err.Clear
    On Error GoTo 0
    Dim reBodyL1 As Object, reFrontL1 As Object, reLoa As Object
    Set reBodyL1 = CreateObject("VBScript.RegExp")
    reBodyL1.IgnoreCase = True
    reBodyL1.Pattern = "^(?:(?:chapter|" & ModConstants.Uni("ch\u01b0\u01a1ng") & ")\s+1\b|1\.\s*(?:introduction|" & ModConstants.Uni("m\u1edf \u0111\u1ea7u|gi\u1edbi thi\u1ec7u|t\u1ed5ng quan") & "))"
    Set reFrontL1 = CreateObject("VBScript.RegExp")
    reFrontL1.IgnoreCase = True
    reFrontL1.Pattern = "^(?:tables? of contents?|list of (?:tables?|figures?|acronyms?|abbreviations?)|abstract|acknowledg(?:e)?ments?|declaration|" & ModConstants.Uni("m\u1ee5c l\u1ee5c|danh m\u1ee5c|l\u1eddi c[\u1ea3\u00e1]m \u01a1n|l\u1eddi cam \u0111oan|t\u00f3m t\u1eaft") & ")\b"
    Set reLoa = CreateObject("VBScript.RegExp")
    reLoa.IgnoreCase = True
    reLoa.Pattern = "^(?:list of (?:acronyms?|abbreviations?)|" & ModConstants.Uni("danh m\u1ee5c (?:c\u00e1c )?(?:t\u1eeb|ch\u1eef|k\u00fd hi\u1ec7u)") & ")"
    Dim p As Paragraph, pIdx As Long, afterFrontIdx As Long, bodyL1Idx As Long, limitScan As Long, maxLoaScan As Long
    Dim r As Range, cleanTxt As String, oLvl As Long, inTbl As Boolean
    afterFrontIdx = 1
    maxLoaScan = CLng(totalParas * 0.35)
    If maxLoaScan < 30 Then maxLoaScan = 30
    limitScan = CLng(totalParas * 0.5)
    If limitScan < 60 Then limitScan = 60
    For Each p In doc.Paragraphs
        pIdx = pIdx + 1
        Set r = p.Range
        If maxFrontPos > 0 And r.Start < maxFrontPos Then
            afterFrontIdx = pIdx + 1
        Else
            cleanTxt = NormalizeStr(r.Text)
            If Len(cleanTxt) > 0 Then
                inTbl = False
                On Error Resume Next
                inTbl = CBool(r.Information(12))
                Err.Clear
                On Error GoTo 0
                If Not inTbl And pIdx <= maxLoaScan And Len(cleanTxt) <= 100 And reLoa.Test(cleanTxt) Then
                    On Error Resume Next
                    If pIdx < totalParas Then
                        If CBool(doc.Paragraphs(pIdx + 1).Range.Information(12)) Then
                            If doc.Paragraphs(pIdx + 1).Range.Tables(1).Range.End > maxFrontPos Then maxFrontPos = doc.Paragraphs(pIdx + 1).Range.Tables(1).Range.End
                        End If
                    End If
                    Err.Clear
                    On Error GoTo 0
                    afterFrontIdx = pIdx + 1
                ElseIf Not inTbl And Len(cleanTxt) <= 150 Then
                    If reBodyL1.Test(cleanTxt) Then
                        bodyL1Idx = pIdx
                        Exit For
                    End If
                    oLvl = 10
                    On Error Resume Next
                    oLvl = p.OutlineLevel
                    Err.Clear
                    On Error GoTo 0
                    If oLvl = 1 And pIdx > 1 And Not reFrontL1.Test(cleanTxt) Then
                        bodyL1Idx = pIdx
                        Exit For
                    End If
                End If
            End If
            If pIdx > limitScan And afterFrontIdx > 1 Then Exit For
        End If
    Next p
    If bodyL1Idx > 0 Then
        InferBodyStartPara = bodyL1Idx
    ElseIf afterFrontIdx <= totalParas Then
        InferBodyStartPara = afterFrontIdx
    End If
End Function

Public Function ScanDocForPatch(ByVal doc As Document, ByRef outMarkers As MarkerLocations, _
                                ByRef outParas() As ScannedPara, ByRef outCount As Long, _
                                ByRef outMaxRefNum As Long, ByRef outErr As String) As Boolean
    On Error GoTo ScanErr
    outCount = 0
    outMaxRefNum = 0
    outErr = ""
    If doc Is Nothing Then
        outErr = "No active document."
        GoTo CleanScanExit
    End If
    Dim scanMkOk As Boolean, mkErr As String, hasExpMk As Boolean, startP As Long
    scanMkOk = ModMarkerLexer.ScanMarkers(doc, outMarkers, mkErr)
    If Not scanMkOk Then
        If outMarkers.HasContentStart Or outMarkers.HasContentEnd Then
            outErr = mkErr
            GoTo CleanScanExit
        End If
    End If
    If outMarkers.HasContentStart Then
        hasExpMk = True
        startP = outMarkers.ContentStartParaIndex + 1
    Else
        startP = InferBodyStartPara(doc, hasExpMk)
    End If
    g_LastInferredBodyStart = startP
    g_LastHasExplicitMarkers = hasExpMk
    Dim bm As Bookmark, bmNum As Long, bmCap As Long
    mRefBmDocPtr = ObjPtr(doc)
    mRefBmCount = 0
    bmCap = 32
    ReDim mRefBmNames(1 To bmCap)
    ReDim mRefBmStarts(1 To bmCap)
    For Each bm In doc.Bookmarks
        If Left$(bm.Name, 7) = BM_PREFIX_REF Then
            bmNum = ExtractRefBmNum(bm.Name)
            If bmNum > outMaxRefNum Then outMaxRefNum = bmNum
            mRefBmCount = mRefBmCount + 1
            If mRefBmCount > bmCap Then
                bmCap = bmCap * 2
                ReDim Preserve mRefBmNames(1 To bmCap)
                ReDim Preserve mRefBmStarts(1 To bmCap)
            End If
            mRefBmNames(mRefBmCount) = bm.Name
            mRefBmStarts(mRefBmCount) = bm.Range.Start
        End If
    Next bm
    ModTableIndex.BuildTableIndex doc
    ModCustomProfile.LoadTagCache doc
    Dim hasTags As Boolean, tagCode As Long
    hasTags = ModCustomProfile.HasActiveTags()
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

    Dim totalParas As Long
    totalParas = doc.Paragraphs.Count
    If startP > totalParas Then
        outErr = "No body paragraphs found after =contentStart()."
        GoTo CleanScanExit
    End If
    Dim cap As Long
    cap = totalParas
    If cap < 8 Then cap = 8
    ReDim outParas(1 To cap)

    Dim docHasMath As Boolean, inRefZone As Boolean, p As Paragraph, pIdx As Long, refCls As Long
    Dim rng As Range, rawTxt As String, normTxt As String, ps() As Long, pe() As Long, nLen As Long
    Dim rStart As Long, rEnd As Long, checkH1Style As Boolean
    Dim inTbl As Boolean, isL1 As Boolean, isHead As Boolean, isExcl As Boolean, stName As String
    Dim exclS() As Long, exclE() As Long, exclCount As Long
    ModPatchDom.CollectExcludedSpans doc, exclS, exclE, exclCount
    docHasMath = (doc.OMaths.Count > 0)
    If docHasMath Then ModTableIndex.BuildOMathIndex doc
    inRefZone = False
    pIdx = 0
    For Each p In doc.Paragraphs
        pIdx = pIdx + 1
        Set rng = p.Range
        rawTxt = rng.Text
        NormalizeWithPosMap rawTxt, normTxt, ps, pe, nLen
        If nLen > 0 And Not reMark.Test(normTxt) Then
            rStart = rng.Start
            rEnd = rng.End
            inTbl = ModTableIndex.IsPosInTable(rStart)
            isL1 = False
            isHead = False
            isExcl = False
            If exclCount > 0 Then isExcl = ModPatchDom.IsInExcludedSpans(rStart, rEnd, exclS, exclE, exclCount)
            If Not isExcl And docHasMath Then
                If ModTableIndex.HasOMathInSpan(rStart, rEnd) Then isExcl = IsMathBlockOnly(rng, rawTxt)
            End If
            If Not isExcl And Not inTbl And nLen <= MAX_HEADING_LEN Then
                If reL1.Test(normTxt) Or ModLocale.IsRomanLevel1(normTxt) Then isL1 = True
                If Not isL1 Then
                    On Error Resume Next
                    If p.OutlineLevel = 1 Then isL1 = True
                    Err.Clear
                    On Error GoTo ScanErr
                End If
                If isL1 Or reL2.Test(normTxt) Or reL3.Test(normTxt) Or reL4.Test(normTxt) Or reTbl.Test(normTxt) Or reFig.Test(normTxt) Then
                    isHead = True
                End If
            End If
            Dim tagForcedZone As Long, pOutLvl As Long
            tagForcedZone = 0
            If Not isExcl And hasTags Then
                tagCode = ModCustomProfile.LookupTagForRange(rStart, rEnd)
                If tagCode = TAG_CODE_SKIP Then
                    isL1 = False
                    isHead = False
                ElseIf tagCode = TAG_CODE_REF Then
                    isL1 = True
                    isHead = True
                    inRefZone = True
                    tagForcedZone = 1
                ElseIf tagCode = TAG_CODE_REF_SUB Then
                    isHead = True
                    tagForcedZone = 2
                ElseIf tagCode = TAG_CODE_APP Then
                    isL1 = True
                    isHead = True
                    inRefZone = False
                    tagForcedZone = -1
                ElseIf tagCode = ElementTypeLevel1 Then
                    isL1 = True
                    isHead = True
                ElseIf tagCode >= ElementTypeLevel2 And tagCode <= ElementTypeFigureTitle Then
                    isHead = True
                ElseIf tagCode = ElementTypeSourceNote Then
                    isL1 = False
                    isHead = False
                End If
            End If
            If tagForcedZone = 0 And Not isExcl And Not inTbl And nLen <= MAX_HEADING_LEN And pIdx >= startP Then
                If isL1 Then
                    refCls = ModLocale.ClassifyRefZoneHeading(normTxt)
                    If refCls = 1 Then
                        inRefZone = True
                    ElseIf refCls = -1 Then
                        inRefZone = False
                    End If
                ElseIf inRefZone And nLen <= 100 And InStr(1, normTxt, "(") = 0 Then
                    If ModLocale.ClassifyRefZoneHeading(normTxt) = -1 Then inRefZone = False
                End If
            End If
            If inRefZone And Not isHead And Not isExcl And Not inTbl And nLen <= 120 Then
                If ModLocale.IsRefSubheading(normTxt) Then
                    isHead = True
                Else
                    pOutLvl = 10
                    On Error Resume Next
                    pOutLvl = p.OutlineLevel
                    Err.Clear
                    On Error GoTo ScanErr
                    If pOutLvl >= 2 And pOutLvl <= 4 Then isHead = True
                End If
            End If
            outCount = outCount + 1
            outParas(outCount).ParaIndex = pIdx
            outParas(outCount).RangeStart = rStart
            outParas(outCount).RangeEnd = rEnd
            outParas(outCount).RawText = rawTxt
            outParas(outCount).NormText = normTxt
            outParas(outCount).PosStart = ps
            outParas(outCount).PosEnd = pe
            outParas(outCount).NormLen = nLen
            outParas(outCount).InBodyScope = (pIdx >= startP And Not inRefZone And Not isExcl)
            outParas(outCount).IsLevel1 = isL1
            outParas(outCount).IsHeading = isHead
            outParas(outCount).IsInTable = inTbl
            outParas(outCount).IsInRefZone = inRefZone
            outParas(outCount).IsExcluded = isExcl
        End If
    Next p
    ScanDocForPatch = True
CleanScanExit:
    On Error Resume Next
    ModCustomProfile.ResetTagCache
    Err.Clear
    Exit Function
ScanErr:
    If Len(outErr) = 0 Then outErr = "Scan error (" & Err.Number & "): " & Err.Description
    ScanDocForPatch = False
    Resume CleanScanExit
End Function

Private Function FindCachedRefBmInPara(ByVal doc As Document, ByVal pStart As Long, ByVal pEnd As Long) As String
    Dim k As Long, bm As Bookmark
    If mRefBmDocPtr = ObjPtr(doc) Then
        For k = 1 To mRefBmCount
            If mRefBmStarts(k) >= pStart And mRefBmStarts(k) < pEnd Then
                FindCachedRefBmInPara = mRefBmNames(k)
                Exit Function
            End If
        Next k
        Exit Function
    End If
    For Each bm In doc.Bookmarks
        If Left$(bm.Name, 7) = BM_PREFIX_REF Then
            If bm.Range.Start >= pStart And bm.Range.Start < pEnd Then
                FindCachedRefBmInPara = bm.Name
                Exit Function
            End If
        End If
    Next bm
End Function

Private Function SafeLeftUtf16(ByVal s As String, ByVal maxUnits As Long) As String
    If maxUnits <= 0 Then Exit Function
    If Len(s) <= maxUnits Then
        SafeLeftUtf16 = s
        Exit Function
    End If
    Dim cutLen As Long, lastCode As Long
    cutLen = maxUnits
    lastCode = AscW(Mid$(s, cutLen, 1)) And &HFFFF&
    If lastCode >= &HD800& And lastCode <= &HDBFF& Then cutLen = cutLen - 1
    If cutLen > 0 Then SafeLeftUtf16 = Left$(s, cutLen)
End Function

Private Function BuildShortRefTip(ByVal normTxt As String) As String
    Dim s As String
    s = Trim$(normTxt)
    If Len(s) > 80 Then s = SafeLeftUtf16(s, 77) & "..."
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
            existBm = FindCachedRefBmInPara(doc, paras(hitParaArrIdx).RangeStart, paras(hitParaArrIdx).RangeEnd)
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
