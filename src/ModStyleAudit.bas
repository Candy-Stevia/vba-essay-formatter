Attribute VB_Name = "ModStyleAudit"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModStyleAudit
' Purpose: Wave 2 - Document Topology X-Ray & 8-Scope Pre-Submission Style Linter
' Architecture: Academic Document Format Engine (Read-Only, 0 DOM Mutation, 7-bit ASCII)
' ==============================================================================

Private Const MAX_HEADING_LEN As Long = 350
Private Const WD_UNDEF As Long = 9999999

Public Type StyleSample
    ParaIdx As Long
    PageNum As Long
    ScopeIdx As Long
    FontName As String
    SizeHalf As Long
    BoldVal As Long
    ItalicVal As Long
    UnderVal As Long
    ColorKey As String
    AlignVal As Long
    LeftInd20 As Long
    FirstInd20 As Long
    SpaceBef As Long
    SpaceAft As Long
    LineSp20 As Long
    IsUpper As Boolean
    IsSrcCanonical As Boolean
    RawLen As Long
    IsExplicitTag As Boolean
    PreviewText As String
End Type

Public Type ScopeVoteResult
    ScopeKey As String
    SampleCount As Long
    WinnerCount As Long
    SecondCount As Long
    Status As String
    WinFont As String
    WinSizeHalf As Long
    WinBold As Long
    WinItalic As Long
    WinUnder As Long
    WinColor As String
    WinAlign As Long
    WinLeft20 As Long
    WinFirst20 As Long
    WinSpBef As Long
    WinSpAft As Long
    WinLineSp20 As Long
    WinUpper As Boolean
    WinSrcCanonical As Boolean
    TieSummary As String
    DeviationsText As String
    DevCount As Long
End Type

Public Type VisualSuspectItem
    ParaIdx As Long
    PageNum As Long
    FullText As String
    StyleName As String
    OutlineLvl As Long
    FontName As String
    SizeHalf As Long
    BoldVal As Long
    ColorKey As String
End Type

Public Function ScopeIdxToKey(ByVal idx As Long) As String
    Dim keys As Variant
    keys = Array("", "L1", "L2", "L3", "L4", "TBL", "FIG", "SRC", "BODY")
    If idx >= 1 And idx <= 8 Then ScopeIdxToKey = CStr(keys(idx))
End Function

Public Function AlignToName(ByVal aVal As Long) As String
    If aVal = 1 Then
        AlignToName = "center"
    ElseIf aVal = 2 Then
        AlignToName = "right"
    ElseIf aVal = 3 Then
        AlignToName = "justify"
    Else
        AlignToName = "left"
    End If
End Function

Public Function NormalizeColorKey(ByVal cVal As Long) As String
    If cVal = WD_UNDEF Then
        NormalizeColorKey = "mixed"
    ElseIf cVal = -16777216 Or cVal = 0 Then
        NormalizeColorKey = "auto"
    Else
        Dim r As Long, g As Long, b As Long
        r = cVal And &HFF&
        g = (cVal \ 256&) And &HFF&
        b = (cVal \ 65536&) And &HFF&
        NormalizeColorKey = "#" & Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
    End If
End Function

Private Function QuantSizeHalf(ByVal sVal As Single) As Long
    If sVal >= 999999# Or sVal <= 0# Then
        QuantSizeHalf = WD_UNDEF
    Else
        QuantSizeHalf = CLng(Int(sVal * 2# + 0.5))
    End If
End Function

Private Function QuantCm20(ByVal ptsVal As Single) As Long
    If ptsVal >= 999999# Then
        QuantCm20 = WD_UNDEF
    Else
        Dim cmVal As Double
        cmVal = CDbl(ptsVal) / 28.3464567
        If cmVal < 0# Then cmVal = 0#
        QuantCm20 = CLng(Int(cmVal * 20# + 0.5))
    End If
End Function

Private Function QuantPt1(ByVal ptsVal As Single) As Long
    If ptsVal >= 999999# Or ptsVal < 0# Then
        QuantPt1 = WD_UNDEF
    Else
        QuantPt1 = CLng(Int(ptsVal + 0.5))
    End If
End Function

Private Function QuantLineSp20(ByVal pf As ParagraphFormat) As Long
    On Error Resume Next
    Dim ruleVal As Long, spVal As Single, multVal As Double
    ruleVal = pf.LineSpacingRule
    spVal = pf.LineSpacing
    If Err.Number <> 0 Or ruleVal = WD_UNDEF Or spVal >= 999999# Then
        Err.Clear
        QuantLineSp20 = WD_UNDEF
        Exit Function
    End If
    If ruleVal = 0 Then
        multVal = 1#
    ElseIf ruleVal = 1 Then
        multVal = 1.5
    ElseIf ruleVal = 2 Then
        multVal = 2#
    ElseIf ruleVal = 5 Then
        multVal = CDbl(spVal) / 12#
    Else
        QuantLineSp20 = -1
        Exit Function
    End If
    QuantLineSp20 = CLng(Int(multVal * 20# + 0.5))
End Function

Private Function ReadSingleSample(ByVal doc As Document, ByVal p As Paragraph, ByVal pRng As Range, ByVal pIdx As Long, _
                                  ByVal scopeIdx As Long, ByVal cleanTxt As String, _
                                  Optional ByVal isExplicitTag As Boolean = False) As StyleSample
    Dim sm As StyleSample, rNoCr As Range
    sm.ParaIdx = pIdx
    sm.ScopeIdx = scopeIdx
    sm.RawLen = Len(cleanTxt)
    sm.IsExplicitTag = isExplicitTag
    If sm.RawLen <= 120 Then sm.PreviewText = cleanTxt Else sm.PreviewText = Left$(cleanTxt, 60)
    On Error Resume Next
    sm.PageNum = pRng.Information(3)
    If pRng.End > pRng.Start + 1 Then
        Set rNoCr = doc.Range(pRng.Start, pRng.End - 1)
    Else
        Set rNoCr = pRng
    End If
    sm.FontName = Trim$(rNoCr.Font.Name)
    If Len(sm.FontName) = 0 Then sm.FontName = "mixed"
    sm.SizeHalf = QuantSizeHalf(rNoCr.Font.Size)
    sm.BoldVal = rNoCr.Font.Bold
    sm.ItalicVal = rNoCr.Font.Italic
    sm.UnderVal = rNoCr.Font.Underline
    sm.ColorKey = NormalizeColorKey(rNoCr.Font.Color)
    sm.AlignVal = pRng.ParagraphFormat.Alignment
    If sm.AlignVal = WD_UNDEF Then sm.AlignVal = 0
    sm.LeftInd20 = QuantCm20(pRng.ParagraphFormat.LeftIndent)
    sm.FirstInd20 = QuantCm20(pRng.ParagraphFormat.FirstLineIndent)
    sm.SpaceBef = QuantPt1(pRng.ParagraphFormat.SpaceBefore)
    sm.SpaceAft = QuantPt1(pRng.ParagraphFormat.SpaceAfter)
    sm.LineSp20 = QuantLineSp20(pRng.ParagraphFormat)
    sm.IsUpper = (StrComp(cleanTxt, UCase$(cleanTxt), vbBinaryCompare) = 0) Or (rNoCr.Font.AllCaps = -1)
    If scopeIdx = 7 Then
        Dim colonPos As Long, pfxRng As Range, bodyRng As Range
        colonPos = InStr(1, cleanTxt, ":", vbBinaryCompare)
        If colonPos > 1 And colonPos < Len(cleanTxt) And (rNoCr.End - rNoCr.Start) = Len(cleanTxt) Then
            Set pfxRng = doc.Range(rNoCr.Start, rNoCr.Start + colonPos)
            Set bodyRng = doc.Range(rNoCr.Start + colonPos, rNoCr.End)
            If pfxRng.Font.Bold = -1 And bodyRng.Font.Italic = -1 Then
                sm.IsSrcCanonical = True
                sm.BoldVal = 0
                sm.ItalicVal = 0
                sm.UnderVal = 0
            End If
        End If
    End If
    Err.Clear
    ReadSingleSample = sm
End Function

Private Function TryBatchProbeBody(ByVal doc As Document, ByVal bStart As Long, ByVal bEnd As Long, _
                                   ByRef outSm As StyleSample) As Boolean
    If bEnd <= bStart + 1 Then Exit Function
    Dim r As Range
    On Error Resume Next
    Set r = doc.Range(bStart, bEnd - 1)
    Dim fn As String, sz As Long, al As Long, li As Long, fi As Long, ls As Long
    Dim bVal As Long, iVal As Long, uVal As Long, cKey As String
    fn = Trim$(r.Font.Name)
    sz = QuantSizeHalf(r.Font.Size)
    al = r.ParagraphFormat.Alignment
    li = QuantCm20(r.ParagraphFormat.LeftIndent)
    fi = QuantCm20(r.ParagraphFormat.FirstLineIndent)
    ls = QuantLineSp20(r.ParagraphFormat)
    bVal = r.Font.Bold
    iVal = r.Font.Italic
    uVal = r.Font.Underline
    cKey = NormalizeColorKey(r.Font.Color)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    If Len(fn) = 0 Or sz = WD_UNDEF Or al = WD_UNDEF Or li = WD_UNDEF Or fi = WD_UNDEF Or ls = WD_UNDEF Then Exit Function
    If bVal = WD_UNDEF Or iVal = WD_UNDEF Or uVal = WD_UNDEF Or cKey = "mixed" Then Exit Function
    outSm.ScopeIdx = 8
    outSm.FontName = fn
    outSm.SizeHalf = sz
    outSm.BoldVal = bVal
    outSm.ItalicVal = iVal
    outSm.UnderVal = uVal
    outSm.ColorKey = cKey
    outSm.AlignVal = al
    outSm.LeftInd20 = li
    outSm.FirstInd20 = fi
    outSm.LineSp20 = ls
    TryBatchProbeBody = True
End Function

Public Function RunStyleAuditCore(ByVal doc As Document, ByRef outVotes() As ScopeVoteResult, _
                                  ByRef outCompactSummary As String, ByRef outFullReport As String, _
                                  ByRef outSparseBlock As String, ByRef outErr As String) As Boolean
    On Error GoTo AuditErr
    ReDim outVotes(1 To 8)
    If doc Is Nothing Then
        outErr = "No active document."
        Exit Function
    End If
    Dim totalParas As Long
    totalParas = doc.Paragraphs.Count
    If totalParas < 1 Then
        outErr = "Empty document."
        Exit Function
    End If

    Dim mk As MarkerLocations, mkErr As String, hasExp As Boolean, startP As Long
    ModMarkerLexer.ScanMarkers doc, mk, mkErr
    startP = ModPatchMatcher.InferBodyStartPara(doc, hasExp)
    Dim exclS() As Long, exclE() As Long, exclCount As Long
    ModPatchDom.CollectExcludedSpans doc, exclS, exclE, exclCount
    ModTableIndex.BuildTableIndex doc
    ModTableIndex.BuildListIndex doc
    ModTableIndex.BuildOMathIndex doc
    ModCustomProfile.LoadTagCache doc
    Dim hasTags As Boolean, tblValid As Boolean, listValid As Boolean, docHasMath As Boolean
    hasTags = ModCustomProfile.HasActiveTags()
    tblValid = ModTableIndex.IsTableIndexValidForDoc(doc)
    listValid = ModTableIndex.IsListIndexValidForDoc(doc)
    docHasMath = (doc.OMaths.Count > 0)

    Dim reL1 As Object, reL2 As Object, reL3 As Object, reL4 As Object
    Dim reTbl As Object, reFig As Object, reSrc As Object, reMark As Object
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
    Set reSrc = CreateObject("VBScript.RegExp")
    reSrc.IgnoreCase = True
    reSrc.Pattern = ModLocale.GetPatternSourceNote()
    Set reMark = CreateObject("VBScript.RegExp")
    reMark.IgnoreCase = True
    reMark.Pattern = ModConstants.GetPatternMarker()

    Dim samples() As StyleSample, smCnt As Long, smCap As Long
    Dim suspects() As VisualSuspectItem, susCnt As Long, susCap As Long
    Dim bodyParas() As Paragraph, bodyIdx() As Long, bodyClean() As String, bodyIsExp() As Boolean, bodyCnt As Long, bodyCap As Long
    smCap = 64
    ReDim samples(1 To smCap)
    susCap = 16
    ReDim suspects(1 To susCap)
    bodyCap = 128
    ReDim bodyParas(1 To bodyCap)
    ReDim bodyIdx(1 To bodyCap)
    ReDim bodyClean(1 To bodyCap)
    ReDim bodyIsExp(1 To bodyCap)

    Dim p As Paragraph, pRng As Range, pIdx As Long, rawTxt As String, cleanTxt As String
    Dim rStart As Long, rEnd As Long, isExcl As Boolean, inTbl As Boolean, inList As Boolean
    Dim inRefZone As Boolean, tagCode As Long, oLvl As Long, scIdx As Long, byTagOrRx As Boolean
    Dim refCls As Long, sm As StyleSample
    inRefZone = False
    pIdx = 0
    For Each p In doc.Paragraphs
        pIdx = pIdx + 1
        If pIdx < startP Then GoTo NextAuditPara
        Set pRng = p.Range
        rawTxt = pRng.Text
        rStart = pRng.Start
        rEnd = pRng.End
        cleanTxt = ModPatchMatcher.NormalizeStr(rawTxt)
        If Len(cleanTxt) = 0 Then GoTo NextAuditPara
        If reMark.Test(cleanTxt) Then GoTo NextAuditPara
        isExcl = False
        If exclCount > 0 Then isExcl = ModPatchDom.IsInExcludedSpans(rStart, rEnd, exclS, exclE, exclCount)
        If isExcl Then GoTo NextAuditPara
        If tblValid Then inTbl = ModTableIndex.IsPosInTable(rStart) Else inTbl = CBool(pRng.Information(12))
        If inTbl Then GoTo NextAuditPara
        If docHasMath Then
            If ModTableIndex.HasOMathInSpan(rStart, rEnd) Then
                If pRng.OMaths.Count > 0 And Len(Replace(cleanTxt, " ", "")) <= 2 Then GoTo NextAuditPara
            End If
        End If
        If listValid Then inList = ModTableIndex.IsPosInList(rStart) Else inList = (pRng.ListFormat.ListType <> 0)
        oLvl = 10
        On Error Resume Next
        oLvl = p.OutlineLevel
        Err.Clear
        On Error GoTo AuditErr

        tagCode = 0
        If hasTags Then tagCode = ModCustomProfile.LookupTagForRange(rStart, rEnd)
        scIdx = 0
        byTagOrRx = False
        If tagCode = TAG_CODE_REF Then
            inRefZone = True
            scIdx = 1
            byTagOrRx = True
        ElseIf tagCode = TAG_CODE_APP Then
            inRefZone = False
            scIdx = 1
            byTagOrRx = True
        ElseIf tagCode = TAG_CODE_REF_SUB Then
            scIdx = 2
            byTagOrRx = True
        ElseIf tagCode = TAG_CODE_SKIP Then
            scIdx = 8
            byTagOrRx = True
        ElseIf tagCode >= 1 And tagCode <= 7 Then
            scIdx = tagCode
            byTagOrRx = True
        ElseIf inList And oLvl > 4 Then
            GoTo NextAuditPara
        ElseIf Len(cleanTxt) <= MAX_HEADING_LEN Then
            If reTbl.Test(cleanTxt) Then
                scIdx = 5
                byTagOrRx = True
            ElseIf reFig.Test(cleanTxt) Then
                scIdx = 6
                byTagOrRx = True
            ElseIf reSrc.Test(cleanTxt) Then
                scIdx = 7
                byTagOrRx = True
            ElseIf oLvl >= 1 And oLvl <= 4 Then
                scIdx = oLvl
            ElseIf reL1.Test(cleanTxt) Or ModLocale.IsRomanLevel1(cleanTxt) Then
                scIdx = 1
                byTagOrRx = True
            ElseIf reL2.Test(cleanTxt) Then
                scIdx = 2
                byTagOrRx = True
            ElseIf reL3.Test(cleanTxt) Then
                scIdx = 3
                byTagOrRx = True
            ElseIf reL4.Test(cleanTxt) Then
                scIdx = 4
                byTagOrRx = True
            End If
        End If
        If tagCode <> TAG_CODE_REF And tagCode <> TAG_CODE_APP And tagCode <> TAG_CODE_REF_SUB Then
            If scIdx = 1 Then
                refCls = ModLocale.ClassifyRefZoneHeading(cleanTxt)
                If refCls = 1 Then
                    inRefZone = True
                ElseIf refCls = -1 Then
                    inRefZone = False
                End If
            ElseIf inRefZone And Len(cleanTxt) <= 100 And InStr(1, cleanTxt, "(") = 0 Then
                If ModLocale.ClassifyRefZoneHeading(cleanTxt) = -1 Then inRefZone = False
            End If
        End If
        If inRefZone And scIdx <> 1 Then
            If ModLocale.IsRefSubheading(cleanTxt) Then GoTo NextAuditPara
            GoTo NextAuditPara
        End If
        If scIdx = 0 Then scIdx = 8

        If scIdx = 8 Then
            If Not byTagOrRx And Len(cleanTxt) <= 120 Then
                sm = ReadSingleSample(doc, p, pRng, pIdx, 8, cleanTxt, False)
                If oLvl <= 4 Or sm.BoldVal = -1 Then
                    susCnt = susCnt + 1
                    If susCnt > susCap Then
                        susCap = susCap * 2
                        ReDim Preserve suspects(1 To susCap)
                    End If
                    suspects(susCnt).ParaIdx = pIdx
                    suspects(susCnt).PageNum = sm.PageNum
                    suspects(susCnt).FullText = cleanTxt
                    On Error Resume Next
                    suspects(susCnt).StyleName = p.Style.NameLocal
                    Err.Clear
                    On Error GoTo AuditErr
                    suspects(susCnt).OutlineLvl = oLvl
                    suspects(susCnt).FontName = sm.FontName
                    suspects(susCnt).SizeHalf = sm.SizeHalf
                    suspects(susCnt).BoldVal = sm.BoldVal
                    suspects(susCnt).ColorKey = sm.ColorKey
                    GoTo NextAuditPara
                End If
            End If
            bodyCnt = bodyCnt + 1
            If bodyCnt > bodyCap Then
                bodyCap = bodyCap * 2
                ReDim Preserve bodyParas(1 To bodyCap)
                ReDim Preserve bodyIdx(1 To bodyCap)
                ReDim Preserve bodyClean(1 To bodyCap)
                ReDim Preserve bodyIsExp(1 To bodyCap)
            End If
            Set bodyParas(bodyCnt) = p
            bodyIdx(bodyCnt) = pIdx
            bodyClean(bodyCnt) = cleanTxt
            bodyIsExp(bodyCnt) = (tagCode = TAG_CODE_SKIP)
        Else
            sm = ReadSingleSample(doc, p, pRng, pIdx, scIdx, cleanTxt, (tagCode <> 0))
            smCnt = smCnt + 1
            If smCnt > smCap Then
                smCap = smCap * 2
                ReDim Preserve samples(1 To smCap)
            End If
            samples(smCnt) = sm
        End If
NextAuditPara:
    Next p

    ' Batch-Probe wdUndefined on contiguous BODY chunks (up to 60 paras/batch)
    Dim bi As Long, bj As Long, bStartPos As Long, bEndPos As Long, batchSm As StyleSample, bk As Long
    bi = 1
    Do While bi <= bodyCnt
        bj = bi
        bStartPos = bodyParas(bi).Range.Start
        bEndPos = bodyParas(bi).Range.End
        Do While bj < bodyCnt And (bj - bi + 1) < 60
            If bodyParas(bj + 1).Range.Start = bEndPos Then
                bj = bj + 1
                bEndPos = bodyParas(bj).Range.End
            Else
                Exit Do
            End If
        Loop
        If (bj - bi + 1) >= 4 And TryBatchProbeBody(doc, bStartPos, bEndPos, batchSm) Then
            For bk = bi To bj
                smCnt = smCnt + 1
                If smCnt > smCap Then
                    smCap = smCap * 2
                    ReDim Preserve samples(1 To smCap)
                End If
                samples(smCnt) = batchSm
                samples(smCnt).ParaIdx = bodyIdx(bk)
                samples(smCnt).PageNum = 0
                samples(smCnt).RawLen = Len(bodyClean(bk))
                samples(smCnt).IsExplicitTag = bodyIsExp(bk)
                If samples(smCnt).RawLen <= 120 Then
                    samples(smCnt).PreviewText = bodyClean(bk)
                Else
                    samples(smCnt).PreviewText = Left$(bodyClean(bk), 60)
                End If
            Next bk
        Else
            For bk = bi To bj
                sm = ReadSingleSample(doc, bodyParas(bk), bodyParas(bk).Range, bodyIdx(bk), 8, bodyClean(bk), bodyIsExp(bk))
                smCnt = smCnt + 1
                If smCnt > smCap Then
                    smCap = smCap * 2
                    ReDim Preserve samples(1 To smCap)
                End If
                samples(smCnt) = sm
            Next bk
        End If
        bi = bj + 1
    Loop

    Dim sc As Long
    For sc = 1 To 8
        ModStyleCapture.VoteSingleScope doc, sc, samples, smCnt, outVotes(sc)
    Next sc

    ' Second pass: promote short BODY samples with Font.Size > bodyModeSize into Visual Suspects
    Dim bodyModeSize As Long, susPg As Long
    bodyModeSize = outVotes(8).WinSizeHalf
    If bodyModeSize > 0 And bodyModeSize <> WD_UNDEF Then
        For bi = 1 To smCnt
            If samples(bi).ScopeIdx = 8 And samples(bi).RawLen <= 120 And Not samples(bi).IsExplicitTag Then
                If samples(bi).SizeHalf <> WD_UNDEF And samples(bi).SizeHalf > bodyModeSize Then
                    susCnt = susCnt + 1
                    If susCnt > susCap Then
                        susCap = susCap * 2
                        ReDim Preserve suspects(1 To susCap)
                    End If
                    susPg = samples(bi).PageNum
                    If susPg = 0 And samples(bi).ParaIdx >= 1 And samples(bi).ParaIdx <= totalParas Then
                        On Error Resume Next
                        susPg = doc.Paragraphs(samples(bi).ParaIdx).Range.Information(3)
                        Err.Clear
                        On Error GoTo AuditErr
                    End If
                    suspects(susCnt).ParaIdx = samples(bi).ParaIdx
                    suspects(susCnt).PageNum = susPg
                    suspects(susCnt).FullText = samples(bi).PreviewText
                    suspects(susCnt).StyleName = "Body"
                    suspects(susCnt).OutlineLvl = 10
                    suspects(susCnt).FontName = samples(bi).FontName
                    suspects(susCnt).SizeHalf = samples(bi).SizeHalf
                    suspects(susCnt).BoldVal = samples(bi).BoldVal
                    suspects(susCnt).ColorKey = samples(bi).ColorKey
                End If
            End If
        Next bi
    End If

    outSparseBlock = ModStyleCapture.BuildSparseStyleBlock(outVotes)
    ModStyleCapture.FormatAuditReports doc, mk, hasExp, startP, outVotes, suspects, susCnt, outSparseBlock, outCompactSummary, outFullReport
    RunStyleAuditCore = True

CleanAudit:
    On Error Resume Next
    ModTableIndex.InvalidateAllIndexes
    ModCustomProfile.ResetTagCache
    Err.Clear
    Exit Function
AuditErr:
    outErr = "StyleAudit Error " & Err.Number & ": " & Err.Description
    Resume CleanAudit
End Function

Public Function RunStyleAuditForTest(ByVal doc As Document) As String
    Dim votes() As ScopeVoteResult, compactSum As String, fullRep As String, sparseBlk As String, errMsg As String
    If RunStyleAuditCore(doc, votes, compactSum, fullRep, sparseBlk, errMsg) Then
        RunStyleAuditForTest = fullRep
    Else
        RunStyleAuditForTest = "ERR|" & errMsg
    End If
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
