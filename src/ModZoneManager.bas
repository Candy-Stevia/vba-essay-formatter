Attribute VB_Name = "ModZoneManager"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModZoneManager
' Purpose: Tri-Zone management (Wipe & Rebuild), Bookmark preservation & Submission cleaner
' Architecture: Academic Document Format Engine (Idempotent Architecture)
' ==============================================================================

Public Type QuickZoneSlot
    Found As Boolean
    Kind As Long
    ZoneBm As String
    HeaderBm As String
    AnchorBm As String
    HeaderTrackBm As String
    MarkerPara As Long
    HdrStart As Long
    BodyStart As Long
    BodyEnd As Long
    EntryCount As Long
    ItemCount As Long
End Type

Private Function IsMarkerString(ByVal txt As String) As Boolean
    Dim c As String
    c = LCase$(Trim$(Replace(Replace(Replace(txt, vbCr, ""), Chr(7), ""), Chr(11), "")))
    Select Case c
        Case "=contentstart()", "=contentend()", "=listcontent()", "=listtablehere()", "=listfigurehere()", "=listacronyms()"
            IsMarkerString = True
    End Select
End Function

Private Function TrimTrailingBreak(ByVal doc As Document, ByVal sPos As Long, ByVal ePos As Long) As Long
    If ePos <= sPos Then GoTo Done
    On Error Resume Next
    Dim dEnd As Long, rTxt As String, lp As Range
    dEnd = doc.Range.End
    If ePos >= dEnd And dEnd > 0 Then ePos = dEnd - 1
    If ePos <= sPos Then GoTo Done
    rTxt = doc.Range(sPos, ePos).Text
    If Err.Number = 0 And InStr(1, rTxt, Chr(12)) > 0 Then
        Dim firstBrk As Long
        firstBrk = InStr(1, rTxt, Chr(12))
        If firstBrk = Len(rTxt) And Right$(rTxt, 1) = Chr(12) Then
            ePos = ePos - 1
        ElseIf firstBrk = Len(rTxt) - 1 And Right$(rTxt, 2) = Chr(12) & vbCr Then
            ePos = ePos - 2
        Else
            Dim bp As Paragraph
            For Each bp In doc.Range(sPos, ePos).Paragraphs
                If InStr(1, bp.Range.Text, Chr(12)) > 0 Then
                    ePos = bp.Range.Start
                    Exit For
                End If
            Next bp
        End If
    End If
    Err.Clear
Done:
    TrimTrailingBreak = ePos
End Function

Public Sub WipeZone(ByVal doc As Document, ByVal bmZoneName As String, ByVal bmHeaderName As String)
    On Error Resume Next
    Dim zRng As Range, p As Paragraph, delStart As Long, delEnd As Long
    If doc.Bookmarks.Exists(bmZoneName) Then
        Set zRng = doc.Bookmarks(bmZoneName).Range
        If zRng.End > zRng.Start Then
            delStart = zRng.Start
            For Each p In zRng.Paragraphs
                If IsMarkerString(p.Range.Text) Or Not IsQuickEntryPara(p) Then Exit For
                delEnd = p.Range.End
                If InStr(1, p.Range.Text, Chr(12)) > 0 Then Exit For
            Next p
            delEnd = TrimTrailingBreak(doc, delStart, delEnd)
            If delEnd > delStart Then doc.Range(delStart, delEnd).Delete
        End If
    End If
    If doc.Bookmarks.Exists(bmHeaderName) Then
        Set zRng = doc.Bookmarks(bmHeaderName).Range.Paragraphs(1).Range
        If Not IsMarkerString(zRng.Text) Then
            delEnd = TrimTrailingBreak(doc, zRng.Start, zRng.End)
            If delEnd > zRng.Start Then doc.Range(zRng.Start, delEnd).Delete
        End If
    End If
    If Err.Number <> 0 Then ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
    Err.Clear
End Sub

Public Sub WipeAllZones(ByVal doc As Document)
    WipeZone doc, BM_ZONE_LOF, BM_LOF_HEADER
    WipeZone doc, BM_ZONE_LOT, BM_LOT_HEADER
    WipeZone doc, BM_ZONE_TOC, BM_TOC_HEADER
    On Error Resume Next
    If doc.Bookmarks.Exists(BM_SYS_GENERATED_BLOCK) Then doc.Bookmarks(BM_SYS_GENERATED_BLOCK).Delete
    If Err.Number <> 0 Then ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
    Err.Clear
End Sub

Public Sub EnsureTargetBookmarks(ByVal doc As Document, ByRef snapshots() As ParagraphSnapshot, ByVal totalCount As Long)
    Dim i As Long, tType As Long, countToc As Long, countTbl As Long, countFig As Long, bmName As String, dEnd As Long
    dEnd = doc.Range.End
    For i = 1 To totalCount
        tType = snapshots(i).ElementType
        bmName = ""
        If tType >= ElementTypeLevel1 And tType <= ElementTypeLevel4 Then
            countToc = countToc + 1
            bmName = FormatBookmarkId(BM_PREFIX_TOC, countToc)
        ElseIf tType = ElementTypeTableTitle Then
            countTbl = countTbl + 1
            bmName = FormatBookmarkId(BM_PREFIX_TBL, countTbl)
        ElseIf tType = ElementTypeFigureTitle Then
            countFig = countFig + 1
            bmName = FormatBookmarkId(BM_PREFIX_FIG, countFig)
        End If
        If Len(bmName) > 0 And snapshots(i).RangeStart < dEnd Then AddBookmark doc, bmName, doc.Range(snapshots(i).RangeStart, snapshots(i).RangeEnd)
    Next i
End Sub

Public Sub CleanMarkersForSubmission(ByVal doc As Document)
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then Exit Sub
    If doc Is Nothing And Documents.Count > 0 Then Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub
    Dim errMsg As String
    If Not ModBackup.ValidateWritableDocument(doc, errMsg) Then
        ModUI.MsgBoxW errMsg, vbExclamation, ModConstants.Uni("Cannot Edit (Kh\u00f4ng th\u1ec3 ch\u1ec9nh s\u1eeda)")
        Exit Sub
    End If
    If Application.DisplayAlerts <> 0 Then
        If ModUI.MsgBoxW(ModConstants.Uni("Remove all =...() markers and finalize math shading for submission? (Undoable with Ctrl+Z)") & vbCrLf & _
                         ModConstants.Uni("X\u00f3a to\u00e0n b\u1ed9 th\u1ebb =...() v\u00e0 ch\u1ed1t m\u00e0u n\u1ec1n to\u00e1n \u0111\u1ec3 n\u1ed9p b\u00e0i? (C\u00f3 th\u1ec3 ho\u00e0n t\u00e1c b\u1eb1ng Ctrl+Z)"), _
                         vbYesNo + vbQuestion + vbDefaultButton2, _
                         ModConstants.Uni("Confirm Clean Markers (X\u00e1c nh\u1eadn d\u1ecdn th\u1ebb)")) <> vbYes Then
            Exit Sub
        End If
    End If
    Dim wasTracking As Boolean
    If Not ModBackup.HandleTrackChanges(doc, wasTracking) Then Exit Sub
    ModMain.g_IsEngineRunning = True
    ModRibbonCallbacks.InvalidateRibbonSafe
    Dim objUndo As Object, weStartedUndo As Boolean, hadErr As Boolean, errDesc As String
    weStartedUndo = False
    hadErr = False
    errDesc = ""
    On Error GoTo ErrorHandler
    weStartedUndo = ModBackup.StartUndoSafe("Clean Markers for Submission", objUndo)
    Dim mStarts() As Long, mEnds() As Long, mCount As Long, dEnd As Long, mEnd As Long, k As Long
    Dim fRng As Range, pRng As Range, pTxt As String, prevEnd As Long, lastParaStart As Long
    ReDim mStarts(1 To 64)
    ReDim mEnds(1 To 64)
    dEnd = doc.Range.End
    Set fRng = doc.Content.Duplicate
    ModMathHelpers.InitDomFind fRng.Find, "="
    fRng.Find.MatchCase = False
    prevEnd = -1
    lastParaStart = -1
    Do While fRng.Find.Execute
        If fRng.End <= prevEnd Then Exit Do
        prevEnd = fRng.End
        Set pRng = fRng.Paragraphs(1).Range
        If pRng.Start > lastParaStart Then
            lastParaStart = pRng.Start
            pTxt = pRng.Text
            If IsMarkerString(pTxt) Then
                If Right$(pTxt, 2) = vbCr & Chr(7) Then
                    mEnd = pRng.End - 2
                Else
                    mEnd = pRng.End
                End If
                If mEnd >= dEnd And dEnd > 0 Then mEnd = dEnd - 1
                mEnd = TrimTrailingBreak(doc, pRng.Start, mEnd)
                If mEnd > pRng.Start Then
                    mCount = mCount + 1
                    If mCount > UBound(mStarts) Then
                        ReDim Preserve mStarts(1 To mCount + 32)
                        ReDim Preserve mEnds(1 To mCount + 32)
                    End If
                    mStarts(mCount) = pRng.Start
                    mEnds(mCount) = mEnd
                End If
            End If
        End If
        If pRng.End > prevEnd Then prevEnd = pRng.End
        If pRng.End >= dEnd Then Exit Do
        fRng.SetRange pRng.End, dEnd
    Loop
    Dim oldLen As Long
    For k = mCount To 1 Step -1
        oldLen = doc.Range.End
        doc.Range(mStarts(k), mEnds(k)).Delete
        If (oldLen - doc.Range.End) < (mEnds(k) - mStarts(k)) And mStarts(k) + 1 < doc.Range.End Then
            If doc.Range(mStarts(k), mStarts(k) + 1).Text = vbCr Then doc.Range(mStarts(k), mStarts(k) + 1).Delete
        End If
    Next k
    On Error Resume Next
    If doc.Bookmarks.Exists(BM_ZONE_TOC) Then doc.Bookmarks(BM_ZONE_TOC).Delete
    If doc.Bookmarks.Exists(BM_ZONE_LOT) Then doc.Bookmarks(BM_ZONE_LOT).Delete
    If doc.Bookmarks.Exists(BM_ZONE_LOF) Then doc.Bookmarks(BM_ZONE_LOF).Delete
    If doc.Bookmarks.Exists(BM_ZONE_LOA) Then doc.Bookmarks(BM_ZONE_LOA).Delete
    If doc.Bookmarks.Exists(BM_SYS_GENERATED_BLOCK) Then doc.Bookmarks(BM_SYS_GENERATED_BLOCK).Delete
    If Err.Number <> 0 Then ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
    Err.Clear
    On Error GoTo ErrorHandler
    ModMathController.FinalizeAllMath doc, True
CleanExit:
    On Error Resume Next
    ModBackup.EndUndoSafe objUndo, weStartedUndo
    ModBackup.RestoreTrackChanges doc, wasTracking
    ModTableIndex.InvalidateAllIndexes
    ModMain.g_IsEngineRunning = False
    ModRibbonCallbacks.InvalidateRibbonSafe
    If hadErr Then
        ModUI.MsgBoxW ModConstants.Uni("Marker cleanup error (L\u1ed7i d\u1ecdn th\u1ebb): ") & errDesc, _
                      vbCritical, ModConstants.Uni("Error (L\u1ed7i)")
    Else
        ModUI.MsgBoxW ModConstants.Uni("Cleaned ") & mCount & _
                      ModConstants.Uni(" markers. Bookmarks and headings preserved. Press Ctrl+Z to undo. (\u0110\u00e3 x\u00f3a ") & mCount & _
                      ModConstants.Uni(" th\u1ebb. Li\u00ean k\u1ebft v\u00e0 ti\u00eau \u0111\u1ec1 \u0111\u01b0\u1ee3c gi\u1eef nguy\u00ean. Nh\u1ea5n Ctrl+Z \u0111\u1ec3 ho\u00e0n t\u00e1c.)"), _
                      vbInformation, ModConstants.Uni("Clean Complete (D\u1ecdn d\u1eb9p xong)")
    End If
    Exit Sub
ErrorHandler:
    hadErr = True
    errDesc = Err.Description
    Resume CleanExit
End Sub

Public Sub SafeDelBm(ByVal doc As Document, ByVal bmName As String)
    On Error Resume Next
    If doc.Bookmarks.Exists(bmName) Then doc.Bookmarks(bmName).Delete
    On Error GoTo 0
End Sub

Private Function FindParaByPos(ByRef pStarts() As Long, ByVal maxIdx As Long, ByVal posVal As Long) As Long
    Dim lo As Long, hi As Long, mid As Long
    If maxIdx <= 0 Then Exit Function
    lo = 1
    hi = maxIdx
    Do While lo < hi
        mid = (lo + hi + 1) \ 2
        If pStarts(mid) <= posVal Then lo = mid Else hi = mid - 1
    Loop
    If pStarts(lo) <= posVal Then FindParaByPos = lo
End Function

Public Function FindParaIndexByBookmark(ByVal doc As Document, ByVal bmName As String) As Long
    If Not doc.Bookmarks.Exists(bmName) Then Exit Function
    Dim bmRng As Range, pos As Long, docEnd As Long
    Set bmRng = doc.Bookmarks(bmName).Range
    If bmRng.End - 1 > bmRng.Start Then pos = bmRng.End - 1 Else pos = bmRng.Start + 1
    docEnd = doc.Range.End
    If pos > docEnd Then pos = docEnd
    FindParaIndexByBookmark = doc.Range(0, pos).Paragraphs.Count
End Function

Private Function IsQuickEntryPara(ByVal p As Paragraph) As Boolean
    On Error GoTo ErrCheck
    If p.Range.Hyperlinks.Count <= 0 Then Exit Function
    Dim subAddr As String
    On Error Resume Next
    subAddr = LCase$(p.Range.Hyperlinks(1).SubAddress)
    On Error GoTo ErrCheck
    If Left$(subAddr, 7) = "bm_toc_" Or Left$(subAddr, 7) = "bm_tbl_" Or Left$(subAddr, 7) = "bm_fig_" _
       Or subAddr = BM_TOC_HEADER Or subAddr = BM_LOT_HEADER Or subAddr = BM_LOF_HEADER Then IsQuickEntryPara = True
    Exit Function
ErrCheck:
    IsQuickEntryPara = False
End Function

Private Sub AddDelSpan(ByRef dStarts() As Long, ByRef dEnds() As Long, ByRef dCount As Long, _
                       ByRef pStarts() As Long, ByRef pEnds() As Long, ByVal pS As Long, ByVal pE As Long, ByVal skipP As Long)
    If pS <= 0 Or pE < pS Then Exit Sub
    If skipP >= pS And skipP <= pE Then
        If skipP > pS Then AddDelSpan dStarts, dEnds, dCount, pStarts, pEnds, pS, skipP - 1, 0
        If skipP < pE Then AddDelSpan dStarts, dEnds, dCount, pStarts, pEnds, skipP + 1, pE, 0
        Exit Sub
    End If
    dCount = dCount + 1
    dStarts(dCount) = pStarts(pS)
    dEnds(dCount) = pEnds(pE)
End Sub

Private Sub InitSlot(ByRef s As QuickZoneSlot, ByVal kindVal As Long, ByVal zBm As String, _
                     ByVal hBm As String, ByVal aBm As String, ByVal tBm As String, ByVal mPara As Long)
    s.Kind = kindVal
    s.ZoneBm = zBm
    s.HeaderBm = hBm
    s.AnchorBm = aBm
    s.HeaderTrackBm = tBm
    s.MarkerPara = mPara
    s.Found = (mPara > 0)
End Sub

Private Function IsSlotMarker(ByVal pIdx As Long, ByRef markers As MarkerLocations) As Boolean
    IsSlotMarker = (pIdx = markers.ContentStartParaIndex Or pIdx = markers.ContentEndParaIndex Or _
                    pIdx = markers.ListContentParaIndex Or pIdx = markers.ListTableParaIndex Or _
                    pIdx = markers.ListFigureParaIndex Or pIdx = markers.ListAcronymsParaIndex)
End Function

Private Function FindCompactIdxByPos(ByRef cStarts() As Long, ByRef cEnds() As Long, _
                                     ByVal maxIdx As Long, ByVal posVal As Long) As Long
    Dim ci As Long
    ci = FindParaByPos(cStarts, maxIdx, posVal)
    If ci > 0 Then
        If posVal < cEnds(ci) Then FindCompactIdxByPos = ci
    End If
End Function

Private Function ParaIdxToCompact(ByRef cParaIdx() As Long, ByVal maxIdx As Long, ByVal pIdx As Long) As Long
    If maxIdx <= 0 Or pIdx <= 0 Then Exit Function
    Dim lo As Long, hi As Long, mid As Long
    lo = 1
    hi = maxIdx
    Do While lo <= hi
        mid = lo + (hi - lo) \ 2
        If cParaIdx(mid) = pIdx Then
            ParaIdxToCompact = mid
            Exit Function
        ElseIf cParaIdx(mid) < pIdx Then
            lo = mid + 1
        Else
            hi = mid - 1
        End If
    Loop
End Function

Private Function IsWordTocOrTofAfterPara(ByVal doc As Document, ByVal pIdx As Long) As Boolean
    On Error Resume Next
    If doc.TablesOfContents.Count = 0 And doc.TablesOfFigures.Count = 0 Then Exit Function
    Dim checkStart As Long, checkEnd As Long, toc As TableOfContents, tof As TableOfFigures
    If pIdx < 1 Or pIdx > doc.Paragraphs.Count Then Exit Function
    checkStart = doc.Paragraphs(pIdx).Range.Start
    If pIdx < doc.Paragraphs.Count Then
        checkEnd = doc.Paragraphs(pIdx + 1).Range.Start
    Else
        checkEnd = doc.Paragraphs(pIdx).Range.End
    End If
    For Each toc In doc.TablesOfContents
        If (checkStart >= toc.Range.Start And checkStart < toc.Range.End) Or _
           (checkEnd >= toc.Range.Start And checkEnd < toc.Range.End) Then
            IsWordTocOrTofAfterPara = True
            Exit Function
        End If
    Next toc
    For Each tof In doc.TablesOfFigures
        If (checkStart >= tof.Range.Start And checkStart < tof.Range.End) Or _
           (checkEnd >= tof.Range.Start And checkEnd < tof.Range.End) Then
            IsWordTocOrTofAfterPara = True
            Exit Function
        End If
    Next tof
    Err.Clear
End Function

Public Function DiscoverAndPrepareQuickZones(ByVal doc As Document, ByRef markers As MarkerLocations, _
                                             ByRef slots() As QuickZoneSlot) As Boolean
    Dim totalParas As Long, csIdx As Long, ceIdx As Long
    totalParas = doc.Paragraphs.Count
    If totalParas < 1 Then Exit Function
    csIdx = markers.ContentStartParaIndex
    ceIdx = markers.ContentEndParaIndex
    ReDim slots(1 To 3)
    InitSlot slots(1), 1, BM_ZONE_TOC, BM_TOC_HEADER, "bm_qu_anc_toc", "bm_qu_hdr_toc", markers.ListContentParaIndex
    InitSlot slots(2), 2, BM_ZONE_LOT, BM_LOT_HEADER, "bm_qu_anc_lot", "bm_qu_hdr_lot", markers.ListTableParaIndex
    InitSlot slots(3), 3, BM_ZONE_LOF, BM_LOF_HEADER, "bm_qu_anc_lof", "bm_qu_hdr_lof", markers.ListFigureParaIndex
    ModTableIndex.EnsureTableIndex doc
    Dim preCount As Long, postCount As Long, maxK As Long
    If csIdx > 1 Then
        preCount = csIdx - 1
        If preCount > totalParas Then preCount = totalParas
    End If
    If ceIdx > 0 And ceIdx < totalParas Then
        postCount = totalParas - ceIdx
    End If
    maxK = preCount + postCount
    Dim cStarts() As Long, cEnds() As Long, cParaIdx() As Long, cTexts() As String, cFlags() As Long
    Dim kCount As Long, actualPreCount As Long
    kCount = 0
    actualPreCount = 0
    If maxK > 0 Then
        ReDim cStarts(1 To maxK)
        ReDim cEnds(1 To maxK)
        ReDim cParaIdx(1 To maxK)
        ReDim cTexts(1 To maxK)
        ReDim cFlags(1 To maxK)
        Dim scanRng As Range, p As Paragraph, pRng As Range, rawTxt As String, flg As Long
        Dim csStart As Long, ceEnd As Long, docEnd As Long, pNum As Long
        If preCount > 0 And csIdx <= totalParas Then
            csStart = doc.Paragraphs(csIdx).Range.Start
            If csStart > 0 Then
                Set scanRng = doc.Range(0, csStart)
                pNum = 0
                For Each p In scanRng.Paragraphs
                    Set pRng = p.Range
                    If pRng.Start >= csStart Then Exit For
                    pNum = pNum + 1
                    If pNum > preCount Then Exit For
                    kCount = kCount + 1
                    cStarts(kCount) = pRng.Start
                    cEnds(kCount) = pRng.End
                    cParaIdx(kCount) = pNum
                    rawTxt = pRng.Text
                    cTexts(kCount) = Trim$(Replace(Replace(Replace(rawTxt, vbCr, ""), Chr(7), ""), Chr(11), " "))
                    flg = 0
                    If ModTableIndex.IsPosInTable(cStarts(kCount)) Then flg = flg Or 1
                    If InStr(1, rawTxt, Chr(12)) > 0 Then flg = flg Or 8
                    If pRng.Hyperlinks.Count > 0 Then
                        flg = flg Or 2
                        If IsQuickEntryPara(p) Then flg = flg Or 4
                    End If
                    cFlags(kCount) = flg
                Next p
                actualPreCount = kCount
            End If
        End If
        If postCount > 0 Then
            ceEnd = doc.Paragraphs(ceIdx).Range.End
            docEnd = doc.Range.End
            If ceEnd < docEnd Then
                Set scanRng = doc.Range(ceEnd, docEnd)
                pNum = ceIdx
                For Each p In scanRng.Paragraphs
                    Set pRng = p.Range
                    If pRng.Start < ceEnd Then GoTo SkipPostPara
                    pNum = pNum + 1
                    If pNum > totalParas Then Exit For
                    kCount = kCount + 1
                    If kCount > maxK Then
                        maxK = kCount
                        ReDim Preserve cStarts(1 To maxK)
                        ReDim Preserve cEnds(1 To maxK)
                        ReDim Preserve cParaIdx(1 To maxK)
                        ReDim Preserve cTexts(1 To maxK)
                        ReDim Preserve cFlags(1 To maxK)
                    End If
                    cStarts(kCount) = pRng.Start
                    cEnds(kCount) = pRng.End
                    cParaIdx(kCount) = pNum
                    rawTxt = pRng.Text
                    cTexts(kCount) = Trim$(Replace(Replace(Replace(rawTxt, vbCr, ""), Chr(7), ""), Chr(11), " "))
                    flg = 0
                    If ModTableIndex.IsPosInTable(cStarts(kCount)) Then flg = flg Or 1
                    If InStr(1, rawTxt, Chr(12)) > 0 Then flg = flg Or 8
                    If pRng.Hyperlinks.Count > 0 Then
                        flg = flg Or 2
                        If IsQuickEntryPara(p) Then flg = flg Or 4
                    End If
                    cFlags(kCount) = flg
SkipPostPara:
                Next p
            End If
        End If
    End If
    Dim regEx As Object, patKind(1 To 3) As String
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False
    patKind(1) = ModLocale.GetPatternManualToc()
    patKind(2) = ModLocale.GetPatternManualLot()
    patKind(3) = ModLocale.GetPatternManualLof()
    Dim hdrC(1 To 3) As Long, bodySC(1 To 3) As Long, bodyEC(1 To 3) As Long
    Dim k As Long, zRng As Range, ci As Long, j As Long, bStart As Long, bEnd As Long, validEnd As Long, scanEnd As Long
    If kCount > 0 Then
        For k = 1 To 3
            If doc.Bookmarks.Exists(slots(k).HeaderBm) Then
                ci = FindCompactIdxByPos(cStarts, cEnds, kCount, doc.Bookmarks(slots(k).HeaderBm).Range.Start)
                If ci > 0 Then
                    hdrC(k) = ci
                    slots(k).HdrStart = cParaIdx(ci)
                End If
            End If
            If doc.Bookmarks.Exists(slots(k).ZoneBm) Then
                Set zRng = doc.Bookmarks(slots(k).ZoneBm).Range
                If zRng.End > zRng.Start Then
                    bStart = FindCompactIdxByPos(cStarts, cEnds, kCount, zRng.Start)
                    If bStart > 0 Then
                        bEnd = FindParaByPos(cStarts, kCount, zRng.End - 1)
                        If bStart <= actualPreCount And bEnd > actualPreCount Then bEnd = actualPreCount
                        validEnd = 0
                        For j = bStart To bEnd
                            If IsSlotMarker(cParaIdx(j), markers) Or (cFlags(j) And 4) = 0 Then Exit For
                            validEnd = j
                            If (cFlags(j) And 8) <> 0 Then Exit For
                        Next j
                        If validEnd >= bStart Then
                            bodySC(k) = bStart
                            bodyEC(k) = validEnd
                            slots(k).BodyStart = cParaIdx(bStart)
                            slots(k).BodyEnd = cParaIdx(validEnd)
                        End If
                    End If
                End If
            End If
            If hdrC(k) = 0 Then
                regEx.Pattern = patKind(k)
                For ci = 1 To kCount
                    If Not IsSlotMarker(cParaIdx(ci), markers) And (cFlags(ci) And 3) = 0 Then
                        If regEx.Test(cTexts(ci)) Then
                            If Not IsWordTocOrTofAfterPara(doc, cParaIdx(ci)) Then
                                hdrC(k) = ci
                                slots(k).HdrStart = cParaIdx(ci)
                                Exit For
                            End If
                        End If
                    End If
                Next ci
            End If
            If hdrC(k) > 0 And bodyEC(k) = 0 Then
                regEx.Pattern = ModLocale.GetPatternManualList()
                If hdrC(k) <= actualPreCount Then scanEnd = actualPreCount Else scanEnd = kCount
                For j = hdrC(k) + 1 To scanEnd
                    If cParaIdx(j) <> slots(k).MarkerPara Then
                        If IsSlotMarker(cParaIdx(j), markers) Then Exit For
                        If (cFlags(j) And 2) = 0 And regEx.Test(cTexts(j)) Then Exit For
                        If (cFlags(j) And 4) = 0 Then Exit For
                        If bodySC(k) = 0 Then
                            bodySC(k) = j
                            slots(k).BodyStart = cParaIdx(j)
                        End If
                        bodyEC(k) = j
                        slots(k).BodyEnd = cParaIdx(j)
                        If (cFlags(j) And 8) <> 0 Then Exit For
                    End If
                Next j
            End If
            If hdrC(k) > 0 Or bodyEC(k) > 0 Then slots(k).Found = True
        Next k
    End If
    If Not (slots(1).Found Or slots(2).Found Or slots(3).Found) Then Exit Function
    Dim dStarts(1 To 12) As Long, dEnds(1 To 12) As Long, dCount As Long, ancC As Long, bS As Long, skipC As Long
    For k = 1 To 3
        If slots(k).Found And slots(k).MarkerPara > 0 Then
            skipC = ParaIdxToCompact(cParaIdx, kCount, slots(k).MarkerPara)
            AddDelSpan dStarts, dEnds, dCount, cStarts, cEnds, bodySC(k), bodyEC(k), skipC
            If hdrC(k) > 0 Then
                If cParaIdx(hdrC(k)) <> slots(k).MarkerPara Then
                    AddDelSpan dStarts, dEnds, dCount, cStarts, cEnds, hdrC(k), hdrC(k), skipC
                End If
            End If
        ElseIf slots(k).Found Then
            If hdrC(k) > 0 Then ancC = hdrC(k) Else ancC = bodySC(k)
            If ancC > 0 Then
                ModBookmark.AddBookmark doc, slots(k).AnchorBm, doc.Range(cStarts(ancC), cEnds(ancC))
                bS = bodySC(k)
                If hdrC(k) = 0 Then bS = bS + 1
                AddDelSpan dStarts, dEnds, dCount, cStarts, cEnds, bS, bodyEC(k), ancC
            End If
        End If
    Next k
    If dCount > 0 Then
        Dim m As Long, n As Long, tmpS As Long, tmpE As Long, mCount As Long, delEnd As Long
        For m = 1 To dCount - 1
            For n = m + 1 To dCount
                If dStarts(n) < dStarts(m) Then
                    tmpS = dStarts(m)
                    dStarts(m) = dStarts(n)
                    dStarts(n) = tmpS
                    tmpE = dEnds(m)
                    dEnds(m) = dEnds(n)
                    dEnds(n) = tmpE
                End If
            Next n
        Next m
        mCount = 1
        For m = 2 To dCount
            If dStarts(m) < dEnds(mCount) Then
                If dEnds(m) > dEnds(mCount) Then dEnds(mCount) = dEnds(m)
            ElseIf dStarts(m) = dEnds(mCount) And InStr(1, doc.Range(dStarts(mCount), dEnds(mCount)).Text, Chr(12)) = 0 Then
                If dEnds(m) > dEnds(mCount) Then dEnds(mCount) = dEnds(m)
            Else
                mCount = mCount + 1
                dStarts(mCount) = dStarts(m)
                dEnds(mCount) = dEnds(m)
            End If
        Next m
        For m = mCount To 1 Step -1
            delEnd = TrimTrailingBreak(doc, dStarts(m), dEnds(m))
            If delEnd > dStarts(m) Then doc.Range(dStarts(m), delEnd).Delete
        Next m
    End If
    For k = 1 To 3
        If slots(k).Found And slots(k).MarkerPara = 0 And doc.Bookmarks.Exists(slots(k).AnchorBm) Then
            ModQuickUpdate.ResetAnchorPara doc, FindParaIndexByBookmark(doc, slots(k).AnchorBm), slots(k).AnchorBm
        End If
        SafeDelBm doc, slots(k).ZoneBm
        SafeDelBm doc, slots(k).HeaderBm
    Next k
    SafeDelBm doc, BM_SYS_GENERATED_BLOCK
    ModTableIndex.InvalidateAllIndexes
    DiscoverAndPrepareQuickZones = True
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
