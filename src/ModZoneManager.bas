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
        If Right$(rTxt, 1) = Chr(12) Then
            ePos = ePos - 1
        ElseIf Right$(rTxt, 2) = Chr(12) & vbCr Then
            ePos = ePos - 2
        Else
            Set lp = doc.Range(ePos - 1, ePos).Paragraphs(1).Range
            If InStr(1, lp.Text, Chr(12)) > 0 Then ePos = lp.Start
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
    If doc Is Nothing And Documents.Count > 0 Then Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub
    Dim objUndo As Object, weStartedUndo As Boolean
    On Error Resume Next
    Set objUndo = Application.UndoRecord
    On Error GoTo ErrorHandler
    If Not objUndo Is Nothing Then
        If Not objUndo.IsRecordingCustomRecord Then
            objUndo.StartCustomRecord "Clean Markers for Submission"
            weStartedUndo = True
        End If
    End If
    Dim mStarts() As Long, mEnds() As Long, mCount As Long, dEnd As Long, mEnd As Long, p As Paragraph, k As Long
    ReDim mStarts(1 To 64)
    ReDim mEnds(1 To 64)
    dEnd = doc.Range.End
    For Each p In doc.Paragraphs
        If IsMarkerString(p.Range.Text) Then
            If p.Range.Information(12) Then mEnd = p.Range.End - 2 Else mEnd = p.Range.End
            If mEnd >= dEnd And dEnd > 0 Then mEnd = dEnd - 1
            If mEnd > p.Range.Start Then
                mCount = mCount + 1
                If mCount > UBound(mStarts) Then
                    ReDim Preserve mStarts(1 To mCount + 32)
                    ReDim Preserve mEnds(1 To mCount + 32)
                End If
                mStarts(mCount) = p.Range.Start
                mEnds(mCount) = mEnd
            End If
        End If
    Next p
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
    If weStartedUndo And Not (objUndo Is Nothing) Then If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
    ModUI.MsgBoxW ModConstants.Uni("Cleaned ") & mCount & _
                  ModConstants.Uni(" markers. Bookmarks and headings preserved. Press Ctrl+Z to undo. (\u0110\u00e3 x\u00f3a ") & mCount & _
                  ModConstants.Uni(" th\u1ebb. Li\u00ean k\u1ebft v\u00e0 ti\u00eau \u0111\u1ec1 \u0111\u01b0\u1ee3c gi\u1eef nguy\u00ean. Nh\u1ea5n Ctrl+Z \u0111\u1ec3 ho\u00e0n t\u00e1c.)"), _
                  vbInformation, ModConstants.Uni("Clean Complete (D\u1ecdn d\u1eb9p xong)")
    Exit Sub
ErrorHandler:
    If weStartedUndo And Not (objUndo Is Nothing) Then If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
    ModUI.MsgBoxW ModConstants.Uni("Marker cleanup error (L\u1ed7i d\u1ecdn th\u1ebb): ") & Err.Description, _
                  vbCritical, ModConstants.Uni("Error (L\u1ed7i)")
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
    Dim pStarts() As Long, pEnds() As Long, pTexts() As String, pFlags() As Long
    ReDim pStarts(1 To totalParas)
    ReDim pEnds(1 To totalParas)
    ReDim pTexts(1 To totalParas)
    ReDim pFlags(1 To totalParas)
    Dim i As Long, p As Paragraph, rawTxt As String, flg As Long
    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        pStarts(i) = p.Range.Start
        pEnds(i) = p.Range.End
        If i < csIdx Or (ceIdx > 0 And i > ceIdx) Then
            rawTxt = p.Range.Text
            pTexts(i) = Trim$(Replace(Replace(Replace(rawTxt, vbCr, ""), Chr(7), ""), Chr(11), " "))
            flg = 0
            If CBool(p.Range.Information(12)) Then flg = flg Or 1
            If InStr(1, rawTxt, Chr(12)) > 0 Then flg = flg Or 8
            If p.Range.Hyperlinks.Count > 0 Then
                flg = flg Or 2
                If IsQuickEntryPara(p) Then flg = flg Or 4
            End If
            pFlags(i) = flg
        End If
    Next p
    Dim regEx As Object, patKind(1 To 3) As String
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False
    patKind(1) = ModLocale.GetPatternManualToc()
    patKind(2) = ModLocale.GetPatternManualLot()
    patKind(3) = ModLocale.GetPatternManualLof()
    Dim k As Long, zRng As Range, pIdx As Long, j As Long, bStart As Long, bEnd As Long, validEnd As Long, scanEnd As Long
    For k = 1 To 3
        If doc.Bookmarks.Exists(slots(k).HeaderBm) Then
            pIdx = FindParaByPos(pStarts, totalParas, doc.Bookmarks(slots(k).HeaderBm).Range.Start)
            If pIdx > 0 And (pIdx < csIdx Or (ceIdx > 0 And pIdx > ceIdx)) Then slots(k).HdrStart = pIdx
        End If
        If doc.Bookmarks.Exists(slots(k).ZoneBm) Then
            Set zRng = doc.Bookmarks(slots(k).ZoneBm).Range
            If zRng.End > zRng.Start Then
                bStart = FindParaByPos(pStarts, totalParas, zRng.Start)
                bEnd = FindParaByPos(pStarts, totalParas, zRng.End - 1)
                validEnd = 0
                For j = bStart To bEnd
                    If j <= 0 Or (j >= csIdx And (ceIdx = 0 Or j <= ceIdx)) Or IsSlotMarker(j, markers) Or (pFlags(j) And 4) = 0 Then Exit For
                    validEnd = j
                    If (pFlags(j) And 8) <> 0 Then Exit For
                Next j
                If validEnd >= bStart And bStart > 0 Then
                    slots(k).BodyStart = bStart
                    slots(k).BodyEnd = validEnd
                End If
            End If
        End If
        If slots(k).HdrStart = 0 Then
            regEx.Pattern = patKind(k)
            For pIdx = 1 To totalParas
                If (pIdx < csIdx Or (ceIdx > 0 And pIdx > ceIdx)) And Not IsSlotMarker(pIdx, markers) And (pFlags(pIdx) And 3) = 0 Then
                    If regEx.Test(pTexts(pIdx)) Then
                        slots(k).HdrStart = pIdx
                        Exit For
                    End If
                End If
            Next pIdx
        End If
        If slots(k).HdrStart > 0 And slots(k).BodyEnd = 0 Then
            regEx.Pattern = ModLocale.GetPatternManualList()
            If slots(k).HdrStart < csIdx Then scanEnd = csIdx - 1 Else scanEnd = totalParas
            For j = slots(k).HdrStart + 1 To scanEnd
                If j <> slots(k).MarkerPara Then
                    If IsSlotMarker(j, markers) Then Exit For
                    If (pFlags(j) And 2) = 0 And regEx.Test(pTexts(j)) Then Exit For
                    If (pFlags(j) And 4) = 0 Then Exit For
                    If slots(k).BodyStart = 0 Then slots(k).BodyStart = j
                    slots(k).BodyEnd = j
                    If (pFlags(j) And 8) <> 0 Then Exit For
                End If
            Next j
        End If
        If slots(k).HdrStart > 0 Or slots(k).BodyEnd > 0 Then slots(k).Found = True
    Next k
    If Not (slots(1).Found Or slots(2).Found Or slots(3).Found) Then Exit Function
    Dim dStarts(1 To 12) As Long, dEnds(1 To 12) As Long, dCount As Long, ancPara As Long, bS As Long
    For k = 1 To 3
        If slots(k).Found And slots(k).MarkerPara > 0 Then
            AddDelSpan dStarts, dEnds, dCount, pStarts, pEnds, slots(k).BodyStart, slots(k).BodyEnd, slots(k).MarkerPara
            If slots(k).HdrStart > 0 And slots(k).HdrStart <> slots(k).MarkerPara Then
                AddDelSpan dStarts, dEnds, dCount, pStarts, pEnds, slots(k).HdrStart, slots(k).HdrStart, slots(k).MarkerPara
            End If
        ElseIf slots(k).Found Then
            If slots(k).HdrStart > 0 Then ancPara = slots(k).HdrStart Else ancPara = slots(k).BodyStart
            If ancPara > 0 Then
                ModBookmark.AddBookmark doc, slots(k).AnchorBm, doc.Range(pStarts(ancPara), pEnds(ancPara))
                bS = slots(k).BodyStart
                If slots(k).HdrStart = 0 Then bS = bS + 1
                AddDelSpan dStarts, dEnds, dCount, pStarts, pEnds, bS, slots(k).BodyEnd, ancPara
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
            If dStarts(m) <= dEnds(mCount) Then
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
    DiscoverAndPrepareQuickZones = True
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
