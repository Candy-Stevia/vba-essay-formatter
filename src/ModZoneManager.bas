Attribute VB_Name = "ModZoneManager"
Option Explicit

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

Public Sub WipeZone(ByVal doc As Document, ByVal bmZoneName As String, ByVal bmHeaderName As String)
    On Error Resume Next
    If doc.Bookmarks.Exists(bmZoneName) Then
        Dim zRng As Range
        Set zRng = doc.Bookmarks(bmZoneName).Range
        If zRng.End > zRng.Start Then
            zRng.Delete
        End If
    End If
    If doc.Bookmarks.Exists(bmHeaderName) Then
        doc.Bookmarks(bmHeaderName).Range.Paragraphs(1).Range.Delete
    End If
    If Err.Number <> 0 Then
        ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
        Err.Clear
    End If
End Sub

Public Sub WipeAllZones(ByVal doc As Document)
    ' Wipes LOF, LOT, TOC generated zones and headers in preparation for clean rebuild
    WipeZone doc, BM_ZONE_LOF, BM_LOF_HEADER
    WipeZone doc, BM_ZONE_LOT, BM_LOT_HEADER
    WipeZone doc, BM_ZONE_TOC, BM_TOC_HEADER
    
    ' Clear legacy single-block bookmark if present from v1
    On Error Resume Next
    If doc.Bookmarks.Exists(BM_SYS_GENERATED_BLOCK) Then
        doc.Bookmarks(BM_SYS_GENERATED_BLOCK).Delete
    End If
    If Err.Number <> 0 Then
        ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
        Err.Clear
    End If
End Sub

Public Sub EnsureTargetBookmarks(ByVal doc As Document, _
                                ByRef snapshots() As ParagraphSnapshot, _
                                ByVal totalCount As Long)
    ' Mode B (Lists Only): Ensures all heading, table, and figure bookmarks exist
    ' without altering body paragraph styles, fonts, or alignments
    Dim i As Long, tType As Long, pIdx As Long
    Dim countToc As Long
    countToc = 0
    Dim countTbl As Long
    countTbl = 0
    Dim countFig As Long
    countFig = 0
    Dim bmName As String
    
    For i = 1 To totalCount
        tType = snapshots(i).ElementType
        pIdx = snapshots(i).ParaIndex
        bmName = ""
        
        Select Case tType
            Case ElementTypeLevel1 To ElementTypeLevel4
                countToc = countToc + 1
                bmName = FormatBookmarkId(BM_PREFIX_TOC, countToc)
            Case ElementTypeTableTitle
                countTbl = countTbl + 1
                bmName = FormatBookmarkId(BM_PREFIX_TBL, countTbl)
            Case ElementTypeFigureTitle
                countFig = countFig + 1
                bmName = FormatBookmarkId(BM_PREFIX_FIG, countFig)
        End Select
        
        If Len(bmName) > 0 And snapshots(i).RangeStart < doc.Range.End Then
            AddBookmark doc, bmName, doc.Range(snapshots(i).RangeStart, snapshots(i).RangeEnd)
        End If
    Next i
End Sub

Public Sub CleanMarkersForSubmission(ByVal doc As Document)
    ' Submission Cleaner: Removes the 5 boundary markers and temporary zone bookmarks
    ' strictly preserves bm_toc_header, bm_lot_header, bm_lof_header and all target bookmarks
    If doc Is Nothing Then Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub
    
    Dim objUndo As Object
    On Error Resume Next
    Set objUndo = Application.UndoRecord
    On Error GoTo ErrorHandler
    If Not objUndo Is Nothing Then objUndo.StartCustomRecord "Clean Markers for Submission"
    
    Dim mStarts() As Long, mEnds() As Long, mCount As Long
    mCount = 0
    ReDim mStarts(1 To 20)
    ReDim mEnds(1 To 20)
    
    Dim p As Paragraph, cleanText As String
    For Each p In doc.Paragraphs
        cleanText = LCase$(Trim$(Replace(Replace(Replace(p.Range.Text, vbCr, ""), Chr(7), ""), Chr(11), "")))
        Select Case cleanText
            Case "=contentstart()", "=contentend()", "=listcontent()", "=listtablehere()", "=listfigurehere()"
                mCount = mCount + 1
                If mCount > UBound(mStarts) Then
                    ReDim Preserve mStarts(1 To mCount + 20)
                    ReDim Preserve mEnds(1 To mCount + 20)
                End If
                mStarts(mCount) = p.Range.Start
                mEnds(mCount) = p.Range.End
        End Select
    Next p
    
    Dim k As Long
    For k = mCount To 1 Step -1
        doc.Range(mStarts(k), mEnds(k)).Delete
    Next k
    
    ' Remove zone bookmarks but preserve bm_*_header and target bookmarks
    On Error Resume Next
    If doc.Bookmarks.Exists(BM_ZONE_TOC) Then doc.Bookmarks(BM_ZONE_TOC).Delete
    If doc.Bookmarks.Exists(BM_ZONE_LOT) Then doc.Bookmarks(BM_ZONE_LOT).Delete
    If doc.Bookmarks.Exists(BM_ZONE_LOF) Then doc.Bookmarks(BM_ZONE_LOF).Delete
    If doc.Bookmarks.Exists(BM_SYS_GENERATED_BLOCK) Then doc.Bookmarks(BM_SYS_GENERATED_BLOCK).Delete
    If Err.Number <> 0 Then
        ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
        Err.Clear
    End If
    
    ModMathController.FinalizeAllMath doc, True
    
CleanExit:
    If Not objUndo Is Nothing Then
        If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
    End If
    ModUI.MsgBoxW ModConstants.Uni("Cleaned ") & mCount & _
                  ModConstants.Uni(" markers. Bookmarks and headings preserved. Press Ctrl+Z to undo. (\u0110\u00e3 x\u00f3a ") & mCount & _
                  ModConstants.Uni(" th\u1ebb. Li\u00ean k\u1ebft v\u00e0 ti\u00eau \u0111\u1ec1 \u0111\u01b0\u1ee3c gi\u1eef nguy\u00ean. Nh\u1ea5n Ctrl+Z \u0111\u1ec3 ho\u00e0n t\u00e1c.)"), _
                  vbInformation, ModConstants.Uni("Clean Complete (D\u1ecdn d\u1eb9p xong)")
    Exit Sub
    
ErrorHandler:
    If Not objUndo Is Nothing Then
        If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
    End If
    ModUI.MsgBoxW ModConstants.Uni("Marker cleanup error (L\u1ed7i d\u1ecdn th\u1ebb): ") & Err.Description, _
                  vbCritical, ModConstants.Uni("Error (L\u1ed7i)")
End Sub

Public Sub SafeDelBm(ByVal doc As Document, ByVal bmName As String)
    On Error Resume Next
    If doc.Bookmarks.Exists(bmName) Then
        doc.Bookmarks(bmName).Delete
    End If
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
    Dim targetEnd As Long, j As Long, p As Paragraph
    targetEnd = doc.Bookmarks(bmName).Range.End
    j = 0
    For Each p In doc.Paragraphs
        j = j + 1
        If p.Range.End >= targetEnd Then
            FindParaIndexByBookmark = j
            Exit Function
        End If
    Next p
End Function

Private Function IsQuickEntryPara(ByVal p As Paragraph) As Boolean
    Dim subAddr As String
    On Error GoTo ErrCheck
    If p.Range.Hyperlinks.Count <= 0 Then Exit Function
    On Error Resume Next
    subAddr = LCase$(p.Range.Hyperlinks(1).SubAddress)
    On Error GoTo ErrCheck
    If Left$(subAddr, 7) = "bm_toc_" Or Left$(subAddr, 7) = "bm_tbl_" Or Left$(subAddr, 7) = "bm_fig_" Then
        IsQuickEntryPara = True
    ElseIf subAddr = BM_TOC_HEADER Or subAddr = BM_LOT_HEADER Or subAddr = BM_LOF_HEADER Then
        IsQuickEntryPara = True
    End If
    Exit Function
ErrCheck:
    IsQuickEntryPara = False
End Function

Private Sub AddDelSpan(ByRef dStarts() As Long, ByRef dEnds() As Long, ByRef dCount As Long, _
                       ByVal doc As Document, ByVal pStart As Long, ByVal pEnd As Long, ByVal skipPara As Long)
    Dim curS As Long, k As Long
    If pStart <= 0 Or pEnd < pStart Then Exit Sub
    curS = 0
    For k = pStart To pEnd
        If k = skipPara Then
            If curS > 0 Then
                dCount = dCount + 1
                dStarts(dCount) = doc.Paragraphs(curS).Range.Start
                dEnds(dCount) = doc.Paragraphs(k - 1).Range.End
                curS = 0
            End If
        ElseIf curS = 0 Then
            curS = k
        End If
    Next k
    If curS > 0 Then
        dCount = dCount + 1
        dStarts(dCount) = doc.Paragraphs(curS).Range.Start
        dEnds(dCount) = doc.Paragraphs(pEnd).Range.End
    End If
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
                    pIdx = markers.ListFigureParaIndex)
End Function

Public Function DiscoverAndPrepareQuickZones(ByVal doc As Document, ByRef markers As MarkerLocations, _
                                             ByRef slots() As QuickZoneSlot) As Boolean
    Dim frontLimit As Long
    frontLimit = markers.ContentStartParaIndex - 1
    If frontLimit < 1 Then Exit Function
    ReDim slots(1 To 3)
    InitSlot slots(1), 1, BM_ZONE_TOC, BM_TOC_HEADER, "bm_qu_anc_toc", "bm_qu_hdr_toc", markers.ListContentParaIndex
    InitSlot slots(2), 2, BM_ZONE_LOT, BM_LOT_HEADER, "bm_qu_anc_lot", "bm_qu_hdr_lot", markers.ListTableParaIndex
    InitSlot slots(3), 3, BM_ZONE_LOF, BM_LOF_HEADER, "bm_qu_anc_lof", "bm_qu_hdr_lof", markers.ListFigureParaIndex

    Dim pStarts() As Long, pEnds() As Long, pTexts() As String, i As Long, p As Paragraph
    ReDim pStarts(1 To frontLimit)
    ReDim pEnds(1 To frontLimit)
    ReDim pTexts(1 To frontLimit)
    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        If i > frontLimit Then Exit For
        pStarts(i) = p.Range.Start
        pEnds(i) = p.Range.End
        pTexts(i) = Trim$(Replace(Replace(Replace(p.Range.Text, vbCr, ""), Chr(7), ""), Chr(11), " "))
    Next p

    Dim regEx As Object, patAny As String, patKind(1 To 3) As String
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False
    patAny = ModLocale.GetPatternManualList()
    patKind(1) = ModLocale.GetPatternManualToc()
    patKind(2) = ModLocale.GetPatternManualLot()
    patKind(3) = ModLocale.GetPatternManualLof()

    Dim k As Long, zRng As Range, hRng As Range, pIdx As Long, j As Long
    For k = 1 To 3
        If doc.Bookmarks.Exists(slots(k).HeaderBm) Then
            Set hRng = doc.Bookmarks(slots(k).HeaderBm).Range
            If hRng.Start < pEnds(frontLimit) Then
                slots(k).HdrStart = FindParaByPos(pStarts, frontLimit, hRng.Start)
                If slots(k).HdrStart > 0 Then slots(k).Found = True
            End If
        End If
        If doc.Bookmarks.Exists(slots(k).ZoneBm) Then
            Set zRng = doc.Bookmarks(slots(k).ZoneBm).Range
            If zRng.End > zRng.Start And zRng.Start < pEnds(frontLimit) Then
                slots(k).BodyStart = FindParaByPos(pStarts, frontLimit, zRng.Start)
                slots(k).BodyEnd = FindParaByPos(pStarts, frontLimit, zRng.End - 1)
                If slots(k).BodyStart > 0 And slots(k).BodyEnd >= slots(k).BodyStart Then slots(k).Found = True
            End If
        End If
        If slots(k).HdrStart = 0 Then
            regEx.Pattern = patKind(k)
            For pIdx = 1 To frontLimit
                If Not IsSlotMarker(pIdx, markers) Then
                    If doc.Paragraphs(pIdx).Range.Hyperlinks.Count = 0 And regEx.Test(pTexts(pIdx)) Then
                        slots(k).HdrStart = pIdx
                        slots(k).Found = True
                        Exit For
                    End If
                End If
            Next pIdx
        End If
        If slots(k).HdrStart > 0 And slots(k).BodyEnd = 0 Then
            regEx.Pattern = patAny
            For j = slots(k).HdrStart + 1 To frontLimit
                If j <> slots(k).MarkerPara Then
                    If IsSlotMarker(j, markers) Then Exit For
                    If doc.Paragraphs(j).Range.Hyperlinks.Count = 0 And regEx.Test(pTexts(j)) Then Exit For
                    If Not IsQuickEntryPara(doc.Paragraphs(j)) Then Exit For
                    If slots(k).BodyStart = 0 Then slots(k).BodyStart = j
                    slots(k).BodyEnd = j
                End If
            Next j
        End If
    Next k
    If Not (slots(1).Found Or slots(2).Found Or slots(3).Found) Then Exit Function

    Dim dStarts(1 To 12) As Long, dEnds(1 To 12) As Long, dCount As Long, ancPara As Long
    dCount = 0
    For k = 1 To 3
        If slots(k).Found And slots(k).MarkerPara > 0 Then
            AddDelSpan dStarts, dEnds, dCount, doc, slots(k).BodyStart, slots(k).BodyEnd, slots(k).MarkerPara
            If slots(k).HdrStart > 0 And slots(k).HdrStart <> slots(k).MarkerPara Then
                AddDelSpan dStarts, dEnds, dCount, doc, slots(k).HdrStart, slots(k).HdrStart, slots(k).MarkerPara
            End If
        ElseIf slots(k).Found Then
            If slots(k).HdrStart > 0 Then ancPara = slots(k).HdrStart Else ancPara = slots(k).BodyStart
            If ancPara > 0 Then
                ModBookmark.AddBookmark doc, slots(k).AnchorBm, doc.Paragraphs(ancPara).Range
                If slots(k).HdrStart > 0 Then
                    AddDelSpan dStarts, dEnds, dCount, doc, slots(k).BodyStart, slots(k).BodyEnd, ancPara
                Else
                    AddDelSpan dStarts, dEnds, dCount, doc, slots(k).BodyStart + 1, slots(k).BodyEnd, ancPara
                End If
            End If
        End If
    Next k

    Dim m As Long, n As Long, tmpS As Long, tmpE As Long, prevCut As Long
    For m = 1 To dCount - 1
        For n = m + 1 To dCount
            If dStarts(n) > dStarts(m) Then
                tmpS = dStarts(m)
                dStarts(m) = dStarts(n)
                dStarts(n) = tmpS
                tmpE = dEnds(m)
                dEnds(m) = dEnds(n)
                dEnds(n) = tmpE
            End If
        Next n
    Next m
    prevCut = doc.Range.End + 1
    For m = 1 To dCount
        If dEnds(m) <= prevCut And dEnds(m) > dStarts(m) Then
            doc.Range(dStarts(m), dEnds(m)).Delete
            prevCut = dStarts(m)
        End If
    Next m

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