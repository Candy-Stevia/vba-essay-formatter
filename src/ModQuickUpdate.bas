Attribute VB_Name = "ModQuickUpdate"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModQuickUpdate
' Purpose: High-speed Quick Update for TOC, LOT, LOF (Marker & Markerless 3-Tier)
' ==============================================================================

Private Type QuickItem
    ElementType As Long
    CleanText As String
    BookmarkName As String
End Type

Public Sub ResetAnchorPara(ByVal doc As Document, ByVal pIdx As Long, ByVal anchorBm As String)
    Dim pRng As Range
    Set pRng = doc.Paragraphs(pIdx).Range
    If pRng.End - 1 > pRng.Start And InStr(1, pRng.Text, Chr(12)) = 0 Then
        doc.Range(pRng.Start, pRng.End - 1).Delete
    End If
    Set pRng = doc.Paragraphs(pIdx).Range
    ModFormatSingle.EnsureOutlineLevelDetachHeading doc, pRng, 10
    pRng.HighlightColorIndex = 0
    pRng.Case = 0
    pRng.Font.Bold = False
    pRng.Font.Italic = False
    pRng.ParagraphFormat.Alignment = 0
    pRng.ParagraphFormat.FirstLineIndent = 0#
    pRng.ParagraphFormat.LeftIndent = 0#
    pRng.ParagraphFormat.TabStops.ClearAll
    ModBookmark.AddBookmark doc, anchorBm, pRng
End Sub

Private Function IsMathBlockPara(ByVal rng As Range, ByVal docHasMath As Boolean) As Boolean
    If Not docHasMath Then Exit Function
    On Error GoTo ErrMath
    If rng.OMaths.Count > 0 Then If rng.OMaths(1).Type = 0 Or rng.OMaths(1).Range.Text = rng.Text Then IsMathBlockPara = True
    Exit Function
ErrMath:
    IsMathBlockPara = False
End Function

Private Function IsOldTocPara(ByVal rStart As Long, ByRef sArr() As Long, ByRef eArr() As Long, ByVal n As Long) As Boolean
    Dim k As Long
    For k = 1 To n
        If rStart >= sArr(k) And rStart < eArr(k) Then GoTo Hit
    Next k
    Exit Function
Hit:
    IsOldTocPara = True
End Function

Private Function CleanQuickText(ByVal raw As String) As String
    Dim s As String
    s = Replace(Replace(Replace(raw, vbCr, ""), Chr(7), ""), Chr(11), " ")
    s = Replace(Replace(Replace(s, Chr(1), ""), Chr(2), ""), Chr(5), "")
    CleanQuickText = Trim$(Replace(Replace(Replace(s, vbTab, " "), Chr(10), " "), Chr(12), " "))
End Function

Private Sub ScanAndFormatQuickItems(ByVal doc As Document, ByRef markers As MarkerLocations, ByRef slots() As QuickZoneSlot, _
                                    ByRef items() As QuickItem, ByRef itemCount As Long, ByRef cntHead As Long, _
                                    ByRef cntTbl As Long, ByRef cntFig As Long, ByRef cntNotes As Long, _
                                    ByRef qSnaps() As ParagraphSnapshot, ByRef qCount As Long)
    Dim qp As QuickRegexPool, patFig As String
    ModScanner.InitQuickRegexPool qp
    patFig = ModLocale.GetPatternFigure()
    Dim wantXref As Boolean, reXrefQuick As Object
    wantXref = (g_Config.CrossRefColor <> "xref_off")
    If wantXref Then
        Set reXrefQuick = CreateObject("VBScript.RegExp")
        reXrefQuick.IgnoreCase = True
        reXrefQuick.Global = False
        reXrefQuick.Pattern = ModCrossRef.BuildCrossRefPattern()
    End If
    Dim totalParas As Long, latexFlags() As Boolean, latexWarn As Long, docHasMath As Boolean
    totalParas = doc.Paragraphs.Count
    qCount = 0
    Dim qCap As Long
    qCap = 64
    ReDim items(1 To 64)
    ReDim qSnaps(1 To qCap)
    ModTableIndex.BuildTableIndex doc
    ModTableIndex.BuildListIndex doc
    ModTableIndex.BuildOMathIndex doc
    ModLatex.ComputeLatexFlags doc, latexFlags, latexWarn
    docHasMath = (doc.OMaths.Count > 0)
    Dim hasTags As Boolean, numOldTocs As Long, tocS() As Long, tocE() As Long
    hasTags = ModCustomProfile.HasActiveTags()
    ModScanner.CollectOldListRanges doc, tocS, tocE, numOldTocs
    Dim i As Long, p As Paragraph, rng As Range, raw As String, clean As String
    Dim inContent As Boolean, inPost As Boolean, tType As Long, bmName As String, inTbl As Boolean, isSkip As Boolean, isListPara As Boolean
    Dim isCandLen As Boolean, hasXref As Boolean, tagCode As Long, rStart As Long, rEnd As Long, hasParaMath As Boolean
    Dim sbPos As Long, gateLen As Long
    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        ModProgress.ProgressUpdate i, totalParas
        If ModProgress.g_CancelRequested Then Exit Sub
        Set rng = p.Range
        raw = rng.Text
        If Len(raw) <= 2 And Not hasTags Then GoTo NextQuickPara
        sbPos = InStr(1, raw, Chr(11))
        If sbPos > 0 Then
            gateLen = sbPos - 1
        Else
            gateLen = Len(raw)
        End If
        isCandLen = (gateLen > 2 And gateLen <= 350)
        hasXref = False
        If wantXref And i > markers.ContentStartParaIndex And Len(raw) > 5 Then hasXref = reXrefQuick.Test(raw)
        tagCode = 0
        If hasTags Then
            rStart = rng.Start
            rEnd = rng.End
            tagCode = ModCustomProfile.LookupTagForRange(rStart, rEnd)
        End If
        If isCandLen Or (tagCode <> 0) Or hasXref Then
            If Not hasTags Then
                rStart = rng.Start
                rEnd = rng.End
            End If
            inContent = (i > markers.ContentStartParaIndex And i < markers.ContentEndParaIndex)
            inPost = (i > markers.ContentEndParaIndex)
            tType = ElementTypeUnknown
            hasParaMath = False
            If docHasMath Then hasParaMath = ModTableIndex.HasOMathInSpan(rStart, rEnd)
            isSkip = latexFlags(i) Or IsOldTocPara(rStart, tocS, tocE, numOldTocs)
            If Not isSkip And hasParaMath Then isSkip = IsMathBlockPara(rng, True)
            inTbl = ModTableIndex.IsPosInTable(rStart)
            isListPara = ModTableIndex.IsPosInList(rStart)
            If Not isSkip And (isCandLen Or tagCode <> 0) Then
                clean = CleanQuickText(raw)
                If Len(clean) > 0 Then
                    If tagCode > 0 Then
                        tType = tagCode
                    ElseIf tagCode = TAG_CODE_REF Or tagCode = TAG_CODE_APP Then
                        tType = ElementTypeLevel1
                    ElseIf tagCode = TAG_CODE_REF_SUB Then
                        tType = ElementTypeLevel2
                    ElseIf tagCode = TAG_CODE_SKIP Then
                        tType = ElementTypeUnknown
                        ModFormatSingle.EnsureOutlineLevelDetachHeading doc, rng, 10
                    ElseIf isListPara Then
                        ' List items are excluded from heading/caption classification, but stay eligible for Cross-Ref
                    ElseIf Left$(clean, 1) = "=" And qp.ReMrk.Test(clean) Then
                    ElseIf inTbl And Not ModScanner.IsLayoutFigureCaption(rng, clean, qp.ReFig, patFig) Then
                    ElseIf qp.ReManToc.Test(clean) Then
                        If Not slots(1).Found Then tType = ElementTypeLevel1
                    ElseIf qp.ReManLot.Test(clean) Then
                        If Not slots(2).Found Then tType = ElementTypeLevel1
                    ElseIf qp.ReManLof.Test(clean) Then
                        If Not slots(3).Found Then tType = ElementTypeLevel1
                    ElseIf qp.ReL1.Test(clean) Or (inContent And ModLocale.IsRomanLevel1(clean)) Then
                        tType = ElementTypeLevel1
                    ElseIf inContent Or inPost Then
                        If qp.ReTbl.Test(clean) Then
                            tType = ElementTypeTableTitle
                        ElseIf qp.ReFig.Test(clean) Then
                            tType = ElementTypeFigureTitle
                        ElseIf qp.ReSrc.Test(clean) Then
                            tType = ElementTypeSourceNote
                        ElseIf inContent Then
                            If qp.ReL2.Test(clean) Then
                                tType = ElementTypeLevel2
                            ElseIf qp.ReL3.Test(clean) Then
                                tType = ElementTypeLevel3
                            ElseIf qp.ReL4.Test(clean) Then
                                tType = ElementTypeLevel4
                            End If
                        End If
                    End If
                    If tType <> ElementTypeUnknown Then
                        bmName = ""
                        If tType >= ElementTypeLevel1 And tType <= ElementTypeLevel4 Then
                            cntHead = cntHead + 1
                            bmName = ModBookmark.FormatBookmarkId(BM_PREFIX_TOC, cntHead)
                        ElseIf tType = ElementTypeTableTitle Then
                            cntTbl = cntTbl + 1
                            bmName = ModBookmark.FormatBookmarkId(BM_PREFIX_TBL, cntTbl)
                        ElseIf tType = ElementTypeFigureTitle Then
                            cntFig = cntFig + 1
                            bmName = ModBookmark.FormatBookmarkId(BM_PREFIX_FIG, cntFig)
                        ElseIf tType = ElementTypeSourceNote Then
                            cntNotes = cntNotes + 1
                        End If
                        ModFormatSingle.FormatSingleItem doc, rng, tType, bmName, g_Config.FontName, g_Config.FontSize, g_Config.IndentPts, hasParaMath
                        If g_Config.LineSpacingMode <> "ls_keep" And tType >= ElementTypeLevel1 And tType <= ElementTypeSourceNote Then
                            ModFormatCore.ApplyLineSpacing rng, g_Config.LineSpacingMode
                        End If
                        ModFormatBatch.AppendQuickSnap qSnaps, qCount, qCap, i, rng.Start, rng.End, tType, raw, clean
                        If tType <> ElementTypeSourceNote Then
                            itemCount = itemCount + 1
                            If itemCount > UBound(items) Then ReDim Preserve items(1 To itemCount + 64)
                            items(itemCount).ElementType = tType
                            items(itemCount).CleanText = clean
                            items(itemCount).BookmarkName = bmName
                        End If
                    End If
                End If
            End If
            If tType = ElementTypeUnknown And Not isSkip And hasXref Then
                Dim candType As Long
                If inTbl Then
                    candType = ElementTypeTableCell
                ElseIf isListPara Then
                    candType = ElementTypeList
                Else
                    candType = ElementTypeNormalParagraph
                End If
                ModFormatBatch.AppendQuickSnap qSnaps, qCount, qCap, i, rStart, rEnd, candType, raw, CleanQuickText(raw)
            End If
        End If
NextQuickPara:
    Next p
    ModBookmark.CleanExcessBookmarks doc, cntHead, cntTbl, cntFig
End Sub

Private Sub CollectTocHeadersSorted(ByVal doc As Document, ByVal itemPos As Long, _
                                    ByRef hdrs() As String, ByRef bms() As String, _
                                    ByVal hCount As Long, ByRef hIdx As Long, _
                                    ByRef entries() As BatchListEntry, ByRef eCnt As Long, ByRef eCap As Long)
    Dim doIns As Boolean
    Do While hIdx <= hCount
        doIns = (itemPos < 0)
        If Not doIns And doc.Bookmarks.Exists(bms(hIdx)) Then doIns = (doc.Bookmarks(bms(hIdx)).Range.Start < itemPos)
        If doIns Then
            ModFormatBatch.AppendBatchEntry entries, eCnt, eCap, hdrs(hIdx), bms(hIdx), 0#, True
            hIdx = hIdx + 1
        ElseIf itemPos >= 0 And Not doc.Bookmarks.Exists(bms(hIdx)) Then
            hIdx = hIdx + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub BuildQuickSingleList(ByVal doc As Document, ByRef slot As QuickZoneSlot, ByRef items() As QuickItem, _
                                 ByVal itemCount As Long, ByVal expectedCnt As Long)
    If Not slot.Found Then Exit Sub
    If (slot.Kind = 2 Or slot.Kind = 3) And expectedCnt <= 0 And slot.MarkerPara > 0 Then Exit Sub
    Dim pIdx As Long, hdrTitle As String, mText As String, i As Long, wantType As Long
    If slot.Kind = 3 Then
        hdrTitle = ModLocale.GetLofHeaderTitle(g_Config.ListLang)
        mText = "=listfigurehere()"
        wantType = ElementTypeFigureTitle
    ElseIf slot.Kind = 2 Then
        hdrTitle = ModLocale.GetLotHeaderTitle(g_Config.ListLang)
        mText = "=listtablehere()"
        wantType = ElementTypeTableTitle
    Else
        hdrTitle = ModLocale.GetTocHeaderTitle(g_Config.ListLang)
        mText = "=listcontent()"
    End If
    If slot.MarkerPara > 0 Then
        pIdx = ModMarkerLexer.FindMarkerParagraph(doc, mText)
    Else
        pIdx = ModZoneManager.FindParaIndexByBookmark(doc, slot.AnchorBm)
        ModZoneManager.SafeDelBm doc, slot.AnchorBm
    End If
    If pIdx <= 0 Then Exit Sub
    ModListBuilder.InsertListHeader doc, pIdx, hdrTitle, slot.HeaderTrackBm, g_Config.FontName, g_Config.FontSize
    ModBookmark.AddBookmark doc, slot.HeaderBm, doc.Paragraphs(pIdx).Range
    If pIdx + 1 = doc.Paragraphs.Count Then doc.Paragraphs(pIdx + 1).Range.InsertAfter vbCr
    Dim tabPos As Single, curRng As Range, entries() As BatchListEntry, built As Long, eCap As Long, itemPos As Long
    tabPos = ModPageNumberSync.CalculateRightTab(doc, doc.Paragraphs(pIdx).Range)
    Set curRng = doc.Paragraphs(pIdx + 1).Range
    If slot.MarkerPara = 0 Then curRng.Collapse 1
    built = 0
    If slot.Kind = 3 Or slot.Kind = 2 Then
        For i = 1 To itemCount
            If items(i).ElementType = wantType Then
                ModFormatBatch.AppendBatchEntry entries, built, eCap, items(i).CleanText, items(i).BookmarkName, 0#, True
            End If
        Next i
        Set curRng = ModFormatBatch.InsertListEntriesBatch(doc, curRng, entries, built, g_Config.FontName, g_Config.FontSize, tabPos, 0#, True)
    Else
        Dim hdrs(1 To 4) As String, bms(1 To 4) As String, hPos(1 To 4) As Long, hCount As Long
        ModListBuilder.CollectSortedHeaderTargets doc, hdrs, bms, hPos, hCount
        Dim maxAllowed As Long, hIdx As Long, indPts As Single, isB As Boolean
        If g_Config.TocDepth = 3 Then maxAllowed = ElementTypeLevel3 Else maxAllowed = ElementTypeLevel4
        hIdx = 1
        For i = 1 To itemCount
            If items(i).ElementType >= ElementTypeLevel1 And items(i).ElementType <= ElementTypeLevel4 Then
                If doc.Bookmarks.Exists(items(i).BookmarkName) Then
                    itemPos = doc.Bookmarks(items(i).BookmarkName).Range.Start
                    CollectTocHeadersSorted doc, itemPos, hdrs, bms, hCount, hIdx, entries, built, eCap
                End If
                If items(i).ElementType <= maxAllowed Then
                    indPts = 0#
                    isB = (items(i).ElementType = ElementTypeLevel1) Or g_Config.IsBoldTOC
                    If items(i).ElementType = ElementTypeLevel2 Then indPts = CmToPoints(0.5)
                    If items(i).ElementType = ElementTypeLevel3 Then indPts = CmToPoints(0.75)
                    If items(i).ElementType = ElementTypeLevel4 Then indPts = CmToPoints(1#)
                    ModFormatBatch.AppendBatchEntry entries, built, eCap, items(i).CleanText, items(i).BookmarkName, indPts, isB
                End If
            End If
        Next i
        CollectTocHeadersSorted doc, -1, hdrs, bms, hCount, hIdx, entries, built, eCap
        Set curRng = ModFormatBatch.InsertListEntriesBatch(doc, curRng, entries, built, g_Config.FontName, g_Config.FontSize, tabPos, 0#, g_Config.IsBoldTOC)
    End If
    slot.ItemCount = built
    slot.EntryCount = built
    If slot.MarkerPara = 0 And pIdx + 1 + built < doc.Paragraphs.Count Then
        Dim ancText As String
        ancText = doc.Paragraphs(pIdx + 1 + built).Range.Text
        If InStr(1, ancText, Chr(12)) = 0 Then
            doc.Paragraphs(pIdx + 1 + built).Range.Delete
        End If
    End If
End Sub

Private Sub FinalizeQuickZoneBookmarks(ByVal doc As Document, ByRef slots() As QuickZoneSlot)
    Dim k As Long, hdrPara As Long, hdrRng As Range, firstEntry As Long, lastEntry As Long, zStart As Long, zEnd As Long
    For k = 1 To 3
        If slots(k).Found And doc.Bookmarks.Exists(slots(k).HeaderTrackBm) Then
            hdrPara = ModZoneManager.FindParaIndexByBookmark(doc, slots(k).HeaderTrackBm)
            If hdrPara > 0 Then
                Set hdrRng = doc.Paragraphs(hdrPara).Range
                ModBookmark.AddBookmark doc, slots(k).HeaderBm, hdrRng
                ModZoneManager.SafeDelBm doc, slots(k).HeaderTrackBm
                If slots(k).EntryCount > 0 Then
                    If slots(k).MarkerPara > 0 Then firstEntry = hdrPara + 2 Else firstEntry = hdrPara + 1
                    lastEntry = firstEntry + slots(k).EntryCount - 1
                    If lastEntry <= doc.Paragraphs.Count Then
                        zStart = doc.Paragraphs(firstEntry).Range.Start
                        zEnd = doc.Paragraphs(lastEntry).Range.End - 1
                        If zEnd > zStart Then ModBookmark.AddBookmark doc, slots(k).ZoneBm, doc.Range(zStart, zEnd)
                    End If
                Else
                    ModBookmark.AddBookmark doc, slots(k).ZoneBm, doc.Range(hdrRng.End, hdrRng.End)
                End If
            End If
        End If
    Next k
End Sub

Public Sub RunQuickUpdate(Optional ByVal targetDoc As Document = Nothing)
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then Exit Sub
    ModMain.g_IsEngineRunning = True
    ModRibbonCallbacks.InvalidateRibbonSafe
    On Error GoTo MainErr
    Dim wasCancelled As Boolean
    wasCancelled = False
    Dim doc As Document
    If targetDoc Is Nothing Then
        If Documents.Count > 0 Then Set doc = ActiveDocument
    Else
        Set doc = targetDoc
    End If
    If doc Is Nothing Then GoTo CleanExit
    ModRibbonCallbacks.EnsureConfigInitialized
    ModConstants.g_SilentErrors = 0
    ModRibbonCallbacks.g_LastError = ""
    ModCustomProfile.LoadStyleProfile doc

    Dim errMsg As String, markers As MarkerLocations, wasTracking As Boolean, isSuccess As Boolean
    Dim deferMsg As String, deferStyle As Long, deferTitle As String
    If Not ModBackup.ValidatePreFlightConditions(doc, errMsg) Then
        ModRibbonCallbacks.g_LastError = errMsg
        deferMsg = errMsg
        deferStyle = vbExclamation
        deferTitle = ModConstants.Uni("Pre-Flight Check (Ki\u1ec3m tra tr\u01b0\u1edbc khi ch\u1ea1y)")
        GoTo CleanExit
    End If
    If Not ModMarkerLexer.ScanMarkers(doc, markers, errMsg) Then
        ModRibbonCallbacks.g_LastError = errMsg
        deferMsg = errMsg
        deferStyle = vbExclamation
        deferTitle = ModConstants.Uni("Marker Error (L\u1ed7i Th\u1ebb)")
        GoTo CleanExit
    End If
    If Not ModBackup.HandleTrackChanges(doc, wasTracking) Then
        deferMsg = ModConstants.Uni("Stopped safely as requested (Track Changes). (\u0110\u00e3 d\u1eebng an to\u00e0n theo y\u00eau c\u1ea7u.)")
        deferStyle = vbInformation
        deferTitle = ModConstants.Uni("Stopped (\u0110\u00e3 d\u1eebng)")
        GoTo CleanExit
    End If
    Dim prevScreen As Boolean, screenDisabled As Boolean, prevPag As Boolean, hasPag As Boolean
    Dim prevView As Long, hasView As Boolean, prevSeek As Long, hasSeek As Boolean
    prevScreen = Application.ScreenUpdating
    Application.EnableCancelKey = 0
    Application.ScreenUpdating = False
    screenDisabled = True
    On Error Resume Next
    prevPag = Application.Options.Pagination
    If Err.Number = 0 Then hasPag = True
    If hasPag Then Application.Options.Pagination = False
    Err.Clear
    If doc.Windows.Count > 0 Then
        If Not doc.ActiveWindow Is Nothing Then
            prevSeek = doc.ActiveWindow.View.SeekView
            If Err.Number = 0 Then hasSeek = True
            Err.Clear
            If hasSeek And prevSeek <> 0 Then
                doc.ActiveWindow.View.SeekView = 0
                Err.Clear
            End If
            prevView = doc.ActiveWindow.View.Type
            If Err.Number = 0 Then hasView = True
            If hasView Then doc.ActiveWindow.View.Type = 1
            Err.Clear
        End If
    End If
    Dim objUndo As Object, undoStarted As Boolean
    Set objUndo = Application.UndoRecord
    On Error GoTo MainErr
    undoStarted = ModBackup.StartUndoSafe("Quick Update TOC & Lists", objUndo)
    ModProgress.ProgressInit "Quick Update"
    ModProgress.ProgressSetPhase ModConstants.Uni("Preparing Zones (Chu\u1ea9n b\u1ecb v\u00f9ng)"), 5#, 15#
    Dim slots() As QuickZoneSlot, prevParaCount As Long, newCsIdx As Long, csDelta As Long
    prevParaCount = doc.Paragraphs.Count
    If Not ModZoneManager.DiscoverAndPrepareQuickZones(doc, markers, slots) Then
        deferMsg = ModConstants.Uni("No TOC, LOT, or LOF found to update. (Kh\u00f4ng t\u00ecm th\u1ea5y M\u1ee5c l\u1ee5c hay Danh m\u1ee5c \u0111\u1ec3 c\u1eadp nh\u1eadt.)")
        deferStyle = vbExclamation
        deferTitle = ModConstants.Uni("Not Found (Kh\u00f4ng t\u00ecm th\u1ea5y)")
        GoTo CleanExit
    End If
    If doc.Paragraphs.Count <> prevParaCount Then
        newCsIdx = ModMarkerLexer.FindMarkerParagraph(doc, "=contentstart()")
        If newCsIdx <= 0 Then Err.Raise vbObjectError + 1000, "RunQuickUpdate", "Missing =contentStart() after zone cleanup."
        csDelta = newCsIdx - markers.ContentStartParaIndex
        markers.ContentStartParaIndex = newCsIdx
        If markers.HasContentEnd Then markers.ContentEndParaIndex = markers.ContentEndParaIndex + csDelta
    End If
    ModCustomProfile.LoadTagCache doc
    ModProgress.ProgressSetPhase ModConstants.Uni("Scanning Items (Qu\u00e9t ti\u00eau \u0111\u1ec1)"), 20#, 35#
    Dim items() As QuickItem, itemCount As Long, cntHead As Long, cntTbl As Long, cntFig As Long, cntNotes As Long, qSnaps() As ParagraphSnapshot, qCount As Long
    ScanAndFormatQuickItems doc, markers, slots, items, itemCount, cntHead, cntTbl, cntFig, cntNotes, qSnaps, qCount
    If ModProgress.g_CancelRequested Then
        wasCancelled = True
        GoTo CleanExit
    End If
    ModProgress.ProgressSetPhase ModConstants.Uni("Cross-References (Li\u00ean k\u1ebft ch\u00e9o)"), 55#, 10#
    ModCrossRef.ApplyCrossReferences doc, markers, qSnaps, qCount, g_Config.CrossRefColor
    ModProgress.ProgressSetPhase ModConstants.Uni("Building Lists (D\u1ef1ng danh m\u1ee5c)"), 65#, 15#
    BuildQuickSingleList doc, slots(3), items, itemCount, cntFig
    BuildQuickSingleList doc, slots(2), items, itemCount, cntTbl
    BuildQuickSingleList doc, slots(1), items, itemCount, cntHead
    FinalizeQuickZoneBookmarks doc, slots
    ModProgress.ProgressSetPhase ModConstants.Uni("Syncing Pages (\u0110\u1ed3ng b\u1ed9 s\u1ed1 trang)"), 80#, 18#
    Dim pageMode As Long
    If g_Config.IsFooterPage Then pageMode = PageSyncRealFooter Else pageMode = PageSyncPhysicalPage
    If hasView And doc.Windows.Count > 0 Then
        If Not doc.ActiveWindow Is Nothing Then doc.ActiveWindow.View.Type = 3
    End If
    doc.Repaginate
    If Not ModPageNumberSync.SyncAllListPages(doc, pageMode) Then Err.Raise vbObjectError + 1006, "RunQuickUpdate", "SyncAllListPages failed"
    ModMarkerLexer.HighlightMarkers doc
    isSuccess = True
    GoTo CleanExit
MainErr:
    ModRibbonCallbacks.g_LastError = "QuickUpdate Error " & Err.Number & ": " & Err.Description
    deferMsg = ModConstants.Uni("Quick Update error (L\u1ed7i c\u1eadp nh\u1eadt nhanh) (") & Err.Number & "): " & Err.Description
    deferStyle = vbCritical
    deferTitle = ModConstants.Uni("Error (L\u1ed7i)")
CleanExit:
    On Error Resume Next
    Application.EnableCancelKey = 1
    ModProgress.ProgressFinish
    If hasView And Not doc Is Nothing Then
        If doc.Windows.Count > 0 Then
            If Not doc.ActiveWindow Is Nothing Then
                doc.ActiveWindow.View.Type = prevView
                If hasSeek And prevSeek <> 0 Then doc.ActiveWindow.View.SeekView = prevSeek
            End If
        End If
    End If
    If hasPag Then Application.Options.Pagination = prevPag
    If undoStarted And Not (objUndo Is Nothing) Then If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
    If wasCancelled And undoStarted And Not (doc Is Nothing) Then
        If Not doc.Saved Then doc.Undo
    End If
    If Not doc Is Nothing Then ModBackup.RestoreTrackChanges doc, wasTracking
    If screenDisabled Then
        Application.ScreenUpdating = prevScreen
        Application.ScreenRefresh
    End If
    If wasCancelled Then
        If Application.DisplayAlerts <> 0 Then
            ModUI.MsgBoxW ModConstants.Uni("Operation cancelled safely. (\u0110\u00e3 h\u1ee7y thao t\u00e1c an to\u00e0n.)"), vbInformation, ModConstants.Uni("Cancelled (\u0110\u00e3 h\u1ee7y)")
        End If
    ElseIf Len(deferMsg) > 0 Then
        ModUI.MsgBoxW deferMsg, deferStyle, deferTitle
    ElseIf isSuccess Then
        ModUI.ShowQuickUpdateReport slots(1).Found, slots(1).ItemCount, slots(2).Found, slots(2).ItemCount, slots(3).Found, slots(3).ItemCount, cntHead, cntTbl, cntFig, cntNotes
    End If
    ModCustomProfile.ResetStyleProfileCache
    ModCustomProfile.ResetTagCache
    ModTableIndex.InvalidateAllIndexes
    ModMain.g_IsEngineRunning = False
    ModRibbonCallbacks.InvalidateRibbonSafe
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
