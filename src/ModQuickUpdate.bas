Attribute VB_Name = "ModQuickUpdate"
Option Explicit

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
    If pRng.End - 1 > pRng.Start Then doc.Range(pRng.Start, pRng.End - 1).Delete
    Set pRng = doc.Paragraphs(pIdx).Range
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
    If rng.OMaths.Count > 0 Then
        If rng.OMaths(1).Type = 0 Or rng.OMaths(1).Range.Text = rng.Text Then IsMathBlockPara = True
    End If
    Exit Function
ErrMath:
    IsMathBlockPara = False
End Function

Private Function IsOldTocPara(ByVal rng As Range, ByRef sArr() As Long, ByRef eArr() As Long, ByVal n As Long) As Boolean
    Dim k As Long
    For k = 1 To n
        If rng.Start >= sArr(k) And rng.Start < eArr(k) Then
            IsOldTocPara = True
            Exit Function
        End If
    Next k
End Function

Private Function MatchesPat(ByVal re As Object, ByVal pat As String, ByVal s As String) As Boolean
    re.Pattern = pat
    MatchesPat = re.Test(s)
End Function

Private Sub ScanAndFormatQuickItems(ByVal doc As Document, ByRef markers As MarkerLocations, ByRef slots() As QuickZoneSlot, _
                                    ByRef items() As QuickItem, ByRef itemCount As Long, ByRef cntHead As Long, _
                                    ByRef cntTbl As Long, ByRef cntFig As Long, ByRef cntNotes As Long)
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False
    Dim patL1 As String, patL2 As String, patL3 As String, patL4 As String
    Dim patTbl As String, patFig As String, patSrc As String, patMkr As String
    Dim patToc As String, patLot As String, patLof As String
    patL1 = ModLocale.GetPatternLevel1()
    patL2 = ModConstants.GetPatternLevel2()
    patL3 = ModConstants.GetPatternLevel3()
    patL4 = ModConstants.GetPatternLevel4()
    patTbl = ModLocale.GetPatternTable()
    patFig = ModLocale.GetPatternFigure()
    patSrc = ModLocale.GetPatternSourceNote()
    patMkr = ModConstants.GetPatternMarker()
    patToc = ModLocale.GetPatternManualToc()
    patLot = ModLocale.GetPatternManualLot()
    patLof = ModLocale.GetPatternManualLof()

    Dim totalParas As Long, latexFlags() As Boolean, latexWarn As Long, docHasMath As Boolean
    totalParas = doc.Paragraphs.Count
    itemCount = 0
    ReDim items(1 To 64)
    ModTableIndex.BuildTableIndex doc
    ModLatex.ComputeLatexFlags doc, latexFlags, latexWarn
    docHasMath = (doc.OMaths.Count > 0)
    Dim numOldTocs As Long, tocS() As Long, tocE() As Long, k As Long
    numOldTocs = doc.TablesOfContents.Count
    If numOldTocs > 0 Then
        ReDim tocS(1 To numOldTocs)
        ReDim tocE(1 To numOldTocs)
        For k = 1 To numOldTocs
            tocS(k) = doc.TablesOfContents(k).Range.Start
            tocE(k) = doc.TablesOfContents(k).Range.End
        Next k
    End If

    Dim i As Long, p As Paragraph, rng As Range, raw As String, clean As String
    Dim inContent As Boolean, inPost As Boolean, tType As Long, bmName As String
    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        ModProgress.ProgressUpdate i, totalParas
        Set rng = p.Range
        raw = rng.Text
        If Len(raw) <= 550 Then
            clean = Trim$(Replace(Replace(Replace(raw, vbCr, ""), Chr(7), ""), Chr(11), " "))
            If Len(clean) > 0 And Len(clean) <= 500 And Not latexFlags(i) Then
                inContent = (i > markers.ContentStartParaIndex And i < markers.ContentEndParaIndex)
                inPost = (i > markers.ContentEndParaIndex)
                tType = ElementTypeUnknown
                If IsOldTocPara(rng, tocS, tocE, numOldTocs) Then
                ElseIf IsMathBlockPara(rng, docHasMath) Then
                ElseIf rng.ListFormat.ListType <> 0 Then
                ElseIf Left$(clean, 1) = "=" And MatchesPat(regEx, patMkr, clean) Then
                ElseIf ModTableIndex.IsRangeInTable(rng) And Not ModScanner.IsLayoutFigureCaption(rng, clean, regEx, patFig) Then
                ElseIf MatchesPat(regEx, patToc, clean) Then
                    If Not slots(1).Found Then tType = ElementTypeLevel1
                ElseIf MatchesPat(regEx, patLot, clean) Then
                    If Not slots(2).Found Then tType = ElementTypeLevel1
                ElseIf MatchesPat(regEx, patLof, clean) Then
                    If Not slots(3).Found Then tType = ElementTypeLevel1
                ElseIf MatchesPat(regEx, patL1, clean) Then
                    tType = ElementTypeLevel1
                ElseIf inContent Or inPost Then
                    If MatchesPat(regEx, patTbl, clean) Then
                        tType = ElementTypeTableTitle
                    ElseIf MatchesPat(regEx, patFig, clean) Then
                        tType = ElementTypeFigureTitle
                    ElseIf MatchesPat(regEx, patSrc, clean) Then
                        tType = ElementTypeSourceNote
                    ElseIf inContent Then
                        If MatchesPat(regEx, patL2, clean) Then
                            tType = ElementTypeLevel2
                        ElseIf MatchesPat(regEx, patL3, clean) Then
                            tType = ElementTypeLevel3
                        ElseIf MatchesPat(regEx, patL4, clean) Then
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
                    ModFormatSingle.FormatSingleItem doc, rng, tType, bmName, g_Config.FontName, g_Config.FontSize, g_Config.IndentPts, docHasMath
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
    Next p
    ModBookmark.CleanExcessBookmarks doc, cntHead, cntTbl, cntFig
End Sub

Private Sub InsertTocHeadersSorted(ByVal doc As Document, ByRef curRng As Range, ByVal itemPos As Long, _
                                   ByRef hdrs() As String, ByRef bms() As String, ByRef hPos() As Long, _
                                   ByVal hCount As Long, ByRef hIdx As Long, ByVal tabPos As Single, ByRef entryCnt As Long)
    Do While hIdx <= hCount
        If itemPos < 0 Or hPos(hIdx) < itemPos Then
            Set curRng = ModListBuilder.InsertEntry(doc, curRng, hdrs(hIdx), bms(hIdx), g_Config.FontName, g_Config.FontSize, tabPos, 0#, True)
            entryCnt = entryCnt + 1
            hIdx = hIdx + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub BuildQuickSingleList(ByVal doc As Document, ByRef slot As QuickZoneSlot, ByRef items() As QuickItem, ByVal itemCount As Long)
    If Not slot.Found Then Exit Sub
    Dim pIdx As Long, hdrTitle As String, mText As String
    If slot.Kind = 3 Then
        hdrTitle = ModLocale.GetLofHeaderTitle(g_Config.ListLang)
        mText = "=listfigurehere()"
    ElseIf slot.Kind = 2 Then
        hdrTitle = ModLocale.GetLotHeaderTitle(g_Config.ListLang)
        mText = "=listtablehere()"
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
    Dim tabPos As Single, curRng As Range, i As Long, built As Long, itemPos As Long
    tabPos = ModPageNumberSync.CalculateRightTab(doc)
    Set curRng = doc.Paragraphs(pIdx + 1).Range
    built = 0
    If slot.Kind = 3 Or slot.Kind = 2 Then
        Dim wantType As Long
        If slot.Kind = 3 Then wantType = ElementTypeFigureTitle Else wantType = ElementTypeTableTitle
        For i = 1 To itemCount
            If items(i).ElementType = wantType Then
                Set curRng = ModListBuilder.InsertEntry(doc, curRng, items(i).CleanText, items(i).BookmarkName, g_Config.FontName, g_Config.FontSize, tabPos, 0#, True)
                built = built + 1
            End If
        Next i
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
                    InsertTocHeadersSorted doc, curRng, itemPos, hdrs, bms, hPos, hCount, hIdx, tabPos, built
                End If
                If items(i).ElementType <= maxAllowed Then
                    indPts = 0#
                    isB = g_Config.IsBoldTOC
                    If items(i).ElementType = ElementTypeLevel1 Then
                        isB = True
                    ElseIf items(i).ElementType = ElementTypeLevel2 Then
                        indPts = CmToPoints(0.5)
                    ElseIf items(i).ElementType = ElementTypeLevel3 Then
                        indPts = CmToPoints(0.75)
                    Else
                        indPts = CmToPoints(1#)
                    End If
                    Set curRng = ModListBuilder.InsertEntry(doc, curRng, items(i).CleanText, items(i).BookmarkName, g_Config.FontName, g_Config.FontSize, tabPos, indPts, isB)
                    built = built + 1
                End If
            End If
        Next i
        InsertTocHeadersSorted doc, curRng, -1, hdrs, bms, hPos, hCount, hIdx, tabPos, built
    End If
    slot.ItemCount = built
    slot.EntryCount = built
    If slot.MarkerPara = 0 Then doc.Paragraphs(pIdx + 1).Range.Delete
End Sub

Private Sub FinalizeQuickZoneBookmarks(ByVal doc As Document, ByRef slots() As QuickZoneSlot)
    Dim k As Long, hdrPara As Long, hdrRng As Range, firstEntry As Long, lastEntry As Long
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
                        ModBookmark.AddBookmark doc, slots(k).ZoneBm, doc.Range(doc.Paragraphs(firstEntry).Range.Start, doc.Paragraphs(lastEntry).Range.End)
                    End If
                Else
                    ModBookmark.AddBookmark doc, slots(k).ZoneBm, doc.Range(hdrRng.End, hdrRng.End)
                End If
            End If
        End If
    Next k
End Sub

Public Sub RunQuickUpdate(Optional ByVal targetDoc As Document = Nothing)
    If ModMain.g_IsEngineRunning Then Exit Sub
    ModMain.g_IsEngineRunning = True
    On Error GoTo MainErr
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

    Dim errMsg As String
    If Not ModBackup.ValidatePreFlightConditions(doc, errMsg) Then
        ModRibbonCallbacks.g_LastError = errMsg
        ModUI.MsgBoxW errMsg, vbExclamation, ModConstants.Uni("Pre-Flight Check (Ki\u1ec3m tra tr\u01b0\u1edbc khi ch\u1ea1y)")
        GoTo CleanExit
    End If
    Dim markers As MarkerLocations
    If Not ModMarkerLexer.ScanMarkers(doc, markers, errMsg) Then
        ModRibbonCallbacks.g_LastError = errMsg
        ModUI.MsgBoxW errMsg, vbExclamation, ModConstants.Uni("Marker Error (L\u1ed7i Th\u1ebb)")
        GoTo CleanExit
    End If
    Dim wasTracking As Boolean
    If Not ModBackup.HandleTrackChanges(doc, wasTracking) Then
        ModUI.MsgBoxW ModConstants.Uni("Stopped safely as requested (Track Changes). (\u0110\u00e3 d\u1eebng an to\u00e0n theo y\u00eau c\u1ea7u.)"), vbInformation, ModConstants.Uni("Stopped (\u0110\u00e3 d\u1eebng)")
        GoTo CleanExit
    End If

    Dim prevScreen As Boolean, screenDisabled As Boolean, prevPag As Boolean, hasPag As Boolean
    Dim prevView As Long, hasView As Boolean
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    screenDisabled = True
    On Error Resume Next
    prevPag = Application.Options.Pagination
    If Err.Number = 0 Then
        hasPag = True
        Application.Options.Pagination = False
    Else
        Err.Clear
    End If
    If Not doc.ActiveWindow Is Nothing Then
        prevView = doc.ActiveWindow.View.Type
        hasView = True
        doc.ActiveWindow.View.Type = 1
    End If
    On Error GoTo MainErr

    Dim objUndo As Object, undoStarted As Boolean
    On Error Resume Next
    Set objUndo = Application.UndoRecord
    On Error GoTo MainErr
    If Not objUndo Is Nothing Then
        objUndo.StartCustomRecord "Quick Update TOC & Lists"
        undoStarted = True
    End If
    ModProgress.ProgressInit "Quick Update"
    ModProgress.ProgressSetPhase ModConstants.Uni("Preparing Zones (Chu\u1ea9n b\u1ecb v\u00f9ng)"), 5#, 15#
    Dim slots() As QuickZoneSlot
    If Not ModZoneManager.DiscoverAndPrepareQuickZones(doc, markers, slots) Then
        ModUI.MsgBoxW ModConstants.Uni("No TOC, LOT, or LOF found to update. (Kh\u00f4ng t\u00ecm th\u1ea5y M\u1ee5c l\u1ee5c hay Danh m\u1ee5c \u0111\u1ec3 c\u1eadp nh\u1eadt.)"), vbExclamation, ModConstants.Uni("Not Found (Kh\u00f4ng t\u00ecm th\u1ea5y)")
        GoTo CleanExit
    End If
    If Not ModMarkerLexer.ScanMarkers(doc, markers, errMsg) Then Err.Raise vbObjectError + 1000, "RunQuickUpdate", errMsg

    ModProgress.ProgressSetPhase ModConstants.Uni("Scanning Items (Qu\u00e9t ti\u00eau \u0111\u1ec1)"), 20#, 40#
    Dim items() As QuickItem, itemCount As Long, cntHead As Long, cntTbl As Long, cntFig As Long, cntNotes As Long
    ScanAndFormatQuickItems doc, markers, slots, items, itemCount, cntHead, cntTbl, cntFig, cntNotes
    ModProgress.ProgressSetPhase ModConstants.Uni("Building Lists (D\u1ef1ng danh m\u1ee5c)"), 60#, 20#
    BuildQuickSingleList doc, slots(3), items, itemCount
    BuildQuickSingleList doc, slots(2), items, itemCount
    BuildQuickSingleList doc, slots(1), items, itemCount
    FinalizeQuickZoneBookmarks doc, slots

    ModProgress.ProgressSetPhase ModConstants.Uni("Syncing Pages (\u0110\u1ed3ng b\u1ed9 s\u1ed1 trang)"), 80#, 18#
    Dim pageMode As Long
    If g_Config.IsFooterPage Then pageMode = PageSyncRealFooter Else pageMode = PageSyncPhysicalPage
    If hasView Then doc.ActiveWindow.View.Type = 3
    doc.Repaginate
    If Not ModPageNumberSync.SyncAllListPages(doc, pageMode) Then
        Err.Raise vbObjectError + 1006, "RunQuickUpdate", "SyncAllListPages failed"
    End If
    ModMarkerLexer.HighlightMarkers doc
    Dim isSuccess As Boolean
    isSuccess = True
    GoTo CleanExit

MainErr:
    ModRibbonCallbacks.g_LastError = "QuickUpdate Error " & Err.Number & ": " & Err.Description
    ModUI.MsgBoxW ModConstants.Uni("Quick Update error (L\u1ed7i c\u1eadp nh\u1eadt nhanh) (") & Err.Number & "): " & Err.Description, vbCritical, ModConstants.Uni("Error (L\u1ed7i)")
CleanExit:
    On Error Resume Next
    ModProgress.ProgressFinish
    If hasView Then doc.ActiveWindow.View.Type = prevView
    If hasPag Then Application.Options.Pagination = prevPag
    If undoStarted Then
        If Not objUndo Is Nothing Then
            If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
        End If
    End If
    If Not doc Is Nothing Then ModBackup.RestoreTrackChanges doc, wasTracking
    If screenDisabled Then Application.ScreenUpdating = prevScreen
    If isSuccess Then
        ModUI.ShowQuickUpdateReport slots(1).Found, slots(1).ItemCount, slots(2).Found, slots(2).ItemCount, slots(3).Found, slots(3).ItemCount, cntHead, cntTbl, cntFig, cntNotes
    End If
    ModMain.g_IsEngineRunning = False
End Sub
