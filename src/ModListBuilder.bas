Attribute VB_Name = "ModListBuilder"
Option Explicit

' ==============================================================================
' Module: ModListBuilder
' Purpose: Construction of LOF, LOT, TOC with dynamic markers & 2-Phase page sync
' Architecture: Academic Document Format Engine (Bold TOC & Dynamic List Headers)
' ==============================================================================

Public Sub InsertListHeader(ByVal doc As Document, ByVal pIdx As Long, ByVal title As String, _
                            ByVal bm As String, ByVal fontName As String, ByVal fontSize As Single)
    Dim r As Range
    Set r = doc.Paragraphs(pIdx).Range
    r.HighlightColorIndex = 0
    r.InsertBefore title & vbCr
    Set r = doc.Paragraphs(pIdx).Range
    r.HighlightColorIndex = 0
    r.Font.Name = fontName
    r.Font.Size = fontSize
    r.Font.Bold = True
    r.Case = 1
    r.ParagraphFormat.Alignment = 1
    r.ParagraphFormat.FirstLineIndent = 0
    r.ParagraphFormat.LeftIndent = 0
    AddBookmark doc, bm, r
End Sub

Private Function HasElementType(ByRef snapshots() As ParagraphSnapshot, ByVal totalCount As Long, ByVal targetType As Long) As Boolean
    Dim i As Long
    For i = 1 To totalCount
        If snapshots(i).ElementType = targetType Then
            HasElementType = True
            Exit Function
        End If
    Next i
End Function

Public Function GetFreshTitleText(ByVal doc As Document, ByVal bmName As String, ByVal fallback As String) As String
    On Error Resume Next
    If bmName = BM_TOC_HEADER Or bmName = BM_LOT_HEADER Or bmName = BM_LOF_HEADER Then
        GetFreshTitleText = fallback
        Exit Function
    End If
    If Len(bmName) > 0 Then
        If doc.Bookmarks.Exists(bmName) Then
            Dim t As String
            t = Trim$(Replace(Replace(Replace(doc.Bookmarks(bmName).Range.Text, vbCr, ""), Chr(7), ""), Chr(11), " "))
            Do While InStr(t, "  ") > 0
                t = Replace(t, "  ", " ")
            Loop
            If Len(t) > 0 Then
                GetFreshTitleText = t
                Exit Function
            End If
        End If
    End If
    GetFreshTitleText = fallback
End Function

Public Function InsertEntry(ByVal doc As Document, ByVal atRange As Range, ByVal titleText As String, _
                             ByVal targetBm As String, ByVal fontName As String, ByVal fontSize As Single, _
                             ByVal tabPos As Single, ByVal leftIndentPts As Single, ByVal isBold As Boolean) As Range
    atRange.HighlightColorIndex = 0
    Dim actualTitle As String
    actualTitle = GetFreshTitleText(doc, targetBm, titleText)
    Dim entryText As String
    entryText = actualTitle & vbTab & DUMMY_PAGE_STRING & vbCr
    Dim startPos As Long
    startPos = atRange.End
    atRange.InsertAfter entryText
    Dim lineRng As Range
    Set lineRng = doc.Range(startPos, atRange.End - 1)
    lineRng.HighlightColorIndex = 0
    ApplyListLineFormat lineRng, fontName, fontSize, tabPos, leftIndentPts, isBold
    If Len(targetBm) > 0 Then
        If doc.Bookmarks.Exists(targetBm) Then
            InsertCleanHyperlink doc, lineRng, targetBm
        End If
    End If
    Dim nextRng As Range
    Set nextRng = doc.Range(atRange.End, atRange.End)
    nextRng.Collapse 0
    nextRng.HighlightColorIndex = 0
    Set InsertEntry = nextRng
End Function

Public Function BuildListOfFigures(ByVal doc As Document, ByRef markers As MarkerLocations, _
                                   ByRef snapshots() As ParagraphSnapshot, ByVal totalCount As Long, _
                                   ByVal fontName As String, ByVal fontSize As Single, ByVal pageMode As Long, _
                                   Optional ByVal autoSync As Boolean = True) As Boolean
    Dim pIdx As Long
    pIdx = ModMarkerLexer.FindMarkerParagraph(doc, "=listfigurehere()")
    If pIdx = 0 Then
        BuildListOfFigures = True
        Exit Function
    End If
    If Not HasElementType(snapshots, totalCount, ElementTypeFigureTitle) Then
        BuildListOfFigures = True
        Exit Function
    End If
    Dim objUndo As Object, weStartedUndo As Boolean
    weStartedUndo = False
    On Error Resume Next
    Set objUndo = Application.UndoRecord
    On Error GoTo ErrorHandler
    If Not objUndo Is Nothing Then
        If Not objUndo.IsRecordingCustomRecord Then
            objUndo.StartCustomRecord "Record 2: List of Figures"
            weStartedUndo = True
        End If
    End If
    
    Dim tabPos As Single
    tabPos = CalculateRightTab(doc)
    InsertListHeader doc, pIdx, ModLocale.GetLofHeaderTitle(g_Config.ListLang), BM_LOF_HEADER, fontName, fontSize
    
    Dim startEntriesPos As Long
    startEntriesPos = doc.Paragraphs(pIdx + 1).Range.End
    Dim curRng As Range
    Set curRng = doc.Paragraphs(pIdx + 1).Range
    Dim i As Long, figCount As Long
    figCount = 0
    For i = 1 To totalCount
        If snapshots(i).ElementType = ElementTypeFigureTitle Then
            figCount = figCount + 1
            Set curRng = InsertEntry(doc, curRng, snapshots(i).CleanText, FormatBookmarkId(BM_PREFIX_FIG, figCount), fontName, fontSize, tabPos, 0#, True)
        End If
    Next i
    If curRng.End > startEntriesPos Then AddBookmark doc, BM_ZONE_LOF, doc.Range(startEntriesPos, curRng.End)
    If autoSync Then
        doc.Repaginate
        If Not SyncAllListPages(doc, pageMode) Then
            Err.Raise vbObjectError + 1006, "BuildListOfFigures", "SyncAllListPages failed"
        End If
    End If
CleanExit:
    If weStartedUndo Then If Not objUndo Is Nothing Then If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
    BuildListOfFigures = True
    Exit Function
ErrorHandler:
    ModRibbonCallbacks.g_LastError = "BuildLOF Error " & Err.Number & ": " & Err.Description
    If weStartedUndo Then If Not objUndo Is Nothing Then If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
    BuildListOfFigures = False
End Function

Public Function BuildListOfTables(ByVal doc As Document, ByRef markers As MarkerLocations, _
                                  ByRef snapshots() As ParagraphSnapshot, ByVal totalCount As Long, _
                                  ByVal fontName As String, ByVal fontSize As Single, ByVal pageMode As Long, _
                                  Optional ByVal autoSync As Boolean = True) As Boolean
    Dim pIdx As Long
    pIdx = ModMarkerLexer.FindMarkerParagraph(doc, "=listtablehere()")
    If pIdx = 0 Then
        BuildListOfTables = True
        Exit Function
    End If
    If Not HasElementType(snapshots, totalCount, ElementTypeTableTitle) Then
        BuildListOfTables = True
        Exit Function
    End If
    Dim objUndo As Object, weStartedUndo As Boolean
    weStartedUndo = False
    On Error Resume Next
    Set objUndo = Application.UndoRecord
    On Error GoTo ErrorHandler
    If Not objUndo Is Nothing Then
        If Not objUndo.IsRecordingCustomRecord Then
            objUndo.StartCustomRecord "Record 3: List of Tables"
            weStartedUndo = True
        End If
    End If
    
    Dim tabPos As Single
    tabPos = CalculateRightTab(doc)
    InsertListHeader doc, pIdx, ModLocale.GetLotHeaderTitle(g_Config.ListLang), BM_LOT_HEADER, fontName, fontSize
    
    Dim startEntriesPos As Long
    startEntriesPos = doc.Paragraphs(pIdx + 1).Range.End
    Dim curRng As Range
    Set curRng = doc.Paragraphs(pIdx + 1).Range
    Dim i As Long, tblCount As Long
    tblCount = 0
    For i = 1 To totalCount
        If snapshots(i).ElementType = ElementTypeTableTitle Then
            tblCount = tblCount + 1
            Set curRng = InsertEntry(doc, curRng, snapshots(i).CleanText, FormatBookmarkId(BM_PREFIX_TBL, tblCount), fontName, fontSize, tabPos, 0#, True)
        End If
    Next i
    If curRng.End > startEntriesPos Then AddBookmark doc, BM_ZONE_LOT, doc.Range(startEntriesPos, curRng.End)
    If autoSync Then
        doc.Repaginate
        If Not SyncAllListPages(doc, pageMode) Then
            Err.Raise vbObjectError + 1006, "BuildListOfTables", "SyncAllListPages failed"
        End If
    End If
CleanExit:
    If weStartedUndo Then If Not objUndo Is Nothing Then If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
    BuildListOfTables = True
    Exit Function
ErrorHandler:
    ModRibbonCallbacks.g_LastError = "BuildLOT Error " & Err.Number & ": " & Err.Description
    If weStartedUndo Then If Not objUndo Is Nothing Then If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
    BuildListOfTables = False
End Function

Public Sub CollectSortedHeaderTargets(ByVal doc As Document, ByRef hdrs() As String, _
                                      ByRef bms() As String, ByRef hPos() As Long, ByRef hCount As Long)
    hCount = 0
    If doc.Bookmarks.Exists(BM_TOC_HEADER) Then
        hCount = hCount + 1
        hdrs(hCount) = ModLocale.GetTocHeaderTitle(g_Config.ListLang)
        bms(hCount) = BM_TOC_HEADER
        hPos(hCount) = doc.Bookmarks(BM_TOC_HEADER).Range.Start
    End If
    If doc.Bookmarks.Exists(BM_LOT_HEADER) Then
        hCount = hCount + 1
        hdrs(hCount) = ModLocale.GetLotHeaderTitle(g_Config.ListLang)
        bms(hCount) = BM_LOT_HEADER
        hPos(hCount) = doc.Bookmarks(BM_LOT_HEADER).Range.Start
    End If
    If doc.Bookmarks.Exists(BM_LOF_HEADER) Then
        hCount = hCount + 1
        hdrs(hCount) = ModLocale.GetLofHeaderTitle(g_Config.ListLang)
        bms(hCount) = BM_LOF_HEADER
        hPos(hCount) = doc.Bookmarks(BM_LOF_HEADER).Range.Start
    End If
    Dim m As Long, n As Long, tmpH As String, tmpB As String, tmpP As Long
    For m = 1 To hCount - 1
        For n = m + 1 To hCount
            If hPos(n) < hPos(m) Then
                tmpP = hPos(m)
                hPos(m) = hPos(n)
                hPos(n) = tmpP
                tmpH = hdrs(m)
                hdrs(m) = hdrs(n)
                hdrs(n) = tmpH
                tmpB = bms(m)
                bms(m) = bms(n)
                bms(n) = tmpB
            End If
        Next n
    Next m
End Sub

Public Function BuildTableOfContents(ByVal doc As Document, ByRef markers As MarkerLocations, _
                                     ByRef snapshots() As ParagraphSnapshot, ByVal totalCount As Long, _
                                     ByVal fontName As String, ByVal fontSize As Single, _
                                     ByVal pageMode As Long, ByVal isBoldAll As Boolean, _
                                     ByVal tocDepth As Long, _
                                     Optional ByVal autoSync As Boolean = True) As Boolean
    Dim pIdx As Long
    pIdx = ModMarkerLexer.FindMarkerParagraph(doc, "=listcontent()")
    If pIdx = 0 Then
        BuildTableOfContents = True
        Exit Function
    End If
    Dim objUndo As Object, weStartedUndo As Boolean
    weStartedUndo = False
    On Error Resume Next
    Set objUndo = Application.UndoRecord
    On Error GoTo ErrorHandler
    If Not objUndo Is Nothing Then
        If Not objUndo.IsRecordingCustomRecord Then
            objUndo.StartCustomRecord "Record 4: Table of Contents & Final Sync"
            weStartedUndo = True
        End If
    End If
    
    Dim tabPos As Single
    tabPos = CalculateRightTab(doc)
    InsertListHeader doc, pIdx, ModLocale.GetTocHeaderTitle(g_Config.ListLang), BM_TOC_HEADER, fontName, fontSize
    
    Dim hdrs(1 To 4) As String, bms(1 To 4) As String, hPos(1 To 4) As Long, hCount As Long
    CollectSortedHeaderTargets doc, hdrs, bms, hPos, hCount
    
    Dim startEntriesPos As Long
    startEntriesPos = doc.Paragraphs(pIdx + 1).Range.End
    Dim curRng As Range
    Set curRng = doc.Paragraphs(pIdx + 1).Range
    Dim i As Long, tocCount As Long, hIdx As Long
    tocCount = 0
    hIdx = 1
    Dim bmName As String, itemPos As Long, maxAllowed As Long
    If tocDepth = 3 Then maxAllowed = ElementTypeLevel3 Else maxAllowed = ElementTypeLevel4
    
    For i = 1 To totalCount
        If snapshots(i).ElementType >= ElementTypeLevel1 And snapshots(i).ElementType <= ElementTypeLevel4 Then
            tocCount = tocCount + 1
            bmName = FormatBookmarkId(BM_PREFIX_TOC, tocCount)
            If doc.Bookmarks.Exists(bmName) Then
                itemPos = doc.Bookmarks(bmName).Range.Start
                Do While hIdx <= hCount
                    If hPos(hIdx) < itemPos Then
                        Set curRng = InsertEntry(doc, curRng, hdrs(hIdx), bms(hIdx), fontName, fontSize, tabPos, 0#, True)
                        hIdx = hIdx + 1
                    Else
                        Exit Do
                    End If
                Loop
            End If
            If snapshots(i).ElementType <= maxAllowed Then
                Select Case snapshots(i).ElementType
                    Case ElementTypeLevel1
                        Set curRng = InsertEntry(doc, curRng, snapshots(i).CleanText, bmName, fontName, fontSize, tabPos, 0#, True)
                    Case ElementTypeLevel2
                        Set curRng = InsertEntry(doc, curRng, snapshots(i).CleanText, bmName, fontName, fontSize, tabPos, CmToPoints(0.5), isBoldAll)
                    Case ElementTypeLevel3
                        Set curRng = InsertEntry(doc, curRng, snapshots(i).CleanText, bmName, fontName, fontSize, tabPos, CmToPoints(0.75), isBoldAll)
                    Case ElementTypeLevel4
                        Set curRng = InsertEntry(doc, curRng, snapshots(i).CleanText, bmName, fontName, fontSize, tabPos, CmToPoints(1#), isBoldAll)
                End Select
            End If
        End If
    Next i
    Do While hIdx <= hCount
        Set curRng = InsertEntry(doc, curRng, hdrs(hIdx), bms(hIdx), fontName, fontSize, tabPos, 0#, True)
        hIdx = hIdx + 1
    Loop
    
    If curRng.End > startEntriesPos Then AddBookmark doc, BM_ZONE_TOC, doc.Range(startEntriesPos, curRng.End)
    Dim pM As Long
    pM = ModMarkerLexer.FindMarkerParagraph(doc, "=listfigurehere()")
    If pM > 1 And doc.Bookmarks.Exists(BM_LOF_HEADER) Then
        AddBookmark doc, BM_LOF_HEADER, doc.Paragraphs(pM - 1).Range
    End If
    pM = ModMarkerLexer.FindMarkerParagraph(doc, "=listtablehere()")
    If pM > 1 And doc.Bookmarks.Exists(BM_LOT_HEADER) Then
        AddBookmark doc, BM_LOT_HEADER, doc.Paragraphs(pM - 1).Range
    End If
    pM = ModMarkerLexer.FindMarkerParagraph(doc, "=listcontent()")
    If pM > 1 And doc.Bookmarks.Exists(BM_TOC_HEADER) Then
        AddBookmark doc, BM_TOC_HEADER, doc.Paragraphs(pM - 1).Range
    End If
    If doc.Bookmarks.Exists(BM_TOC_HEADER) Then AddBookmark doc, BM_SYS_GENERATED_BLOCK, doc.Range(doc.Bookmarks(BM_TOC_HEADER).Range.Start, curRng.End)
    If autoSync Then
        doc.Repaginate
        If Not SyncAllListPages(doc, pageMode) Then
            Err.Raise vbObjectError + 1006, "BuildTableOfContents", "SyncAllListPages failed"
        End If
    End If
CleanExit:
    If weStartedUndo Then If Not objUndo Is Nothing Then If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
    BuildTableOfContents = True
    Exit Function
ErrorHandler:
    ModRibbonCallbacks.g_LastError = "BuildTOC Error " & Err.Number & ": " & Err.Description
    If weStartedUndo Then If Not objUndo Is Nothing Then If objUndo.IsRecordingCustomRecord Then objUndo.EndCustomRecord
    BuildTableOfContents = False
End Function