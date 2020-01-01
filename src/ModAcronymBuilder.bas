Attribute VB_Name = "ModAcronymBuilder"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModAcronymBuilder
' Purpose: Builds 2- or 3-column List of Acronyms table at =listAcronyms()
' Architecture: Academic Document Format Engine (Zero-Drift Idempotent Builder)
' ==============================================================================

Private Function CleanSingleLineText(ByVal s As String) As String
    CleanSingleLineText = LCase$(Trim$(Replace(Replace(Replace(s, vbCr, ""), Chr(7), ""), Chr(11), "")))
End Function

Private Function IsManualLoaHeading(ByVal txt As String) As Boolean
    Dim clean As String
    clean = CleanSingleLineText(txt)
    If Len(clean) = 0 Or Len(clean) > 80 Then Exit Function
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True
        re.Global = False
        re.Pattern = "^(?:(?:list of )?(?:acronyms?|abbreviations?)|" & _
                     ModConstants.Uni("(?:danh m\u1ee5c )?(?:c\u00e1c )?(?:t\u1eeb|ch\u1eef) vi\u1ebft t\u1eaft|danh m\u1ee5c k\u00fd hi\u1ec7u v\u00e0 ch\u1eef vi\u1ebft t\u1eaft") & ")$"
    End If
    IsManualLoaHeading = re.Test(clean)
End Function

Private Function FindManualLoaHeadingPara(ByVal doc As Document) As Long
    Dim csIdx As Long, i As Long, p As Paragraph
    csIdx = ModMarkerLexer.FindMarkerParagraph(doc, "=contentstart()")
    If csIdx <= 1 Then csIdx = doc.Paragraphs.Count + 1
    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        If i >= csIdx Then Exit For
        If Not CBool(p.Range.Information(12)) And p.Range.Hyperlinks.Count = 0 Then
            If IsManualLoaHeading(p.Range.Text) Then
                FindManualLoaHeadingPara = i
                Exit Function
            End If
        End If
    Next p
End Function

Public Function CanLocateAcronymTarget(ByVal doc As Document) As Boolean
    If doc Is Nothing Then Exit Function
    If ModMarkerLexer.FindMarkerParagraph(doc, "=listacronyms()") > 0 Then
        CanLocateAcronymTarget = True
    ElseIf doc.Bookmarks.Exists(BM_LOA_HEADER) Or doc.Bookmarks.Exists(BM_ZONE_LOA) Then
        CanLocateAcronymTarget = True
    Else
        CanLocateAcronymTarget = (FindManualLoaHeadingPara(doc) > 0)
    End If
End Function

Public Sub WipeAcronymZone(ByVal doc As Document, ByVal removeHeaderToo As Boolean)
    On Error Resume Next
    Dim zRng As Range, t As Table, dEnd As Long, hRng As Range
    If doc.Bookmarks.Exists(BM_ZONE_LOA) Then
        Set zRng = doc.Bookmarks(BM_ZONE_LOA).Range
        Do While zRng.Tables.Count > 0
            zRng.Tables(1).Delete
        Loop
        Set zRng = doc.Bookmarks(BM_ZONE_LOA).Range
        dEnd = doc.Range.End
        If zRng.End > zRng.Start And zRng.End < dEnd Then
            If InStr(1, zRng.Text, Chr(12)) = 0 Then doc.Range(zRng.Start, zRng.End).Delete
        End If
        If doc.Bookmarks.Exists(BM_ZONE_LOA) Then doc.Bookmarks(BM_ZONE_LOA).Delete
    End If
    If removeHeaderToo And doc.Bookmarks.Exists(BM_LOA_HEADER) Then
        Set hRng = doc.Bookmarks(BM_LOA_HEADER).Range.Paragraphs(1).Range
        dEnd = doc.Range.End
        If hRng.End < dEnd And InStr(1, hRng.Text, Chr(12)) = 0 Then hRng.Delete
        If doc.Bookmarks.Exists(BM_LOA_HEADER) Then doc.Bookmarks(BM_LOA_HEADER).Delete
    End If
    Err.Clear
End Sub

Private Sub SortAndDedupeAcronyms(ByRef rawItems() As PatchAcronymItem, ByVal rawCount As Long, _
                                  ByRef outItems() As PatchAcronymItem, ByRef outCount As Long, ByRef outHasCol3 As Boolean)
    Dim dIdx As Object, i As Long, k As String, slot As Long
    Dim j As Long, tmp As PatchAcronymItem
    Set dIdx = CreateObject("Scripting.Dictionary")
    outCount = 0
    outHasCol3 = False
    ReDim outItems(1 To rawCount)
    For i = 1 To rawCount
        k = LCase$(Trim$(rawItems(i).Abbr))
        If Len(k) > 0 And Len(Trim$(rawItems(i).FullForm)) > 0 Then
            If dIdx.Exists(k) Then
                slot = CLng(dIdx(k))
                outItems(slot) = rawItems(i)
            Else
                outCount = outCount + 1
                dIdx.Add k, outCount
                outItems(outCount) = rawItems(i)
            End If
        End If
    Next i
    For i = 1 To outCount - 1
        For j = i + 1 To outCount
            If StrComp(outItems(j).Abbr, outItems(i).Abbr, vbTextCompare) < 0 Then
                tmp = outItems(i)
                outItems(i) = outItems(j)
                outItems(j) = tmp
            End If
        Next j
    Next i
    For i = 1 To outCount
        If outItems(i).HasNote Then
            outHasCol3 = True
            Exit For
        End If
    Next i
End Sub

Private Sub FormatLoaHeadingPara(ByVal rng As Range, ByVal fontName As String, ByVal fontSize As Single)
    rng.HighlightColorIndex = 0
    rng.Font.Name = fontName
    rng.Font.Size = fontSize
    rng.Font.Bold = True
    rng.Font.Italic = False
    rng.Case = 1
    rng.ParagraphFormat.Alignment = 1
    rng.ParagraphFormat.FirstLineIndent = 0#
    rng.ParagraphFormat.LeftIndent = 0#
    rng.ParagraphFormat.SpaceBefore = 6#
    rng.ParagraphFormat.SpaceAfter = 6#
    rng.ParagraphFormat.KeepWithNext = True
End Sub

Private Sub ApplyCol3NoteWithBoldLabels(ByVal cellRng As Range, ByVal noteText As String, _
                                        ByVal fontName As String, ByVal fontSize As Single)
    Dim crRng As Range, p As Paragraph, pTxt As String, colonPos As Long, lbl As String, bLen As Long
    Dim rawLines() As String, lnIdx As Long, tLine As String, cleaned As String
    rawLines = Split(Replace(Replace(noteText, vbCrLf, vbCr), vbLf, vbCr), vbCr)
    cleaned = ""
    For lnIdx = LBound(rawLines) To UBound(rawLines)
        tLine = Trim$(rawLines(lnIdx))
        If Len(tLine) > 0 Then
            If Len(cleaned) > 0 Then cleaned = cleaned & vbCr
            cleaned = cleaned & tLine
        End If
    Next lnIdx
    Set crRng = cellRng.Duplicate
    crRng.End = crRng.End - 1
    crRng.Text = cleaned
    crRng.Font.Name = fontName
    crRng.Font.Size = fontSize
    crRng.Font.Bold = False
    crRng.ParagraphFormat.Alignment = 0
    crRng.ParagraphFormat.FirstLineIndent = 0#
    crRng.ParagraphFormat.LeftIndent = 0#
    crRng.ParagraphFormat.SpaceBefore = 2#
    crRng.ParagraphFormat.SpaceAfter = 2#
    For Each p In crRng.Paragraphs
        pTxt = Trim$(Replace(Replace(p.Range.Text, vbCr, ""), Chr(7), ""))
        colonPos = InStr(1, pTxt, ":")
        If colonPos > 1 And colonPos <= 16 Then
            lbl = Left$(pTxt, colonPos)
            If StrComp(lbl, ModConstants.Uni("\u00dd ngh\u0129a:"), vbTextCompare) = 0 Or _
               StrComp(lbl, ModConstants.Uni("Ngu\u1ed3n:"), vbTextCompare) = 0 Or _
               lbl = ModConstants.Uni("\u00dd NGH\u0128A:") Or lbl = ModConstants.Uni("NGU\u1ed2N:") Or _
               StrComp(lbl, "Definition:", vbTextCompare) = 0 Or _
               StrComp(lbl, "Meaning:", vbTextCompare) = 0 Or _
               StrComp(lbl, "Source:", vbTextCompare) = 0 Then
                bLen = InStr(1, pTxt, ":")
                If bLen > 0 Then cellRng.Document.Range(p.Range.Start, p.Range.Start + bLen).Font.Bold = True
            End If
        End If
    Next p
End Sub

Private Sub StyleAcronymCellText(ByVal cellRng As Range, ByVal txt As String, ByVal fontName As String, _
                                 ByVal fontSize As Single, ByVal isBold As Boolean, ByVal alignVal As Long)
    Dim r As Range
    Set r = cellRng.Duplicate
    r.End = r.End - 1
    r.Text = txt
    r.Font.Name = fontName
    r.Font.Size = fontSize
    r.Font.Bold = isBold
    r.Font.Italic = False
    r.ParagraphFormat.Alignment = alignVal
    r.ParagraphFormat.FirstLineIndent = 0#
    r.ParagraphFormat.LeftIndent = 0#
    r.ParagraphFormat.SpaceBefore = 2#
    r.ParagraphFormat.SpaceAfter = 2#
End Sub

Private Sub ConfigureTableGrid(ByVal tbl As Table, ByVal hasCol3 As Boolean, ByVal usableW As Single)
    tbl.AllowAutoFit = False
    tbl.Borders.InsideLineStyle = 1
    tbl.Borders.OutsideLineStyle = 1
    tbl.Borders.InsideLineWidth = 4
    tbl.Borders.OutsideLineWidth = 4
    tbl.Borders.InsideColor = RGB(191, 191, 191)
    tbl.Borders.OutsideColor = RGB(191, 191, 191)
    tbl.TopPadding = 4#
    tbl.BottomPadding = 4#
    tbl.LeftPadding = 6#
    tbl.RightPadding = 6#
    If usableW < 200# Then usableW = 460#
    If hasCol3 Then
        tbl.Columns(1).Width = usableW * 0.16#
        tbl.Columns(2).Width = usableW * 0.32#
        tbl.Columns(3).Width = usableW - tbl.Columns(1).Width - tbl.Columns(2).Width
    Else
        tbl.Columns(1).Width = usableW * 0.24#
        tbl.Columns(2).Width = usableW - tbl.Columns(1).Width
    End If
End Sub

Private Sub DeleteLegacyTableAfterHeader(ByVal doc As Document, ByVal hdrRng As Range)
    If hdrRng.End >= doc.Range.End Then Exit Sub
    Dim checkRng As Range, pTxt As String
    Set checkRng = doc.Range(hdrRng.End, hdrRng.End + 1)
    If Not CBool(checkRng.Information(12)) Then
        pTxt = CleanSingleLineText(checkRng.Paragraphs(1).Range.Text)
        If Len(pTxt) = 0 And checkRng.Paragraphs(1).Range.End < doc.Range.End Then
            Set checkRng = doc.Range(checkRng.Paragraphs(1).Range.End, checkRng.Paragraphs(1).Range.End + 1)
        End If
    End If
    If CBool(checkRng.Information(12)) And checkRng.Tables.Count > 0 Then
        Dim colCnt As Long
        colCnt = checkRng.Tables(1).Columns.Count
        If colCnt = 2 Or colCnt = 3 Then checkRng.Tables(1).Delete
    End If
End Sub

Private Sub RefreshLoaHeaderTitle(ByVal doc As Document, ByRef ioHdrRng As Range, ByVal hdrTitle As String)
    Dim curClean As String, txtRng As Range
    curClean = CleanSingleLineText(ioHdrRng.Text)
    If StrComp(curClean, CleanSingleLineText(hdrTitle), vbTextCompare) <> 0 Then
        If ioHdrRng.End - 1 >= ioHdrRng.Start Then
            Set txtRng = doc.Range(ioHdrRng.Start, ioHdrRng.End - 1)
            txtRng.Text = hdrTitle
            Set ioHdrRng = doc.Range(txtRng.Start, txtRng.Start).Paragraphs(1).Range
        End If
    End If
End Sub

Private Function ResolveLoaInsertRange(ByVal doc As Document, ByVal hdrTitle As String, _
                                       ByVal fontName As String, ByVal fontSize As Single, _
                                       ByRef outHdrRng As Range) As Range
    Dim mPara As Long, hPara As Long, anchorEnd As Long, hRng As Range, nextRng As Range
    mPara = ModMarkerLexer.FindMarkerParagraph(doc, "=listacronyms()")
    If mPara > 0 Then
        WipeAcronymZone doc, True
        mPara = ModMarkerLexer.FindMarkerParagraph(doc, "=listacronyms()")
        anchorEnd = doc.Paragraphs(mPara).Range.End
        If anchorEnd >= doc.Range.End Then doc.Paragraphs(mPara).Range.InsertAfter vbCr
        Set hRng = doc.Range(doc.Paragraphs(mPara).Range.End, doc.Paragraphs(mPara).Range.End)
        hRng.InsertBefore hdrTitle & vbCr
        Set outHdrRng = doc.Paragraphs(mPara + 1).Range
        FormatLoaHeadingPara outHdrRng, fontName, fontSize
        ModBookmark.AddBookmark doc, BM_LOA_HEADER, outHdrRng
        If outHdrRng.End >= doc.Range.End Then outHdrRng.InsertAfter vbCr
        Set nextRng = doc.Range(outHdrRng.End, outHdrRng.End)
        nextRng.Collapse 1
        Set ResolveLoaInsertRange = nextRng
        Exit Function
    End If
    WipeAcronymZone doc, False
    If doc.Bookmarks.Exists(BM_LOA_HEADER) Then
        Set outHdrRng = doc.Bookmarks(BM_LOA_HEADER).Range.Paragraphs(1).Range
        RefreshLoaHeaderTitle doc, outHdrRng, hdrTitle
        FormatLoaHeadingPara outHdrRng, fontName, fontSize
        ModBookmark.AddBookmark doc, BM_LOA_HEADER, outHdrRng
        DeleteLegacyTableAfterHeader doc, outHdrRng
        If outHdrRng.End >= doc.Range.End Then outHdrRng.InsertAfter vbCr
        Set nextRng = doc.Range(outHdrRng.End, outHdrRng.End)
        nextRng.Collapse 1
        Set ResolveLoaInsertRange = nextRng
        Exit Function
    End If
    hPara = FindManualLoaHeadingPara(doc)
    If hPara > 0 Then
        Set outHdrRng = doc.Paragraphs(hPara).Range
        FormatLoaHeadingPara outHdrRng, fontName, fontSize
        ModBookmark.AddBookmark doc, BM_LOA_HEADER, outHdrRng
        DeleteLegacyTableAfterHeader doc, outHdrRng
        If outHdrRng.End >= doc.Range.End Then outHdrRng.InsertAfter vbCr
        Set nextRng = doc.Range(outHdrRng.End, outHdrRng.End)
        nextRng.Collapse 1
        Set ResolveLoaInsertRange = nextRng
    End If
End Function

Public Function BuildAcronymTable(ByVal doc As Document, ByRef rawItems() As PatchAcronymItem, _
                                  ByVal rawCount As Long, ByRef outWrittenCount As Long) As Boolean
    outWrittenCount = 0
    If doc Is Nothing Or rawCount <= 0 Then Exit Function
    Dim items() As PatchAcronymItem, count As Long, hasCol3 As Boolean
    SortAndDedupeAcronyms rawItems, rawCount, items, count, hasCol3
    If count <= 0 Then Exit Function
    ModRibbonCallbacks.EnsureConfigInitialized
    Dim fontName As String, bodySize As Single, cellSize As Single, lang As String
    fontName = g_Config.FontName
    bodySize = g_Config.FontSize
    cellSize = bodySize
    lang = g_Config.ListLang
    Dim hdrTitle As String, hdrRng As Range, insRng As Range
    hdrTitle = ModLocale.GetLoaHeaderTitle(lang)
    Set insRng = ResolveLoaInsertRange(doc, hdrTitle, fontName, bodySize, hdrRng)
    If insRng Is Nothing Then Exit Function
    Dim colCnt As Long, tbl As Table, usableW As Single, r As Long
    colCnt = IIf(hasCol3, 3, 2)
    usableW = ModPageNumberSync.CalculateRightTab(doc, hdrRng)
    Set tbl = doc.Tables.Add(insRng, count + 1, colCnt)
    ConfigureTableGrid tbl, hasCol3, usableW
    tbl.Rows(1).HeadingFormat = -1
    tbl.Rows(1).AllowBreakAcrossPages = False
    If lang = LANG_VI Then
        StyleAcronymCellText tbl.Cell(1, 1).Range, ModConstants.Uni("Ch\u1eef vi\u1ebft t\u1eaft"), fontName, cellSize, True, 1
        StyleAcronymCellText tbl.Cell(1, 2).Range, ModConstants.Uni("D\u1ea1ng \u0111\u1ea7y \u0111\u1ee7"), fontName, cellSize, True, 1
        If hasCol3 Then StyleAcronymCellText tbl.Cell(1, 3).Range, ModConstants.Uni("Gi\u1ea3i th\u00edch & Ngu\u1ed3n"), fontName, cellSize, True, 1
    Else
        StyleAcronymCellText tbl.Cell(1, 1).Range, "Acronym", fontName, cellSize, True, 1
        StyleAcronymCellText tbl.Cell(1, 2).Range, "Full Term", fontName, cellSize, True, 1
        If hasCol3 Then StyleAcronymCellText tbl.Cell(1, 3).Range, "Definition & Source", fontName, cellSize, True, 1
    End If
    For r = 1 To count
        tbl.Rows(r + 1).AllowBreakAcrossPages = False
        StyleAcronymCellText tbl.Cell(r + 1, 1).Range, items(r).Abbr, fontName, cellSize, True, 1
        StyleAcronymCellText tbl.Cell(r + 1, 2).Range, items(r).FullForm, fontName, cellSize, False, 0
        If hasCol3 Then
            If items(r).HasNote Then
                ApplyCol3NoteWithBoldLabels tbl.Cell(r + 1, 3).Range, items(r).NoteText, fontName, cellSize
            Else
                StyleAcronymCellText tbl.Cell(r + 1, 3).Range, "", fontName, cellSize, False, 0
            End If
        End If
    Next r
    tbl.Range.Cells.VerticalAlignment = 1
    tbl.Range.HighlightColorIndex = 0
    ModBookmark.AddBookmark doc, BM_ZONE_LOA, tbl.Range
    outWrittenCount = count
    BuildAcronymTable = True
End Function

Public Function ApplyAcronymsFromPatchForTest(ByVal doc As Document, ByVal rawPatch As String) As String
    Dim payload As PatchPayload, written As Long, colCnt As Long
    If Not ModPatchParser.ParseAiPatch(rawPatch, payload) Then
        ApplyAcronymsFromPatchForTest = "PARSE_EMPTY"
        Exit Function
    End If
    If payload.AcronymCount <= 0 Then
        ApplyAcronymsFromPatchForTest = "ZERO_ACRONYMS|" & payload.SkippedAcronyms & "|" & payload.CiteCount & "|" & payload.SkippedCites & "|" & payload.RefCount & "|" & payload.FixCount
        Exit Function
    End If
    If Not BuildAcronymTable(doc, payload.Acronyms, payload.AcronymCount, written) Then
        ApplyAcronymsFromPatchForTest = "NO_ANCHOR|" & payload.SkippedAcronyms
        Exit Function
    End If
    colCnt = doc.Bookmarks(BM_ZONE_LOA).Range.Tables(1).Columns.Count
    ApplyAcronymsFromPatchForTest = "OK|" & written & "|" & colCnt & "|" & payload.SkippedAcronyms & "|" & payload.CiteCount & "|" & payload.SkippedCites & "|" & payload.RefCount & "|" & payload.FixCount
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
