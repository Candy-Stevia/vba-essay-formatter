Attribute VB_Name = "ModMarkerLexer"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModMarkerLexer
' Purpose: Identification, validation, and highlighting of user-placed boundary markers
' Architecture: Academic Document Format Engine (Win32 Unicode Safe)
' ==============================================================================

Public Type MarkerLocations
    HasContentStart As Boolean
    ContentStartParaIndex As Long
    HasContentEnd As Boolean
    ContentEndParaIndex As Long
    HasListContent As Boolean
    ListContentParaIndex As Long
    HasListTable As Boolean
    ListTableParaIndex As Long
    HasListFigure As Boolean
    ListFigureParaIndex As Long
    HasListAcronyms As Boolean
    ListAcronymsParaIndex As Long
End Type

' --- Marker Scanner & Validator ---

Private Function FindAllExactMarkerHits(ByVal doc As Document, ByVal markerLower As String, _
                                        ByVal displayName As String, ByVal isListMarker As Boolean, _
                                        ByRef outHas As Boolean, ByRef outIdx As Long, _
                                        ByRef listNames() As String, ByRef listIndices() As Long, _
                                        ByRef listCnt As Long, ByRef listCap As Long, _
                                        ByRef outErr As String) As Boolean
    Dim fRng As Range, pRng As Range, docEnd As Long, prevEnd As Long, cleanText As String, pIdx As Long
    Set fRng = doc.Content.Duplicate
    ModMathHelpers.InitDomFind fRng.Find, markerLower
    fRng.Find.MatchCase = False
    docEnd = doc.Content.End
    prevEnd = 0
    Do While fRng.Find.Execute
        If fRng.End <= prevEnd Then Exit Do
        prevEnd = fRng.End
        Set pRng = fRng.Paragraphs(1).Range
        cleanText = LCase$(Trim$(Replace(Replace(Replace(pRng.Text, vbCr, ""), Chr(7), ""), Chr(11), "")))
        If cleanText = markerLower Then
            pIdx = doc.Range(0, pRng.End).Paragraphs.Count
            If CBool(pRng.Information(12)) Then
                outErr = BuildTableMarkerError(displayName, pIdx)
                Exit Function
            End If
            If outHas Then
                outErr = BuildDuplicateMarkerError(displayName, outIdx, pIdx)
                Exit Function
            End If
            outHas = True
            outIdx = pIdx
            If isListMarker Then
                RecordListMarker listNames, listIndices, listCnt, listCap, displayName, pIdx
            End If
        End If
        If pRng.End >= docEnd Then Exit Do
        fRng.SetRange pRng.End, docEnd
    Loop
    FindAllExactMarkerHits = True
End Function

Public Function ScanMarkers(ByVal doc As Document, _
                            ByRef outMarkers As MarkerLocations, _
                            ByRef outErrorMessage As String) As Boolean
    Dim listMarkerNames() As String, listMarkerIndices() As Long
    Dim listMarkerCount As Long, listMarkerCap As Long
    Dim blankM As MarkerLocations
    outMarkers = blankM
    outErrorMessage = ""
    listMarkerCount = 0
    listMarkerCap = 16
    ReDim listMarkerNames(1 To listMarkerCap)
    ReDim listMarkerIndices(1 To listMarkerCap)

    If Not FindAllExactMarkerHits(doc, "=contentstart()", "=contentStart()", False, _
        outMarkers.HasContentStart, outMarkers.ContentStartParaIndex, _
        listMarkerNames, listMarkerIndices, listMarkerCount, listMarkerCap, outErrorMessage) Then Exit Function
    If Not FindAllExactMarkerHits(doc, "=contentend()", "=contentEnd()", False, _
        outMarkers.HasContentEnd, outMarkers.ContentEndParaIndex, _
        listMarkerNames, listMarkerIndices, listMarkerCount, listMarkerCap, outErrorMessage) Then Exit Function
    If Not FindAllExactMarkerHits(doc, "=listcontent()", "=listContent()", True, _
        outMarkers.HasListContent, outMarkers.ListContentParaIndex, _
        listMarkerNames, listMarkerIndices, listMarkerCount, listMarkerCap, outErrorMessage) Then Exit Function
    If Not FindAllExactMarkerHits(doc, "=listtablehere()", "=listTableHere()", True, _
        outMarkers.HasListTable, outMarkers.ListTableParaIndex, _
        listMarkerNames, listMarkerIndices, listMarkerCount, listMarkerCap, outErrorMessage) Then Exit Function
    If Not FindAllExactMarkerHits(doc, "=listfigurehere()", "=listFigureHere()", True, _
        outMarkers.HasListFigure, outMarkers.ListFigureParaIndex, _
        listMarkerNames, listMarkerIndices, listMarkerCount, listMarkerCap, outErrorMessage) Then Exit Function
    If Not FindAllExactMarkerHits(doc, "=listacronyms()", "=listAcronyms()", True, _
        outMarkers.HasListAcronyms, outMarkers.ListAcronymsParaIndex, _
        listMarkerNames, listMarkerIndices, listMarkerCount, listMarkerCap, outErrorMessage) Then Exit Function

    If Not outMarkers.HasContentStart Then
        outErrorMessage = ModConstants.Uni("Missing =contentStart(). Please add it on its own line. (Thi\u1ebfu th\u1ebb =contentStart(), h\u00e3y th\u00eam tr\u00ean m\u1ed9t d\u00f2ng ri\u00eang.)")
        Exit Function
    End If
    If Not outMarkers.HasContentEnd Then
        outErrorMessage = ModConstants.Uni("Missing =contentEnd(). Please add it on its own line. (Thi\u1ebfu th\u1ebb =contentEnd(), h\u00e3y th\u00eam tr\u00ean m\u1ed9t d\u00f2ng ri\u00eang.)")
        Exit Function
    End If
    If outMarkers.ContentStartParaIndex >= outMarkers.ContentEndParaIndex Then
        outErrorMessage = ModConstants.Uni("=contentStart() (para ") & outMarkers.ContentStartParaIndex & _
                          ModConstants.Uni(") must be before =contentEnd() (para ") & outMarkers.ContentEndParaIndex & _
                          ModConstants.Uni("). (=contentStart() ph\u1ea3i n\u1eb1m tr\u01b0\u1edbc =contentEnd().)")
        Exit Function
    End If

    Dim k As Long, badMarkers As String
    badMarkers = ""
    For k = 1 To listMarkerCount
        If listMarkerIndices(k) > outMarkers.ContentStartParaIndex And _
           listMarkerIndices(k) < outMarkers.ContentEndParaIndex Then
            If Len(badMarkers) > 0 Then badMarkers = badMarkers & ", "
            badMarkers = badMarkers & listMarkerNames(k) & " (line " & listMarkerIndices(k) & ")"
        End If
    Next k
    If Len(badMarkers) > 0 Then
        outErrorMessage = BuildMarkerPlacementWarning(badMarkers)
        Exit Function
    End If
    ScanMarkers = True
End Function

Public Function HasAnySystemMarker(ByVal doc As Document) As Boolean
    Dim mNames As Variant, k As Long
    mNames = Array("=contentstart()", "=contentend()", "=listcontent()", "=listtablehere()", "=listfigurehere()", "=listacronyms()")
    For k = LBound(mNames) To UBound(mNames)
        If FindMarkerParagraph(doc, CStr(mNames(k))) > 0 Then
            HasAnySystemMarker = True
            Exit Function
        End If
    Next k
End Function

Private Function BuildDuplicateMarkerError(ByVal markerText As String, ByVal idx1 As Long, ByVal idx2 As Long) As String
    BuildDuplicateMarkerError = ModConstants.Uni("Duplicate ") & markerText & _
        ModConstants.Uni(" at paragraphs ") & idx1 & " & " & idx2 & _
        ModConstants.Uni(". (Tr\u00f9ng th\u1ebb ") & markerText & _
        ModConstants.Uni(" \u1edf \u0111o\u1ea1n ") & idx1 & " & " & idx2 & ".)"
End Function

Private Sub RecordListMarker(ByRef names() As String, ByRef indices() As Long, _
                             ByRef cnt As Long, ByRef cap As Long, _
                             ByVal mName As String, ByVal pIdx As Long)
    cnt = cnt + 1
    If cnt > cap Then
        cap = cap * 2
        ReDim Preserve names(1 To cap)
        ReDim Preserve indices(1 To cap)
    End If
    names(cnt) = mName
    indices(cnt) = pIdx
End Sub

Private Function BuildTableMarkerError(ByVal markerText As String, ByVal paraIdx As Long) As String
    BuildTableMarkerError = ModConstants.Uni("Marker ") & markerText & _
        ModConstants.Uni(" at paragraph ") & paraIdx & _
        ModConstants.Uni(" cannot be inside a table. (Th\u1ebb ") & markerText & _
        ModConstants.Uni(" \u1edf \u0111o\u1ea1n ") & paraIdx & _
        ModConstants.Uni(" kh\u00f4ng \u0111\u01b0\u1ee3c n\u1eb1m trong b\u1ea3ng.)")
End Function

Private Function BuildMarkerPlacementWarning(ByVal badList As String) As String
    BuildMarkerPlacementWarning = "List marker(s) found inside body: " & badList & vbCrLf & _
        ModConstants.Uni("(Ph\u00e1t hi\u1ec7n th\u1ebb danh m\u1ee5c n\u1eb1m trong th\u00e2n b\u00e0i.)") & vbCrLf & vbCrLf & _
        "Recommended order (Th" & ModConstants.Uni("\u1ee9 t\u1ef1 \u0111\u00fang") & "):" & vbCrLf & _
        ModConstants.Uni("  TOC -> Lists -> =contentStart() -> [Body / Th\u00e2n b\u00e0i] -> =contentEnd()")
End Function

' --- Marker Highlighting Routine (wdYellow = 7, only exact markers) ---

Public Sub HighlightMarkers(ByVal doc As Document)
    On Error Resume Next
    Dim mNames As Variant, k As Long, fRng As Range, pRng As Range, docEnd As Long, prevEnd As Long
    mNames = Array("=contentstart()", "=contentend()", "=listcontent()", "=listtablehere()", "=listfigurehere()", "=listacronyms()")
    docEnd = doc.Content.End
    For k = LBound(mNames) To UBound(mNames)
        Set fRng = doc.Content.Duplicate
        ModMathHelpers.InitDomFind fRng.Find, CStr(mNames(k))
        fRng.Find.MatchCase = False
        prevEnd = 0
        Do While fRng.Find.Execute
            If fRng.End <= prevEnd Then Exit Do
            prevEnd = fRng.End
            Set pRng = fRng.Paragraphs(1).Range
            If LCase$(Trim$(Replace(Replace(Replace(pRng.Text, vbCr, ""), Chr(7), ""), Chr(11), ""))) = CStr(mNames(k)) And Not CBool(pRng.Information(12)) Then
                pRng.HighlightColorIndex = 7
                ModFormatSingle.SafeSetOutlineLevel pRng, 10
            End If
            If pRng.End >= docEnd Then Exit Do
            fRng.SetRange pRng.End, docEnd
        Loop
    Next k
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function FindMarkerParagraph(ByVal doc As Document, ByVal markerText As String) As Long
    Dim fRng As Range, pRng As Range, docEnd As Long, prevEnd As Long
    Set fRng = doc.Content.Duplicate
    ModMathHelpers.InitDomFind fRng.Find, markerText
    fRng.Find.MatchCase = False
    docEnd = doc.Content.End
    Do While fRng.Find.Execute
        If fRng.End <= prevEnd Then Exit Do
        prevEnd = fRng.End
        Set pRng = fRng.Paragraphs(1).Range
        If LCase$(Trim$(Replace(Replace(Replace(pRng.Text, vbCr, ""), Chr(7), ""), Chr(11), ""))) = markerText And Not CBool(pRng.Information(12)) Then
            FindMarkerParagraph = doc.Range(0, pRng.End).Paragraphs.Count
            Exit Function
        End If
        If pRng.End >= docEnd Then Exit Do
        fRng.SetRange pRng.End, docEnd
    Loop
End Function

Public Sub NormalizeContentBreaks(ByVal doc As Document)
    Dim pStart As Long, pEnd As Long
    pStart = FindMarkerParagraph(doc, "=contentstart()")
    pEnd = FindMarkerParagraph(doc, "=contentend()")
    If pStart = 0 Or pEnd = 0 Or pEnd <= pStart + 1 Then Exit Sub
    
    Dim totalParas As Long
    totalParas = doc.Paragraphs.Count
    Dim hasTables As Boolean, hasOMaths As Boolean
    hasTables = (doc.Tables.Count > 0)
    hasOMaths = (doc.OMaths.Count > 0)
    If hasTables Then
        ModTableIndex.BuildTableIndex doc
    End If
    ModTableIndex.BuildListIndex doc
    ModTableIndex.BuildOMathIndex doc
    
    Dim latexFlags() As Boolean, latexWarn As Long
    ModLatex.ComputeLatexFlags doc, latexFlags, latexWarn
    
    Dim i As Long, pRng As Range, rawText As String, firstLine As String
    Dim p As Paragraph, cRng As Range, splitCount As Long
    Set cRng = doc.Range(doc.Paragraphs(pStart).Range.End, doc.Paragraphs(pEnd).Range.Start)
    splitCount = 0
    
    ' Phase 1: Handle Chr(11) (Shift + Enter) via bottom-up coordinate execution (F-10)
    If cRng.Start < cRng.End Then
        Dim findRng As Range, hasSoftBreak As Boolean
        Set findRng = cRng.Duplicate
        ModMathHelpers.InitDomFind findRng.Find, "^l"
        findRng.Find.Replacement.ClearFormatting
        hasSoftBreak = findRng.Find.Execute
        If hasSoftBreak Then
            Dim splitStarts() As Long, splitEnds() As Long
            ReDim splitStarts(1 To totalParas)
            ReDim splitEnds(1 To totalParas)
            Dim pIdx As Long, skipPara As Boolean
            pIdx = pStart
            For Each p In cRng.Paragraphs
                pIdx = pIdx + 1
                If pIdx >= pEnd Then Exit For
                skipPara = False
                If pIdx <= totalParas Then
                    If latexFlags(pIdx) Then skipPara = True
                End If
                Set pRng = p.Range
                If Not skipPara Then
                    If ModTableIndex.IsPosInList(pRng.Start) Then
                        skipPara = True
                    ElseIf hasTables Then
                        If ModTableIndex.IsPosInTable(pRng.Start) Then skipPara = True
                    End If
                    If (Not skipPara) And hasOMaths Then
                        If ModTableIndex.HasOMathInSpan(pRng.Start, pRng.End) Then skipPara = True
                    End If
                End If
                If Not skipPara Then
                    rawText = pRng.Text
                    If InStr(rawText, Chr(11)) > 0 Then
                        firstLine = Trim$(Split(Replace(rawText, vbCr, ""), Chr(11))(0))
                        If Not ModLocale.IsPrefixOnlyHeading(firstLine) Then
                            splitCount = splitCount + 1
                            splitStarts(splitCount) = pRng.Start
                            splitEnds(splitCount) = pRng.End
                        End If
                    End If
                End If
            Next p
            Dim j As Long
            For j = splitCount To 1 Step -1
                Set pRng = doc.Range(splitStarts(j), splitEnds(j))
                ModMathHelpers.InitDomFind pRng.Find, "^l"
                pRng.Find.Replacement.ClearFormatting
                pRng.Find.Replacement.Text = "^p"
                pRng.Find.Execute Replace:=2
            Next j
            If splitCount > 0 Then
                ModTableIndex.InvalidateAllIndexes
            End If
        End If
    End If
    
    ' Phase 2: Safely remove blank paragraphs from bottom up (F-03, R-06)
    If splitCount > 0 Then
        pStart = FindMarkerParagraph(doc, "=contentstart()")
        pEnd = FindMarkerParagraph(doc, "=contentend()")
        If pStart = 0 Or pEnd = 0 Or pEnd <= pStart + 1 Then Exit Sub
        If hasTables Then
            ModTableIndex.BuildTableIndex doc
        End If
    End If
    
    Dim blankStarts() As Long, blankCount As Long
    blankCount = 0
    ReDim blankStarts(1 To doc.Paragraphs.Count)
    
    Dim hasShapes As Boolean
    hasShapes = (doc.Shapes.Count > 0)
    Dim hasInline As Boolean
    hasInline = (doc.InlineShapes.Count > 0)
    Dim r As Range, curInTbl As Boolean, prevInTbl As Boolean, nextInTbl As Boolean, isBlank As Boolean, rText As String
    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        If i >= pEnd Then Exit For
        If i > pStart And i < pEnd Then
            Set r = p.Range
            If (r.End - r.Start = 1) Then
                rText = r.Text
                If Len(rText) = 1 And rText = vbCr Then
                    curInTbl = False
                    If hasTables Then curInTbl = ModTableIndex.IsPosInTable(r.Start)
                    If Not curInTbl Then
                        isBlank = (r.Bookmarks.Count = 0 And (Not r.ParagraphFormat.PageBreakBefore))
                        If isBlank And hasInline Then isBlank = (r.InlineShapes.Count = 0)
                        If isBlank And hasShapes Then isBlank = (r.ShapeRange.Count = 0)
                        If isBlank Then
                            prevInTbl = False
                            nextInTbl = False
                            If hasTables Then
                                If Not p.Previous Is Nothing Then prevInTbl = ModTableIndex.IsPosInTable(p.Previous.Range.Start)
                                If Not p.Next Is Nothing Then nextInTbl = ModTableIndex.IsPosInTable(p.Next.Range.Start)
                            End If
                            If (Not prevInTbl) And (Not nextInTbl) Then
                                blankCount = blankCount + 1
                                blankStarts(blankCount) = r.Start
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next p
    
    If blankCount > 0 Then
        Dim k As Long
        For k = blankCount To 1 Step -1
            doc.Range(blankStarts(k), blankStarts(k) + 1).Delete
        Next k
        ModTableIndex.InvalidateAllIndexes
    End If
End Sub

Public Function ScanMarkersSummaryForTest(ByVal doc As Document) As String
    Dim m As MarkerLocations, errMsg As String
    If ScanMarkers(doc, m, errMsg) Then
        ScanMarkersSummaryForTest = "OK|" & CStr(m.ContentStartParaIndex) & "|" & CStr(m.ContentEndParaIndex) & "|" & CStr(m.ListContentParaIndex)
    Else
        ScanMarkersSummaryForTest = "ERR|" & errMsg
    End If
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
