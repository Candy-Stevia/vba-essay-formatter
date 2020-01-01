Attribute VB_Name = "ModMarkerLexer"
Option Explicit

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
End Type

' --- Marker Scanner & Validator ---

Public Function ScanMarkers(ByVal doc As Document, _
                            ByRef outMarkers As MarkerLocations, _
                            ByRef outErrorMessage As String) As Boolean
    Dim i As Long, rawText As String, cleanText As String
    Dim listMarkerNames() As String
    Dim listMarkerIndices() As Long
    Dim listMarkerCount As Long
    Dim listMarkerCap As Long
    
    outErrorMessage = ""
    outMarkers.HasContentStart = False
    outMarkers.ContentStartParaIndex = 0
    outMarkers.HasContentEnd = False
    outMarkers.ContentEndParaIndex = 0
    outMarkers.HasListContent = False
    outMarkers.ListContentParaIndex = 0
    outMarkers.HasListTable = False
    outMarkers.ListTableParaIndex = 0
    outMarkers.HasListFigure = False
    outMarkers.ListFigureParaIndex = 0
    
    listMarkerCount = 0
    listMarkerCap = 16
    ReDim listMarkerNames(1 To listMarkerCap)
    ReDim listMarkerIndices(1 To listMarkerCap)

    Dim p As Paragraph
    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        rawText = p.Range.Text
        cleanText = LCase$(Trim$(Replace(Replace(rawText, vbCr, ""), Chr(11), "")))
        
        Select Case cleanText
            Case "=contentstart()"
                If outMarkers.HasContentStart Then
                    outErrorMessage = ModConstants.Uni("Duplicate =contentStart() at paragraphs ") & _
                                      outMarkers.ContentStartParaIndex & " & " & i & _
                                      ModConstants.Uni(". (Tr\u00f9ng th\u1ebb =contentStart() \u1edf \u0111o\u1ea1n ") & _
                                      outMarkers.ContentStartParaIndex & " & " & i & ".)"
                    ScanMarkers = False
                    Exit Function
                End If
                outMarkers.HasContentStart = True
                outMarkers.ContentStartParaIndex = i
                
            Case "=contentend()"
                If outMarkers.HasContentEnd Then
                    outErrorMessage = ModConstants.Uni("Duplicate =contentEnd() at paragraphs ") & _
                                      outMarkers.ContentEndParaIndex & " & " & i & _
                                      ModConstants.Uni(". (Tr\u00f9ng th\u1ebb =contentEnd() \u1edf \u0111o\u1ea1n ") & _
                                      outMarkers.ContentEndParaIndex & " & " & i & ".)"
                    ScanMarkers = False
                    Exit Function
                End If
                outMarkers.HasContentEnd = True
                outMarkers.ContentEndParaIndex = i
                
            Case "=listcontent()"
                outMarkers.HasListContent = True
                outMarkers.ListContentParaIndex = i
                listMarkerCount = listMarkerCount + 1
                If listMarkerCount > listMarkerCap Then
                    listMarkerCap = listMarkerCap * 2
                    ReDim Preserve listMarkerNames(1 To listMarkerCap)
                    ReDim Preserve listMarkerIndices(1 To listMarkerCap)
                End If
                listMarkerNames(listMarkerCount) = "=listContent()"
                listMarkerIndices(listMarkerCount) = i

            Case "=listtablehere()"
                outMarkers.HasListTable = True
                outMarkers.ListTableParaIndex = i
                listMarkerCount = listMarkerCount + 1
                If listMarkerCount > listMarkerCap Then
                    listMarkerCap = listMarkerCap * 2
                    ReDim Preserve listMarkerNames(1 To listMarkerCap)
                    ReDim Preserve listMarkerIndices(1 To listMarkerCap)
                End If
                listMarkerNames(listMarkerCount) = "=listTableHere()"
                listMarkerIndices(listMarkerCount) = i

            Case "=listfigurehere()"
                outMarkers.HasListFigure = True
                outMarkers.ListFigureParaIndex = i
                listMarkerCount = listMarkerCount + 1
                If listMarkerCount > listMarkerCap Then
                    listMarkerCap = listMarkerCap * 2
                    ReDim Preserve listMarkerNames(1 To listMarkerCap)
                    ReDim Preserve listMarkerIndices(1 To listMarkerCap)
                End If
                listMarkerNames(listMarkerCount) = "=listFigureHere()"
                listMarkerIndices(listMarkerCount) = i
        End Select
    Next p
    
    If Not outMarkers.HasContentStart Then
        outErrorMessage = ModConstants.Uni("Missing =contentStart(). Please add it on its own line. (Thi\u1ebfu th\u1ebb =contentStart(), h\u00e3y th\u00eam tr\u00ean m\u1ed9t d\u00f2ng ri\u00eang.)")
        ScanMarkers = False
        Exit Function
    End If
    
    If Not outMarkers.HasContentEnd Then
        outErrorMessage = ModConstants.Uni("Missing =contentEnd(). Please add it on its own line. (Thi\u1ebfu th\u1ebb =contentEnd(), h\u00e3y th\u00eam tr\u00ean m\u1ed9t d\u00f2ng ri\u00eang.)")
        ScanMarkers = False
        Exit Function
    End If
    
    If outMarkers.ContentStartParaIndex >= outMarkers.ContentEndParaIndex Then
        outErrorMessage = ModConstants.Uni("=contentStart() (para ") & outMarkers.ContentStartParaIndex & _
                          ModConstants.Uni(") must be before =contentEnd() (para ") & outMarkers.ContentEndParaIndex & _
                          ModConstants.Uni("). (=contentStart() ph\u1ea3i n\u1eb1m tr\u01b0\u1edbc =contentEnd().)")
        ScanMarkers = False
        Exit Function
    End If
    
    Dim k As Long, badMarkers As String
    badMarkers = ""
    For k = 1 To listMarkerCount
        If listMarkerIndices(k) > outMarkers.ContentStartParaIndex And _
           listMarkerIndices(k) < outMarkers.ContentEndParaIndex Then
            If Len(badMarkers) > 0 Then
                badMarkers = badMarkers & ", "
            End If
            badMarkers = badMarkers & listMarkerNames(k) & " (line " & listMarkerIndices(k) & ")"
        End If
    Next k

    If Len(badMarkers) > 0 Then
        outErrorMessage = BuildMarkerPlacementWarning(badMarkers)
        ScanMarkers = False
        Exit Function
    End If

    ScanMarkers = True
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
    Dim p As Paragraph, cleanText As String, pRng As Range
    For Each p In doc.Paragraphs
        Set pRng = p.Range
        cleanText = LCase$(Trim$(Replace(Replace(Replace(pRng.Text, vbCr, ""), Chr(7), ""), Chr(11), "")))
        Select Case cleanText
            Case "=contentstart()", "=contentend()", "=listcontent()", "=listtablehere()", "=listfigurehere()"
                pRng.HighlightColorIndex = 7 ' wdYellow = 7
        End Select
    Next p
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function FindMarkerParagraph(ByVal doc As Document, ByVal markerText As String) As Long
    Dim i As Long, p As Paragraph, cleanText As String
    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        cleanText = LCase$(Trim$(Replace(Replace(Replace(p.Range.Text, vbCr, ""), Chr(7), ""), Chr(11), "")))
        If cleanText = markerText Then
            FindMarkerParagraph = i
            Exit Function
        End If
    Next p
    FindMarkerParagraph = 0
End Function

Public Sub NormalizeContentBreaks(ByVal doc As Document)
    Dim pStart As Long, pEnd As Long
    pStart = FindMarkerParagraph(doc, "=contentstart()")
    pEnd = FindMarkerParagraph(doc, "=contentend()")
    If pStart = 0 Or pEnd = 0 Or pStart >= pEnd Then Exit Sub
    
    Dim i As Long, pRng As Range, rawText As String, firstLine As String
    Dim p As Paragraph
    Dim cRng As Range
    Set cRng = doc.Range(doc.Paragraphs(pStart).Range.End, doc.Paragraphs(pEnd).Range.Start)
    
    ' Phase 1: Handle Chr(11) (Shift + Enter) via bottom-up coordinate execution (F-10)
    If InStr(cRng.Text, Chr(11)) > 0 Then
        Dim splitStarts() As Long, splitEnds() As Long, splitCount As Long
        splitCount = 0
        ReDim splitStarts(1 To doc.Paragraphs.Count)
        ReDim splitEnds(1 To doc.Paragraphs.Count)
        For Each p In cRng.Paragraphs
            rawText = p.Range.Text
            If InStr(rawText, Chr(11)) > 0 Then
                firstLine = Trim$(Split(Replace(rawText, vbCr, ""), Chr(11))(0))
                If Not ModLocale.IsPrefixOnlyHeading(firstLine) Then
                    splitCount = splitCount + 1
                    splitStarts(splitCount) = p.Range.Start
                    splitEnds(splitCount) = p.Range.End
                End If
            End If
        Next p
        Dim j As Long
        For j = splitCount To 1 Step -1
            Set pRng = doc.Range(splitStarts(j), splitEnds(j))
            pRng.Find.ClearFormatting
            pRng.Find.Replacement.ClearFormatting
            pRng.Find.Execute FindText:="^l", ReplaceWith:="^p", Replace:=2, Wrap:=0
        Next j
    End If
    
    ' Phase 2: Safely remove blank paragraphs from bottom up (F-03, R-06)
    pStart = FindMarkerParagraph(doc, "=contentstart()")
    pEnd = FindMarkerParagraph(doc, "=contentend()")
    If pStart = 0 Or pEnd = 0 Or pStart >= pEnd Then Exit Sub
    
    Dim blankStarts() As Long, blankCount As Long
    blankCount = 0
    ReDim blankStarts(1 To doc.Paragraphs.Count)
    
    Dim hasTables As Boolean
    hasTables = (doc.Tables.Count > 0)
    If hasTables Then
        ModTableIndex.BuildTableIndex doc
    End If
    Dim hasShapes As Boolean
    hasShapes = (doc.Shapes.Count > 0)
    Dim hasInline As Boolean
    hasInline = (doc.InlineShapes.Count > 0)
    Dim r As Range, prevInTbl As Boolean
    prevInTbl = False
    Dim curInTbl As Boolean, nextInTbl As Boolean, isBlank As Boolean
    
    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        Set r = p.Range
        curInTbl = False
        If hasTables Then curInTbl = ModTableIndex.IsRangeInTable(r)
        If i > pStart And i < pEnd Then
            If r.Text = vbCr And (Not curInTbl) Then
                isBlank = True
                If hasInline Then If r.InlineShapes.Count > 0 Then isBlank = False
                If hasShapes And isBlank Then If r.ShapeRange.Count > 0 Then isBlank = False
                If isBlank Then
                    nextInTbl = False
                    If hasTables And (Not p.Next Is Nothing) Then nextInTbl = ModTableIndex.IsRangeInTable(p.Next.Range)
                    If (Not prevInTbl) And (Not nextInTbl) Then
                        blankCount = blankCount + 1
                        blankStarts(blankCount) = r.Start
                    End If
                End If
            End If
        End If
        prevInTbl = curInTbl
    Next p
    
    Dim k As Long
    For k = blankCount To 1 Step -1
        doc.Range(blankStarts(k), blankStarts(k) + 1).Delete
    Next k
End Sub
