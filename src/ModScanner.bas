Attribute VB_Name = "ModScanner"
Option Explicit

' ==============================================================================
' Module: ModScanner
' Purpose: Phase 1 Read-only snapshot scanning & element classification
' Architecture: Academic Document Format Engine
' ==============================================================================

Public Type ParagraphSnapshot
    ParaIndex As Long
    RangeStart As Long
    RangeEnd As Long
    RawText As String
    CleanText As String
    ElementType As Long ' TargetElementType
    IsExcluded As Boolean
    IsManualListWarning As Boolean
    HasSoftBreak As Boolean
    FirstLineText As String
End Type

Private Function Matches(ByVal re As Object, ByVal pat As String, ByVal s As String) As Boolean
    re.Pattern = pat
    Matches = re.Test(s)
End Function

Public Function IsLayoutFigureCaption(ByVal rng As Range, ByVal cleanText As String, ByVal regEx As Object, ByVal patFig As String) As Boolean
    On Error Resume Next
    If Matches(regEx, patFig, cleanText) Then
        If rng.Tables.Count > 0 Then
            If rng.Tables(1).Range.InlineShapes.Count > 0 Then
                IsLayoutFigureCaption = True
            End If
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Function

' --- Phase 1: Read-Only Document Snapshot Scanner ---

Public Function ScanDocumentSnapshot(ByVal doc As Document, _
                                     ByRef markers As MarkerLocations, _
                                     ByRef outSnapshots() As ParagraphSnapshot, _
                                     ByRef outCount As Long, _
                                     ByRef outLatexWarningCount As Long) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False
    
    Dim patL1 As String, patL2 As String, patL3 As String, patL4 As String
    Dim patTbl As String, patFig As String, patSrc As String, patMarker As String
    Dim patToc As String, patLot As String, patLof As String
    patL1 = ModLocale.GetPatternLevel1()
    patToc = ModLocale.GetPatternManualToc()
    patLot = ModLocale.GetPatternManualLot()
    patLof = ModLocale.GetPatternManualLof()
    patL2 = ModConstants.GetPatternLevel2()
    patL3 = ModConstants.GetPatternLevel3()
    patL4 = ModConstants.GetPatternLevel4()
    patTbl = ModLocale.GetPatternTable()
    patFig = ModLocale.GetPatternFigure()
    patSrc = ModLocale.GetPatternSourceNote()
    patMarker = ModConstants.GetPatternMarker()
    
    Dim totalParas As Long
    totalParas = doc.Paragraphs.Count
    outCount = totalParas
    outLatexWarningCount = 0
    If totalParas = 0 Then
        ScanDocumentSnapshot = True
        Exit Function
    End If
    ReDim outSnapshots(1 To totalParas)
    ModTableIndex.BuildTableIndex doc
    
    Dim latexFlags() As Boolean
    ModLatex.ComputeLatexFlags doc, latexFlags, outLatexWarningCount
    
    Dim docHasMath As Boolean
    docHasMath = (doc.OMaths.Count > 0)
    
    Dim numOldTocs As Long
    numOldTocs = doc.TablesOfContents.Count
    Dim oldTocStarts() As Long, oldTocEnds() As Long
    If numOldTocs > 0 Then
        ReDim oldTocStarts(1 To numOldTocs)
        ReDim oldTocEnds(1 To numOldTocs)
        Dim tIdx As Long
        For tIdx = 1 To numOldTocs
            oldTocStarts(tIdx) = doc.TablesOfContents(tIdx).Range.Start
            oldTocEnds(tIdx) = doc.TablesOfContents(tIdx).Range.End
        Next tIdx
    End If
    
    Dim i As Long, p As Paragraph, rng As Range
    Dim raw As String, clean As String, firstLine As String, hasSoft As Boolean
    Dim inContentZone As Boolean, inPostZone As Boolean
    
    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        ModProgress.ProgressUpdate i, totalParas
        Set rng = p.Range
        raw = rng.Text
        hasSoft = (InStr(raw, Chr(11)) > 0)
        clean = Trim$(Replace(Replace(Replace(raw, vbCr, ""), Chr(7), ""), Chr(11), " "))
        If hasSoft Then firstLine = Trim$(Split(Replace(raw, vbCr, ""), Chr(11))(0)) Else firstLine = clean
        
        outSnapshots(i).ParaIndex = i
        outSnapshots(i).RangeStart = rng.Start
        outSnapshots(i).RangeEnd = rng.End
        outSnapshots(i).RawText = raw
        outSnapshots(i).CleanText = clean
        outSnapshots(i).HasSoftBreak = hasSoft
        outSnapshots(i).FirstLineText = firstLine
        outSnapshots(i).ElementType = ElementTypeUnknown
        outSnapshots(i).IsExcluded = False
        outSnapshots(i).IsManualListWarning = False
        
        inContentZone = (markers.HasContentStart And markers.HasContentEnd And _
                         i > markers.ContentStartParaIndex And i < markers.ContentEndParaIndex)
        inPostZone = (markers.HasContentEnd And i > markers.ContentEndParaIndex)
        
        ' Old Word TOC field exclusion
        Dim isInsideOldToc As Boolean
        isInsideOldToc = False
        If numOldTocs > 0 Then
            For tIdx = 1 To numOldTocs
                If rng.Start >= oldTocStarts(tIdx) And rng.Start < oldTocEnds(tIdx) Then
                    isInsideOldToc = True
                    Exit For
                End If
            Next tIdx
        End If
        If isInsideOldToc Then
            outSnapshots(i).IsExcluded = True
            outSnapshots(i).ElementType = ElementTypeOldToc
        End If
        
        ' LaTeX exclusion from ModLatex
        If latexFlags(i) And Not isInsideOldToc Then
            outSnapshots(i).IsExcluded = True
            outSnapshots(i).ElementType = ElementTypeLatexBlock
        End If
        
        ' Non-LaTeX Classification
        If Not outSnapshots(i).IsExcluded Then
            Dim isMathBlock As Boolean
            isMathBlock = False
            If docHasMath Then
                If rng.OMaths.Count > 0 Then
                    If rng.OMaths(1).Type = 0 Or rng.OMaths(1).Range.Text = rng.Text Then
                        isMathBlock = True
                    End If
                End If
            End If
            If isMathBlock Then
                outSnapshots(i).IsExcluded = True
                outSnapshots(i).ElementType = ElementTypeWordMath
            ElseIf rng.ListFormat.ListType <> 0 Then
                outSnapshots(i).IsExcluded = True
                outSnapshots(i).ElementType = ElementTypeList
            ElseIf Matches(regEx, patMarker, clean) Then
                outSnapshots(i).IsExcluded = True
            ElseIf ModTableIndex.IsRangeInTable(rng) = True And Not IsLayoutFigureCaption(rng, clean, regEx, patFig) Then
                If inContentZone Or inPostZone Then
                    outSnapshots(i).ElementType = ElementTypeTableCell
                End If
            ElseIf Matches(regEx, patToc, clean) Then
                If markers.HasListContent Then
                    outSnapshots(i).IsExcluded = True
                    outSnapshots(i).IsManualListWarning = True
                    outSnapshots(i).ElementType = ElementTypeIgnoredManualList
                Else
                    outSnapshots(i).ElementType = ElementTypeLevel1
                End If
            ElseIf Matches(regEx, patLot, clean) Then
                If markers.HasListTable Then
                    outSnapshots(i).IsExcluded = True
                    outSnapshots(i).IsManualListWarning = True
                    outSnapshots(i).ElementType = ElementTypeIgnoredManualList
                Else
                    outSnapshots(i).ElementType = ElementTypeLevel1
                End If
            ElseIf Matches(regEx, patLof, clean) Then
                If markers.HasListFigure Then
                    outSnapshots(i).IsExcluded = True
                    outSnapshots(i).IsManualListWarning = True
                    outSnapshots(i).ElementType = ElementTypeIgnoredManualList
                Else
                    outSnapshots(i).ElementType = ElementTypeLevel1
                End If
            ElseIf Matches(regEx, patL1, clean) Then
                outSnapshots(i).ElementType = ElementTypeLevel1
            ElseIf inContentZone Or inPostZone Then
                If Matches(regEx, patTbl, clean) Then
                    outSnapshots(i).ElementType = ElementTypeTableTitle
                ElseIf Matches(regEx, patFig, clean) Then
                    outSnapshots(i).ElementType = ElementTypeFigureTitle
                ElseIf Matches(regEx, patSrc, clean) Then
                    outSnapshots(i).ElementType = ElementTypeSourceNote
                ElseIf rng.InlineShapes.Count > 0 Then
                    outSnapshots(i).ElementType = ElementTypeInlinePicture
                ElseIf inContentZone Then
                    If Matches(regEx, patL2, clean) Then
                        outSnapshots(i).ElementType = ElementTypeLevel2
                    ElseIf Matches(regEx, patL3, clean) Then
                        outSnapshots(i).ElementType = ElementTypeLevel3
                    ElseIf Matches(regEx, patL4, clean) Then
                        outSnapshots(i).ElementType = ElementTypeLevel4
                    Else
                        outSnapshots(i).ElementType = ElementTypeNormalParagraph
                    End If
                End If
            End If
        End If
    Next p
    
    ScanDocumentSnapshot = True
End Function
