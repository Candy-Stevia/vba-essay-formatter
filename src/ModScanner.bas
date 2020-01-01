Attribute VB_Name = "ModScanner"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

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

Public Type QuickRegexPool
    ReL1 As Object
    ReL2 As Object
    ReL3 As Object
    ReL4 As Object
    ReTbl As Object
    ReFig As Object
    ReSrc As Object
    ReMrk As Object
    ReManToc As Object
    ReManLot As Object
    ReManLof As Object
End Type

Private Const MAX_HEADING_LEN As Long = 350

Public Sub InitQuickRegexPool(ByRef pool As QuickRegexPool)
    Set pool.ReL1 = CreateRegex(ModLocale.GetPatternLevel1())
    Set pool.ReL2 = CreateRegex(ModConstants.GetPatternLevel2())
    Set pool.ReL3 = CreateRegex(ModConstants.GetPatternLevel3())
    Set pool.ReL4 = CreateRegex(ModConstants.GetPatternLevel4())
    Set pool.ReTbl = CreateRegex(ModLocale.GetPatternTable())
    Set pool.ReFig = CreateRegex(ModLocale.GetPatternFigure())
    Set pool.ReSrc = CreateRegex(ModLocale.GetPatternSourceNote())
    Set pool.ReMrk = CreateRegex(ModConstants.GetPatternMarker())
    Set pool.ReManToc = CreateRegex(ModLocale.GetPatternManualToc())
    Set pool.ReManLot = CreateRegex(ModLocale.GetPatternManualLot())
    Set pool.ReManLof = CreateRegex(ModLocale.GetPatternManualLof())
End Sub

Public Sub CollectOldListRanges(ByVal doc As Document, ByRef exclStarts() As Long, ByRef exclEnds() As Long, ByRef outNum As Long)
    outNum = doc.TablesOfContents.Count + doc.TablesOfFigures.Count
    If outNum <= 0 Then Exit Sub
    ReDim exclStarts(1 To outNum)
    ReDim exclEnds(1 To outNum)
    Dim tIdx As Long, tocObj As TableOfContents, tofObj As TableOfFigures
    tIdx = 0
    For Each tocObj In doc.TablesOfContents
        tIdx = tIdx + 1
        exclStarts(tIdx) = tocObj.Range.Start
        exclEnds(tIdx) = tocObj.Range.End
    Next tocObj
    For Each tofObj In doc.TablesOfFigures
        tIdx = tIdx + 1
        exclStarts(tIdx) = tofObj.Range.Start
        exclEnds(tIdx) = tofObj.Range.End
    Next tofObj
End Sub

Private Function CreateRegex(ByVal pat As String, Optional ByVal ignoreCase As Boolean = True) As Object
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = ignoreCase
    re.Global = False
    re.Pattern = pat
    Set CreateRegex = re
End Function

Private Function IsMathBlockPara(ByVal rng As Range, ByVal raw As String) As Boolean
    If rng.OMaths.Count = 0 Then Exit Function
    Dim omIdx As Long, omChars As Long, paraNoSpace As String, paraTrim As String
    paraTrim = Trim$(Replace(Replace(raw, vbCr, ""), Chr(11), ""))
    omChars = 0
    For omIdx = 1 To rng.OMaths.Count
        If rng.OMaths(omIdx).Type = 0 Then
            IsMathBlockPara = True
            Exit Function
        End If
        omChars = omChars + Len(Replace(Trim$(rng.OMaths(omIdx).Range.Text), " ", ""))
    Next omIdx
    If Len(paraTrim) > 0 Then
        paraNoSpace = Replace(paraTrim, " ", "")
        If Len(paraNoSpace) = omChars Then
            IsMathBlockPara = True
        End If
    End If
End Function

Public Function IsLayoutFigureCaption(ByVal rng As Range, ByVal cleanText As String, ByVal regEx As Object, ByVal patFig As String) As Boolean
    If Len(cleanText) = 0 Or Len(cleanText) > MAX_HEADING_LEN Then Exit Function
    On Error Resume Next
    If regEx.Pattern <> patFig Then regEx.Pattern = patFig
    If regEx.Test(cleanText) Then
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
    Dim patFig As String
    patFig = ModLocale.GetPatternFigure()
    
    Dim reL1 As Object, reL2 As Object, reL3 As Object, reL4 As Object
    Dim reTbl As Object, reFig As Object, reSrc As Object, reMarker As Object
    Dim reToc As Object, reLot As Object, reLof As Object
    Set reL1 = CreateRegex(ModLocale.GetPatternLevel1())
    Set reToc = CreateRegex(ModLocale.GetPatternManualToc())
    Set reLot = CreateRegex(ModLocale.GetPatternManualLot())
    Set reLof = CreateRegex(ModLocale.GetPatternManualLof())
    Set reL2 = CreateRegex(ModConstants.GetPatternLevel2())
    Set reL3 = CreateRegex(ModConstants.GetPatternLevel3())
    Set reL4 = CreateRegex(ModConstants.GetPatternLevel4())
    Set reTbl = CreateRegex(ModLocale.GetPatternTable())
    Set reFig = CreateRegex(patFig)
    Set reSrc = CreateRegex(ModLocale.GetPatternSourceNote())
    Set reMarker = CreateRegex(ModConstants.GetPatternMarker())
    
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
    
    Dim docHasMath As Boolean, docHasInline As Boolean
    docHasMath = (doc.OMaths.Count > 0)
    docHasInline = (doc.InlineShapes.Count > 0)
    
    Dim numExclRanges As Long, exclStarts() As Long, exclEnds() As Long
    numExclRanges = doc.TablesOfContents.Count + doc.TablesOfFigures.Count
    If numExclRanges > 0 Then
        ReDim exclStarts(1 To numExclRanges)
        ReDim exclEnds(1 To numExclRanges)
        Dim tIdx As Long, tocObj As TableOfContents, tofObj As TableOfFigures
        tIdx = 0
        For Each tocObj In doc.TablesOfContents
            tIdx = tIdx + 1
            exclStarts(tIdx) = tocObj.Range.Start
            exclEnds(tIdx) = tocObj.Range.End
        Next tocObj
        For Each tofObj In doc.TablesOfFigures
            tIdx = tIdx + 1
            exclStarts(tIdx) = tofObj.Range.Start
            exclEnds(tIdx) = tofObj.Range.End
        Next tofObj
    End If
    
    Dim i As Long, p As Paragraph, rng As Range
    Dim raw As String, clean As String, firstLine As String, hasSoft As Boolean
    Dim inContentZone As Boolean, inPostZone As Boolean
    Dim cleanLen As Long, isShortHeading As Boolean
    
    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        ModProgress.ProgressUpdate i, totalParas
        Set rng = p.Range
        raw = rng.Text
        hasSoft = (InStr(raw, Chr(11)) > 0)
        clean = Replace(Replace(Replace(Replace(raw, vbCr, ""), Chr(7), ""), Chr(1), ""), Chr(12), "")
        clean = Trim$(Replace(Replace(Replace(clean, Chr(11), " "), Chr(10), " "), vbTab, " "))
        cleanLen = Len(clean)
        If hasSoft Then
            firstLine = Replace(Replace(Replace(raw, vbCr, ""), Chr(7), ""), Chr(12), "")
            firstLine = Trim$(Replace(Replace(Split(firstLine, Chr(11))(0), Chr(10), " "), vbTab, " "))
        Else
            firstLine = clean
        End If
        
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
        
        ' Old Word TOC / LOT / LOF field exclusion
        Dim isInsideOldToc As Boolean
        isInsideOldToc = False
        If numExclRanges > 0 Then
            For tIdx = 1 To numExclRanges
                If rng.Start >= exclStarts(tIdx) And rng.Start < exclEnds(tIdx) Then
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
        If latexFlags(i) And (Not isInsideOldToc) Then
            outSnapshots(i).IsExcluded = True
            outSnapshots(i).ElementType = ElementTypeLatexBlock
        End If
        
        ' Break-only Chr(12) exclusion (Page/Section Break without text)
        If (Not outSnapshots(i).IsExcluded) And (cleanLen = 0) Then
            If InStr(raw, Chr(12)) > 0 Then
                outSnapshots(i).IsExcluded = True
            End If
        End If
        
        ' Non-LaTeX Classification
        If Not outSnapshots(i).IsExcluded Then
            Dim isMathBlock As Boolean
            isMathBlock = False
            If docHasMath Then
                isMathBlock = IsMathBlockPara(rng, raw)
            End If
            
            isShortHeading = (cleanLen > 0 And cleanLen <= MAX_HEADING_LEN)
            
            If isMathBlock Then
                outSnapshots(i).IsExcluded = True
                outSnapshots(i).ElementType = ElementTypeWordMath
            ElseIf rng.ListFormat.ListType <> 0 Then
                outSnapshots(i).IsExcluded = True
                outSnapshots(i).ElementType = ElementTypeList
            ElseIf cleanLen > 0 And reMarker.Test(clean) Then
                outSnapshots(i).IsExcluded = True
            ElseIf ModTableIndex.IsRangeInTable(rng) Then
                Dim isLayFig As Boolean
                isLayFig = False
                If isShortHeading Then
                    isLayFig = IsLayoutFigureCaption(rng, clean, reFig, patFig)
                End If
                If isLayFig Then
                    If inContentZone Or inPostZone Then
                        outSnapshots(i).ElementType = ElementTypeFigureTitle
                    End If
                ElseIf inContentZone Or inPostZone Then
                    outSnapshots(i).ElementType = ElementTypeTableCell
                End If
            ElseIf isShortHeading And reToc.Test(clean) Then
                If markers.HasListContent Then
                    outSnapshots(i).IsExcluded = True
                    outSnapshots(i).IsManualListWarning = True
                    outSnapshots(i).ElementType = ElementTypeIgnoredManualList
                Else
                    outSnapshots(i).ElementType = ElementTypeLevel1
                End If
            ElseIf isShortHeading And reLot.Test(clean) Then
                If markers.HasListTable Then
                    outSnapshots(i).IsExcluded = True
                    outSnapshots(i).IsManualListWarning = True
                    outSnapshots(i).ElementType = ElementTypeIgnoredManualList
                Else
                    outSnapshots(i).ElementType = ElementTypeLevel1
                End If
            ElseIf isShortHeading And reLof.Test(clean) Then
                If markers.HasListFigure Then
                    outSnapshots(i).IsExcluded = True
                    outSnapshots(i).IsManualListWarning = True
                    outSnapshots(i).ElementType = ElementTypeIgnoredManualList
                Else
                    outSnapshots(i).ElementType = ElementTypeLevel1
                End If
            ElseIf isShortHeading And reL1.Test(clean) Then
                outSnapshots(i).ElementType = ElementTypeLevel1
            ElseIf inContentZone And isShortHeading And ModLocale.IsRomanLevel1(clean) Then
                outSnapshots(i).ElementType = ElementTypeLevel1
            ElseIf inContentZone Or inPostZone Then
                Dim matchedCap As Boolean
                matchedCap = False
                If isShortHeading Then
                    If reTbl.Test(clean) Then
                        outSnapshots(i).ElementType = ElementTypeTableTitle
                        matchedCap = True
                    ElseIf reFig.Test(clean) Then
                        outSnapshots(i).ElementType = ElementTypeFigureTitle
                        matchedCap = True
                    End If
                End If
                If Not matchedCap Then
                    If cleanLen > 0 And reSrc.Test(clean) Then
                        outSnapshots(i).ElementType = ElementTypeSourceNote
                    ElseIf cleanLen = 0 And docHasInline Then
                        If rng.InlineShapes.Count > 0 Then
                            outSnapshots(i).ElementType = ElementTypeInlinePicture
                        ElseIf inContentZone Then
                            outSnapshots(i).ElementType = ElementTypeNormalParagraph
                        End If
                    ElseIf inContentZone Then
                        Dim matchedSub As Boolean
                        matchedSub = False
                        If isShortHeading Then
                            If reL2.Test(clean) Then
                                outSnapshots(i).ElementType = ElementTypeLevel2
                                matchedSub = True
                            ElseIf reL3.Test(clean) Then
                                outSnapshots(i).ElementType = ElementTypeLevel3
                                matchedSub = True
                            ElseIf reL4.Test(clean) Then
                                outSnapshots(i).ElementType = ElementTypeLevel4
                                matchedSub = True
                            End If
                        End If
                        If Not matchedSub Then
                            outSnapshots(i).ElementType = ElementTypeNormalParagraph
                        End If
                    End If
                End If
            End If
        End If
    Next p
    
    ScanDocumentSnapshot = True
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
