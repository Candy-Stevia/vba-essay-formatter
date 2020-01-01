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
    Static cached As QuickRegexPool
    Static isInit As Boolean
    If Not isInit Then
        Set cached.ReL1 = CreateRegex(ModLocale.GetPatternLevel1())
        Set cached.ReL2 = CreateRegex(ModConstants.GetPatternLevel2())
        Set cached.ReL3 = CreateRegex(ModConstants.GetPatternLevel3())
        Set cached.ReL4 = CreateRegex(ModConstants.GetPatternLevel4())
        Set cached.ReTbl = CreateRegex(ModLocale.GetPatternTable())
        Set cached.ReFig = CreateRegex(ModLocale.GetPatternFigure())
        Set cached.ReSrc = CreateRegex(ModLocale.GetPatternSourceNote())
        Set cached.ReMrk = CreateRegex(ModConstants.GetPatternMarker())
        Set cached.ReManToc = CreateRegex(ModLocale.GetPatternManualToc())
        Set cached.ReManLot = CreateRegex(ModLocale.GetPatternManualLot())
        Set cached.ReManLof = CreateRegex(ModLocale.GetPatternManualLof())
        isInit = True
    End If
    pool = cached
End Sub

Private Function StripRareControls(ByVal s As String) As String
    If InStr(1, s, Chr(7)) > 0 Then s = Replace(s, Chr(7), "")
    If InStr(1, s, Chr(1)) > 0 Then s = Replace(s, Chr(1), "")
    If InStr(1, s, Chr(12)) > 0 Then s = Replace(s, Chr(12), "")
    If InStr(1, s, Chr(2)) > 0 Then s = Replace(s, Chr(2), "")
    If InStr(1, s, Chr(5)) > 0 Then s = Replace(s, Chr(5), "")
    If InStr(1, s, Chr(10)) > 0 Then s = Replace(s, Chr(10), " ")
    If InStr(1, s, vbTab) > 0 Then s = Replace(s, vbTab, " ")
    StripRareControls = s
End Function

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
            Dim tbl As Table
            Set tbl = rng.Tables(1)
            If tbl.Rows.Count <= 3 And tbl.Range.Cells.Count <= 4 Then
                If tbl.Range.InlineShapes.Count > 0 Then
                    IsLayoutFigureCaption = True
                End If
            End If
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Function

Private Function IsParaInListSafe(ByVal rng As Range, ByVal rStart As Long, ByVal listIdxValid As Boolean) As Boolean
    If listIdxValid Then
        IsParaInListSafe = ModTableIndex.IsPosInList(rStart)
    Else
        IsParaInListSafe = (rng.ListFormat.ListType <> 0)
    End If
End Function

' --- Phase 1: Read-Only Document Snapshot Scanner ---

Public Function ScanDocumentSnapshot(ByVal doc As Document, _
                                     ByRef markers As MarkerLocations, _
                                     ByRef outSnapshots() As ParagraphSnapshot, _
                                     ByRef outCount As Long, _
                                     ByRef outLatexWarningCount As Long) As Boolean
    Dim patFig As String
    patFig = ModLocale.GetPatternFigure()
    
    Dim pool As QuickRegexPool
    InitQuickRegexPool pool
    Dim reL1 As Object, reL2 As Object, reL3 As Object, reL4 As Object
    Dim reTbl As Object, reFig As Object, reSrc As Object, reMarker As Object
    Dim reToc As Object, reLot As Object, reLof As Object
    Set reL1 = pool.ReL1
    Set reToc = pool.ReManToc
    Set reLot = pool.ReManLot
    Set reLof = pool.ReManLof
    Set reL2 = pool.ReL2
    Set reL3 = pool.ReL3
    Set reL4 = pool.ReL4
    Set reTbl = pool.ReTbl
    Set reFig = pool.ReFig
    Set reSrc = pool.ReSrc
    Set reMarker = pool.ReMrk
    
    Dim totalParas As Long
    totalParas = doc.Paragraphs.Count
    outCount = totalParas
    outLatexWarningCount = 0
    If totalParas = 0 Then
        ScanDocumentSnapshot = True
        Exit Function
    End If
    ReDim outSnapshots(1 To totalParas)
    ModTableIndex.EnsureTableIndex doc
    ModTableIndex.EnsureListIndex doc
    ModTableIndex.EnsureOMathIndex doc
    Dim tblIdxValid As Boolean, listIdxValid As Boolean
    tblIdxValid = ModTableIndex.IsTableIndexValidForDoc(doc)
    listIdxValid = ModTableIndex.IsListIndexValidForDoc(doc)
    ModCustomProfile.LoadTagCache doc
    Dim hasTags As Boolean
    hasTags = ModCustomProfile.HasActiveTags()
    
    Dim latexFlags() As Boolean
    ModLatex.ComputeLatexFlags doc, latexFlags, outLatexWarningCount
    
    Dim docHasMath As Boolean, docHasInline As Boolean, docHasLists As Boolean
    docHasMath = (doc.OMaths.Count > 0)
    docHasInline = (doc.InlineShapes.Count > 0)
    docHasLists = (doc.ListParagraphs.Count > 0)
    
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
    Dim pOutLvl As Long, pListType As Long, listPfx As String
    
    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        ModProgress.ProgressUpdate i, totalParas
        If ModProgress.g_CancelRequested Then Exit Function
        Set rng = p.Range
        raw = rng.Text
        hasSoft = (InStr(1, raw, Chr(11)) > 0)
        clean = StripRareControls(Replace(raw, vbCr, ""))
        If hasSoft Then
            firstLine = Trim$(Split(clean, Chr(11))(0))
            clean = Trim$(Replace(clean, Chr(11), " "))
        Else
            clean = Trim$(clean)
            firstLine = clean
        End If
        cleanLen = Len(clean)
        If docHasLists And cleanLen > 0 And cleanLen <= MAX_HEADING_LEN Then
            pOutLvl = 10
            On Error Resume Next
            pOutLvl = p.OutlineLevel
            If Err.Number = 0 And pOutLvl <= 4 Then
                pListType = rng.ListFormat.ListType
                If Err.Number = 0 And pListType >= 3 Then
                    listPfx = Trim$(rng.ListFormat.ListString)
                    If Err.Number = 0 And Len(listPfx) > 0 Then
                        If Left$(clean, Len(listPfx)) <> listPfx Then
                            clean = listPfx & " " & clean
                            firstLine = listPfx & " " & firstLine
                            cleanLen = Len(clean)
                        End If
                    End If
                End If
            End If
            Err.Clear
            On Error GoTo 0
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
                If ModTableIndex.HasOMathInSpan(outSnapshots(i).RangeStart, outSnapshots(i).RangeEnd) Then
                    isMathBlock = IsMathBlockPara(rng, raw)
                End If
            End If
            
            isShortHeading = (cleanLen > 0 And cleanLen <= MAX_HEADING_LEN)
            
            Dim tagCode As Long, inTblPos As Boolean
            tagCode = 0
            If hasTags Then tagCode = ModCustomProfile.LookupTagForRange(outSnapshots(i).RangeStart, outSnapshots(i).RangeEnd)
            If tblIdxValid Then
                inTblPos = ModTableIndex.IsPosInTable(outSnapshots(i).RangeStart)
            Else
                inTblPos = CBool(rng.Information(12))
            End If
            If isMathBlock Then
                outSnapshots(i).IsExcluded = True
                outSnapshots(i).ElementType = ElementTypeWordMath
            ElseIf cleanLen > 0 And reMarker.Test(clean) Then
                outSnapshots(i).IsExcluded = True
            ElseIf tagCode > 0 Then
                outSnapshots(i).ElementType = tagCode
            ElseIf tagCode = TAG_CODE_REF Or tagCode = TAG_CODE_APP Then
                outSnapshots(i).ElementType = ElementTypeLevel1
            ElseIf tagCode = TAG_CODE_REF_SUB Then
                outSnapshots(i).ElementType = ElementTypeLevel2
            ElseIf tagCode = TAG_CODE_SKIP Then
                If inTblPos Then
                    If inContentZone Or inPostZone Then outSnapshots(i).ElementType = ElementTypeTableCell
                ElseIf inContentZone Then
                    outSnapshots(i).ElementType = ElementTypeNormalParagraph
                Else
                    outSnapshots(i).ElementType = ElementTypeUnknown
                End If
            ElseIf IsParaInListSafe(rng, outSnapshots(i).RangeStart, listIdxValid) Then
                outSnapshots(i).IsExcluded = True
                outSnapshots(i).ElementType = ElementTypeList
            ElseIf inTblPos Then
                Dim isLayFig As Boolean
                isLayFig = False
                If isShortHeading And docHasInline Then
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
