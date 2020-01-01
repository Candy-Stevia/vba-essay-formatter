Attribute VB_Name = "ModPageNumberSync"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModPageNumberSync
' Purpose: Roman numeral conversion, dynamic tab leader calculation, page sync
' Architecture: Academic Document Format Engine (Multi-Zone Batch Sync & Bounded Convergence)
' ==============================================================================

Private mSecCacheActive As Boolean
Private mSecStyles() As Long
Private mSecHasStyle() As Boolean
Private mSecCap As Long

Public Function ToRomanLowerCase(ByVal num As Long) As String
    If num <= 0 Or num > 3999 Then
        ToRomanLowerCase = CStr(num)
        Exit Function
    End If
    Dim aVals As Variant, rSyms As Variant, res As String, i As Long, cNum As Long
    aVals = Array(1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1)
    rSyms = Array("m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i")
    cNum = num
    res = ""
    For i = LBound(aVals) To UBound(aVals)
        Do While cNum >= CLng(aVals(i))
            res = res & CStr(rSyms(i))
            cNum = cNum - CLng(aVals(i))
        Loop
    Next i
    ToRomanLowerCase = res
End Function

Private Function IntToAlphaPage(ByVal n As Long, ByVal isUpper As Boolean) As String
    If n <= 0 Then
        IntToAlphaPage = CStr(n)
        Exit Function
    End If
    Dim baseCode As Long, cNum As Long, remVal As Long, res As String
    If isUpper Then
        baseCode = 65
    Else
        baseCode = 97
    End If
    cNum = n
    res = ""
    Do While cNum > 0
        remVal = (cNum - 1) Mod 26
        res = Chr$(baseCode + remVal) & res
        cNum = (cNum - 1) \ 26
    Loop
    IntToAlphaPage = res
End Function

Public Sub ResetSectionStyleCache()
    mSecCacheActive = False
    mSecCap = 0
    Erase mSecStyles
    Erase mSecHasStyle
End Sub

Private Function QuerySectionNumberStyle(ByVal r As Range) As Long
    Dim secNum As Long, secStyle As Long, idx As Long
    On Error Resume Next
    secNum = r.Information(2) ' wdActiveEndSectionNumber = 2
    If Err.Number <> 0 Then
        Err.Clear
        secNum = 0
    End If
    If mSecCacheActive And secNum >= 1 And secNum <= 256 Then
        If secNum <= mSecCap Then
            If mSecHasStyle(secNum) Then
                QuerySectionNumberStyle = mSecStyles(secNum)
                Exit Function
            End If
        End If
    End If
    secStyle = r.Sections(1).Footers(1).PageNumbers.NumberStyle
    If Err.Number <> 0 Then
        Err.Clear
        secStyle = r.Sections(1).Headers(1).PageNumbers.NumberStyle
    End If
    If Err.Number <> 0 Then
        secStyle = 0
        Err.Clear
    End If
    If mSecCacheActive And secNum >= 1 And secNum <= 256 Then
        If secNum > mSecCap Then
            mSecCap = secNum + 16
            ReDim Preserve mSecStyles(1 To mSecCap)
            ReDim Preserve mSecHasStyle(1 To mSecCap)
        End If
        mSecStyles(secNum) = secStyle
        mSecHasStyle(secNum) = True
    End If
    QuerySectionNumberStyle = secStyle
End Function

Public Function GetFormattedPageNumber(ByVal targetRng As Range, ByVal pageMode As Long) As String
    On Error Resume Next
    Dim r As Range
    Set r = targetRng.Duplicate
    r.Collapse 1 ' wdCollapseStart = 1
    
    Dim pNum As Long
    pNum = r.Information(pageMode)
    If pNum <= 0 Then pNum = r.Information(3) ' Fallback wdActiveEndPageNumber
    If pNum <= 0 Then
        ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
        pNum = 1
    End If
    
    If pageMode = PageSyncPhysicalPage Then
        GetFormattedPageNumber = CStr(pNum)
        Exit Function
    End If
    
    Dim secNumberStyle As Long
    secNumberStyle = QuerySectionNumberStyle(r)
    
    Select Case secNumberStyle
        Case 1 ' wdPageNumberStyleUppercaseRoman = 1
            GetFormattedPageNumber = UCase$(ToRomanLowerCase(pNum))
        Case 2 ' wdPageNumberStyleLowercaseRoman = 2
            GetFormattedPageNumber = ToRomanLowerCase(pNum)
        Case 3 ' wdPageNumberStyleUppercaseLetter = 3
            GetFormattedPageNumber = IntToAlphaPage(pNum, True)
        Case 4 ' wdPageNumberStyleLowercaseLetter = 4
            GetFormattedPageNumber = IntToAlphaPage(pNum, False)
        Case Else
            GetFormattedPageNumber = CStr(pNum)
    End Select
    If Err.Number <> 0 Then
        ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
        GetFormattedPageNumber = CStr(pNum)
        Err.Clear
    End If
End Function

Public Function CalculateRightTab(ByVal doc As Document, Optional ByVal targetRng As Range = Nothing) As Single
    On Error GoTo FallbackTab
    Dim ps As PageSetup
    If Not targetRng Is Nothing Then
        If targetRng.Sections.Count >= 1 Then Set ps = targetRng.Sections(1).PageSetup
    End If
    If ps Is Nothing Then
        If doc.Sections.Count >= 1 Then Set ps = doc.Sections(1).PageSetup Else Set ps = doc.PageSetup
    End If
    Dim w As Single, colCnt As Long, colW As Single
    colCnt = 1
    colW = 0#
    On Error Resume Next
    colCnt = ps.TextColumns.Count
    If Err.Number = 0 And colCnt > 1 Then
        colW = ps.TextColumns(1).Width
        If Err.Number <> 0 Then colW = 0#
    End If
    Err.Clear
    On Error GoTo FallbackTab
    If colCnt > 1 And colW > 10# And colW < 2000# Then
        CalculateRightTab = colW
        Exit Function
    End If
    w = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    If ps.Gutter > 0 And ps.GutterPos <> 1 Then
        w = w - ps.Gutter
    End If
    If w > 10# And w < 2000# Then
        CalculateRightTab = w
        Exit Function
    End If
FallbackTab:
    On Error Resume Next
    Dim ps1 As PageSetup, w2 As Single
    Set ps1 = doc.Sections(1).PageSetup
    w2 = ps1.PageWidth - ps1.LeftMargin - ps1.RightMargin
    If Err.Number = 0 And ps1.Gutter > 0 And ps1.GutterPos <> 1 Then
        w2 = w2 - ps1.Gutter
    End If
    If Err.Number = 0 And w2 > 10# And w2 < 2000# Then
        CalculateRightTab = w2
    Else
        Err.Clear
        CalculateRightTab = 451.3!
    End If
End Function

Public Sub ApplyListLineFormat(ByVal rng As Range, _
                               ByVal fontName As String, _
                               ByVal fontSize As Single, _
                               ByVal tabPos As Single, _
                               ByVal leftIndentPts As Single, _
                               ByVal isBold As Boolean)
    ModFormatSingle.EnsureOutlineLevelDetachHeading rng.Document, rng, 10
    rng.HighlightColorIndex = 0 ' wdNoHighlight = 0
    With rng.Font
        .Name = fontName
        .Size = fontSize
        .Bold = isBold
        .Italic = False
        .Underline = 0
        .ColorIndex = 0
    End With
    With rng.ParagraphFormat
        .Alignment = 0
        .LineSpacingRule = 0
        .SpaceBefore = 0
        .SpaceAfter = 0
        .FirstLineIndent = 0
        .LeftIndent = leftIndentPts
        .TabStops.ClearAll
        .TabStops.Add Position:=tabPos, Alignment:=2, Leader:=1
    End With
End Sub

Private Sub CaptureZoneBoundaries(ByVal doc As Document, ByRef zNames As Variant, _
                                  ByRef outRngs() As Range, ByRef outPages() As Long, _
                                  ByRef outVerts() As Single, ByRef outCnt As Long)
    Dim k As Long, zName As String, zRng As Range, nextP As Paragraph, bRng As Range
    Dim pg As Long, vPos As Single
    outCnt = 0
    On Error Resume Next
    For k = LBound(zNames) To UBound(zNames)
        zName = CStr(zNames(k))
        If doc.Bookmarks.Exists(zName) Then
            Err.Clear
            Set zRng = doc.Bookmarks(zName).Range
            Set nextP = Nothing
            If zRng.Paragraphs.Count > 0 Then
                Set nextP = zRng.Paragraphs(zRng.Paragraphs.Count).Next
            End If
            If Err.Number = 0 And Not nextP Is Nothing Then
                Set bRng = nextP.Range.Duplicate
                bRng.Collapse 1 ' wdCollapseStart = 1
                pg = bRng.Information(3) ' wdActiveEndPageNumber = 3
                vPos = bRng.Information(6) ' wdVerticalPositionRelativeToPage = 6
                If Err.Number = 0 And pg > 0 Then
                    outCnt = outCnt + 1
                    Set outRngs(outCnt) = bRng
                    outPages(outCnt) = pg
                    outVerts(outCnt) = vPos
                End If
            End If
            Err.Clear
        End If
    Next k
End Sub

Private Function AreZoneBoundariesStable(ByRef bRngs() As Range, ByRef bPages() As Long, _
                                         ByRef bVerts() As Single, ByVal bCnt As Long) As Boolean
    If bCnt <= 0 Then Exit Function
    Dim k As Long, afterPg As Long, afterVert As Single
    On Error Resume Next
    For k = 1 To bCnt
        Err.Clear
        afterPg = bRngs(k).Information(3)
        afterVert = bRngs(k).Information(6)
        If Err.Number <> 0 Or afterPg <= 0 Then
            Err.Clear
            Exit Function
        End If
        If afterPg <> bPages(k) Or Abs(afterVert - bVerts(k)) >= 0.5# Then
            Exit Function
        End If
    Next k
    AreZoneBoundariesStable = True
End Function

Public Function SyncAllListPages(ByVal doc As Document, ByVal pageMode As Long) As Boolean
    On Error GoTo ErrHandler
    SyncAllListPages = True
    Dim zNames As Variant
    zNames = Array(BM_ZONE_LOF, BM_ZONE_LOT, BM_ZONE_TOC)
    
    Dim totalParas As Long
    totalParas = 0
    Dim zIdx As Long
    For zIdx = LBound(zNames) To UBound(zNames)
        If doc.Bookmarks.Exists(CStr(zNames(zIdx))) Then
            totalParas = totalParas + doc.Bookmarks(CStr(zNames(zIdx))).Range.Paragraphs.Count
        End If
    Next zIdx
    
    If totalParas = 0 Then
        Exit Function
    End If
    
    Dim bms() As String
    Dim oldVals() As String
    Dim newVals() As String
    Dim pRngs() As Range
    ReDim bms(1 To totalParas)
    ReDim oldVals(1 To totalParas)
    ReDim newVals(1 To totalParas)
    ReDim pRngs(1 To totalParas)
    
    Dim validCount As Long
    validCount = 0
    
    Dim curZName As String
    Dim zRng As Range
    Dim p As Paragraph
    Dim pRng As Range
    Dim pText As String
    Dim tabIdx As Long
    Dim oldNum As String
    Dim targetBm As String
    
    For zIdx = LBound(zNames) To UBound(zNames)
        curZName = CStr(zNames(zIdx))
        If doc.Bookmarks.Exists(curZName) Then
            Set zRng = doc.Bookmarks(curZName).Range
            For Each p In zRng.Paragraphs
                Set pRng = p.Range
                If pRng.Hyperlinks.Count > 0 Then
                    targetBm = pRng.Hyperlinks(1).SubAddress
                    If Left$(targetBm, 3) = "bm_" And targetBm <> BM_SYS_GENERATED_BLOCK Then
                        pText = pRng.Text
                        tabIdx = InStrRev(pText, vbTab)
                        If tabIdx > 0 Then
                            oldNum = Mid$(pText, tabIdx + 1)
                            oldNum = Replace(Replace(Replace(oldNum, vbCr, ""), Chr(7), ""), Chr(11), "")
                            If Len(oldNum) > 0 Then
                                validCount = validCount + 1
                                bms(validCount) = targetBm
                                oldVals(validCount) = oldNum
                                Set pRngs(validCount) = pRng
                            End If
                        End If
                    End If
                End If
            Next p
        End If
    Next zIdx
    
    If validCount = 0 Then
        Exit Function
    End If
    
    Dim i As Long
    Dim realPage As String
    Dim fRng As Range
    Dim hlRng As Range
    Dim tailRng As Range
    Dim hlEnd As Long
    Dim oldLen As Long
    Dim didReplace As Boolean
    Dim pass As Long
    Dim anyChangedInPass As Boolean
    Dim lastPassChanged As Boolean
    lastPassChanged = False
    Dim writeFailures As Long
    Dim afterZoneRngs(1 To 3) As Range
    Dim beforePage(1 To 3) As Long
    Dim beforeVert(1 To 3) As Single
    Dim zoneBoundCnt As Long
    zoneBoundCnt = 0
    ResetSectionStyleCache
    mSecCacheActive = True
    
    For pass = 1 To 3
        anyChangedInPass = False
        writeFailures = 0
        
        ' Phase A: Batch Read all page numbers across all zones on pristine layout
        For i = 1 To validCount
            If doc.Bookmarks.Exists(bms(i)) Then
                newVals(i) = GetFormattedPageNumber(doc.Bookmarks(bms(i)).Range, pageMode)
            Else
                If oldVals(i) = DUMMY_PAGE_STRING Then
                    realPage = CStr(pRngs(i).Information(3))
                    If Len(realPage) = 0 Or realPage = "0" Then
                        realPage = "1"
                    End If
                    newVals(i) = realPage
                    ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
                Else
                    newVals(i) = oldVals(i)
                End If
            End If
        Next i
        
        If pass = 1 Then
            CaptureZoneBoundaries doc, zNames, afterZoneRngs, beforePage, beforeVert, zoneBoundCnt
        End If
        
        ' Phase B: Batch Write changed page numbers across all zones (hyperlink tail-slice fast-path + Find fallback)
        For i = 1 To validCount
            If oldVals(i) <> newVals(i) Then
                Dim entryReplaced As Boolean
                entryReplaced = False
                If Not pRngs(i) Is Nothing Then
                    If pRngs(i).Hyperlinks.Count = 1 Then
                        Set hlRng = pRngs(i).Hyperlinks(1).Range
                        hlEnd = hlRng.End
                        oldLen = Len(oldVals(i))
                        If hlEnd - oldLen >= hlRng.Start Then
                            Set tailRng = doc.Range(hlEnd - oldLen, hlEnd)
                            If tailRng.Text = oldVals(i) Then
                                tailRng.Text = newVals(i)
                                anyChangedInPass = True
                                oldVals(i) = newVals(i)
                                entryReplaced = True
                            End If
                        End If
                        If (Not entryReplaced) And (hlEnd - 1 - oldLen >= hlRng.Start) Then
                            Set tailRng = doc.Range(hlEnd - 1 - oldLen, hlEnd - 1)
                            If tailRng.Text = oldVals(i) Then
                                tailRng.Text = newVals(i)
                                anyChangedInPass = True
                                oldVals(i) = newVals(i)
                                entryReplaced = True
                            End If
                        End If
                    End If
                    If Not entryReplaced Then
                        Set fRng = pRngs(i).Duplicate
                        ModMathHelpers.InitDomFind fRng.Find, vbTab & oldVals(i)
                        fRng.Find.Replacement.ClearFormatting
                        didReplace = fRng.Find.Execute(FindText:=vbTab & oldVals(i), _
                                                       MatchCase:=False, MatchWholeWord:=False, _
                                                       MatchWildcards:=False, Forward:=False, Wrap:=0, _
                                                       Format:=False)
                        If didReplace Then
                            If fRng.Start >= pRngs(i).Start And fRng.End <= pRngs(i).End Then
                                fRng.Start = fRng.Start + 1
                                fRng.Text = newVals(i)
                                anyChangedInPass = True
                                oldVals(i) = newVals(i)
                                entryReplaced = True
                            End If
                        End If
                    End If
                End If
                If Not entryReplaced Then
                    writeFailures = writeFailures + 1
                End If
            End If
        Next i
        
        If writeFailures > 0 Then
            ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + writeFailures
            ModRibbonCallbacks.g_LastError = "Page sync: " & writeFailures & " entry replacement(s) failed."
            SyncAllListPages = False
        End If
        
        lastPassChanged = anyChangedInPass
        If Not anyChangedInPass Then
            Exit For
        End If
        
        If pass = 1 And zoneBoundCnt > 0 Then
            If AreZoneBoundariesStable(afterZoneRngs, beforePage, beforeVert, zoneBoundCnt) Then
                lastPassChanged = False
                Exit For
            End If
        End If
        
        ' Repaginate before next pass to calculate any layout shifts caused by digit alterations
        If pass < 3 Then
            doc.Repaginate
        End If
    Next pass
    
    ' Phase C: Read-Only Verification if Pass 3 produced changes (zero extra writes)
    If lastPassChanged Then
        doc.Repaginate
        Dim hasResidual As Boolean
        hasResidual = False
        For i = 1 To validCount
            If doc.Bookmarks.Exists(bms(i)) Then
                Dim checkPage As String
                checkPage = GetFormattedPageNumber(doc.Bookmarks(bms(i)).Range, pageMode)
                If checkPage <> oldVals(i) Then
                    hasResidual = True
                    Exit For
                End If
            End If
        Next i
        If hasResidual Then
            ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
            ModRibbonCallbacks.g_LastError = "Page sync bounded at 3 passes; residual page shift detected."
            SyncAllListPages = False
        End If
    End If
    ResetSectionStyleCache
    Exit Function

ErrHandler:
    ResetSectionStyleCache
    SyncAllListPages = False
    ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
    ModRibbonCallbacks.g_LastError = "SyncAllListPages Error " & Err.Number & ": " & Err.Description
End Function

Public Function ApplyEqTagAndWrapRange(ByVal doc As Document, ByVal om As OMath, ByVal eqTag As String) As Range
    Dim dispTag As String, rTab As Single, bStart As Long, tagRng As Range
    Dim cWidth As Single, padL As Single, padR As Single, cellErr As Long
    If Left$(eqTag, 1) = "(" And Right$(eqTag, 1) = ")" Then
        dispTag = eqTag
    Else
        dispTag = "(" & eqTag & ")"
    End If
    If om.Range.Information(12) Then
        cWidth = 0#
        padL = 0#
        padR = 0#
        On Error Resume Next
        cWidth = om.Range.Cells(1).Width
        cellErr = Err.Number
        If cellErr = 0 Then
            padL = om.Range.Cells(1).LeftPadding
            If Err.Number <> 0 Or padL < 0# Or padL > 100# Then
                Err.Clear
                padL = 5.4!
            End If
            padR = om.Range.Cells(1).RightPadding
            If Err.Number <> 0 Or padR < 0# Or padR > 100# Then
                Err.Clear
                padR = 5.4!
            End If
        End If
        Err.Clear
        On Error GoTo 0
        If cellErr <> 0 Or cWidth >= 999999# Or cWidth <= 0# Then
            rTab = CalculateRightTab(doc, om.Range) / 2#
        Else
            rTab = cWidth - padL - padR - 2#
            If rTab < 40# Or rTab > 2000# Then
                rTab = CalculateRightTab(doc, om.Range) / 2#
            End If
        End If
    Else
        rTab = CalculateRightTab(doc, om.Range)
    End If
    om.Range.ParagraphFormat.Alignment = 0
    om.Range.ParagraphFormat.FirstLineIndent = 0#
    om.Range.ParagraphFormat.LeftIndent = 0#
    om.Range.ParagraphFormat.TabStops.ClearAll
    om.Range.ParagraphFormat.TabStops.Add rTab / 2#, 1
    om.Range.ParagraphFormat.TabStops.Add rTab, 2
    bStart = om.Range.Start
    om.Range.InsertBefore vbTab
    Set tagRng = doc.Range(om.Range.End, om.Range.End)
    tagRng.InsertAfter vbTab & dispTag
    Set ApplyEqTagAndWrapRange = doc.Range(bStart, tagRng.End)
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
