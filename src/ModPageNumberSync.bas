Attribute VB_Name = "ModPageNumberSync"
Option Explicit

' ==============================================================================
' Module: ModPageNumberSync
' Purpose: Roman numeral conversion, dynamic tab leader calculation, page sync
' Architecture: Academic Document Format Engine (Multi-Zone Batch Sync & Bounded Convergence)
' ==============================================================================

Public Function ToRomanLowerCase(ByVal num As Long) As String
    If num <= 0 Then
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
    secNumberStyle = 0
    secNumberStyle = r.Sections(1).Footers(1).PageNumbers.NumberStyle
    If Err.Number <> 0 Or secNumberStyle = 0 Then
        Err.Clear
        secNumberStyle = r.Sections(1).Headers(1).PageNumbers.NumberStyle
    End If
    If Err.Number <> 0 Then
        secNumberStyle = 0
        Err.Clear
    End If
    
    Select Case secNumberStyle
        Case 1 ' wdPageNumberStyleUppercaseRoman = 1
            GetFormattedPageNumber = UCase$(ToRomanLowerCase(pNum))
        Case 2 ' wdPageNumberStyleLowercaseRoman = 2
            GetFormattedPageNumber = ToRomanLowerCase(pNum)
        Case Else
            GetFormattedPageNumber = CStr(pNum)
    End Select
    If Err.Number <> 0 Then
        ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
        GetFormattedPageNumber = CStr(pNum)
        Err.Clear
    End If
End Function

Public Function CalculateRightTab(ByVal doc As Document) As Single
    With doc.PageSetup
        CalculateRightTab = .PageWidth - .LeftMargin - .RightMargin
    End With
End Function

Public Sub ApplyListLineFormat(ByVal rng As Range, _
                               ByVal fontName As String, _
                               ByVal fontSize As Single, _
                               ByVal tabPos As Single, _
                               ByVal leftIndentPts As Single, _
                               ByVal isBold As Boolean)
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
        .LineSpacingRule = 0
        .SpaceBefore = 0
        .SpaceAfter = 0
        .FirstLineIndent = 0
        .LeftIndent = leftIndentPts
        .TabStops.ClearAll
        .TabStops.Add Position:=tabPos, Alignment:=2, Leader:=1
    End With
End Sub

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
    Dim zTags() As String
    Dim pIndices() As Long
    ReDim bms(1 To totalParas)
    ReDim oldVals(1 To totalParas)
    ReDim newVals(1 To totalParas)
    ReDim zTags(1 To totalParas)
    ReDim pIndices(1 To totalParas)
    
    Dim validCount As Long
    validCount = 0
    
    Dim curZName As String
    Dim zRng As Range
    Dim p As Paragraph
    Dim k As Long
    Dim pText As String
    Dim tabIdx As Long
    Dim oldNum As String
    Dim targetBm As String
    
    For zIdx = LBound(zNames) To UBound(zNames)
        curZName = CStr(zNames(zIdx))
        If doc.Bookmarks.Exists(curZName) Then
            Set zRng = doc.Bookmarks(curZName).Range
            k = 0
            For Each p In zRng.Paragraphs
                k = k + 1
                If p.Range.Hyperlinks.Count > 0 Then
                    targetBm = p.Range.Hyperlinks(1).SubAddress
                    If Left$(targetBm, 3) = "bm_" And targetBm <> BM_SYS_GENERATED_BLOCK Then
                        pText = p.Range.Text
                        tabIdx = InStr(pText, vbTab)
                        If tabIdx > 0 Then
                            oldNum = Mid$(pText, tabIdx + 1)
                            oldNum = Replace(Replace(Replace(oldNum, vbCr, ""), Chr(7), ""), Chr(11), "")
                            If Len(oldNum) > 0 Then
                                validCount = validCount + 1
                                bms(validCount) = targetBm
                                oldVals(validCount) = oldNum
                                zTags(validCount) = curZName
                                pIndices(validCount) = k
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
    Dim didReplace As Boolean
    Dim pass As Long
    Dim anyChangedInPass As Boolean
    Dim lastPassChanged As Boolean
    lastPassChanged = False
    Dim writeFailures As Long
    writeFailures = 0
    
    For pass = 1 To 2
        anyChangedInPass = False
        
        ' Phase A: Batch Read all page numbers across all zones on pristine layout
        For i = 1 To validCount
            If doc.Bookmarks.Exists(bms(i)) Then
                newVals(i) = GetFormattedPageNumber(doc.Bookmarks(bms(i)).Range, pageMode)
            Else
                If oldVals(i) = DUMMY_PAGE_STRING Then
                    Dim fallbackRng As Range
                    Set fallbackRng = doc.Bookmarks(zTags(i)).Range.Paragraphs(pIndices(i)).Range
                    realPage = CStr(fallbackRng.Information(3))
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
        
        ' Phase B: Batch Write changed page numbers across all zones
        For i = 1 To validCount
            If oldVals(i) <> newVals(i) Then
                Dim entryReplaced As Boolean
                entryReplaced = False
                If doc.Bookmarks.Exists(zTags(i)) Then
                    If pIndices(i) <= doc.Bookmarks(zTags(i)).Range.Paragraphs.Count Then
                        Set fRng = doc.Bookmarks(zTags(i)).Range.Paragraphs(pIndices(i)).Range.Duplicate
                        fRng.Find.ClearFormatting
                        didReplace = fRng.Find.Execute(FindText:=vbTab & oldVals(i), _
                                                       MatchCase:=False, MatchWholeWord:=False, _
                                                       MatchWildcards:=False, Forward:=True, Wrap:=0, _
                                                       Format:=False, ReplaceWith:=vbTab & newVals(i), Replace:=1)
                        If didReplace Then
                            anyChangedInPass = True
                            oldVals(i) = newVals(i)
                            entryReplaced = True
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
        
        ' Repaginate before Pass 2 to calculate any layout shifts caused by digit alterations
        If pass < 2 Then
            doc.Repaginate
        End If
    Next pass
    
    ' Phase C: Read-Only Verification if Pass 2 produced changes (zero extra writes)
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
            ModRibbonCallbacks.g_LastError = "Page sync bounded at 2 passes; residual page shift detected."
            SyncAllListPages = False
        End If
    End If
    Exit Function

ErrHandler:
    SyncAllListPages = False
    ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
    ModRibbonCallbacks.g_LastError = "SyncAllListPages Error " & Err.Number & ": " & Err.Description
End Function
