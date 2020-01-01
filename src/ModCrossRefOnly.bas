Attribute VB_Name = "ModCrossRefOnly"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModCrossRefOnly
' Purpose: Teamwork Mode - Zero-Typography, Zero-TOC, Zero-Outline Cross-Ref Linker
' Architecture: Academic Document Format Engine (100% 7-bit ASCII, < 250 lines)
' ==============================================================================

Private Const MAX_HEADING_LEN As Long = 350

Private Function ExecuteCrossRefOnlyCore(ByVal doc As Document, ByRef outTargets As Long, _
                                         ByRef outLinksAfter As Long, ByRef outErr As String) As Boolean
    Dim prevScreen As Boolean, prevTrack As Boolean, objUndo As Object, weStartedUndo As Boolean
    Dim totalParas As Long, startP As Long, hasExp As Boolean, markers As MarkerLocations
    prevScreen = Application.ScreenUpdating
    On Error GoTo ErrHandler
    If doc Is Nothing Then
        outErr = "No active document."
        Exit Function
    End If
    If Not ModBackup.ValidateWritableDocument(doc, outErr) Then Exit Function
    totalParas = doc.Paragraphs.Count
    If totalParas < 1 Then
        outErr = "Document is empty."
        Exit Function
    End If
    ModMain.g_IsEngineRunning = True
    ModRibbonCallbacks.EnsureConfigInitialized
    ModRibbonCallbacks.InvalidateRibbonSafe
    Application.ScreenUpdating = False
    ModProgress.ProgressInit ModConstants.Uni("Link Cross-Refs Only (Ch\u1ec9 g\u1eafn Link ch\u00e9o)")
    weStartedUndo = ModBackup.StartUndoSafe("AcademicFormatter: Link Cross-Refs Only", objUndo)
    If Not ModBackup.HandleTrackChanges(doc, prevTrack) Then
        outErr = "Cancelled due to Track Changes."
        GoTo CleanExit
    End If

    startP = ModPatchMatcher.InferBodyStartPara(doc, hasExp)
    If startP < 1 Then startP = 1
    markers.HasContentStart = True
    markers.ContentStartParaIndex = startP - 1

    Dim exclS() As Long, exclE() As Long, exclCount As Long
    ModPatchDom.CollectExcludedSpans doc, exclS, exclE, exclCount
    ModTableIndex.BuildTableIndex doc
    ModTableIndex.BuildListIndex doc
    ModTableIndex.BuildOMathIndex doc
    ModCustomProfile.LoadTagCache doc
    Dim hasTags As Boolean, tblValid As Boolean, listValid As Boolean
    hasTags = ModCustomProfile.HasActiveTags()
    tblValid = ModTableIndex.IsTableIndexValidForDoc(doc)
    listValid = ModTableIndex.IsListIndexValidForDoc(doc)

    Dim reL1 As Object, reL2 As Object, reL3 As Object, reL4 As Object
    Dim reTbl As Object, reFig As Object, reSrc As Object, reMark As Object
    Set reL1 = CreateObject("VBScript.RegExp")
    reL1.IgnoreCase = True
    reL1.Pattern = ModLocale.GetPatternLevel1()
    Set reL2 = CreateObject("VBScript.RegExp")
    reL2.IgnoreCase = True
    reL2.Pattern = ModConstants.GetPatternLevel2()
    Set reL3 = CreateObject("VBScript.RegExp")
    reL3.IgnoreCase = True
    reL3.Pattern = ModConstants.GetPatternLevel3()
    Set reL4 = CreateObject("VBScript.RegExp")
    reL4.IgnoreCase = True
    reL4.Pattern = ModConstants.GetPatternLevel4()
    Set reTbl = CreateObject("VBScript.RegExp")
    reTbl.IgnoreCase = True
    reTbl.Pattern = ModLocale.GetPatternTable()
    Set reFig = CreateObject("VBScript.RegExp")
    reFig.IgnoreCase = True
    reFig.Pattern = ModLocale.GetPatternFigure()
    Set reSrc = CreateObject("VBScript.RegExp")
    reSrc.IgnoreCase = True
    reSrc.Pattern = ModLocale.GetPatternSourceNote()
    Set reMark = CreateObject("VBScript.RegExp")
    reMark.IgnoreCase = True
    reMark.Pattern = ModConstants.GetPatternMarker()

    Dim snaps() As ParagraphSnapshot, p As Paragraph, pRng As Range, i As Long
    Dim rawTxt As String, cleanTxt As String, rStart As Long, rEnd As Long
    Dim isExcl As Boolean, inTbl As Boolean, inList As Boolean, tagCode As Long, oLvl As Long
    Dim eType As Long, cntToc As Long, cntTbl As Long, cntFig As Long, listPfx As String, lType As Long
    ReDim snaps(1 To totalParas)
    ModProgress.ProgressSetPhase "Scanning targets...", 5, 45
    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        If (i Mod 20) = 0 Then ModProgress.ProgressUpdate i, totalParas
        Set pRng = p.Range
        rawTxt = pRng.Text
        rStart = pRng.Start
        rEnd = pRng.End
        cleanTxt = ModPatchMatcher.NormalizeStr(rawTxt)
        snaps(i).ParaIndex = i
        snaps(i).RangeStart = rStart
        snaps(i).RangeEnd = rEnd
        snaps(i).RawText = rawTxt
        snaps(i).CleanText = cleanTxt
        snaps(i).ElementType = ElementTypeUnknown
        snaps(i).IsExcluded = False

        isExcl = (i < startP)
        If Not isExcl And exclCount > 0 Then isExcl = ModPatchDom.IsInExcludedSpans(rStart, rEnd, exclS, exclE, exclCount)
        If Not isExcl And Len(cleanTxt) > 0 Then
            If reMark.Test(cleanTxt) Then isExcl = True
        End If
        If isExcl Then
            snaps(i).IsExcluded = True
            snaps(i).ElementType = ElementTypeOldToc
            GoTo NextScanItem
        End If

        If tblValid Then inTbl = ModTableIndex.IsPosInTable(rStart) Else inTbl = CBool(pRng.Information(12))
        If listValid Then inList = ModTableIndex.IsPosInList(rStart) Else inList = (pRng.ListFormat.ListType <> 0)
        oLvl = 10
        On Error Resume Next
        oLvl = p.OutlineLevel
        If Err.Number = 0 And oLvl <= 4 And Len(cleanTxt) > 0 Then
            lType = pRng.ListFormat.ListType
            If Err.Number = 0 And lType >= 3 Then
                listPfx = Trim$(pRng.ListFormat.ListString)
                If Err.Number = 0 And Len(listPfx) > 0 Then
                    If Left$(cleanTxt, Len(listPfx)) <> listPfx Then
                        cleanTxt = listPfx & " " & cleanTxt
                        snaps(i).CleanText = cleanTxt
                    End If
                End If
            End If
        End If
        Err.Clear
        On Error GoTo ErrHandler

        tagCode = 0
        If hasTags Then tagCode = ModCustomProfile.LookupTagForRange(rStart, rEnd)
        eType = ElementTypeUnknown
        If tagCode > 0 Then
            eType = tagCode
        ElseIf tagCode = TAG_CODE_REF Or tagCode = TAG_CODE_APP Then
            eType = ElementTypeLevel1
        ElseIf tagCode = TAG_CODE_REF_SUB Then
            eType = ElementTypeLevel2
        ElseIf tagCode = TAG_CODE_SKIP Then
            If inTbl Then eType = ElementTypeTableCell Else eType = ElementTypeNormalParagraph
        ElseIf inTbl Then
            eType = ElementTypeTableCell
        ElseIf inList And oLvl > 4 Then
            snaps(i).IsExcluded = True
            eType = ElementTypeList
        ElseIf Len(cleanTxt) > 0 And Len(cleanTxt) <= MAX_HEADING_LEN Then
            If reTbl.Test(cleanTxt) Then
                eType = ElementTypeTableTitle
            ElseIf reFig.Test(cleanTxt) Then
                eType = ElementTypeFigureTitle
            ElseIf reSrc.Test(cleanTxt) Then
                eType = ElementTypeSourceNote
            ElseIf oLvl >= 1 And oLvl <= 4 Then
                eType = oLvl
            ElseIf reL1.Test(cleanTxt) Or ModLocale.IsRomanLevel1(cleanTxt) Then
                eType = ElementTypeLevel1
            ElseIf reL2.Test(cleanTxt) Then
                eType = ElementTypeLevel2
            ElseIf reL3.Test(cleanTxt) Then
                eType = ElementTypeLevel3
            ElseIf reL4.Test(cleanTxt) Then
                eType = ElementTypeLevel4
            ElseIf i >= startP Then
                eType = ElementTypeNormalParagraph
            End If
        ElseIf Len(cleanTxt) > 0 And i >= startP Then
            If reSrc.Test(cleanTxt) Then eType = ElementTypeSourceNote Else eType = ElementTypeNormalParagraph
        End If
        snaps(i).ElementType = eType

        If i >= startP Then
            If eType >= ElementTypeLevel1 And eType <= ElementTypeLevel4 Then
                cntToc = cntToc + 1
                ModBookmark.AddBookmark doc, ModBookmark.FormatBookmarkId(BM_PREFIX_TOC, cntToc), pRng
            ElseIf eType = ElementTypeTableTitle Then
                cntTbl = cntTbl + 1
                ModBookmark.AddBookmark doc, ModBookmark.FormatBookmarkId(BM_PREFIX_TBL, cntTbl), pRng
            ElseIf eType = ElementTypeFigureTitle Then
                cntFig = cntFig + 1
                ModBookmark.AddBookmark doc, ModBookmark.FormatBookmarkId(BM_PREFIX_FIG, cntFig), pRng
            End If
        End If
NextScanItem:
    Next p

    ModBookmark.CleanExcessBookmarks doc, cntToc, cntTbl, cntFig
    ModProgress.ProgressSetPhase "Linking cross-references...", 50, 95
    ModCrossRef.ApplyCrossReferences doc, markers, snaps, totalParas, ModRibbonCallbacks.g_Config.CrossRefColor
    outTargets = cntToc + cntTbl + cntFig
    outLinksAfter = doc.Hyperlinks.Count
    ExecuteCrossRefOnlyCore = True

CleanExit:
    On Error Resume Next
    ModBackup.RestoreTrackChanges doc, prevTrack
    ModBackup.EndUndoSafe objUndo, weStartedUndo
    ModProgress.ProgressFinish
    ModTableIndex.InvalidateAllIndexes
    ModCustomProfile.ResetTagCache
    Application.ScreenUpdating = prevScreen
    ModMain.g_IsEngineRunning = False
    ModRibbonCallbacks.InvalidateRibbonSafe
    Err.Clear
    Exit Function

ErrHandler:
    outErr = "CrossRefOnly Error " & Err.Number & ": " & Err.Description
    ModRibbonCallbacks.g_LastError = outErr
    Resume CleanExit
End Function

Public Sub RunLinkCrossRefsOnly(Optional ByVal targetDoc As Document = Nothing)
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then Exit Sub
    Dim doc As Document
    If targetDoc Is Nothing Then
        If Documents.Count = 0 Then Exit Sub
        Set doc = ActiveDocument
    Else
        Set doc = targetDoc
    End If
    Dim tgtCnt As Long, lnkCnt As Long, errMsg As String
    If ExecuteCrossRefOnlyCore(doc, tgtCnt, lnkCnt, errMsg) Then
        If Application.DisplayAlerts <> 0 Then
            ModUI.MsgBoxW ModConstants.Uni("Link Cross-Refs Only completed! (Ho\u00e0n t\u1ea5t g\u1eafn Link ch\u00e9o!)") & vbCrLf & _
                          "- Targets indexed: " & CStr(tgtCnt) & vbCrLf & _
                          "- Document hyperlinks: " & CStr(lnkCnt) & vbCrLf & _
                          ModConstants.Uni("(0 typography/TOC/Outline changes made.)"), _
                          vbInformation, ModConstants.Uni("Link Cross-Refs Only (Ch\u1ec9 g\u1eafn Link ch\u00e9o)")
        End If
    ElseIf Len(errMsg) > 0 Then
        ModUI.MsgBoxW errMsg, vbExclamation, ModConstants.Uni("Link Cross-Refs Only (Ch\u1ec9 g\u1eafn Link ch\u00e9o)")
    End If
End Sub

Public Function RunLinkCrossRefsOnlyForTest(ByVal doc As Document) As String
    Dim tgtCnt As Long, lnkCnt As Long, errMsg As String
    If ExecuteCrossRefOnlyCore(doc, tgtCnt, lnkCnt, errMsg) Then
        RunLinkCrossRefsOnlyForTest = "OK|" & CStr(tgtCnt) & "|" & CStr(lnkCnt)
    Else
        RunLinkCrossRefsOnlyForTest = "ERR|" & errMsg
    End If
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
