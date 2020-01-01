Attribute VB_Name = "ModMathController"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ModMathController: Math conversion, Shading, Ribbon, Undo (vba_lint compliant)
Public g_IsMathBuilding As Boolean
Public Function NextMathSeq(ByVal doc As Document) As Long
    NextMathSeq = ModMathHelpers.NextMathSeq(doc)
End Function
Public Sub SetDocVariable(ByVal doc As Document, ByVal varName As String, ByVal varVal As String)
    ModMathHelpers.SetDocVariable doc, varName, varVal
End Sub
Private Sub CountMathOutcome(ByVal ok As Boolean, ByVal isPink As Boolean, ByRef builtCnt As Long, ByRef pinkCnt As Long)
    If ok Then
        builtCnt = builtCnt + 1
    ElseIf isPink Then
        pinkCnt = pinkCnt + 1
    End If
End Sub

Private Function InsertOpcAndShade(ByVal doc As Document, ByVal targetRng As Range, ByVal ommlXml As String, ByVal isBlock As Boolean, ByVal bmID As String, Optional ByVal eqTag As String = "") As Boolean
    Dim opcXml As String, om As OMath, wrapRng As Range, shadeCol As Long
    opcXml = ModOmmlEmit.WrapOmmlInFlatOpc(ommlXml)
    ModMathInline.TrimTrailingDomMarks targetRng
    On Error GoTo FailInsert
    targetRng.InsertXML opcXml
    If targetRng.OMaths.Count > 0 Then
        On Error Resume Next
        Set om = targetRng.OMaths(1)
        Set wrapRng = om.Range
        If isBlock Then
            om.Type = 0
            shadeCol = ModMathScanner.COLOR_PASTEL_MINT
            If Len(eqTag) > 0 Then
                Set wrapRng = ModPageNumberSync.ApplyEqTagAndWrapRange(doc, om, eqTag)
            Else
                om.Range.ParagraphFormat.Alignment = 1
                om.Range.ParagraphFormat.FirstLineIndent = 0
                om.Range.ParagraphFormat.LeftIndent = 0
            End If
        Else
            om.Type = 1
            shadeCol = ModMathScanner.COLOR_PASTEL_CYAN
        End If
        ModMathHelpers.ApplyPastelShading wrapRng, shadeCol
        doc.Bookmarks.Add bmID, wrapRng
        Err.Clear
        On Error GoTo 0
        InsertOpcAndShade = True
    End If
    Exit Function
FailInsert:
    Err.Clear
End Function
Private Function ConvertWithSeq(ByVal doc As Document, ByVal targetRng As Range, ByVal rawLatex As String, ByVal isBlock As Boolean, ByRef ioSeqId As Long, ByVal prevBm As String, ByRef outIsPink As Boolean) As Boolean
    Dim cleanLatex As String, eqTag As String, ommlXml As String, bmID As String, isAmb As Boolean
    cleanLatex = ModMathScanner.StripOuterDelims(rawLatex)
    If Len(cleanLatex) = 0 Or cleanLatex = "$" Then
        If Len(prevBm) = 0 Then Exit Function
        isAmb = True
    End If
    If ioSeqId <= 0 Then ioSeqId = ModMathHelpers.NextMathSeq(doc)
    bmID = "bm_math_" & Format$(ioSeqId, "00000")
    If Not isAmb Then
        If Len(rawLatex) > 16000 Or ModMathInline.HasForbiddenDomChar(targetRng.Text) Then
            isAmb = True
        Else
            ommlXml = ModOmmlEmit.TranspileLatexToOmml(cleanLatex, isBlock, isAmb, eqTag)
        End If
    End If
    outIsPink = True
    If Not isAmb And Len(ommlXml) > 0 Then
        If InsertOpcAndShade(doc, targetRng, ommlXml, isBlock, bmID, eqTag) Then
            If Len(prevBm) > 0 Then If doc.Bookmarks.Exists(prevBm) Then doc.Bookmarks(prevBm).Delete
            ModMathHelpers.SetDocVariable doc, bmID, rawLatex
            ModMathHelpers.SetDocVariable doc, bmID & "_isblock", IIf(isBlock, "1", "0")
            outIsPink = False
            ConvertWithSeq = True
            Exit Function
        End If
    End If
    If Len(prevBm) = 0 Or Left$(prevBm, 12) = "bm_math_src_" Or Not doc.Bookmarks.Exists("bm_math_pink_" & Format$(ioSeqId, "00000")) Then
        If Len(prevBm) > 0 Then If doc.Bookmarks.Exists(prevBm) Then doc.Bookmarks(prevBm).Delete
        ModMathHelpers.MarkSpanPink doc, targetRng, ioSeqId, rawLatex, isBlock
    End If
End Function
Public Function ConvertRangeToMath(ByVal doc As Document, ByVal targetRng As Range, ByVal rawLatex As String, ByVal isBlock As Boolean, ByRef outSeqId As Long, ByRef outIsPink As Boolean) As Boolean
    outSeqId = 0
    ConvertRangeToMath = ConvertWithSeq(doc, targetRng, rawLatex, isBlock, outSeqId, "", outIsPink)
End Function
Public Sub BuildAllMathInDocument(ByVal doc As Document, Optional ByVal enforceSafety As Boolean = False)
    Dim pairCount As Long, k As Long, dummySeq As Long, totalParas As Long, pStep As Long, docEnd As Long, prevViewType As Long
    Dim builtCount As Long, pinkCount As Long, prevFindEnd As Long, pEnd As Long, spanSteps As Long
    Dim isPink As Boolean, isScanningActive As Boolean, isAmb As Boolean, weStartedUndo As Boolean
    Dim prevScreen As Boolean, prevPagination As Boolean, hasPaginationState As Boolean, hasWindowView As Boolean
    Dim hasErr As Boolean, wasCancelled As Boolean, didStartUndo As Boolean, wasTracking As Boolean, backupPath As String
    Dim pStarts() As Long, pEnds() As Long, clsArr() As Long, candTexts() As String
    Dim rawLatex As String, rawPText As String, pText As String, innerT As String, cAfter As String, spanText As String
    Dim candRng As Range, fRng As Range, p As Paragraph, nextP As Paragraph, endPara As Paragraph, objUndo As Object
    If g_IsMathBuilding Or ModMain.g_IsEngineRunning Or doc Is Nothing Then Exit Sub
    If enforceSafety Then
        If Not ModBackup.BeginProtectedSession(doc, backupPath, wasTracking) Then Exit Sub
    End If
    ModConstants.g_SilentErrors = 0
    ModRibbonCallbacks.g_LastError = ""
    prevScreen = Application.ScreenUpdating
    On Error GoTo ErrHandler
    g_IsMathBuilding = True
    ModRibbonCallbacks.InvalidateRibbonSafe
    Application.EnableCancelKey = 0
    Application.ScreenUpdating = False
    On Error Resume Next
    prevPagination = Application.Options.Pagination
    If Err.Number = 0 Then
        hasPaginationState = True
        Application.Options.Pagination = False
    End If
    Err.Clear
    If doc.Windows.Count > 0 Then
        If Not doc.ActiveWindow Is Nothing Then
            prevViewType = doc.ActiveWindow.View.Type
            If Err.Number = 0 Then
                hasWindowView = True
                doc.ActiveWindow.View.Type = 1
            End If
            Err.Clear
        End If
    End If
    On Error GoTo ErrHandler
    weStartedUndo = ModBackup.StartUndoSafe("Build LaTeX Math (Review Mode)", objUndo)
    ModProgress.ProgressInit "LaTeX Math Engine"
    ModProgress.ProgressSetPhase ModConstants.Uni("Converting Math (Chuy\u1ec3n \u0111\u1ed5i to\u00e1n)"), 5#, 95#
    ModMathHelpers.InitMathSeqCache doc
    ModMathScanner.InitNoFlyZoneCache doc
    totalParas = doc.Paragraphs.Count
    Set p = doc.Paragraphs.First
    docEnd = doc.Content.End
    Set fRng = doc.Content.Duplicate
    ModMathHelpers.InitDomFind fRng.Find, "=contentStart()"
    fRng.Find.MatchCase = False
    prevFindEnd = 0
    Do While fRng.Find.Execute
        If fRng.End <= prevFindEnd Then Exit Do
        prevFindEnd = fRng.End
        pText = LCase$(Trim$(Replace(Replace(Replace(fRng.Paragraphs(1).Range.Text, vbCr, ""), Chr(11), " "), Chr(7), "")))
        If pText = "=contentstart()" Then
            If Not CBool(fRng.Paragraphs(1).Range.Information(12)) Then
                Set p = fRng.Paragraphs(1).Next
                Exit Do
            End If
        End If
        If fRng.End >= docEnd Then Exit Do
        fRng.SetRange fRng.End, docEnd
    Loop
    isScanningActive = True
    Do While Not p Is Nothing
        pStep = pStep + 1
        If pStep <= totalParas Then ModProgress.ProgressUpdate pStep, totalParas
        If ModProgress.g_CancelRequested Then
            wasCancelled = True
            GoTo CleanExit
        End If
        Set nextP = p.Next
        rawPText = p.Range.Text
        pText = LCase$(Trim$(Replace(Replace(Replace(rawPText, vbCr, ""), Chr(11), " "), Chr(7), "")))
        If pText = "=contentend()" Then
            If Not CBool(p.Range.Information(12)) Then isScanningActive = False
        ElseIf InStr(1, rawPText, "$") = 0 And InStr(1, rawPText, "\") = 0 Then
            If Len(rawPText) <= 165 Then ModMathScanner.CheckNoFlyHeading rawPText, isScanningActive
        ElseIf Not ModMathScanner.IsInNoFlyZoneCached(p.Range, rawPText, isScanningActive) Then
            spanSteps = 0
            If ModMathScanner.CheckMultilineBlock(doc, p, endPara, rawLatex, isAmb, spanSteps) Then
                pStep = pStep + spanSteps
                Set nextP = endPara.Next
                CountMathOutcome ConvertRangeToMath(doc, doc.Range(p.Range.Start, endPara.Range.End), rawLatex, True, dummySeq, isPink), isPink, builtCount, pinkCount
            Else
                pText = Trim$(Replace(Replace(Replace(rawPText, vbCr, ""), Chr(11), " "), Chr(7), ""))
                If ModMathInline.IsSingleLineDisplayBlock(pText) Then
                    rawLatex = Trim$(Replace(Replace(rawPText, vbCr, ""), Chr(7), ""))
                    CountMathOutcome ConvertRangeToMath(doc, p.Range, rawLatex, True, dummySeq, isPink), isPink, builtCount, pinkCount
                ElseIf pText <> "$$" And pText <> "\[" And pText <> "\]" And pText <> "$$$" And pText <> "$$$$" And pText <> "\[\]" Then
                    ModMathInline.FindInlinePairs doc, p, pStarts, pEnds, pairCount, rawPText
                    If pairCount > 0 Then
                        pEnd = p.Range.End
                        ReDim clsArr(1 To pairCount)
                        ReDim candTexts(1 To pairCount)
                        For k = 1 To pairCount
                            spanText = doc.Range(pStarts(k), pEnds(k)).Text
                            candTexts(k) = spanText
                            If Len(spanText) >= 2 Then
                                If Left$(spanText, 2) = "\(" Then
                                    clsArr(k) = 1
                                Else
                                    innerT = Mid$(spanText, 2, Len(spanText) - 2)
                                    cAfter = ""
                                    If pEnds(k) < pEnd Then cAfter = doc.Range(pEnds(k), pEnds(k) + 1).Text
                                    clsArr(k) = ModMathScanner.ClassifyDollarSpan(innerT, "", cAfter)
                                End If
                            End If
                        Next k
                        For k = pairCount To 1 Step -1
                            If clsArr(k) = 1 Or clsArr(k) = 2 Then
                                Set candRng = doc.Range(pStarts(k), pEnds(k))
                                If clsArr(k) = 1 Then
                                    CountMathOutcome ConvertRangeToMath(doc, candRng, candTexts(k), False, dummySeq, isPink), isPink, builtCount, pinkCount
                                ElseIf Not ModMathInline.HasForbiddenDomChar(candTexts(k)) Then
                                    dummySeq = ModMathHelpers.NextMathSeq(doc)
                                    ModMathHelpers.MarkSpanPink doc, candRng, dummySeq, candTexts(k), False
                                    pinkCount = pinkCount + 1
                                End If
                            End If
                        Next k
                    End If
                End If
            End If
        End If
        Set p = nextP
    Loop
CleanExit:
    On Error Resume Next
    Application.EnableCancelKey = 1
    ModMathHelpers.FlushMathSeqCache doc
    ModMathHelpers.ResetMathSeqCache
    ModMathScanner.ResetNoFlyZoneCache
    ModProgress.ProgressFinish
    If hasWindowView And doc.Windows.Count > 0 Then doc.ActiveWindow.View.Type = prevViewType
    If hasPaginationState Then Application.Options.Pagination = prevPagination
    didStartUndo = weStartedUndo
    ModBackup.EndUndoSafe objUndo, weStartedUndo
    If wasCancelled And didStartUndo And Not (doc Is Nothing) Then
        If Not doc.Saved Then doc.Undo
    End If
    Application.ScreenUpdating = prevScreen
    If enforceSafety Then ModBackup.EndProtectedSession doc, backupPath, wasTracking, (hasErr And Not wasCancelled)
    g_IsMathBuilding = False
    ModRibbonCallbacks.InvalidateRibbonSafe
    If wasCancelled Then
        If Application.DisplayAlerts <> 0 Then
            ModUI.MsgBoxW ModConstants.Uni("Operation cancelled safely. (\u0110\u00e3 h\u1ee7y thao t\u00e1c an to\u00e0n.)"), vbInformation, ModConstants.Uni("Cancelled (\u0110\u00e3 h\u1ee7y)")
        End If
    ElseIf hasErr And Application.DisplayAlerts <> 0 Then
        ModUI.MsgBoxW ModConstants.Uni("Math Engine error (L\u1ed7i d\u1ef1ng to\u00e1n): ") & ModRibbonCallbacks.g_LastError, vbCritical, ModConstants.Uni("Error (L\u1ed7i)")
    ElseIf Not hasErr And enforceSafety And Application.DisplayAlerts <> 0 Then
        ModUI.ShowMathSummaryReport builtCount, pinkCount
    End If
    Exit Sub
ErrHandler:
    hasErr = True
    ModRibbonCallbacks.g_LastError = "Math Error " & Err.Number & ": " & Err.Description
    Resume CleanExit
End Sub
Private Function IsShadedAt(ByVal doc As Document, ByVal pos As Long, ByVal hasOMath As Boolean, ByRef outCol As Long) As Boolean
    Dim bgCol As Long
    bgCol = doc.Range(pos, pos + 1).Shading.BackgroundPatternColor
    If bgCol = ModMathScanner.COLOR_PASTEL_AMBER Or bgCol = ModMathScanner.COLOR_PASTEL_PINK Then
        If Not hasOMath Then
            outCol = bgCol
            IsShadedAt = True
        ElseIf doc.Range(pos, pos + 1).OMaths.Count = 0 Then
            outCol = bgCol
            IsShadedAt = True
        End If
    End If
End Function

Private Function IsUniformShadedRange(ByVal doc As Document, ByVal sPos As Long, ByVal ePos As Long, ByVal targetCol As Long, ByVal hasOMath As Boolean) As Boolean
    Dim r As Range
    Set r = doc.Range(sPos, ePos)
    If r.Shading.BackgroundPatternColor = targetCol Then
        If Not hasOMath Then
            IsUniformShadedRange = True
        ElseIf r.OMaths.Count = 0 Then
            IsUniformShadedRange = True
        End If
    End If
End Function

Private Function TryFindShadedMathSpan(ByVal doc As Document, ByRef outRng As Range) As Boolean
    On Error GoTo FailShade
    If Not (Selection.Document Is doc) Then Exit Function
    If Selection.Paragraphs.Count = 0 Then Exit Function
    Dim pRng As Range, pMin As Long, pMax As Long, cPos As Long, lPos As Long, rPos As Long
    Dim minL As Long, maxR As Long, stepSz As Long, probePos As Long, lo As Long, hi As Long, midPos As Long
    Dim targetCol As Long, hasOMath As Boolean, ch As String
    Set pRng = Selection.Paragraphs(1).Range
    pMin = pRng.Start
    pMax = pRng.End
    Do While pMax > pMin
        ch = doc.Range(pMax - 1, pMax).Text
        If ch <> vbCr And ch <> Chr(7) Then Exit Do
        pMax = pMax - 1
    Loop
    If pMax <= pMin Then Exit Function
    hasOMath = (pRng.OMaths.Count > 0)
    cPos = Selection.Start
    If cPos < pMin Then cPos = pMin Else If cPos > pMax Then cPos = pMax
    If cPos < pMax Then
        If IsShadedAt(doc, cPos, hasOMath, targetCol) Then GoTo FoundAnchor
    End If
    If cPos > pMin Then
        cPos = cPos - 1
        If IsShadedAt(doc, cPos, hasOMath, targetCol) Then GoTo FoundAnchor
    End If
    Exit Function
FoundAnchor:
    lPos = cPos
    rPos = cPos + 1
    minL = pMin
    If rPos - 2000 > minL Then minL = rPos - 2000
    Dim chkCol As Long
    Do While lPos > minL
        If IsShadedAt(doc, lPos - 1, hasOMath, chkCol) Then
            If chkCol = targetCol Then
                lPos = lPos - 1
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    maxR = pMax
    If lPos + 2000 < maxR Then maxR = lPos + 2000
    Do While rPos < maxR
        If IsShadedAt(doc, rPos, hasOMath, chkCol) Then
            If chkCol = targetCol Then
                rPos = rPos + 1
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    If rPos > lPos Then Set outRng = doc.Range(lPos, rPos)
    TryFindShadedMathSpan = (rPos > lPos)
FailShade:
End Function

Public Sub ToggleMathAtSelection(ByVal doc As Document)
    Dim selRng As Range, objUndo As Object
    Dim foundBm As String, rawLatex As String, idStr As String, errMsg As String, toggleErr As String
    Dim foundKind As Long, seqNum As Long, isBlock As Boolean, isPink As Boolean
    Dim weStartedUndo As Boolean, wasTracking As Boolean
    If g_IsMathBuilding Or ModMain.g_IsEngineRunning Or doc Is Nothing Then Exit Sub
    If Not (Selection.Document Is doc) Then Exit Sub
    If Selection.StoryType <> 1 Then
        If Application.DisplayAlerts <> 0 Then
            Application.StatusBar = ModConstants.Uni("Toggle Math only works in main document body. (Ch\u1ec9 h\u1ed7 tr\u1ee3 \u0111\u1ed5i to\u00e1n trong ph\u1ea7n n\u1ed9i dung ch\u00ednh.)")
        End If
        Exit Sub
    End If
    If Not ModBackup.ValidateWritableDocument(doc, errMsg) Then
        If Application.DisplayAlerts <> 0 Then
            ModUI.MsgBoxW errMsg, vbExclamation, ModConstants.Uni("Cannot Edit (Kh\u00f4ng th\u1ec3 ch\u1ec9nh s\u1eeda)")
        End If
        Exit Sub
    End If
    wasTracking = False
    If Not ModBackup.HandleTrackChanges(doc, wasTracking) Then Exit Sub
    On Error GoTo ErrHandler
    weStartedUndo = ModBackup.StartUndoSafe("Toggle LaTeX Math", objUndo)
    If Not ModMathHelpers.FindEnclosingMathBookmark(doc, Selection.Start, Selection.End, foundBm, foundKind) Then
        If Selection.End > Selection.Start Then
            Set selRng = Selection.Range
            ModMathInline.TrimMathRange selRng
            If selRng.End > selRng.Start Then
                rawLatex = selRng.Text
                ConvertRangeToMath doc, selRng, rawLatex, ModMathHelpers.InferIsBlock(rawLatex, doc, ""), seqNum, isPink
            ElseIf Application.DisplayAlerts <> 0 Then
                Application.StatusBar = ModConstants.Uni("No math formula selected. Place cursor in a formula or select LaTeX. (Kh\u00f4ng c\u00f3 c\u00f4ng th\u1ee9c \u0111\u01b0\u1ee3c ch\u1ecdn.)")
            End If
        ElseIf TryFindShadedMathSpan(doc, selRng) Then
            ModMathHelpers.ApplyPastelShading selRng, ModMathScanner.COLOR_AUTO_CLEAR
            ModMathInline.TrimMathRange selRng
            If selRng.End > selRng.Start Then
                rawLatex = selRng.Text
                ConvertRangeToMath doc, selRng, rawLatex, ModMathHelpers.InferIsBlock(rawLatex, doc, ""), seqNum, isPink
                If isPink Then ModMathHelpers.NotifyUser ModConstants.Uni("Unsupported LaTeX syntax (C\u00fa ph\u00e1p ch\u01b0a h\u1ed7 tr\u1ee3): ") & Left$(rawLatex, 40)
            End If
        ElseIf Application.DisplayAlerts <> 0 Then
            Application.StatusBar = ModConstants.Uni("No math formula at cursor. Place cursor in a formula or select LaTeX. (Kh\u00f4ng t\u00ecm th\u1ea5y c\u00f4ng th\u1ee9c t\u1ea1i con tr\u1ecf.)")
        End If
    Else
        idStr = Mid$(foundBm, InStrRev(foundBm, "_") + 1)
        If foundKind = ModMathScanner.BM_KIND_OMATH Then
            rawLatex = ModMathHelpers.GetDocVariable(doc, foundBm)
            If Len(rawLatex) = 0 Then
                ModMathHelpers.NotifyUser ModConstants.Uni("Original LaTeX source not found. (Kh\u00f4ng t\u00ecm th\u1ea5y m\u00e3 LaTeX g\u1ed1c.)")
            Else
                Set selRng = doc.Bookmarks(foundBm).Range
                doc.Bookmarks(foundBm).Delete
                selRng.Delete
                selRng.Text = rawLatex
                If ModMathHelpers.GetDocVariable(doc, foundBm & "_isblock") = "1" Then
                    selRng.ParagraphFormat.TabStops.ClearAll
                End If
                ModMathHelpers.ApplyPastelShading selRng, ModMathScanner.COLOR_PASTEL_AMBER
                doc.Bookmarks.Add "bm_math_src_" & idStr, selRng
            End If
        ElseIf foundKind = ModMathScanner.BM_KIND_SRC Or foundKind = ModMathScanner.BM_KIND_PINK Then
            Set selRng = doc.Bookmarks(foundBm).Range
            rawLatex = selRng.Text
            isBlock = ModMathHelpers.InferIsBlock(rawLatex, doc, idStr)
            If Len(idStr) > 0 And Len(idStr) <= 9 Then
                If ModMathHelpers.IsAllAsciiDigits(idStr) Then
                    seqNum = CLng(idStr)
                    ConvertWithSeq doc, selRng, rawLatex, isBlock, seqNum, foundBm, isPink
                    If isPink Then
                        ModMathHelpers.NotifyUser ModConstants.Uni("Unsupported LaTeX syntax (C\u00fa ph\u00e1p ch\u01b0a h\u1ed7 tr\u1ee3): ") & Left$(rawLatex, 40)
                    End If
                End If
            End If
        End If
    End If
CleanToggleExit:
    On Error Resume Next
    ModBackup.EndUndoSafe objUndo, weStartedUndo
    ModBackup.RestoreTrackChanges doc, wasTracking
    If Len(toggleErr) > 0 Then
        ModMathHelpers.NotifyUser ModConstants.Uni("Toggle Math error (L\u1ed7i \u0111\u1ed5i to\u00e1n): ") & toggleErr
    End If
    Exit Sub
ErrHandler:
    toggleErr = Err.Description
    Resume CleanToggleExit
End Sub
Public Sub FinalizeAllMath(ByVal doc As Document, ByVal isSilent As Boolean)
    Dim bm As Bookmark, v As Variable, objUndo As Object, i As Long, warnCount As Long
    Dim finMsg As String, errMsg As String, weStartedUndo As Boolean, wasTracking As Boolean
    Dim prevScreen As Boolean, prevPagination As Boolean, hasPaginationState As Boolean, hasErr As Boolean
    If g_IsMathBuilding Or ((Not isSilent) And ModMain.g_IsEngineRunning) Or doc Is Nothing Then Exit Sub
    If Not ModBackup.ValidateWritableDocument(doc, errMsg) Then
        If (Not isSilent) And Application.DisplayAlerts <> 0 Then
            ModUI.MsgBoxW errMsg, vbExclamation, ModConstants.Uni("Cannot Edit (Kh\u00f4ng th\u1ec3 ch\u1ec9nh s\u1eeda)")
        End If
        Exit Sub
    End If
    If (Not isSilent) And Application.DisplayAlerts <> 0 Then
        For i = 1 To doc.Bookmarks.Count
            Set bm = doc.Bookmarks(i)
            If Left$(bm.Name, 13) = "bm_math_pink_" Or Left$(bm.Name, 12) = "bm_math_src_" Then
                warnCount = warnCount + 1
            End If
        Next i
        If warnCount > 0 Then
            finMsg = CStr(warnCount) & ModConstants.Uni(" unconverted Pink/Amber formula(s) remain. Finalize anyway?") & vbCrLf & _
                     ModConstants.Uni("C\u00f2n ") & CStr(warnCount) & ModConstants.Uni(" c\u00f4ng th\u1ee9c H\u1ed3ng/H\u1ed5 ph\u00e1ch ch\u01b0a d\u1ef1ng. V\u1eabn ch\u1ed1t to\u00e1n?")
            If ModUI.MsgBoxW(finMsg, vbYesNo + vbExclamation + vbDefaultButton2, ModConstants.Uni("Confirm Finalize Math (X\u00e1c nh\u1eadn ch\u1ed1t to\u00e1n)")) <> vbYes Then
                Exit Sub
            End If
            warnCount = 0
        End If
    End If
    wasTracking = False
    If isSilent Then
        wasTracking = doc.TrackRevisions
        If wasTracking Then doc.TrackRevisions = False
    Else
        If Not ModBackup.HandleTrackChanges(doc, wasTracking) Then Exit Sub
    End If
    prevScreen = Application.ScreenUpdating
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    On Error Resume Next
    prevPagination = Application.Options.Pagination
    If Err.Number = 0 Then
        hasPaginationState = True
        Application.Options.Pagination = False
    End If
    Err.Clear
    On Error GoTo ErrHandler
    If Not isSilent Then
        weStartedUndo = ModBackup.StartUndoSafe("Finalize Math Equations", objUndo)
    End If
    For i = doc.Bookmarks.Count To 1 Step -1
        Set bm = doc.Bookmarks(i)
        If Left$(bm.Name, 8) = "bm_math_" Then
            If Left$(bm.Name, 13) = "bm_math_pink_" Or Left$(bm.Name, 12) = "bm_math_src_" Then
                warnCount = warnCount + 1
            End If
            ModMathHelpers.ApplyPastelShading bm.Range, ModMathScanner.COLOR_AUTO_CLEAR
            bm.Delete
        End If
    Next i
    ModMathHelpers.ClearOrphanedMathShading doc
    If isSilent Then
        For i = doc.Variables.Count To 1 Step -1
            Set v = doc.Variables(i)
            If Left$(v.Name, 8) = "bm_math_" Or v.Name = "mathcfg_seq" Then v.Delete
        Next i
    End If
CleanFinalizeExit:
    On Error Resume Next
    If hasPaginationState Then Application.Options.Pagination = prevPagination
    ModBackup.EndUndoSafe objUndo, weStartedUndo
    ModBackup.RestoreTrackChanges doc, wasTracking
    Application.ScreenUpdating = prevScreen
    If Not hasErr And Not isSilent And Application.DisplayAlerts <> 0 Then
        finMsg = ModConstants.Uni("Math equations finalized. (\u0110\u00e3 ch\u1ed1t xong c\u00f4ng th\u1ee9c to\u00e1n h\u1ecdc.)")
        If warnCount > 0 Then finMsg = finMsg & vbCrLf & vbCrLf & ModConstants.Uni("Note: ") & CStr(warnCount) & ModConstants.Uni(" Pink/Amber item(s) kept as raw text. (L\u01b0u \u00fd: ") & CStr(warnCount) & ModConstants.Uni(" c\u00f4ng th\u1ee9c H\u1ed3ng/H\u1ed5 ph\u00e1ch ch\u01b0a d\u1ef1ng \u0111\u01b0\u1ee3c gi\u1eef d\u1ea1ng th\u00f4.)")
        ModUI.MsgBoxW finMsg, vbInformation, ModConstants.Uni("Done (Ho\u00e0n t\u1ea5t)")
    End If
    Exit Sub
ErrHandler:
    hasErr = True
    Resume CleanFinalizeExit
End Sub
Public Function NavigateNextPinkMath(ByVal doc As Document) As Long
    Dim bm As Bookmark, i As Long, j As Long, k As Long, cnt As Long, targetIdx As Long
    Dim curEnd As Long, keyStart As Long, keyName As String
    Dim starts() As Long, names() As String
    If doc Is Nothing Then Exit Function
    If ModMain.g_IsEngineRunning Or g_IsMathBuilding Then Exit Function
    If doc.Bookmarks.Count > 0 Then
        ReDim starts(1 To doc.Bookmarks.Count)
        ReDim names(1 To doc.Bookmarks.Count)
        For i = 1 To doc.Bookmarks.Count
            Set bm = doc.Bookmarks(i)
            If Left$(bm.Name, 13) = "bm_math_pink_" Then
                If bm.StoryType = 1 Then
                    cnt = cnt + 1
                    starts(cnt) = bm.Range.Start
                    names(cnt) = bm.Name
                End If
            End If
        Next i
    End If
    If cnt = 0 Then
        If Application.DisplayAlerts <> 0 Then
            Application.StatusBar = "AcademicFormatter: No Pink math items found."
        End If
        NavigateNextPinkMath = 0
        Exit Function
    End If
    For i = 2 To cnt
        keyStart = starts(i)
        keyName = names(i)
        j = i - 1
        Do While j >= 1
            If starts(j) <= keyStart Then Exit Do
            starts(j + 1) = starts(j)
            names(j + 1) = names(j)
            j = j - 1
        Loop
        starts(j + 1) = keyStart
        names(j + 1) = keyName
    Next i
    curEnd = -1
    On Error Resume Next
    If Selection.Document Is doc Then
        curEnd = Selection.End
    End If
    Err.Clear
    On Error GoTo 0
    targetIdx = 1
    For k = 1 To cnt
        If starts(k) > curEnd Then
            targetIdx = k
            Exit For
        End If
    Next k
    doc.Bookmarks(names(targetIdx)).Range.Select
    On Error Resume Next
    doc.ActiveWindow.ScrollIntoView Selection.Range, True
    Err.Clear
    On Error GoTo 0
    Application.StatusBar = "Math Review " & CStr(targetIdx) & "/" & CStr(cnt) & " (Pink)"
    NavigateNextPinkMath = targetIdx
End Function
Public Sub Ribbon_OnMathBuild(ByVal control As IRibbonControl)
    If Documents.Count = 0 Then
        If Application.DisplayAlerts <> 0 Then ModUI.MsgBoxW ModConstants.Uni("No open Word document found. (Kh\u00f4ng t\u00ecm th\u1ea5y t\u00e0i li\u1ec7u Word \u0111ang m\u1edf.)"), vbExclamation, ModConstants.Uni("Error (L\u1ed7i)")
        Exit Sub
    End If
    BuildAllMathInDocument ActiveDocument, True
End Sub
Public Sub Ribbon_OnMathToggle(ByVal control As IRibbonControl)
    If Documents.Count > 0 Then ToggleMathAtSelection ActiveDocument
End Sub
Public Sub Ribbon_OnMathFinalize(ByVal control As IRibbonControl)
    If Documents.Count > 0 Then FinalizeAllMath ActiveDocument, False
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
