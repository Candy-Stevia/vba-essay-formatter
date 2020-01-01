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
        InsertOpcAndShade = True
    End If
    Exit Function
FailInsert:
    Err.Clear
End Function
Private Function ConvertWithSeq(ByVal doc As Document, ByVal targetRng As Range, ByVal rawLatex As String, ByVal isBlock As Boolean, ByRef ioSeqId As Long, ByVal prevBm As String, ByRef outIsPink As Boolean) As Boolean
    Dim cleanLatex As String
    Dim eqTag As String
    Dim ommlXml As String
    Dim bmID As String
    Dim isAmb As Boolean
    cleanLatex = ModMathScanner.StripOuterDelims(rawLatex)
    If Len(cleanLatex) = 0 Or cleanLatex = "$" Then
        If Len(prevBm) = 0 Then
            Exit Function
        End If
        isAmb = True
    End If
    If ioSeqId <= 0 Then
        ioSeqId = ModMathHelpers.NextMathSeq(doc)
    End If
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
            If Len(prevBm) > 0 Then
                If doc.Bookmarks.Exists(prevBm) Then
                    doc.Bookmarks(prevBm).Delete
                End If
            End If
            ModMathHelpers.SetDocVariable doc, bmID, rawLatex
            ModMathHelpers.SetDocVariable doc, bmID & "_isblock", IIf(isBlock, "1", "0")
            outIsPink = False
            ConvertWithSeq = True
            Exit Function
        End If
    End If
    If Len(prevBm) = 0 Or Left$(prevBm, 12) = "bm_math_src_" Or Not doc.Bookmarks.Exists("bm_math_pink_" & Format$(ioSeqId, "00000")) Then
        If Len(prevBm) > 0 Then
            If doc.Bookmarks.Exists(prevBm) Then
                doc.Bookmarks(prevBm).Delete
            End If
        End If
        ModMathHelpers.MarkSpanPink doc, targetRng, ioSeqId, rawLatex
    End If
End Function
Public Function ConvertRangeToMath(ByVal doc As Document, ByVal targetRng As Range, ByVal rawLatex As String, ByVal isBlock As Boolean, ByRef outSeqId As Long, ByRef outIsPink As Boolean) As Boolean
    outSeqId = 0
    ConvertRangeToMath = ConvertWithSeq(doc, targetRng, rawLatex, isBlock, outSeqId, "", outIsPink)
End Function
Public Sub BuildAllMathInDocument(ByVal doc As Document, Optional ByVal enforceSafety As Boolean = False)
    Dim pairCount As Long, k As Long, dummySeq As Long, totalParas As Long, pStep As Long, docEnd As Long, prevViewType As Long
    Dim builtCount As Long, pinkCount As Long, prevFindEnd As Long
    Dim isPink As Boolean, isScanningActive As Boolean, isAmb As Boolean, weStartedUndo As Boolean
    Dim prevScreen As Boolean, prevPagination As Boolean, hasPaginationState As Boolean, hasWindowView As Boolean
    Dim hasErr As Boolean, wasTracking As Boolean, backupPath As String
    Dim pStarts() As Long, pEnds() As Long, clsArr() As Long
    Dim rawLatex As String, rawPText As String, pText As String, innerT As String, cAfter As String
    Dim candRng As Range, fRng As Range, p As Paragraph, nextP As Paragraph, endPara As Paragraph, objUndo As Object
    If g_IsMathBuilding Or doc Is Nothing Then Exit Sub
    If enforceSafety Then
        If Not ModBackup.BeginProtectedSession(doc, backupPath, wasTracking) Then Exit Sub
    End If
    ModConstants.g_SilentErrors = 0
    ModRibbonCallbacks.g_LastError = ""
    prevScreen = Application.ScreenUpdating
    On Error GoTo ErrHandler
    g_IsMathBuilding = True
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
        If fRng.End <= prevFindEnd Then
            Exit Do
        End If
        prevFindEnd = fRng.End
        pText = LCase$(Trim$(Replace(Replace(Replace(fRng.Paragraphs(1).Range.Text, vbCr, ""), Chr(11), " "), Chr(7), "")))
        If pText = "=contentstart()" Then
            Set p = fRng.Paragraphs(1).Next
            Exit Do
        End If
        If fRng.End >= docEnd Then
            Exit Do
        End If
        fRng.SetRange fRng.End, docEnd
    Loop
    isScanningActive = True
    Do While Not p Is Nothing
        pStep = pStep + 1
        If pStep <= totalParas Then
            ModProgress.ProgressUpdate pStep, totalParas
        End If
        Set nextP = p.Next
        rawPText = p.Range.Text
        pText = LCase$(Trim$(Replace(Replace(Replace(rawPText, vbCr, ""), Chr(11), " "), Chr(7), "")))
        If pText = "=contentend()" Then
            isScanningActive = False
        ElseIf Not ModMathScanner.IsInNoFlyZoneCached(p.Range, rawPText, isScanningActive) Then
            If InStr(1, rawPText, "$") > 0 Or InStr(1, rawPText, "\") > 0 Then
                If ModMathScanner.CheckMultilineBlock(doc, p, endPara, rawLatex, isAmb) Then
                    Set nextP = endPara.Next
                    CountMathOutcome ConvertRangeToMath(doc, doc.Range(p.Range.Start, endPara.Range.End), rawLatex, True, dummySeq, isPink), isPink, builtCount, pinkCount
                Else
                    pText = Trim$(Replace(Replace(Replace(rawPText, vbCr, ""), Chr(11), " "), Chr(7), ""))
                    If ModMathInline.IsSingleLineDisplayBlock(pText) Then
                        CountMathOutcome ConvertRangeToMath(doc, p.Range, pText, True, dummySeq, isPink), isPink, builtCount, pinkCount
                    ElseIf pText <> "$$" And pText <> "\[" And pText <> "\]" And pText <> "$$$" And pText <> "$$$$" And pText <> "\[\]" Then
                        ModMathInline.FindInlinePairs doc, p, pStarts, pEnds, pairCount, rawPText
                        If pairCount > 0 Then
                            ReDim clsArr(1 To pairCount)
                            For k = 1 To pairCount
                                Set candRng = doc.Range(pStarts(k), pEnds(k))
                                If Len(candRng.Text) >= 2 Then
                                    If Left$(candRng.Text, 2) = "\(" Then
                                        clsArr(k) = 1
                                    Else
                                        innerT = Mid$(candRng.Text, 2, Len(candRng.Text) - 2)
                                        cAfter = ""
                                        If candRng.End < p.Range.End Then
                                            cAfter = doc.Range(candRng.End, candRng.End + 1).Text
                                        End If
                                        clsArr(k) = ModMathScanner.ClassifyDollarSpan(innerT, "", cAfter)
                                    End If
                                End If
                            Next k
                            For k = pairCount To 1 Step -1
                                If clsArr(k) = 1 Or clsArr(k) = 2 Then
                                    Set candRng = doc.Range(pStarts(k), pEnds(k))
                                    If clsArr(k) = 1 Then
                                        CountMathOutcome ConvertRangeToMath(doc, candRng, candRng.Text, False, dummySeq, isPink), isPink, builtCount, pinkCount
                                    ElseIf Not ModMathInline.HasForbiddenDomChar(candRng.Text) Then
                                        dummySeq = ModMathHelpers.NextMathSeq(doc)
                                        ModMathHelpers.MarkSpanPink doc, candRng, dummySeq, candRng.Text
                                        pinkCount = pinkCount + 1
                                    End If
                                End If
                            Next k
                        End If
                    End If
                End If
            End If
        End If
        Set p = nextP
    Loop
CleanExit:
    On Error Resume Next
    ModMathHelpers.FlushMathSeqCache doc
    ModMathHelpers.ResetMathSeqCache
    ModMathScanner.ResetNoFlyZoneCache
    ModProgress.ProgressFinish
    If hasWindowView And doc.Windows.Count > 0 Then doc.ActiveWindow.View.Type = prevViewType
    If hasPaginationState Then Application.Options.Pagination = prevPagination
    ModBackup.EndUndoSafe objUndo, weStartedUndo
    Application.ScreenUpdating = prevScreen
    If enforceSafety Then ModBackup.EndProtectedSession doc, backupPath, wasTracking, hasErr
    g_IsMathBuilding = False
    If hasErr And Application.DisplayAlerts <> 0 Then
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
Public Sub ToggleMathAtSelection(ByVal doc As Document)
    Dim selRng As Range
    Dim foundBm As String
    Dim rawLatex As String
    Dim idStr As String
    Dim foundKind As Long
    Dim seqNum As Long
    Dim isBlock As Boolean
    Dim isPink As Boolean
    Dim weStartedUndo As Boolean
    Dim objUndo As Object
    If g_IsMathBuilding Or doc Is Nothing Then
        Exit Sub
    End If
    If Selection.StoryType <> 1 Then
        Exit Sub
    End If
    On Error GoTo ErrHandler
    weStartedUndo = ModBackup.StartUndoSafe("Toggle LaTeX Math", objUndo)
    If Not ModMathHelpers.FindEnclosingMathBookmark(doc, Selection.Start, Selection.End, foundBm, foundKind) Then
        If Selection.End > Selection.Start Then
            Set selRng = Selection.Range
            ModMathInline.TrimMathRange selRng
            If selRng.End > selRng.Start Then
                rawLatex = selRng.Text
                ConvertRangeToMath doc, selRng, rawLatex, ModMathHelpers.InferIsBlock(rawLatex, doc, ""), seqNum, isPink
            End If
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
                selRng.ParagraphFormat.TabStops.ClearAll
                ModMathHelpers.ApplyPastelShading selRng, ModMathScanner.COLOR_PASTEL_AMBER
                doc.Bookmarks.Add "bm_math_src_" & idStr, selRng
            End If
        ElseIf foundKind = ModMathScanner.BM_KIND_SRC Or foundKind = ModMathScanner.BM_KIND_PINK Then
            Set selRng = doc.Bookmarks(foundBm).Range
            rawLatex = selRng.Text
            isBlock = ModMathHelpers.InferIsBlock(rawLatex, doc, idStr)
            If IsNumeric(idStr) Then
                seqNum = CLng(idStr)
                ConvertWithSeq doc, selRng, rawLatex, isBlock, seqNum, foundBm, isPink
                If isPink Then
                    ModMathHelpers.NotifyUser ModConstants.Uni("Unsupported LaTeX syntax (C\u00fa ph\u00e1p ch\u01b0a h\u1ed7 tr\u1ee3): ") & Left$(rawLatex, 40)
                End If
            End If
        End If
    End If
CleanToggleExit:
    ModBackup.EndUndoSafe objUndo, weStartedUndo
    Exit Sub
ErrHandler:
    Resume CleanToggleExit
End Sub
Public Sub FinalizeAllMath(ByVal doc As Document, ByVal isSilent As Boolean)
    Dim bm As Bookmark, v As Variable, objUndo As Object, i As Long, warnCount As Long, finMsg As String, weStartedUndo As Boolean
    If g_IsMathBuilding Or doc Is Nothing Then Exit Sub
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
    If isSilent Then
        For i = doc.Variables.Count To 1 Step -1
            Set v = doc.Variables(i)
            If Left$(v.Name, 8) = "bm_math_" Or v.Name = "mathcfg_seq" Then
                v.Delete
            End If
        Next i
    End If
    ModBackup.EndUndoSafe objUndo, weStartedUndo
    If Not isSilent And Application.DisplayAlerts <> 0 Then
        finMsg = ModConstants.Uni("Math equations finalized. (\u0110\u00e3 ch\u1ed1t xong c\u00f4ng th\u1ee9c to\u00e1n h\u1ecdc.)")
        If warnCount > 0 Then
            finMsg = finMsg & vbCrLf & vbCrLf & ModConstants.Uni("Note: ") & CStr(warnCount) & ModConstants.Uni(" Pink/Amber item(s) kept as raw text. (L\u01b0u \u00fd: ") & CStr(warnCount) & ModConstants.Uni(" c\u00f4ng th\u1ee9c H\u1ed3ng/H\u1ed5 ph\u00e1ch ch\u01b0a d\u1ef1ng \u0111\u01b0\u1ee3c gi\u1eef d\u1ea1ng th\u00f4.)")
        End If
        ModUI.MsgBoxW finMsg, vbInformation, ModConstants.Uni("Done (Ho\u00e0n t\u1ea5t)")
    End If
    Exit Sub
ErrHandler:
    ModBackup.EndUndoSafe objUndo, weStartedUndo
End Sub
Public Sub Ribbon_OnMathBuild(ByVal control As IRibbonControl)
    If Documents.Count = 0 Then
        If Application.DisplayAlerts <> 0 Then ModUI.MsgBoxW ModConstants.Uni("No open Word document found. (Kh\u00f4ng t\u00ecm th\u1ea5y t\u00e0i li\u1ec7u Word \u0111ang m\u1edf.)"), vbExclamation, ModConstants.Uni("Error (L\u1ed7i)")
        Exit Sub
    End If
    BuildAllMathInDocument ActiveDocument, True
End Sub
Public Sub Ribbon_OnMathToggle(ByVal control As IRibbonControl)
    If Documents.Count > 0 Then
        ToggleMathAtSelection ActiveDocument
    End If
End Sub
Public Sub Ribbon_OnMathFinalize(ByVal control As IRibbonControl)
    If Documents.Count > 0 Then
        FinalizeAllMath ActiveDocument, False
    End If
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
