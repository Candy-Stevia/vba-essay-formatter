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
Private Function StartUndoSafe(ByVal recName As String, ByRef outUndo As Object) As Boolean
    On Error Resume Next
    Err.Clear
    Set outUndo = Application.UndoRecord
    If Err.Number = 0 And Not outUndo Is Nothing Then
        If Not outUndo.IsRecordingCustomRecord Then
            Err.Clear
            outUndo.StartCustomRecord recName
            StartUndoSafe = (Err.Number = 0)
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function
Private Sub EndUndoSafe(ByRef ioUndo As Object, ByRef ioStarted As Boolean)
    On Error Resume Next
    If ioStarted Then
        If Not ioUndo Is Nothing Then
            If ioUndo.IsRecordingCustomRecord Then
                ioUndo.EndCustomRecord
            End If
        End If
        ioStarted = False
    End If
    Err.Clear
End Sub
Private Function InsertOpcAndShade(ByVal doc As Document, ByVal targetRng As Range, ByVal ommlXml As String, ByVal isBlock As Boolean, ByVal bmID As String) As Boolean
    Dim opcXml As String
    Dim om As OMath
    opcXml = ModOmmlEmit.WrapOmmlInFlatOpc(ommlXml)
    ModMathInline.TrimTrailingDomMarks targetRng
    On Error GoTo FailInsert
    targetRng.InsertXML opcXml
    If targetRng.OMaths.Count > 0 Then
        Set om = targetRng.OMaths(1)
        If isBlock Then
            om.Type = 0
            om.Range.ParagraphFormat.Alignment = 1
            om.Range.ParagraphFormat.FirstLineIndent = 0
            om.Range.ParagraphFormat.LeftIndent = 0
            ModMathHelpers.ApplyPastelShading om.Range, ModMathScanner.COLOR_PASTEL_MINT
        Else
            om.Type = 1
            ModMathHelpers.ApplyPastelShading om.Range, ModMathScanner.COLOR_PASTEL_CYAN
        End If
        doc.Bookmarks.Add bmID, om.Range
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
        If InsertOpcAndShade(doc, targetRng, ommlXml, isBlock, bmID) Then
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
Public Sub BuildAllMathInDocument(ByVal doc As Document)
    Dim pairCount As Long
    Dim k As Long
    Dim dummySeq As Long
    Dim totalParas As Long
    Dim pStep As Long
    Dim docEnd As Long
    Dim prevViewType As Long
    Dim isPink As Boolean
    Dim isScanningActive As Boolean
    Dim isAmb As Boolean
    Dim weStartedUndo As Boolean
    Dim prevScreen As Boolean
    Dim prevPagination As Boolean
    Dim hasPaginationState As Boolean
    Dim hasWindowView As Boolean
    Dim pStarts() As Long
    Dim pEnds() As Long
    Dim clsArr() As Long
    Dim rawLatex As String
    Dim rawPText As String
    Dim pText As String
    Dim innerT As String
    Dim cAfter As String
    Dim candRng As Range
    Dim fRng As Range
    Dim p As Paragraph
    Dim nextP As Paragraph
    Dim endPara As Paragraph
    Dim objUndo As Object
    If g_IsMathBuilding Or doc Is Nothing Then
        Exit Sub
    End If
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
    weStartedUndo = StartUndoSafe("Build LaTeX Math (Review Mode)", objUndo)
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
    Dim prevFindEnd As Long
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
            Exit Do
        End If
        If Not ModMathScanner.IsInNoFlyZoneCached(p.Range, rawPText, isScanningActive) Then
            If InStr(1, rawPText, "$") > 0 Or InStr(1, rawPText, "\") > 0 Then
                If ModMathScanner.CheckMultilineBlock(doc, p, endPara, rawLatex, isAmb) Then
                    Set nextP = endPara.Next
                    ConvertRangeToMath doc, doc.Range(p.Range.Start, endPara.Range.End), rawLatex, True, dummySeq, isPink
                Else
                    pText = Trim$(Replace(Replace(Replace(rawPText, vbCr, ""), Chr(11), " "), Chr(7), ""))
                    If ModMathInline.IsSingleLineDisplayBlock(pText) Then
                        ConvertRangeToMath doc, p.Range, pText, True, dummySeq, isPink
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
                                        ConvertRangeToMath doc, candRng, candRng.Text, False, dummySeq, isPink
                                    ElseIf Not ModMathInline.HasForbiddenDomChar(candRng.Text) Then
                                        dummySeq = ModMathHelpers.NextMathSeq(doc)
                                        ModMathHelpers.MarkSpanPink doc, candRng, dummySeq, candRng.Text
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
    If hasWindowView And doc.Windows.Count > 0 Then
        doc.ActiveWindow.View.Type = prevViewType
    End If
    If hasPaginationState Then
        Application.Options.Pagination = prevPagination
    End If
    EndUndoSafe objUndo, weStartedUndo
    Application.ScreenUpdating = prevScreen
    g_IsMathBuilding = False
    Exit Sub
ErrHandler:
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
    weStartedUndo = StartUndoSafe("Toggle LaTeX Math", objUndo)
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
    EndUndoSafe objUndo, weStartedUndo
    Exit Sub
ErrHandler:
    Resume CleanToggleExit
End Sub
Public Sub FinalizeAllMath(ByVal doc As Document, ByVal isSilent As Boolean)
    Dim bm As Bookmark
    Dim v As Variable
    Dim objUndo As Object
    Dim i As Long
    Dim warnCount As Long
    Dim finMsg As String
    Dim weStartedUndo As Boolean
    If g_IsMathBuilding Or doc Is Nothing Then
        Exit Sub
    End If
    On Error GoTo ErrHandler
    If Not isSilent Then
        weStartedUndo = StartUndoSafe("Finalize Math Equations", objUndo)
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
    EndUndoSafe objUndo, weStartedUndo
    If Not isSilent And Application.DisplayAlerts <> 0 Then
        finMsg = ModConstants.Uni("Math equations finalized. (\u0110\u00e3 ch\u1ed1t xong c\u00f4ng th\u1ee9c to\u00e1n h\u1ecdc.)")
        If warnCount > 0 Then
            finMsg = finMsg & vbCrLf & vbCrLf & ModConstants.Uni("Note: ") & CStr(warnCount) & ModConstants.Uni(" Pink/Amber item(s) kept as raw text. (L\u01b0u \u00fd: ") & CStr(warnCount) & ModConstants.Uni(" c\u00f4ng th\u1ee9c H\u1ed3ng/H\u1ed5 ph\u00e1ch ch\u01b0a d\u1ef1ng \u0111\u01b0\u1ee3c gi\u1eef d\u1ea1ng th\u00f4.)")
        End If
        ModUI.MsgBoxW finMsg, vbInformation, ModConstants.Uni("Done (Ho\u00e0n t\u1ea5t)")
    End If
    Exit Sub
ErrHandler:
    EndUndoSafe objUndo, weStartedUndo
End Sub
Public Sub Ribbon_OnMathBuild(ByVal control As IRibbonControl)
    If Documents.Count > 0 Then
        BuildAllMathInDocument ActiveDocument
    End If
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
