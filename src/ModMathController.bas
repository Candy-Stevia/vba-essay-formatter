Attribute VB_Name = "ModMathController"
Option Explicit

' ModMathController: Math conversion, Shading, Ribbon, Undo (vba_lint compliant)

Private g_IsMathBuilding As Boolean
Public Function NextMathSeq(ByVal doc As Document) As Long
    NextMathSeq = ModMathHelpers.NextMathSeq(doc)
End Function
Public Sub SetDocVariable(ByVal doc As Document, ByVal varName As String, ByVal varVal As String)
    ModMathHelpers.SetDocVariable doc, varName, varVal
End Sub
Private Sub InsertOpcAndShade(ByVal doc As Document, ByVal targetRng As Range, ByVal ommlXml As String, ByVal isBlock As Boolean, ByVal bmID As String)
    Dim opcXml As String
    Dim om As OMath
    opcXml = ModOmmlEmit.WrapOmmlInFlatOpc(ommlXml)
    Do While targetRng.End > targetRng.Start
        If Right$(targetRng.Text, 1) = vbCr Or Right$(targetRng.Text, 1) = Chr(7) Then
            targetRng.End = targetRng.End - 1
        Else
            Exit Do
        End If
    Loop
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
    End If
End Sub

Public Function ConvertRangeToMath(ByVal doc As Document, ByVal targetRng As Range, ByVal rawLatex As String, ByVal isBlock As Boolean, ByRef outSeqId As Long, ByRef outIsPink As Boolean) As Boolean
    Dim cleanLatex As String
    Dim eqTag As String
    Dim ommlXml As String
    Dim bmID As String
    Dim isAmb As Boolean
    Dim seq As Long

    seq = ModMathHelpers.NextMathSeq(doc)
    outSeqId = seq
    cleanLatex = ModMathScanner.StripOuterDelims(rawLatex)
    ommlXml = ModOmmlEmit.TranspileLatexToOmml(cleanLatex, isBlock, isAmb, eqTag)
    bmID = "bm_math_" & Format$(seq, "00000")
    ModMathHelpers.SetDocVariable doc, bmID, rawLatex
    If isBlock Then
        ModMathHelpers.SetDocVariable doc, bmID & "_isblock", "1"
    Else
        ModMathHelpers.SetDocVariable doc, bmID & "_isblock", "0"
    End If
    If isAmb Or Len(ommlXml) = 0 Then
        outIsPink = True
        ModMathHelpers.ApplyPastelShading targetRng, ModMathScanner.COLOR_PASTEL_PINK
        doc.Bookmarks.Add "bm_math_pink_" & Format$(seq, "00000"), targetRng
    Else
        outIsPink = False
        InsertOpcAndShade doc, targetRng, ommlXml, isBlock, bmID
        ConvertRangeToMath = True
    End If
End Function

Public Sub BuildAllMathInDocument(ByVal doc As Document)
    Dim startP As Long
    Dim pIdx As Long
    Dim endPIdx As Long
    Dim pairCount As Long
    Dim k As Long
    Dim cls As Long
    Dim dummySeq As Long
    Dim totalParas As Long
    Dim prevViewType As Long
    Dim isPink As Boolean
    Dim isScanningActive As Boolean
    Dim isAmb As Boolean
    Dim weStartedUndo As Boolean
    Dim prevScreen As Boolean
    Dim screenDisabled As Boolean
    Dim prevPagination As Boolean
    Dim hasPaginationState As Boolean
    Dim hasWindowView As Boolean
    Dim pStarts() As Long
    Dim pEnds() As Long
    Dim rawLatex As String
    Dim rawPText As String
    Dim pText As String
    Dim innerT As String
    Dim cBefore As String
    Dim cAfter As String
    Dim candRng As Range
    Dim p As Paragraph
    Dim objUndo As Object

    If g_IsMathBuilding Then
        Exit Sub
    End If
    g_IsMathBuilding = True
    weStartedUndo = False
    screenDisabled = False
    hasPaginationState = False
    hasWindowView = False
    prevScreen = Application.ScreenUpdating
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    screenDisabled = True
    On Error Resume Next
    prevPagination = Application.Options.Pagination
    If Err.Number = 0 Then
        hasPaginationState = True
        Application.Options.Pagination = False
    Else
        Err.Clear
    End If
    If doc.Windows.Count > 0 Then
        If Not doc.ActiveWindow Is Nothing Then
            prevViewType = doc.ActiveWindow.View.Type
            hasWindowView = True
            doc.ActiveWindow.View.Type = 1
        End If
    End If
    Set objUndo = Application.UndoRecord
    If Not objUndo Is Nothing Then
        If Not objUndo.IsRecordingCustomRecord Then
            objUndo.StartCustomRecord "Build LaTeX Math (Review Mode)"
            weStartedUndo = (Err.Number = 0)
        End If
    End If
    On Error GoTo ErrHandler
    ModProgress.ProgressInit "LaTeX Math Engine"
    ModProgress.ProgressSetPhase ModConstants.Uni("Converting Math (Chuy\u1ec3n \u0111\u1ed5i to\u00e1n)"), 5#, 95#
    ModMathHelpers.InitMathSeqCache doc
    ModMathScanner.InitNoFlyZoneCache doc
    totalParas = doc.Paragraphs.Count
    startP = 1
    For pIdx = 1 To totalParas
        pText = LCase$(Trim$(Replace(Replace(doc.Paragraphs(pIdx).Range.Text, vbCr, ""), Chr(11), "")))
        If pText = "=contentstart()" Then
            startP = pIdx + 1
            Exit For
        End If
    Next pIdx
    isScanningActive = True
    pIdx = startP
    Do While pIdx <= doc.Paragraphs.Count
        ModProgress.ProgressUpdate pIdx, doc.Paragraphs.Count
        Set p = doc.Paragraphs(pIdx)
        rawPText = p.Range.Text
        pText = LCase$(Trim$(Replace(Replace(rawPText, vbCr, ""), Chr(11), "")))
        If pText = "=contentend()" Then
            Exit Do
        End If
        If Not ModMathScanner.IsInNoFlyZoneCached(p.Range, rawPText, isScanningActive) Then
            If InStr(1, rawPText, "$") > 0 Or InStr(1, rawPText, "\") > 0 Then
                If ModMathScanner.CheckMultilineBlock(doc, pIdx, endPIdx, rawLatex, isAmb) Then
                    ConvertRangeToMath doc, doc.Range(doc.Paragraphs(pIdx).Range.Start, doc.Paragraphs(endPIdx).Range.End), rawLatex, True, dummySeq, isPink
                    If isPink Then
                        pIdx = endPIdx
                    End If
                Else
                    pText = Trim$(Replace(Replace(Replace(rawPText, vbCr, ""), Chr(11), ""), Chr(7), ""))
                    If pText = "$$" Or pText = "\[" Or pText = "\]" Then
                        ' A delimiter on its own must never be treated as an empty formula.
                    ElseIf (Left$(pText, 2) = "$$" And Right$(pText, 2) = "$$") Or (Left$(pText, 2) = "\[" And Right$(pText, 2) = "\]") Or (Left$(pText, 6) = "\begin" And InStr(1, pText, "\end{") > 0 And Right$(pText, 1) = "}") Then
                        ConvertRangeToMath doc, p.Range, pText, True, dummySeq, isPink
                    Else
                        ModMathInline.FindInlinePairs doc, p, pStarts, pEnds, pairCount, rawPText
                        If pairCount > 0 Then
                            For k = pairCount To 1 Step -1
                                Set candRng = doc.Range(pStarts(k), pEnds(k))
                                If Len(candRng.Text) >= 2 Then
                                    If Left$(candRng.Text, 2) = "\(" Then
                                        cls = 1
                                    Else
                                        innerT = Mid$(candRng.Text, 2, Len(candRng.Text) - 2)
                                        cBefore = ""
                                        If candRng.Start > p.Range.Start Then
                                            cBefore = doc.Range(candRng.Start - 1, candRng.Start).Text
                                        End If
                                        cAfter = ""
                                        If candRng.End < p.Range.End Then
                                            cAfter = doc.Range(candRng.End, candRng.End + 1).Text
                                        End If
                                        cls = ModMathScanner.ClassifyDollarSpan(innerT, cBefore, cAfter)
                                    End If
                                    If cls = 1 Then
                                        ConvertRangeToMath doc, candRng, candRng.Text, False, dummySeq, isPink
                                    ElseIf cls = 2 Then
                                        dummySeq = ModMathHelpers.NextMathSeq(doc)
                                        ModMathHelpers.ApplyPastelShading candRng, ModMathScanner.COLOR_PASTEL_PINK
                                        doc.Bookmarks.Add "bm_math_pink_" & Format$(dummySeq, "00000"), candRng
                                        ModMathHelpers.SetDocVariable doc, "bm_math_" & Format$(dummySeq, "00000"), candRng.Text
                                    End If
                                End If
                            Next k
                        End If
                    End If
                End If
            End If
        End If
        pIdx = pIdx + 1
    Loop

CleanExit:
    On Error Resume Next
    ModMathHelpers.FlushMathSeqCache doc
    ModMathHelpers.ResetMathSeqCache
    ModMathScanner.ResetNoFlyZoneCache
    ModProgress.ProgressFinish
    If hasWindowView Then
        doc.ActiveWindow.View.Type = prevViewType
    End If
    If hasPaginationState Then
        Application.Options.Pagination = prevPagination
    End If
    If weStartedUndo Then
        If Not objUndo Is Nothing Then
            If objUndo.IsRecordingCustomRecord Then
                objUndo.EndCustomRecord
            End If
        End If
        weStartedUndo = False
    End If
    If screenDisabled Then
        Application.ScreenUpdating = prevScreen
    End If
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
    Dim cleanLatex As String
    Dim eqTag As String
    Dim ommlXml As String
    Dim foundKind As Long
    Dim dummySeq As Long
    Dim isBlock As Boolean
    Dim isAmb As Boolean
    Dim isPink As Boolean

    On Error GoTo ErrHandler
    If Not ModMathHelpers.FindEnclosingMathBookmark(doc, Selection.Start, Selection.End, foundBm, foundKind) Then
        If Selection.End > Selection.Start Then
            Set selRng = Selection.Range
            ModMathInline.TrimMathRange selRng
            If selRng.End > selRng.Start Then
                rawLatex = selRng.Text
                If Left$(rawLatex, 2) = "$$" Or Left$(rawLatex, 2) = "\[" Or (Left$(rawLatex, 6) = "\begin" And Left$(rawLatex, 2) <> "\(") Then
                    isBlock = True
                Else
                    isBlock = False
                End If
                ConvertRangeToMath doc, selRng, rawLatex, isBlock, dummySeq, isPink
            End If
        End If
        Exit Sub
    End If

    idStr = Mid$(foundBm, InStrRev(foundBm, "_") + 1)
    If foundKind = ModMathScanner.BM_KIND_OMATH Then
        rawLatex = ModMathHelpers.GetDocVariable(doc, foundBm)
        If Len(rawLatex) = 0 Then
            ModMathHelpers.NotifyUser ModConstants.Uni("Original LaTeX source not found. (Kh\u00f4ng t\u00ecm th\u1ea5y m\u00e3 LaTeX g\u1ed1c.)")
            Exit Sub
        End If
        Set selRng = doc.Bookmarks(foundBm).Range
        doc.Bookmarks(foundBm).Delete
        selRng.Delete
        selRng.Text = rawLatex
        ModMathHelpers.ApplyPastelShading selRng, ModMathScanner.COLOR_PASTEL_AMBER
        doc.Bookmarks.Add "bm_math_src_" & idStr, selRng
        Exit Sub
    End If

    If foundKind = ModMathScanner.BM_KIND_SRC Or foundKind = ModMathScanner.BM_KIND_PINK Then
        Set selRng = doc.Bookmarks(foundBm).Range
        rawLatex = selRng.Text
        cleanLatex = ModMathScanner.StripOuterDelims(rawLatex)
        If ModMathHelpers.GetDocVariable(doc, "bm_math_" & idStr & "_isblock") = "1" Then
            isBlock = True
        ElseIf ModMathHelpers.GetDocVariable(doc, "bm_math_" & idStr & "_isblock") = "0" Then
            isBlock = False
        Else
            isBlock = (Left$(rawLatex, 2) = "$$" Or Left$(rawLatex, 2) = "\[" Or (Left$(rawLatex, 6) = "\begin" And Left$(rawLatex, 2) <> "\("))
        End If
        ommlXml = ModOmmlEmit.TranspileLatexToOmml(cleanLatex, isBlock, isAmb, eqTag)
        If isAmb Or Len(ommlXml) = 0 Then
            If foundKind = ModMathScanner.BM_KIND_SRC Then
                doc.Bookmarks(foundBm).Delete
                ModMathHelpers.ApplyPastelShading selRng, ModMathScanner.COLOR_PASTEL_PINK
                doc.Bookmarks.Add "bm_math_pink_" & idStr, selRng
                ModMathHelpers.SetDocVariable doc, "bm_math_" & idStr, rawLatex
            End If
            ModMathHelpers.NotifyUser ModConstants.Uni("Unsupported LaTeX syntax (C\u00fa ph\u00e1p ch\u01b0a h\u1ed7 tr\u1ee3): ") & Left$(rawLatex, 40)
            Exit Sub
        End If
        doc.Bookmarks(foundBm).Delete
        InsertOpcAndShade doc, selRng, ommlXml, isBlock, "bm_math_" & idStr
        ModMathHelpers.SetDocVariable doc, "bm_math_" & idStr, rawLatex
        If isBlock Then
            ModMathHelpers.SetDocVariable doc, "bm_math_" & idStr & "_isblock", "1"
        Else
            ModMathHelpers.SetDocVariable doc, "bm_math_" & idStr & "_isblock", "0"
        End If
        Exit Sub
    End If

ErrHandler:
End Sub

Public Sub FinalizeAllMath(ByVal doc As Document, ByVal isSilent As Boolean)
    Dim bm As Bookmark
    Dim v As Variable
    Dim objUndo As Object
    Dim i As Long
    Dim warnCount As Long
    Dim finMsg As String
    Dim weStartedUndo As Boolean
    On Error GoTo ErrHandler
    weStartedUndo = False
    warnCount = 0
    If Not isSilent Then
        On Error Resume Next
        Set objUndo = Application.UndoRecord
        If Not objUndo Is Nothing Then
            If Not objUndo.IsRecordingCustomRecord Then
                objUndo.StartCustomRecord "Finalize Math Equations"
                weStartedUndo = (Err.Number = 0)
            End If
        End If
        On Error GoTo ErrHandler
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
    If weStartedUndo Then
        objUndo.EndCustomRecord
        weStartedUndo = False
    End If
    If Not isSilent And Application.DisplayAlerts <> 0 Then
        finMsg = ModConstants.Uni("Math equations finalized. (\u0110\u00e3 ch\u1ed1t xong c\u00f4ng th\u1ee9c to\u00e1n h\u1ecdc.)")
        If warnCount > 0 Then
            finMsg = finMsg & vbCrLf & vbCrLf & ModConstants.Uni("Note: ") & CStr(warnCount) & ModConstants.Uni(" Pink/Amber item(s) kept as raw text. (L\u01b0u \u00fd: ") & CStr(warnCount) & ModConstants.Uni(" c\u00f4ng th\u1ee9c H\u1ed3ng/H\u1ed5 ph\u00e1ch ch\u01b0a d\u1ef1ng \u0111\u01b0\u1ee3c gi\u1eef d\u1ea1ng th\u00f4.)")
        End If
        ModUI.MsgBoxW finMsg, vbInformation, ModConstants.Uni("Done (Ho\u00e0n t\u1ea5t)")
    End If
    Exit Sub
ErrHandler:
    If weStartedUndo Then
        objUndo.EndCustomRecord
    End If
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
