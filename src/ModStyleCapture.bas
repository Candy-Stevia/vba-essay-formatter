Attribute VB_Name = "ModStyleCapture"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModStyleCapture
' Purpose: Wave 2 - Sparse @STYLE Generator, X-Ray Report Formatter & Capture Style
' Architecture: Academic Document Format Engine (100% 7-bit ASCII, < 250 lines)
' ==============================================================================

Private Const WD_UNDEF As Long = 9999999

Private Function BuildVoteKey(ByRef sm As StyleSample) As String
    Dim bStr As String, iStr As String, uStr As String, colStr As String
    If sm.ScopeIdx = 8 Then
        If sm.BoldVal = WD_UNDEF Then bStr = "0" Else bStr = CStr(sm.BoldVal)
        If sm.ItalicVal = WD_UNDEF Then iStr = "0" Else iStr = CStr(sm.ItalicVal)
        If sm.UnderVal = WD_UNDEF Then uStr = "0" Else uStr = CStr(sm.UnderVal)
        If sm.ColorKey = "mixed" Then colStr = "auto" Else colStr = sm.ColorKey
        BuildVoteKey = sm.FontName & "|" & CStr(sm.SizeHalf) & "|" & CStr(sm.AlignVal) & "|" & _
                       CStr(sm.LeftInd20) & "|" & CStr(sm.FirstInd20) & "|" & CStr(sm.LineSp20) & "|" & _
                       bStr & "|" & iStr & "|" & uStr & "|" & colStr
    ElseIf sm.ScopeIdx = 7 And sm.IsSrcCanonical Then
        BuildVoteKey = sm.FontName & "|" & CStr(sm.SizeHalf) & "|CANON|" & sm.ColorKey & "|" & CStr(sm.AlignVal)
    Else
        BuildVoteKey = sm.FontName & "|" & CStr(sm.SizeHalf) & "|" & CStr(sm.BoldVal) & "|" & _
                       CStr(sm.ItalicVal) & "|" & CStr(sm.UnderVal) & "|" & sm.ColorKey & "|" & _
                       CStr(sm.AlignVal) & "|" & CStr(sm.LeftInd20) & "|" & CStr(sm.FirstInd20)
    End If
End Function

Private Function DescribeDeviation(ByRef sm As StyleSample, ByRef vr As ScopeVoteResult) As String
    Dim d As String
    If StrComp(sm.FontName, vr.WinFont, vbTextCompare) <> 0 Then d = d & "font=" & sm.FontName & " "
    If sm.SizeHalf <> vr.WinSizeHalf Then d = d & "size=" & FormatHalfPt(sm.SizeHalf) & "pt "
    If sm.ScopeIdx = 8 Then
        If sm.BoldVal <> WD_UNDEF And sm.BoldVal <> vr.WinBold Then d = d & "bold=" & IIf(sm.BoldVal = -1, "1", "0") & " "
        If sm.ItalicVal <> WD_UNDEF And sm.ItalicVal <> vr.WinItalic Then d = d & "italic=" & IIf(sm.ItalicVal = -1, "1", "0") & " "
        If sm.UnderVal <> WD_UNDEF And sm.UnderVal <> vr.WinUnder Then d = d & "underline=" & CStr(sm.UnderVal) & " "
        If sm.ColorKey <> "mixed" And sm.ColorKey <> vr.WinColor Then d = d & "color=" & sm.ColorKey & " "
        If sm.AlignVal <> vr.WinAlign Then d = d & "align=" & ModStyleAudit.AlignToName(sm.AlignVal) & " "
        If sm.LeftInd20 <> vr.WinLeft20 Then d = d & "leftIndent=" & FormatCm20(sm.LeftInd20) & "cm "
        If sm.FirstInd20 <> vr.WinFirst20 Then d = d & "firstIndent=" & FormatCm20(sm.FirstInd20) & "cm "
        If sm.LineSp20 <> vr.WinLineSp20 Then d = d & "lineSp=" & FormatCm20(sm.LineSp20) & " "
    Else
        If Not (sm.ScopeIdx = 7 And sm.IsSrcCanonical And vr.WinSrcCanonical) Then
            If sm.BoldVal <> vr.WinBold Then d = d & "bold=" & IIf(sm.BoldVal = -1, "1", "0") & " "
            If sm.ItalicVal <> vr.WinItalic Then d = d & "italic=" & IIf(sm.ItalicVal = -1, "1", "0") & " "
            If sm.UnderVal <> vr.WinUnder Then d = d & "underline=" & CStr(sm.UnderVal) & " "
        End If
        If sm.ColorKey <> vr.WinColor Then d = d & "color=" & sm.ColorKey & " "
        If sm.AlignVal <> vr.WinAlign Then d = d & "align=" & ModStyleAudit.AlignToName(sm.AlignVal) & " "
        If sm.ScopeIdx <> 7 Then
            If sm.LeftInd20 <> vr.WinLeft20 Then d = d & "leftIndent=" & FormatCm20(sm.LeftInd20) & "cm "
            If sm.FirstInd20 <> vr.WinFirst20 Then d = d & "firstIndent=" & FormatCm20(sm.FirstInd20) & "cm "
        End If
    End If
    DescribeDeviation = Trim$(d)
End Function

Public Sub VoteSingleScope(ByVal doc As Document, ByVal scIdx As Long, ByRef samples() As StyleSample, ByVal smCount As Long, _
                           ByRef outVote As ScopeVoteResult)
    outVote.ScopeKey = ModStyleAudit.ScopeIdxToKey(scIdx)
    outVote.SampleCount = 0
    outVote.Status = "EMPTY"
    Dim dictCnt As Object, dictRep As Object, i As Long, k As String, allUpper As Boolean
    Set dictCnt = CreateObject("Scripting.Dictionary")
    Set dictRep = CreateObject("Scripting.Dictionary")
    allUpper = True
    For i = 1 To smCount
        If samples(i).ScopeIdx = scIdx Then
            outVote.SampleCount = outVote.SampleCount + 1
            If scIdx = 1 And Not samples(i).IsUpper Then allUpper = False
            k = BuildVoteKey(samples(i))
            If dictCnt.Exists(k) Then
                dictCnt(k) = CLng(dictCnt(k)) + 1
            Else
                dictCnt.Add k, 1
                dictRep.Add k, i
            End If
        End If
    Next i
    If outVote.SampleCount = 0 Then Exit Sub
    Dim bestKey As String, bestCnt As Long, secKey As String, secCnt As Long, vKey As Variant, cVal As Long
    For Each vKey In dictCnt.Keys
        cVal = CLng(dictCnt(vKey))
        If cVal > bestCnt Then
            secCnt = bestCnt
            secKey = bestKey
            bestCnt = cVal
            bestKey = CStr(vKey)
        ElseIf cVal > secCnt Then
            secCnt = cVal
            secKey = CStr(vKey)
        End If
    Next vKey
    outVote.WinnerCount = bestCnt
    outVote.SecondCount = secCnt
    Dim repIdx As Long
    repIdx = CLng(dictRep(bestKey))
    outVote.WinFont = samples(repIdx).FontName
    outVote.WinSizeHalf = samples(repIdx).SizeHalf
    outVote.WinBold = IIf(samples(repIdx).BoldVal = WD_UNDEF, 0, samples(repIdx).BoldVal)
    outVote.WinItalic = IIf(samples(repIdx).ItalicVal = WD_UNDEF, 0, samples(repIdx).ItalicVal)
    outVote.WinUnder = IIf(samples(repIdx).UnderVal = WD_UNDEF, 0, samples(repIdx).UnderVal)
    outVote.WinColor = IIf(samples(repIdx).ColorKey = "mixed", "auto", samples(repIdx).ColorKey)
    outVote.WinAlign = samples(repIdx).AlignVal
    outVote.WinLeft20 = samples(repIdx).LeftInd20
    outVote.WinFirst20 = samples(repIdx).FirstInd20
    outVote.WinSpBef = samples(repIdx).SpaceBef
    outVote.WinSpAft = samples(repIdx).SpaceAft
    outVote.WinLineSp20 = samples(repIdx).LineSp20
    outVote.WinUpper = (scIdx = 1 And allUpper)
    outVote.WinSrcCanonical = samples(repIdx).IsSrcCanonical
    If (bestCnt * 100 < outVote.SampleCount * 60) Or (secCnt > 0 And bestCnt = secCnt) Then
        outVote.Status = "TIE"
        Dim tLine As String, r2 As Long
        tLine = "Group A (" & bestCnt & "/" & outVote.SampleCount & "): " & outVote.WinFont & " " & FormatHalfPt(outVote.WinSizeHalf) & "pt"
        If Len(secKey) > 0 Then
            r2 = CLng(dictRep(secKey))
            tLine = tLine & " vs Group B (" & secCnt & "/" & outVote.SampleCount & "): " & samples(r2).FontName & " " & FormatHalfPt(samples(r2).SizeHalf) & "pt"
        End If
        outVote.TieSummary = tLine
        Exit Sub
    End If
    Dim devLines As String, diffStr As String, devPg As Long
    For i = 1 To smCount
        If samples(i).ScopeIdx = scIdx Then
            If BuildVoteKey(samples(i)) <> bestKey Then
                diffStr = DescribeDeviation(samples(i), outVote)
                If Len(diffStr) > 0 Then
                    outVote.DevCount = outVote.DevCount + 1
                    devPg = samples(i).PageNum
                    If devPg = 0 And Not doc Is Nothing And samples(i).ParaIdx >= 1 Then
                        On Error Resume Next
                        devPg = doc.Paragraphs(samples(i).ParaIdx).Range.Information(3)
                        Err.Clear
                        On Error GoTo 0
                    End If
                    devLines = devLines & "  - p." & devPg & " [P" & samples(i).ParaIdx & "] " & _
                               diffStr & "| """ & Left$(samples(i).PreviewText, 60) & """" & vbCrLf
                End If
            End If
        End If
    Next i
    outVote.DeviationsText = devLines
    If outVote.DevCount = 0 Then outVote.Status = "PASS" Else outVote.Status = "WARN"
End Sub

Private Function FormatHalfPt(ByVal hVal As Long) As String
    If hVal = 9999999 Then
        FormatHalfPt = "mixed"
    ElseIf (hVal Mod 2) = 0 Then
        FormatHalfPt = CStr(hVal \ 2)
    Else
        FormatHalfPt = CStr(hVal \ 2) & ".5"
    End If
End Function

Private Function FormatCm20(ByVal c20 As Long) As String
    If c20 = 9999999 Then
        FormatCm20 = "0"
    Else
        FormatCm20 = Replace(Format$(CDbl(c20) / 20#, "0.##"), ",", ".")
    End If
End Function

Public Function BuildScopeStyleLine(ByRef v As ScopeVoteResult, ByVal baseFont As String) As String
    If v.SampleCount = 0 Or v.Status = "TIE" Or v.Status = "EMPTY" Then Exit Function
    Dim spec As String, sc As String
    sc = v.ScopeKey
    If Len(v.WinFont) > 0 And v.WinFont <> "mixed" And StrComp(v.WinFont, baseFont, vbTextCompare) <> 0 Then
        spec = spec & "font=" & v.WinFont & " | "
    End If
    If v.WinSizeHalf > 0 And v.WinSizeHalf <> 9999999 Then
        spec = spec & "size=" & FormatHalfPt(v.WinSizeHalf) & " | "
    End If
    If sc = "L1" Then
        spec = spec & "upper=" & IIf(v.WinUpper, "1", "0") & " | "
    End If
    If sc <> "SRC" Or Not v.WinSrcCanonical Then
        If sc <> "BODY" Then
            spec = spec & "bold=" & IIf(v.WinBold = -1, "1", "0") & " | "
            spec = spec & "italic=" & IIf(v.WinItalic = -1, "1", "0") & " | "
        Else
            If v.WinBold = -1 Then spec = spec & "bold=1 | "
            If v.WinItalic = -1 Then spec = spec & "italic=1 | "
        End If
    End If
    If Len(v.WinColor) > 0 And v.WinColor <> "auto" And v.WinColor <> "mixed" Then
        spec = spec & "color=" & v.WinColor & " | "
    End If
    spec = spec & "align=" & ModStyleAudit.AlignToName(v.WinAlign) & " | "
    If sc <> "SRC" And v.WinLeft20 <> 9999999 Then
        spec = spec & "leftIndent=" & FormatCm20(v.WinLeft20) & " | "
    End If
    If sc <> "SRC" And v.WinFirst20 <> 9999999 Then
        spec = spec & "firstIndent=" & FormatCm20(v.WinFirst20) & " | "
    End If
    If Right$(spec, 3) = " | " Then spec = Left$(spec, Len(spec) - 3)
    If Len(spec) > 0 Then BuildScopeStyleLine = sc & " | " & spec
End Function

Public Function BuildSparseStyleBlock(ByRef votes() As ScopeVoteResult) As String
    Dim baseFont As String, sc As Long, ln As String, blk As String
    baseFont = votes(8).WinFont
    If Len(baseFont) = 0 Or baseFont = "mixed" Then
        ModRibbonCallbacks.EnsureConfigInitialized
        baseFont = ModRibbonCallbacks.g_Config.FontName
    End If
    For sc = 1 To 8
        ln = BuildScopeStyleLine(votes(sc), baseFont)
        If Len(ln) > 0 Then
            If Len(blk) > 0 Then blk = blk & vbCrLf
            blk = blk & ln
        End If
    Next sc
    If Len(blk) > 0 Then BuildSparseStyleBlock = "@STYLE" & vbCrLf & blk
End Function

Public Sub FormatAuditReports(ByVal doc As Document, ByRef mk As MarkerLocations, ByVal hasExp As Boolean, _
                              ByVal startP As Long, ByRef votes() As ScopeVoteResult, _
                              ByRef suspects() As VisualSuspectItem, ByVal susCnt As Long, _
                              ByVal sparseBlock As String, ByRef outCompact As String, ByRef outFull As String)
    Dim tocCnt As Long, tofCnt As Long, tblCnt As Long, mathCnt As Long
    Dim tagCnt As Long, styleScopes As String, bm As Bookmark, v As Variable
    On Error Resume Next
    tocCnt = doc.TablesOfContents.Count
    tofCnt = doc.TablesOfFigures.Count
    tblCnt = doc.Tables.Count
    mathCnt = doc.OMaths.Count
    For Each bm In doc.Bookmarks
        If Left$(bm.Name, 6) = "bm_af_" Then tagCnt = tagCnt + 1
    Next bm
    For Each v In doc.Variables
        If Left$(v.Name, 9) = "af_style_" Then
            If Len(styleScopes) > 0 Then styleScopes = styleScopes & ", "
            styleScopes = styleScopes & Mid$(v.Name, 10)
        End If
    Next v
    Err.Clear
    On Error GoTo 0
    If Len(styleScopes) = 0 Then styleScopes = "(none)"

    Dim topo As String
    topo = "=== DOCUMENT TOPOLOGY X-RAY ===" & vbCrLf
    If hasExp Then
        topo = topo & "- Explicit Markers: =contentStart() P" & mk.ContentStartParaIndex & _
               IIf(mk.HasContentEnd, ", =contentEnd() P" & mk.ContentEndParaIndex, ", no =contentEnd()") & vbCrLf
    Else
        topo = topo & "- Explicit Markers: None (Markerless Mode -> InferredBodyStart: P" & startP & ")" & vbCrLf
    End If
    topo = topo & "- Word TOC: " & tocCnt & " | Word TOF: " & tofCnt & " | Tables: " & tblCnt & " | OMaths: " & mathCnt & vbCrLf
    topo = topo & "- Active @TAGs: " & tagCnt & " | Saved @STYLE scopes: " & styleScopes & vbCrLf

    Dim sc As Long, lintSum As String, lintFull As String, stTag As String
    lintSum = "=== 8-SCOPE STYLE LINT ===" & vbCrLf
    lintFull = lintSum
    For sc = 1 To 8
        If votes(sc).SampleCount > 0 Then
            stTag = "[" & votes(sc).ScopeKey & "] " & votes(sc).Status & " (" & votes(sc).WinnerCount & "/" & votes(sc).SampleCount & _
                    ": " & votes(sc).WinFont & " " & FormatHalfPt(votes(sc).WinSizeHalf) & "pt"
            If votes(sc).ScopeKey = "L1" Then stTag = stTag & ", upper=" & IIf(votes(sc).WinUpper, "1", "0")
            stTag = stTag & ")"
            lintSum = lintSum & "- " & stTag & vbCrLf
            lintFull = lintFull & "- " & stTag & vbCrLf
            If votes(sc).Status = "TIE" Then
                lintFull = lintFull & "  TIE / CONFLICT: " & votes(sc).TieSummary & vbCrLf
            ElseIf votes(sc).DevCount > 0 Then
                lintFull = lintFull & votes(sc).DeviationsText
            End If
        End If
    Next sc

    Dim susFull As String, k As Long
    susFull = "=== UNMATCHED VISUAL SUSPECTS (" & susCnt & ") ===" & vbCrLf
    For k = 1 To susCnt
        susFull = susFull & "- p." & suspects(k).PageNum & " [P" & suspects(k).ParaIdx & "] " & _
                  suspects(k).StyleName & "/OL" & suspects(k).OutlineLvl & "/" & suspects(k).FontName & "/" & _
                  FormatHalfPt(suspects(k).SizeHalf) & "pt/B=" & IIf(suspects(k).BoldVal = -1, "1", "0") & "/" & _
                  suspects(k).ColorKey & " | " & suspects(k).FullText & vbCrLf
    Next k

    outCompact = topo & lintSum & "- Unmatched Visual Suspects: " & susCnt & vbCrLf
    outFull = topo & vbCrLf & lintFull & vbCrLf & susFull
    If Len(sparseBlock) > 0 Then
        outFull = outFull & vbCrLf & "=== PROPOSED SPARSE @STYLE ===" & vbCrLf & sparseBlock & vbCrLf
    End If
End Sub

Public Function CaptureStyleFromDocument(Optional ByVal targetDoc As Document = Nothing, _
                                         Optional ByVal silentMode As Boolean = False) As String
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        CaptureStyleFromDocument = "ERR|BUSY"
        Exit Function
    End If
    Dim doc As Document, errMsg As String
    If targetDoc Is Nothing Then
        If Documents.Count = 0 Then
            CaptureStyleFromDocument = "ERR|NODOC"
            Exit Function
        End If
        Set doc = ActiveDocument
    Else
        Set doc = targetDoc
    End If
    If Not ModBackup.ValidateWritableDocument(doc, errMsg) Then
        If Not silentMode And Application.DisplayAlerts <> 0 Then
            ModUI.MsgBoxW errMsg, vbExclamation, ModConstants.Uni("Capture Style (H\u1ecdc Style t\u1eeb file)")
        End If
        CaptureStyleFromDocument = "ERR|" & errMsg
        Exit Function
    End If

    Dim votes() As ScopeVoteResult, compactSum As String, fullRep As String, sparseBlk As String
    If Not ModStyleAudit.RunStyleAuditCore(doc, votes, compactSum, fullRep, sparseBlk, errMsg) Then
        CaptureStyleFromDocument = "ERR|" & errMsg
        Exit Function
    End If

    Dim baseFont As String, sc As Long, rawLn As String, sKey As String, sNorm As String, sRst As Boolean, sErr As String
    Dim planKeys(1 To 8) As String, planSpecs(1 To 8) As String, planCnt As Long, tieList As String, tieCnt As Long
    Dim previewTxt As String
    baseFont = votes(8).WinFont
    If Len(baseFont) = 0 Or baseFont = "mixed" Then
        ModRibbonCallbacks.EnsureConfigInitialized
        baseFont = ModRibbonCallbacks.g_Config.FontName
    End If
    For sc = 1 To 8
        If votes(sc).Status = "TIE" Then
            tieCnt = tieCnt + 1
            If Len(tieList) > 0 Then tieList = tieList & ", "
            tieList = tieList & votes(sc).ScopeKey
        ElseIf votes(sc).SampleCount > 0 Then
            rawLn = BuildScopeStyleLine(votes(sc), baseFont)
            If Len(rawLn) > 0 Then
                If ModCustomProfile.ValidateStyleDirective(rawLn, sKey, sNorm, sRst, sErr) Then
                    planCnt = planCnt + 1
                    planKeys(planCnt) = sKey
                    planSpecs(planCnt) = sNorm
                    previewTxt = previewTxt & "- " & rawLn & vbCrLf
                End If
            End If
        End If
    Next sc

    If planCnt = 0 Then
        If Not silentMode And Application.DisplayAlerts <> 0 Then
            ModUI.MsgBoxW ModConstants.Uni("No unambiguous @STYLE scopes to capture. (Kh\u00f4ng c\u00f3 nh\u00f3m @STYLE r\u00f5 r\u00e0ng \u0111\u1ec3 h\u1ecdc.)"), _
                          vbInformation, ModConstants.Uni("Capture Style (H\u1ecdc Style t\u1eeb file)")
        End If
        CaptureStyleFromDocument = "OK|0|" & CStr(tieCnt) & "|"
        Exit Function
    End If

    If Not silentMode And Application.DisplayAlerts <> 0 Then
        Dim qMsg As String, ans As Long
        qMsg = ModConstants.Uni("Save these Sparse @STYLE scopes into this document? (L\u01b0u c\u00e1c @STYLE sau v\u00e0o file?)") & vbCrLf & vbCrLf & previewTxt
        If tieCnt > 0 Then
            qMsg = qMsg & vbCrLf & ModConstants.Uni("Skipped TIE / CONFLICT scopes (B\u1ecf qua do h\u00f2a phi\u1ebfu): ") & tieList
        End If
        ans = ModUI.MsgBoxW(qMsg, vbYesNo + vbQuestion + vbDefaultButton1, ModConstants.Uni("Capture Style (H\u1ecdc Style t\u1eeb file)"))
        If ans <> 6 Then
            CaptureStyleFromDocument = "CANCEL"
            Exit Function
        End If
    End If

    Dim objUndo As Object, isUndo As Boolean, savedCsv As String
    isUndo = ModBackup.StartUndoSafe("AcademicFormatter: Capture Style", objUndo)
    For sc = 1 To planCnt
        ModCustomProfile.ApplyStyleDirective doc, planKeys(sc), planSpecs(sc), False
        If Len(savedCsv) > 0 Then savedCsv = savedCsv & ","
        savedCsv = savedCsv & planKeys(sc)
    Next sc
    ModBackup.EndUndoSafe objUndo, isUndo
    ModCustomProfile.ResetStyleProfileCache
    CaptureStyleFromDocument = "OK|" & CStr(planCnt) & "|" & CStr(tieCnt) & "|" & savedCsv
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
