Attribute VB_Name = "ModCitationScan"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModCitationScan
' Purpose: Read-only Pre-Audit Citation Counter (APA English + Vietnamese)
' Architecture: Academic Document Format Engine (100% 7-bit ASCII, < 400 lines)
' ==============================================================================

Private Type ChapterCountBucket
    Title As String
    CountCites As Long
End Type

Private g_LastPreAuditReport As String

Public Function GetLastPreAuditReport() As String
    GetLastPreAuditReport = g_LastPreAuditReport
End Function

Private Function BuildApaYearToken(Optional ByVal allowBareKn As Boolean = True) As String
    Dim yr As String, enNoYr As String, viNoYr As String
    yr = "(?:18|19|20)\d{2}[a-z]?"
    enNoYr = "n\.d\.?|in\s+press|forthcoming"
    If allowBareKn Then
        viNoYr = ModConstants.Uni("k\.n\.x\.b\.|k\.nxb|k\.n\.|k\.n|kn|kh\u00f4ng\s+r\u00f5\s+n\u0103m|kh\u00f4ng\s+n\u0103m|kh\u00f4ng\s+ghi\s+n\u0103m|ch\u01b0a\s+r\u00f5\s+n\u0103m|\u0111ang\s+in|\u0111ang\s+xu\u1ea5t\s+b\u1ea3n|s\u1eafp\s+xu\u1ea5t\s+b\u1ea3n")
    Else
        viNoYr = ModConstants.Uni("k\.n\.x\.b\.|k\.nxb|k\.n\.|k\.n|kh\u00f4ng\s+r\u00f5\s+n\u0103m|kh\u00f4ng\s+n\u0103m|kh\u00f4ng\s+ghi\s+n\u0103m|ch\u01b0a\s+r\u00f5\s+n\u0103m|\u0111ang\s+in|\u0111ang\s+xu\u1ea5t\s+b\u1ea3n|s\u1eafp\s+xu\u1ea5t\s+b\u1ea3n")
    End If
    BuildApaYearToken = "(?:" & yr & "|" & enNoYr & "|" & viNoYr & ")"
End Function

Private Function IsUpperLetterStart(ByVal s As String) As Boolean
    Dim t As String, ch As String, code As Long
    t = Trim$(s)
    If Len(t) = 0 Then Exit Function
    ch = Left$(t, 1)
    code = AscW(ch) And &HFFFF&
    If code >= 65 And code <= 90 Then
        IsUpperLetterStart = True
    ElseIf InStr(1, ModConstants.Uni("\u0102\u0110\u0128\u0168\u01a0\u01af"), ch, vbBinaryCompare) > 0 Then
        IsUpperLetterStart = True
    ElseIf code >= &H1EA0 And code <= &H1EF8 Then
        IsUpperLetterStart = ((code And 1) = 0)
    ElseIf UCase$(ch) = ch And LCase$(ch) <> ch Then
        IsUpperLetterStart = True
    End If
End Function

Private Function EndsWithBareNumber(ByVal s As String) As Boolean
    Dim t As String, i As Long, ch As Long, dCnt As Long
    t = Trim$(s)
    If Len(t) = 0 Then Exit Function
    i = Len(t)
    Do While i >= 1
        ch = AscW(Mid$(t, i, 1))
        If ch >= 48 And ch <= 57 Then
            dCnt = dCnt + 1
            i = i - 1
        Else
            Exit Do
        End If
    Loop
    If dCnt = 0 Then Exit Function
    If i = 0 Then
        EndsWithBareNumber = True
    Else
        ch = AscW(Mid$(t, i, 1))
        If ch = 32 Or ch = 44 Or ch = 45 Or ch = 47 Then EndsWithBareNumber = True
    End If
End Function

Private Sub EnsureScanRegexes(ByRef reParen As Object, ByRef reSeg As Object, ByRef reIntro As Object, _
                              ByRef reNonCite As Object, ByRef reNarrYr As Object, ByRef reEtAlEnd As Object, _
                              ByRef reStopWord As Object)
    Dim yrTok As String, narrYrTok As String, locTok As String, secTok As String
    yrTok = BuildApaYearToken(True)
    narrYrTok = BuildApaYearToken(False)
    secTok = ModConstants.Uni("(?:\s*,\s*(?:as\s+cited\s+in|cited\s+in|quoted\s+in|d\u1eabn\s+theo|tr\u00edch\s+d\u1eabn\s+t\u1eeb)\s+[^;()]{2,90})?")
    locTok = "(?:\s*,\s*(?:(?:18|19|20)\d{2}[a-z]?|(?:p|pp|tr|trang)\.?\s*\d+(?:\s*[-]\s*\d+)?|[A-Za-z0-9.\s]{1,25}))?" & secTok
    Set reParen = CreateObject("VBScript.RegExp")
    reParen.Global = True
    reParen.Pattern = "\([^()]{4,260}\)"

    Set reIntro = CreateObject("VBScript.RegExp")
    reIntro.IgnoreCase = True
    reIntro.Pattern = ModConstants.Uni("^(?:see(?:\s+also)?|cf\.?|e\.g\.?,?|for\s+example,?|xem(?:\s+th\u00eam)?|v\u00ed\s+d\u1ee5[:,]?|theo)\s+")

    Set reNonCite = CreateObject("VBScript.RegExp")
    reNonCite.IgnoreCase = True
    reNonCite.Pattern = ModConstants.Uni("^(?:b\u1ea3ng|h\u00ecnh|bi\u1ec3u\s+\u0111\u1ed3|s\u01a1\s+\u0111\u1ed3|\u0111\u1ed3\s+th\u1ecb|ch\u01b0\u01a1ng|m\u1ee5c|ph\u1ea7n|ph\u1ee5\s+l\u1ee5c|table|figure|fig\.?|chart|diagram|chapter|section|appendix|p\.|pp\.|tr\.|trang)\b")

    Set reSeg = CreateObject("VBScript.RegExp")
    reSeg.IgnoreCase = True
    reSeg.Pattern = "^([^;()]{2,110}?)\s*,\s*(" & yrTok & "(?:\s*,\s*(?:" & yrTok & "|[a-z](?![A-Za-z])))*)" & locTok & "$"

    Set reNarrYr = CreateObject("VBScript.RegExp")
    reNarrYr.IgnoreCase = True
    reNarrYr.Global = True
    reNarrYr.Pattern = "\(\s*(" & narrYrTok & "(?:\s*,\s*(?:" & narrYrTok & "|[a-z](?![A-Za-z])))*)(?:\s*,\s*(?:p|pp|tr|trang)\.?\s*\d+(?:\s*[-]\s*\d+)?)?\s*\)"

    Set reEtAlEnd = CreateObject("VBScript.RegExp")
    reEtAlEnd.IgnoreCase = True
    reEtAlEnd.Pattern = ModConstants.Uni("\s+(?:et\s+al\.?(?:'s|\u2019s)?|v\u00e0\s+c\u1ed9ng\s+s\u1ef1|v\u00e0\s+cs\.?|&\s*cs\.?|v\u00e0\s+nnk\.?|v\u00e0\s+nh\u1eefng\s+ng\u01b0\u1eddi\s+kh\u00e1c|c\u00f9ng\s+c\u1ed9ng\s+s\u1ef1)\s*$")

    Set reStopWord = CreateObject("VBScript.RegExp")
    reStopWord.IgnoreCase = True
    reStopWord.Pattern = ModConstants.Uni("^(?:n\u0103m|th\u00e1ng|ng\u00e0y|giai\s+\u0111o\u1ea1n|b\u1ea3ng|h\u00ecnh|bi\u1ec3u\s+\u0111\u1ed3|s\u01a1\s+\u0111\u1ed3|\u0111\u1ed3\s+th\u1ecb|ch\u01b0\u01a1ng|m\u1ee5c|ph\u1ea7n|ph\u1ee5\s+l\u1ee5c|ph\u01b0\u01a1ng\s+tr\u00ecnh|c\u00f4ng\s+th\u1ee9c|k\u1ebft\s+qu\u1ea3|d\u1eef\s+li\u1ec7u|phi\u00ean\s+b\u1ea3n|l\u1ea7n|t\u1eadp|s\u1ed1|quy\u1ec3n|trang|year|years|table|figure|chapter|section|appendix|equation|formula|version|phase|step)$")
End Sub

Private Function CountYearItemsInList(ByVal yrList As String) As Long
    Dim parts() As String, p As Long, t As String
    parts = Split(yrList, ",")
    For p = LBound(parts) To UBound(parts)
        t = Trim$(parts(p))
        If Len(t) > 0 Then CountYearItemsInList = CountYearItemsInList + 1
    Next p
    If CountYearItemsInList < 1 Then CountYearItemsInList = 1
End Function

Private Function CountParenSegs(ByVal innerTxt As String, ByVal reSeg As Object, ByVal reIntro As Object, _
                                ByVal reNonCite As Object) As Long
    Dim segs() As String, k As Long, seg As String, ms As Object, authPart As String, yrPart As String
    segs = Split(innerTxt, ";")
    For k = LBound(segs) To UBound(segs)
        seg = Trim$(segs(k))
        seg = Trim$(reIntro.Replace(seg, ""))
        If Len(seg) >= 6 And Not reNonCite.Test(seg) Then
            Set ms = reSeg.Execute(seg)
            If ms.Count = 1 Then
                authPart = Trim$(ms(0).SubMatches(0))
                yrPart = Trim$(ms(0).SubMatches(1))
                If IsUpperLetterStart(authPart) And Not EndsWithBareNumber(authPart) Then
                    CountParenSegs = CountParenSegs + CountYearItemsInList(yrPart)
                End If
            End If
        End If
    Next k
End Function

Private Function IsViInstitutionPrefix(ByVal w As String) As Boolean
    Static dInst As Object
    If dInst Is Nothing Then
        Set dInst = CreateObject("Scripting.Dictionary")
        dInst.CompareMode = 1
        Dim arr As Variant, k As Long
        arr = Array(ModConstants.Uni("B\u1ed9"), ModConstants.Uni("Ch\u00ednh"), ModConstants.Uni("Qu\u1ed1c"), _
                    ModConstants.Uni("Th\u1ee7"), ModConstants.Uni("Ch\u1ee7"), ModConstants.Uni("T\u00f2a"), _
                    ModConstants.Uni("Vi\u1ec7n"), ModConstants.Uni("Ng\u00e2n"), ModConstants.Uni("T\u1ed5ng"), _
                    ModConstants.Uni("C\u1ee5c"), ModConstants.Uni("S\u1edf"), ModConstants.Uni("\u1ee6y"), _
                    ModConstants.Uni("H\u1ed9i"), ModConstants.Uni("Li\u00ean"), ModConstants.Uni("T\u1ed5"), _
                    ModConstants.Uni("Qu\u1ef9"), ModConstants.Uni("Hi\u1ec7p"), ModConstants.Uni("C\u00f4ng"), _
                    ModConstants.Uni("T\u1eadp"), ModConstants.Uni("Tr\u01b0\u1eddng"), ModConstants.Uni("\u0110\u1ea1i"), _
                    "Ban", ModConstants.Uni("Ph\u00f2ng"), "Trung", "Chi", "Thanh", ModConstants.Uni("Ki\u1ec3m"), "Kho")
        For k = LBound(arr) To UBound(arr)
            dInst(CStr(arr(k))) = True
        Next k
    End If
    IsViInstitutionPrefix = dInst.Exists(Trim$(w))
End Function

Private Function HasValidAuthorBeforeParen(ByVal txt As String, ByVal parenZeroIdx As Long, _
                                           ByVal reEtAlEnd As Object, ByVal reStopWord As Object) As Boolean
    If parenZeroIdx <= 1 Then Exit Function
    Dim leftSlice As String, hadEtAl As Boolean, i As Long, ch As String, nextCh As String
    Dim words() As String, wIdx As Long, w As String, checked As Long, candW As String
    leftSlice = Left$(txt, parenZeroIdx)
    If reEtAlEnd.Test(leftSlice) Then
        hadEtAl = True
        leftSlice = reEtAlEnd.Replace(leftSlice, "")
    End If
    leftSlice = Trim$(leftSlice)
    If Len(leftSlice) = 0 Then Exit Function
    If Right$(leftSlice, 1) = "." Then leftSlice = Trim$(Left$(leftSlice, Len(leftSlice) - 1))
    If Len(leftSlice) = 0 Then Exit Function
    For i = Len(leftSlice) To 1 Step -1
        ch = Mid$(leftSlice, i, 1)
        If ch = "?" Or ch = "!" Or ch = ";" Or ch = "(" Or ch = ")" Or ch = ":" Then
            leftSlice = Trim$(Mid$(leftSlice, i + 1))
            Exit For
        ElseIf ch = "." And i < Len(leftSlice) Then
            nextCh = Mid$(leftSlice, i + 1, 1)
            If nextCh = " " Then
                If i < 3 Or Mid$(leftSlice, i - 2, 1) <> " " Then
                    leftSlice = Trim$(Mid$(leftSlice, i + 1))
                    Exit For
                End If
            End If
        End If
    Next i
    If Len(leftSlice) = 0 Then Exit Function
    words = Split(leftSlice, " ")
    If UBound(words) < 0 Then Exit Function
    wIdx = UBound(words)
    Do While wIdx >= LBound(words)
        w = Trim$(Replace(Replace(Replace(words(wIdx), ",", ""), "[", ""), "]", ""))
        If Len(w) > 0 Then Exit Do
        wIdx = wIdx - 1
    Loop
    If Len(w) = 0 Or reStopWord.Test(w) Then Exit Function
    If IsUpperLetterStart(w) Or hadEtAl Then
        HasValidAuthorBeforeParen = True
        Exit Function
    End If
    checked = 1
    wIdx = wIdx - 1
    Do While wIdx >= LBound(words) And checked < 5
        If Right$(Trim$(words(wIdx)), 1) = "," Then Exit Do
        candW = Trim$(Replace(Replace(Replace(words(wIdx), ",", ""), "[", ""), "]", ""))
        If Len(candW) > 0 Then
            checked = checked + 1
            If IsUpperLetterStart(candW) And IsViInstitutionPrefix(candW) Then
                HasValidAuthorBeforeParen = True
                Exit Function
            End If
        End If
        wIdx = wIdx - 1
    Loop
End Function

Private Sub ScanParaCitations(ByVal normTxt As String, ByVal reParen As Object, ByVal reSeg As Object, _
                              ByVal reIntro As Object, ByVal reNonCite As Object, ByVal reNarrYr As Object, _
                              ByVal reEtAlEnd As Object, ByVal reStopWord As Object, _
                              ByRef outClusters As Long, ByRef outParenCites As Long, ByRef outNarrCites As Long)
    Dim workTxt As String, ms As Object, m As Object, segHits As Long
    Dim innerTxt As String, narrMs As Object, nm As Object
    workTxt = normTxt
    Set ms = reParen.Execute(normTxt)
    If ms.Count > 0 Then
        For Each m In ms
            innerTxt = Mid$(normTxt, m.FirstIndex + 2, m.Length - 2)
            segHits = CountParenSegs(innerTxt, reSeg, reIntro, reNonCite)
            If segHits > 0 Then
                outClusters = outClusters + 1
                outParenCites = outParenCites + segHits
                Mid$(workTxt, m.FirstIndex + 1, m.Length) = Space$(m.Length)
            End If
        Next m
    End If
    Set narrMs = reNarrYr.Execute(workTxt)
    If narrMs.Count > 0 Then
        For Each nm In narrMs
            If HasValidAuthorBeforeParen(workTxt, nm.FirstIndex, reEtAlEnd, reStopWord) Then
                outNarrCites = outNarrCites + CountYearItemsInList(CStr(nm.SubMatches(0)))
            End If
        Next nm
    End If
End Sub

Private Sub AddChapterCites(ByRef buckets() As ChapterCountBucket, ByRef bCount As Long, _
                            ByVal chapTitle As String, ByVal addCites As Long)
    If bCount = 0 Then
        bCount = 1
        ReDim buckets(1 To 16)
        buckets(1).Title = chapTitle
    ElseIf buckets(bCount).Title <> chapTitle Then
        bCount = bCount + 1
        If bCount > UBound(buckets) Then ReDim Preserve buckets(1 To bCount + 16)
        buckets(bCount).Title = chapTitle
    End If
    buckets(bCount).CountCites = buckets(bCount).CountCites + addCites
End Sub

Private Function BuildPreAuditReportText(ByVal totalCites As Long, ByVal clusters As Long, _
                                         ByVal parenCites As Long, ByVal narrCites As Long, _
                                         ByVal refEntries As Long, ByRef buckets() As ChapterCountBucket, _
                                         ByVal bCount As Long) As String
    Dim s As String, k As Long
    s = ModConstants.Uni("PRE-AUDIT SCAN REPORT (B\u00c1O C\u00c1O \u0110\u1ebeM TR\u00cdCH D\u1eaaN):") & vbCrLf & vbCrLf & _
        ModConstants.Uni("- Total citations (T\u1ed5ng l\u01b0\u1ee3t tr\u00edch d\u1eabn): ") & totalCites & vbCrLf & _
        "  + Parenthetical: " & parenCites & " (" & clusters & " clusters)" & vbCrLf & _
        "  + Narrative: " & narrCites & vbCrLf & _
        ModConstants.Uni("- References entries (M\u1ee5c T\u00e0i li\u1ec7u tham kh\u1ea3o): ") & refEntries & vbCrLf
    If bCount > 0 Then
        s = s & vbCrLf & ModConstants.Uni("Breakdown by Chapter (Ph\u00e2n b\u1ed5 theo Ch\u01b0\u01a1ng):") & vbCrLf
        For k = 1 To bCount
            s = s & "  - " & buckets(k).Title & ": " & buckets(k).CountCites & vbCrLf
        Next k
    End If
    BuildPreAuditReportText = s
End Function

Public Function ExecutePreAuditScan(ByVal doc As Document, ByVal showUi As Boolean) As String
    Dim markers As MarkerLocations, paras() As ScannedPara, pCount As Long, maxRefNum As Long, errMsg As String
    g_LastPreAuditReport = ""
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        ExecutePreAuditScan = "ERR|BUSY"
        Exit Function
    End If
    If doc Is Nothing Then
        ExecutePreAuditScan = "ERR|NODOC"
        Exit Function
    End If
    ModMain.g_IsEngineRunning = True
    ModRibbonCallbacks.InvalidateRibbonSafe
    On Error GoTo ScanFail
    Application.StatusBar = "Pre-Audit Scan: Scanning document..."
    If Not ModPatchMatcher.ScanDocForPatch(doc, markers, paras, pCount, maxRefNum, errMsg) Then
        Application.StatusBar = False
        If showUi Then ModUI.MsgBoxW errMsg, vbExclamation, ModConstants.Uni("Pre-Audit Scan Error")
        ExecutePreAuditScan = "ERR|SCAN|" & errMsg
        GoTo CleanScan
    End If
    Application.StatusBar = "Pre-Audit Scan: Counting citations..."
    Dim reParen As Object, reSeg As Object, reIntro As Object, reNonCite As Object
    Dim reNarrYr As Object, reEtAlEnd As Object, reStopWord As Object
    EnsureScanRegexes reParen, reSeg, reIntro, reNonCite, reNarrYr, reEtAlEnd, reStopWord

    Dim clusters As Long, parenCites As Long, narrCites As Long, refEntries As Long
    Dim buckets() As ChapterCountBucket, bCount As Long, curChap As String
    Dim j As Long, cClust As Long, cParen As Long, cNarr As Long
    curChap = "Body (Intro / Main)"
    For j = 1 To pCount
        If (j Mod 50) = 0 Then Application.StatusBar = "Pre-Audit Scan: Counting citations (" & j & "/" & pCount & ")..."
        If paras(j).IsInRefZone Then
            If Not paras(j).IsHeading And Not paras(j).IsInTable And Not paras(j).IsExcluded Then
                refEntries = refEntries + 1
            End If
        ElseIf paras(j).InBodyScope Then
            If paras(j).IsLevel1 Then
                curChap = paras(j).NormText
                If Len(curChap) > 52 Then curChap = Left$(curChap, 49) & "..."
                AddChapterCites buckets, bCount, curChap, 0
            ElseIf Not paras(j).IsHeading And Not paras(j).IsExcluded Then
                cClust = 0
                cParen = 0
                cNarr = 0
                ScanParaCitations paras(j).NormText, reParen, reSeg, reIntro, reNonCite, _
                                  reNarrYr, reEtAlEnd, reStopWord, cClust, cParen, cNarr
                clusters = clusters + cClust
                parenCites = parenCites + cParen
                narrCites = narrCites + cNarr
                If (cParen + cNarr) > 0 Then AddChapterCites buckets, bCount, curChap, cParen + cNarr
            End If
        End If
    Next j
    Dim totalCites As Long, reportTxt As String, auditVotes() As ScopeVoteResult
    Dim auditCompact As String, auditFull As String, auditSparse As String, auditErr As String, fullClipTxt As String
    totalCites = parenCites + narrCites
    reportTxt = BuildPreAuditReportText(totalCites, clusters, parenCites, narrCites, refEntries, buckets, bCount)
    fullClipTxt = reportTxt
    If ModStyleAudit.RunStyleAuditCore(doc, auditVotes, auditCompact, auditFull, auditSparse, auditErr) Then
        fullClipTxt = auditFull & vbCrLf & reportTxt
    End If
    g_LastPreAuditReport = fullClipTxt
    ExecutePreAuditScan = "OK|" & totalCites & "|" & clusters & "|" & parenCites & "|" & narrCites & "|" & refEntries & "|" & bCount
    Application.StatusBar = False
    If showUi Then
        Dim promptMsg As String, resp As Long
        If Len(auditCompact) > 0 Then
            promptMsg = auditCompact & vbCrLf & reportTxt & vbCrLf & _
                        ModConstants.Uni("Copy full X-Ray & Style Lint report to Clipboard? (Ch\u00e9p to\u00e0n b\u1ed9 b\u00e1o c\u00e1o X-Ray & Lint v\u00e0o Clipboard?)")
        Else
            promptMsg = reportTxt & vbCrLf & ModConstants.Uni("Copy this Pre-Audit summary to Clipboard? (Ch\u00e9p b\u00e1o c\u00e1o v\u00e0o Clipboard?)")
        End If
        resp = ModUI.MsgBoxW(promptMsg, vbYesNo + vbQuestion + vbDefaultButton1, ModConstants.Uni("Pre-Audit X-Ray & Lint"))
        If resp = 6 Then ModClipboard.SetClipboardTextUnicode fullClipTxt
    End If
CleanScan:
    On Error Resume Next
    Application.StatusBar = False
    ModTableIndex.InvalidateAllIndexes
    ModCustomProfile.ResetStyleProfileCache
    ModCustomProfile.ResetTagCache
    ModPatchMatcher.ResetRefBmCache
    Err.Clear
    ModMain.g_IsEngineRunning = False
    ModRibbonCallbacks.InvalidateRibbonSafe
    Exit Function
ScanFail:
    Dim failNum As Long, failDesc As String
    failNum = Err.Number
    failDesc = Err.Description
    ExecutePreAuditScan = "ERR|RUNTIME|" & failDesc
    If showUi Then
        On Error Resume Next
        Application.StatusBar = False
        ModUI.MsgBoxW "Pre-Audit Scan error (" & failNum & "): " & failDesc, vbCritical, ModConstants.Uni("Pre-Audit Scan Error")
    End If
    Resume CleanScan
End Function

Public Sub RunPreAuditScan()
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then Exit Sub
    If Application.Documents.Count = 0 Then
        ModUI.MsgBoxW ModConstants.Uni("No open Word document found. (Kh\u00f4ng t\u00ecm th\u1ea5y t\u00e0i li\u1ec7u Word \u0111ang m\u1edf.)"), _
                      vbExclamation, ModConstants.Uni("Error (L\u1ed7i)")
        Exit Sub
    End If
    ExecutePreAuditScan ActiveDocument, True
End Sub

Public Function RunPreAuditScanSilentForTest(ByVal doc As Document) As String
    RunPreAuditScanSilentForTest = ExecutePreAuditScan(doc, False)
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
