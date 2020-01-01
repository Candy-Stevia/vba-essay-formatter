Attribute VB_Name = "ModLatex"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModLatex
' Purpose: Robust LaTeX block detection, nested environment tracking & runaway protection
' Architecture: Academic Document Format Engine (Win32 Unicode Safe)
' ==============================================================================

Public Const LATEX_MAX_SPAN As Long = 40

Private Function CountUnescapedDoubleDollar(ByVal s As String) As Long
    Dim pos As Long, cnt As Long, sLen As Long
    Dim slashPos As Long, slashCnt As Long
    sLen = Len(s)
    pos = 1
    cnt = 0
    Do While pos <= sLen - 1
        If Mid$(s, pos, 2) = "$$" Then
            slashCnt = 0
            slashPos = pos - 1
            Do While slashPos >= 1
                If Mid$(s, slashPos, 1) = "\" Then
                    slashCnt = slashCnt + 1
                    slashPos = slashPos - 1
                Else
                    Exit Do
                End If
            Loop
            If (slashCnt Mod 2) = 0 Then
                cnt = cnt + 1
                pos = pos + 2
            Else
                pos = pos + 1
            End If
        Else
            pos = pos + 1
        End If
    Loop
    CountUnescapedDoubleDollar = cnt
End Function

Private Function CountSubstringOccurrences(ByVal s As String, ByVal needle As String) As Long
    Dim pos As Long, cnt As Long, nLen As Long
    nLen = Len(needle)
    If nLen = 0 Then Exit Function
    pos = InStr(1, s, needle, vbBinaryCompare)
    Do While pos > 0
        cnt = cnt + 1
        pos = InStr(pos + nLen, s, needle, vbBinaryCompare)
    Loop
    CountSubstringOccurrences = cnt
End Function

Private Sub CollectLatexCandidateStarts(ByVal doc As Document, ByVal needle As String, _
                                        ByVal docEnd As Long, ByRef cStarts() As Long, _
                                        ByRef cCount As Long, ByRef cCap As Long)
    Dim fRng As Range, pRng As Range, prevEnd As Long, pStart As Long
    Set fRng = doc.Content.Duplicate
    ModMathHelpers.InitDomFind fRng.Find, needle
    fRng.Find.MatchCase = False
    prevEnd = -1
    Do While fRng.Find.Execute
        If fRng.End <= prevEnd Then Exit Do
        prevEnd = fRng.End
        Set pRng = fRng.Paragraphs(1).Range
        pStart = pRng.Start
        cCount = cCount + 1
        If cCount > cCap Then
            cCap = cCap * 2
            ReDim Preserve cStarts(1 To cCap)
        End If
        cStarts(cCount) = pStart
        If pRng.End > prevEnd Then prevEnd = pRng.End
        If pRng.End >= docEnd Then Exit Do
        fRng.SetRange pRng.End, docEnd
    Loop
End Sub

Private Sub SortAndDedupLongs(ByRef arr() As Long, ByRef cnt As Long)
    If cnt <= 1 Then Exit Sub
    Dim i As Long, j As Long, keyVal As Long, outCnt As Long
    For i = 2 To cnt
        keyVal = arr(i)
        j = i - 1
        Do While j >= 1
            If arr(j) > keyVal Then
                arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = keyVal
    Next i
    outCnt = 1
    For i = 2 To cnt
        If arr(i) <> arr(outCnt) Then
            outCnt = outCnt + 1
            arr(outCnt) = arr(i)
        End If
    Next i
    cnt = outCnt
End Sub

Private Sub AppendLatexBlock(ByRef bStarts() As Long, ByRef bEnds() As Long, _
                             ByRef bCount As Long, ByRef bCap As Long, _
                             ByVal sPos As Long, ByVal ePos As Long)
    bCount = bCount + 1
    If bCount > bCap Then
        bCap = bCap * 2
        ReDim Preserve bStarts(1 To bCap)
        ReDim Preserve bEnds(1 To bCap)
    End If
    bStarts(bCount) = sPos
    bEnds(bCount) = ePos
End Sub

Public Sub ComputeLatexFlags(ByVal doc As Document, _
                             ByRef outIsLatex() As Boolean, _
                             ByRef outLatexWarningCount As Long)
    Dim totalParas As Long
    totalParas = doc.Paragraphs.Count
    outLatexWarningCount = 0
    If totalParas = 0 Then Exit Sub
    ReDim outIsLatex(1 To totalParas)

    Dim docEnd As Long, candStarts() As Long, candCount As Long, candCap As Long
    docEnd = doc.Content.End
    candCount = 0
    candCap = 32
    ReDim candStarts(1 To candCap)
    CollectLatexCandidateStarts doc, "$$", docEnd, candStarts, candCount, candCap
    CollectLatexCandidateStarts doc, "\[", docEnd, candStarts, candCount, candCap
    CollectLatexCandidateStarts doc, "\begin{", docEnd, candStarts, candCount, candCap
    If candCount = 0 Then Exit Sub
    SortAndDedupLongs candStarts, candCount

    ModTableIndex.EnsureTableIndex doc

    Dim blkStarts() As Long, blkEnds() As Long, blkCount As Long, blkCap As Long
    blkCount = 0
    blkCap = 32
    ReDim blkStarts(1 To blkCap)
    ReDim blkEnds(1 To blkCap)

    Dim cIdx As Long, coveredUntilPos As Long
    Dim curP As Paragraph, curRng As Range, rawTxt As String, pText As String
    Dim openType As Long, openTarget As String, closeTarget As String
    Dim braceClose As Long, envName As String, depth As Long
    Dim isSingleLine As Boolean, ddCount As Long
    Dim stepCnt As Long, foundEnd As Boolean, endPos As Long, startInTbl As Boolean
    Dim nextP As Paragraph, nRng As Range, nRaw As String, scanText As String, sDd As Long
    coveredUntilPos = -1

    For cIdx = 1 To candCount
        If candStarts(cIdx) < coveredUntilPos Then GoTo NextCandidate
        Set curP = doc.Range(candStarts(cIdx), candStarts(cIdx)).Paragraphs(1)
        Set curRng = curP.Range
        rawTxt = curRng.Text
        pText = Trim$(Replace(Replace(Replace(rawTxt, vbCr, ""), Chr(7), ""), Chr(11), " "))

        openType = 0
        openTarget = ""
        closeTarget = ""
        depth = 0

        If Left$(pText, 2) = "$$" And Left$(pText, 3) <> "$$$" Then
            openType = 1
            closeTarget = "$$"
        ElseIf Left$(pText, 2) = "\[" Then
            openType = 2
            closeTarget = "\]"
        ElseIf Left$(pText, 7) = "\begin{" Then
            braceClose = InStr(8, pText, "}")
            If braceClose > 8 Then
                openType = 3
                envName = Mid$(pText, 8, braceClose - 8)
                openTarget = "\begin{" & envName & "}"
                closeTarget = "\end{" & envName & "}"
            End If
        End If

        If openType > 0 Then
            isSingleLine = False
            Select Case openType
                Case 1
                    ddCount = CountUnescapedDoubleDollar(pText)
                    If ddCount >= 2 And (ddCount Mod 2) = 0 Then isSingleLine = True
                Case 2
                    If ModMathInline.FindUnescapedToken(pText, 3, "\]") > 0 Then isSingleLine = True
                Case 3
                    depth = CountSubstringOccurrences(pText, openTarget) - CountSubstringOccurrences(pText, closeTarget)
                    If depth <= 0 Then isSingleLine = True
            End Select

            If isSingleLine Then
                AppendLatexBlock blkStarts, blkEnds, blkCount, blkCap, curRng.Start, curRng.End
                coveredUntilPos = curRng.End
            Else
                foundEnd = False
                endPos = curRng.End
                If InStr(1, rawTxt, Chr(7)) = 0 Then
                    startInTbl = ModTableIndex.IsPosInTable(curRng.Start)
                    Set nextP = curP.Next
                    For stepCnt = 1 To LATEX_MAX_SPAN - 1
                        If nextP Is Nothing Then Exit For
                        Set nRng = nextP.Range
                        nRaw = nRng.Text
                        scanText = Trim$(Replace(Replace(Replace(nRaw, vbCr, ""), Chr(7), ""), Chr(11), " "))
                        If LCase$(scanText) = "=contentend()" Then Exit For
                        If ModTableIndex.IsPosInTable(nRng.Start) <> startInTbl Then Exit For
                        If openType = 1 Or openType = 2 Then
                            If Left$(scanText, 2) = "$$" And Left$(scanText, 3) <> "$$$" And Len(scanText) > 2 Then
                                sDd = CountUnescapedDoubleDollar(scanText)
                                If sDd >= 2 And (sDd Mod 2) = 0 Then Exit For
                            ElseIf Left$(scanText, 2) = "\[" And Len(scanText) > 2 Then
                                If ModMathInline.FindUnescapedToken(scanText, 3, "\]") > 0 Then Exit For
                            End If
                        End If
                        Select Case openType
                            Case 1
                                If CountUnescapedDoubleDollar(scanText) > 0 Then
                                    foundEnd = True
                                    endPos = nRng.End
                                    Exit For
                                End If
                            Case 2
                                If ModMathInline.FindUnescapedToken(scanText, 1, "\]") > 0 Then
                                    foundEnd = True
                                    endPos = nRng.End
                                    Exit For
                                End If
                            Case 3
                                depth = depth + CountSubstringOccurrences(scanText, openTarget) - CountSubstringOccurrences(scanText, closeTarget)
                                If depth <= 0 Then
                                    foundEnd = True
                                    endPos = nRng.End
                                    Exit For
                                End If
                        End Select
                        If InStr(1, nRaw, Chr(7)) > 0 Then Exit For
                        Set nextP = nextP.Next
                    Next stepCnt
                End If

                If foundEnd Then
                    AppendLatexBlock blkStarts, blkEnds, blkCount, blkCap, curRng.Start, endPos
                    coveredUntilPos = endPos
                Else
                    AppendLatexBlock blkStarts, blkEnds, blkCount, blkCap, curRng.Start, curRng.End
                    outLatexWarningCount = outLatexWarningCount + 1
                    coveredUntilPos = curRng.End
                End If
            End If
        End If
NextCandidate:
    Next cIdx

    If blkCount = 0 Then Exit Sub
    Dim p As Paragraph, kIdx As Long, curBlk As Long, pPos As Long
    kIdx = 0
    curBlk = 1
    For Each p In doc.Paragraphs
        kIdx = kIdx + 1
        pPos = p.Range.Start
        Do While curBlk <= blkCount
            If pPos >= blkEnds(curBlk) Then
                curBlk = curBlk + 1
            Else
                Exit Do
            End If
        Loop
        If curBlk > blkCount Then Exit For
        If pPos >= blkStarts(curBlk) And pPos < blkEnds(curBlk) Then
            outIsLatex(kIdx) = True
        End If
    Next p
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
