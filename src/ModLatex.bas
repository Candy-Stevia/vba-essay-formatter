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

Public Sub ComputeLatexFlags(ByVal doc As Document, _
                             ByRef outIsLatex() As Boolean, _
                             ByRef outLatexWarningCount As Long)
    Dim totalParas As Long
    totalParas = doc.Paragraphs.Count
    outLatexWarningCount = 0
    If totalParas = 0 Then Exit Sub
    ReDim outIsLatex(1 To totalParas)

    Dim hasToken As Boolean, fRng As Range
    Set fRng = doc.Content.Duplicate
    ModMathHelpers.InitDomFind fRng.Find, "$$"
    fRng.Find.MatchCase = False
    hasToken = fRng.Find.Execute
    If Not hasToken Then
        Set fRng = doc.Content.Duplicate
        ModMathHelpers.InitDomFind fRng.Find, "\["
        fRng.Find.MatchCase = False
        hasToken = fRng.Find.Execute
    End If
    If Not hasToken Then
        Set fRng = doc.Content.Duplicate
        ModMathHelpers.InitDomFind fRng.Find, "\begin{"
        fRng.Find.MatchCase = False
        hasToken = fRng.Find.Execute
    End If
    If Not hasToken Then Exit Sub

    ModTableIndex.EnsureTableIndex doc
    Dim p As Paragraph, pTexts() As String, pStarts() As Long, kIdx As Long, pRng As Range
    ReDim pTexts(1 To totalParas)
    ReDim pStarts(1 To totalParas)
    kIdx = 0
    For Each p In doc.Paragraphs
        kIdx = kIdx + 1
        Set pRng = p.Range
        pStarts(kIdx) = pRng.Start
        pTexts(kIdx) = Trim$(Replace(Replace(Replace(pRng.Text, vbCr, ""), Chr(7), ""), Chr(11), " "))
    Next p
    
    Dim i As Long
    i = 1
    Do While i <= totalParas
        Dim pText As String
        pText = pTexts(i)
        
        Dim openType As Long ' 1: $$, 2: \[, 3: \begin{env}
        Dim openTarget As String, closeTarget As String
        Dim braceClose As Long, envName As String
        Dim depth As Long
        openType = 0
        openTarget = ""
        closeTarget = ""
        depth = 0
        
        If Left$(pText, 2) = "$$" Then
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
            Dim isSingleLine As Boolean
            isSingleLine = False
            Select Case openType
                Case 1 ' $$
                    Dim ddCount As Long
                    ddCount = CountUnescapedDoubleDollar(pText)
                    If ddCount >= 2 And (ddCount Mod 2) = 0 Then isSingleLine = True
                Case 2 ' \[
                    Dim lastClose As Long, lastOpen As Long
                    lastClose = InStrRev(pText, "\]")
                    lastOpen = InStrRev(pText, "\[")
                    If lastClose >= 3 And lastOpen < lastClose Then isSingleLine = True
                Case 3 ' \begin{env}
                    depth = CountSubstringOccurrences(pText, openTarget) - CountSubstringOccurrences(pText, closeTarget)
                    If depth <= 0 Then isSingleLine = True
            End Select
            
            If isSingleLine Then
                outIsLatex(i) = True
                i = i + 1
            Else
                Dim maxLimit As Long, j As Long, foundEnd As Long, startInTbl As Boolean
                maxLimit = i + LATEX_MAX_SPAN
                If maxLimit > totalParas Then maxLimit = totalParas
                foundEnd = 0
                startInTbl = ModTableIndex.IsPosInTable(pStarts(i))
                
                For j = i + 1 To maxLimit
                    Dim scanText As String
                    scanText = pTexts(j)
                    If LCase$(scanText) = "=contentend()" Then Exit For
                    If ModTableIndex.IsPosInTable(pStarts(j)) <> startInTbl Then Exit For
                    If openType = 1 Or openType = 2 Then
                        Dim sDd As Long, sClose As Long, sOpen As Long
                        If Left$(scanText, 2) = "$$" And Len(scanText) > 2 Then
                            sDd = CountUnescapedDoubleDollar(scanText)
                            If sDd >= 2 And (sDd Mod 2) = 0 Then Exit For
                        ElseIf Left$(scanText, 2) = "\[" And Len(scanText) > 2 Then
                            sClose = InStrRev(scanText, "\]")
                            sOpen = InStrRev(scanText, "\[")
                            If sClose >= 3 And sOpen < sClose Then Exit For
                        End If
                    End If
                    Select Case openType
                        Case 1 ' $$
                            If CountUnescapedDoubleDollar(scanText) > 0 Then
                                foundEnd = j
                                Exit For
                            End If
                        Case 2 ' \[
                            If InStr(1, scanText, "\]", vbBinaryCompare) > 0 Then
                                foundEnd = j
                                Exit For
                            End If
                        Case 3 ' \begin{env}
                            depth = depth + CountSubstringOccurrences(scanText, openTarget) - CountSubstringOccurrences(scanText, closeTarget)
                            If depth <= 0 Then
                                foundEnd = j
                                Exit For
                            End If
                    End Select
                Next j
                
                If foundEnd > 0 Then
                    Dim k As Long
                    For k = i To foundEnd
                        outIsLatex(k) = True
                    Next k
                    i = foundEnd + 1
                Else
                    outIsLatex(i) = True
                    outLatexWarningCount = outLatexWarningCount + 1
                    i = i + 1
                End If
            End If
        Else
            i = i + 1
        End If
    Loop
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
