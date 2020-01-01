Attribute VB_Name = "ModLatex"
Option Explicit

' ==============================================================================
' Module: ModLatex
' Purpose: Robust LaTeX block detection, nested environment tracking & runaway protection
' Architecture: Academic Document Format Engine (Win32 Unicode Safe)
' ==============================================================================

Public Const LATEX_MAX_SPAN As Long = 40

Public Sub ComputeLatexFlags(ByVal doc As Document, _
                             ByRef outIsLatex() As Boolean, _
                             ByRef outLatexWarningCount As Long)
    Dim totalParas As Long
    totalParas = doc.Paragraphs.Count
    outLatexWarningCount = 0
    If totalParas = 0 Then Exit Sub
    ReDim outIsLatex(1 To totalParas)
    
    Dim p As Paragraph, pTexts() As String, kIdx As Long
    ReDim pTexts(1 To totalParas)
    kIdx = 0
    For Each p In doc.Paragraphs
        kIdx = kIdx + 1
        pTexts(kIdx) = Trim$(Replace(Replace(Replace(p.Range.Text, vbCr, ""), Chr(7), ""), Chr(11), " "))
    Next p
    
    Dim i As Long
    i = 1
    Do While i <= totalParas
        Dim pText As String
        pText = pTexts(i)
        
        Dim openType As Long ' 1: $$, 2: \[, 3: \begin{env}
        Dim closeTarget As String
        Dim braceClose As Long
        openType = 0
        closeTarget = ""
        
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
                closeTarget = "\end{" & Mid$(pText, 8, braceClose - 8) & "}"
            End If
        End If
        
        If openType > 0 Then
            Dim isSingleLine As Boolean
            isSingleLine = False
            Select Case openType
                Case 1 ' $$
                    If Len(pText) >= 4 And Right$(pText, 2) = "$$" Then isSingleLine = True
                Case 2 ' \[
                    If Len(pText) >= 4 And Right$(pText, 2) = "\]" Then isSingleLine = True
                Case 3 ' \begin{env}
                    If InStr(braceClose + 1, pText, closeTarget) > 0 Then isSingleLine = True
            End Select
            
            If isSingleLine Then
                outIsLatex(i) = True
                i = i + 1
            Else
                Dim maxLimit As Long, j As Long, foundEnd As Long
                maxLimit = i + LATEX_MAX_SPAN
                If maxLimit > totalParas Then maxLimit = totalParas
                foundEnd = 0
                
                For j = i + 1 To maxLimit
                    Dim scanText As String
                    scanText = pTexts(j)
                    Select Case openType
                        Case 1 ' $$
                            If InStr(scanText, "$$") > 0 Then
                                foundEnd = j
                                Exit For
                            End If
                        Case 2 ' \[
                            If InStr(scanText, "\]") > 0 Then
                                foundEnd = j
                                Exit For
                            End If
                        Case 3 ' \begin{env}
                            If InStr(scanText, closeTarget) > 0 Then
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
