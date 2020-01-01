Attribute VB_Name = "ModFormatHeading"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModFormatHeading
' Purpose: Heading capitalization, prefix formatting, and source note layout
' Architecture: Academic Document Format Engine (Baseline Parity)
' ==============================================================================

Public Sub CapChar(ByVal doc As Document, ByVal charPos As Long)
    If charPos >= 0 And charPos < doc.Range.End Then
        doc.Range(charPos, charPos + 1).Case = 1 ' wdUpperCase = 1
    End If
End Sub

Public Sub CapPrefixAndContent(ByVal doc As Document, ByVal rng As Range, ByVal sep As String)
    CapChar doc, rng.Start
    If ModLocale.IsRomanLevel1(rng.Text) Then
        Dim rPos As Long, rLen As Long, rCh As String
        rLen = Len(rng.Text)
        For rPos = 1 To rLen
            rCh = Mid$(rng.Text, rPos, 1)
            If rCh = "." Or rCh = ":" Then
                Do While rPos + 1 <= rLen And Mid$(rng.Text, rPos + 1, 1) = " "
                    rPos = rPos + 1
                Loop
                If rPos < rLen Then CapChar doc, rng.Start + rPos
                Exit Sub
            End If
        Next rPos
    End If
    Dim reg As Object, mc As Object, matchLen As Long
    Set reg = CreateObject("VBScript.RegExp")
    reg.IgnoreCase = True
    reg.Global = False
    reg.Pattern = ModLocale.GetPrefixCapPattern()
    Set mc = reg.Execute(rng.Text)
    If mc.Count > 0 Then
        matchLen = mc.Item(0).Length
        If matchLen < Len(rng.Text) Then
            CapChar doc, rng.Start + matchLen
        End If
    Else
        Dim k As Long, ch As String, t As String
        t = rng.Text
        For k = 1 To Len(t)
            ch = Mid$(t, k, 1)
            If ch <> " " And ch <> "." And ch <> ":" And (ch < "0" Or ch > "9") Then
                CapChar doc, rng.Start + k - 1
                Exit Sub
            End If
        Next k
    End If
End Sub

Public Sub FormatSourceNoteRange(ByVal doc As Document, ByVal rng As Range, _
                                ByVal fontName As String, ByVal fontSize As Single)
    CapPrefixAndContent doc, rng, ": "
    ModFormatCore.ClearYellowAndRedBleed rng
    rng.Font.Name = fontName
    rng.Font.Size = fontSize
    rng.ParagraphFormat.Alignment = 0
    rng.ParagraphFormat.FirstLineIndent = 0#
    rng.ParagraphFormat.LeftIndent = 0#
    Dim colonPos As Long
    colonPos = InStr(rng.Text, ":")
    If colonPos > 0 Then
        Dim rngPrefix As Range, rngBody As Range
        Set rngPrefix = rng.Duplicate
        rngPrefix.End = rng.Start + colonPos
        rngPrefix.Font.Bold = True
        rngPrefix.Font.Italic = False
        Set rngBody = rng.Duplicate
        rngBody.Start = rng.Start + colonPos
        rngBody.Font.Bold = False
        rngBody.Font.Italic = True
    End If
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
