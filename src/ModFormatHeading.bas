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

Private Sub CapCharAtTextPos(ByVal doc As Document, ByVal rng As Range, ByVal txt As String, ByVal textPos As Long)
    Dim tLen As Long
    tLen = Len(txt)
    If textPos < 1 Or textPos > tLen Then Exit Sub
    If rng.Fields.Count = 0 Then
        CapChar doc, rng.Start + textPos - 1
        Exit Sub
    End If
    
    Dim snippet As String, ch As String, k As Long, stopPos As Long
    stopPos = textPos + 15
    If stopPos > tLen Then stopPos = tLen
    snippet = ""
    For k = textPos To stopPos
        ch = Mid$(txt, k, 1)
        If ch = vbCr Or ch = Chr(11) Or ch = Chr(7) Then Exit For
        snippet = snippet & ch
    Next k
    If Len(snippet) = 0 Then Exit Sub
    
    Dim fRng As Range
    Set fRng = rng.Duplicate
    fRng.Find.ClearFormatting
    fRng.Find.Replacement.ClearFormatting
    If fRng.Find.Execute(FindText:=snippet, MatchWildcards:=False, Format:=False, Wrap:=0) Then
        fRng.Characters(1).Case = 1
    End If
End Sub

Public Sub CapPrefixAndContent(ByVal doc As Document, ByVal rng As Range, ByVal sep As String)
    Dim t As String, tLen As Long, leadSpaces As Long, ch As String
    t = rng.Text
    tLen = Len(t)
    leadSpaces = 0
    Do While leadSpaces < tLen
        ch = Mid$(t, leadSpaces + 1, 1)
        If ch = " " Or ch = vbTab Then
            leadSpaces = leadSpaces + 1
        Else
            Exit Do
        End If
    Loop
    If leadSpaces < tLen Then
        CapCharAtTextPos doc, rng, t, leadSpaces + 1
    End If
    
    If ModLocale.IsRomanLevel1(t) Then
        Dim rPos As Long, rCh As String
        For rPos = leadSpaces + 1 To tLen
            rCh = Mid$(t, rPos, 1)
            If rCh = "." Or rCh = ":" Then
                Do While rPos + 1 <= tLen
                    If Mid$(t, rPos + 1, 1) = " " Or Mid$(t, rPos + 1, 1) = vbTab Then
                        rPos = rPos + 1
                    Else
                        Exit Do
                    End If
                Loop
                If rPos < tLen Then CapCharAtTextPos doc, rng, t, rPos + 1
                Exit Sub
            End If
        Next rPos
    End If
    
    Static reg As Object
    If reg Is Nothing Then
        Set reg = CreateObject("VBScript.RegExp")
        reg.IgnoreCase = True
        reg.Global = False
        reg.Pattern = ModLocale.GetPrefixCapPattern()
    End If
    
    Dim mc As Object, matchLen As Long
    Set mc = reg.Execute(t)
    If mc.Count > 0 Then
        matchLen = mc.Item(0).Length
        If matchLen < tLen Then
            CapCharAtTextPos doc, rng, t, matchLen + 1
        End If
    Else
        Dim k As Long
        For k = leadSpaces + 1 To tLen
            ch = Mid$(t, k, 1)
            If ch <> " " And ch <> vbTab And ch <> "." And ch <> ":" And (ch < "0" Or ch > "9") Then
                CapCharAtTextPos doc, rng, t, k
                Exit Sub
            End If
        Next k
    End If
End Sub

Public Sub FormatSourceNoteRange(ByVal doc As Document, ByVal rng As Range, _
                                ByVal fontName As String, ByVal fontSize As Single, _
                                Optional ByVal hasOMaths As Boolean = False)
    CapPrefixAndContent doc, rng, ": "
    ModFormatCore.ClearYellowAndRedBleed rng
    ModFormatSingle.ApplyFontNameSkippingMath doc, rng, fontName, hasOMaths
    rng.Font.Size = fontSize
    rng.ParagraphFormat.Alignment = 0
    rng.ParagraphFormat.FirstLineIndent = 0#
    rng.ParagraphFormat.LeftIndent = 0#
    rng.ParagraphFormat.KeepWithNext = False
    Dim colonSplitPos As Long
    colonSplitPos = 0
    If rng.Fields.Count = 0 Then
        Dim colonPos As Long
        colonPos = InStr(rng.Text, ":")
        If colonPos > 0 Then colonSplitPos = rng.Start + colonPos
    Else
        Dim fRng As Range
        Set fRng = rng.Duplicate
        fRng.Find.ClearFormatting
        fRng.Find.Replacement.ClearFormatting
        If fRng.Find.Execute(FindText:=":", MatchWildcards:=False, Format:=False, Wrap:=0) Then
            colonSplitPos = fRng.End
        End If
    End If
    If colonSplitPos > rng.Start Then
        Dim rngPrefix As Range, rngBody As Range
        Set rngPrefix = rng.Duplicate
        rngPrefix.End = colonSplitPos
        rngPrefix.Font.Bold = True
        rngPrefix.Font.Italic = False
        Set rngBody = rng.Duplicate
        rngBody.Start = colonSplitPos
        rngBody.Font.Bold = False
        rngBody.Font.Italic = True
    End If
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
