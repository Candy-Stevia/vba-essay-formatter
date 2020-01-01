Attribute VB_Name = "ModFormatSingle"
Option Explicit

' ==============================================================================
' Module: ModFormatSingle
' Purpose: Formats individual non-batchable elements (Headings, Captions, Cells)
' Architecture: Academic Document Format Engine (Modular Single Item Dispatcher)
' ==============================================================================

Public Sub FormatSingleItem(ByVal doc As Document, ByVal rng As Range, _
                           ByVal tType As Long, ByVal bmName As String, _
                           ByVal fontName As String, ByVal fontSize As Single, _
                           ByVal indentPts As Single, ByVal hasOMaths As Boolean)
    Dim mIdx As Long
    Select Case tType
        Case ElementTypeLevel1
            ModFormatCore.ClearYellowAndRedBleed rng
            rng.Font.Name = fontName
            rng.Font.Size = fontSize
            rng.Font.Bold = True
            rng.Case = 1
            If hasOMaths Then
                If rng.OMaths.Count > 0 Then
                    For mIdx = 1 To rng.OMaths.Count
                        rng.OMaths(mIdx).Range.Font.Name = "Cambria Math"
                    Next mIdx
                End If
            End If
            rng.ParagraphFormat.Alignment = 1
            rng.ParagraphFormat.FirstLineIndent = 0#
            rng.ParagraphFormat.LeftIndent = 0#
            If Len(bmName) > 0 Then
                ModBookmark.AddBookmark doc, bmName, rng
            End If
        Case ElementTypeLevel2
            ModFormatHeading.CapPrefixAndContent doc, rng, ""
            ModFormatCore.ClearYellowAndRedBleed rng
            rng.Font.Name = fontName
            rng.Font.Size = fontSize
            rng.Font.Bold = True
            rng.Bold = True
            If hasOMaths Then
                If rng.OMaths.Count > 0 Then
                    For mIdx = 1 To rng.OMaths.Count
                        rng.OMaths(mIdx).Range.Font.Name = "Cambria Math"
                    Next mIdx
                End If
            End If
            rng.ParagraphFormat.Alignment = 0
            rng.ParagraphFormat.FirstLineIndent = indentPts
            rng.ParagraphFormat.LeftIndent = ModConstants.CmToPoints(0.5)
            If Len(bmName) > 0 Then
                ModBookmark.AddBookmark doc, bmName, rng
            End If
        Case ElementTypeLevel3
            ModFormatHeading.CapPrefixAndContent doc, rng, ""
            ModFormatCore.ClearYellowAndRedBleed rng
            rng.Font.Name = fontName
            rng.Font.Size = fontSize
            rng.Font.Bold = True
            rng.Bold = True
            rng.Font.Italic = True
            If hasOMaths Then
                If rng.OMaths.Count > 0 Then
                    For mIdx = 1 To rng.OMaths.Count
                        rng.OMaths(mIdx).Range.Font.Name = "Cambria Math"
                    Next mIdx
                End If
            End If
            rng.ParagraphFormat.Alignment = 0
            rng.ParagraphFormat.FirstLineIndent = indentPts
            rng.ParagraphFormat.LeftIndent = ModConstants.CmToPoints(0.75)
            If Len(bmName) > 0 Then
                ModBookmark.AddBookmark doc, bmName, rng
            End If
        Case ElementTypeLevel4
            ModFormatHeading.CapPrefixAndContent doc, rng, ""
            ModFormatCore.ClearYellowAndRedBleed rng
            rng.Font.Name = fontName
            rng.Font.Size = fontSize
            rng.Font.Bold = False
            rng.Font.Italic = True
            If hasOMaths Then
                If rng.OMaths.Count > 0 Then
                    For mIdx = 1 To rng.OMaths.Count
                        rng.OMaths(mIdx).Range.Font.Name = "Cambria Math"
                    Next mIdx
                End If
            End If
            rng.ParagraphFormat.Alignment = 0
            rng.ParagraphFormat.FirstLineIndent = indentPts
            rng.ParagraphFormat.LeftIndent = ModConstants.CmToPoints(1#)
            If Len(bmName) > 0 Then
                ModBookmark.AddBookmark doc, bmName, rng
            End If
        Case ElementTypeTableTitle
            ModFormatHeading.CapPrefixAndContent doc, rng, ". "
            ModFormatCore.ClearYellowAndRedBleed rng
            rng.Font.Name = fontName
            rng.Font.Size = fontSize
            rng.Font.Bold = True
            rng.Bold = True
            rng.ParagraphFormat.Alignment = 0
            rng.ParagraphFormat.FirstLineIndent = indentPts
            rng.ParagraphFormat.LeftIndent = 0#
            If Len(bmName) > 0 Then
                ModBookmark.AddBookmark doc, bmName, rng
            End If
        Case ElementTypeFigureTitle
            If InStr(rng.Text, ":") > 0 Then
                ModFormatHeading.CapPrefixAndContent doc, rng, ": "
            Else
                ModFormatHeading.CapPrefixAndContent doc, rng, ". "
            End If
            ModFormatCore.ClearYellowAndRedBleed rng
            rng.Font.Name = fontName
            rng.Font.Size = fontSize
            rng.Font.Bold = True
            rng.Bold = True
            rng.ParagraphFormat.Alignment = 1
            rng.ParagraphFormat.FirstLineIndent = 0#
            rng.ParagraphFormat.LeftIndent = 0#
            If Len(bmName) > 0 Then
                ModBookmark.AddBookmark doc, bmName, rng
            End If
        Case ElementTypeInlinePicture
            ModFormatCore.ClearYellowAndRedBleed rng
            rng.ParagraphFormat.Alignment = 1
            rng.ParagraphFormat.FirstLineIndent = 0#
            rng.ParagraphFormat.LeftIndent = 0#
        Case ElementTypeSourceNote
            ModFormatHeading.FormatSourceNoteRange doc, rng, fontName, fontSize
        Case ElementTypeTableCell
            ModFormatCore.ClearYellowAndRedBleed rng
            rng.Font.Name = fontName
            rng.Font.Size = fontSize
            If rng.Cells.Count > 0 Then
                rng.Cells(1).VerticalAlignment = 1
            End If
            rng.ParagraphFormat.Alignment = 1
            rng.ParagraphFormat.FirstLineIndent = 0#
            rng.ParagraphFormat.LeftIndent = 0#
        Case ElementTypeNormalParagraph
            ModFormatCore.ClearYellowAndRedBleed rng
            rng.Font.Name = fontName
            rng.Font.Size = fontSize
            If hasOMaths Then
                If rng.OMaths.Count > 0 Then
                    For mIdx = 1 To rng.OMaths.Count
                        rng.OMaths(mIdx).Range.Font.Name = "Cambria Math"
                    Next mIdx
                End If
            End If
            rng.ParagraphFormat.Alignment = 3
            rng.ParagraphFormat.FirstLineIndent = indentPts
            rng.ParagraphFormat.LeftIndent = 0#
            ModFormatBatch.g_SingleParaCount = ModFormatBatch.g_SingleParaCount + 1
        Case ElementTypeIgnoredManualList
            rng.HighlightColorIndex = 6
    End Select
End Sub
