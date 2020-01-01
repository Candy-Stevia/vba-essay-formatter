Attribute VB_Name = "ModFormatBatch"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModFormatBatch
' Purpose: Contiguous Range Batching Engine for Normal Paragraphs
' Architecture: Formats contiguous runs of homogeneous normal paragraphs in a
'               single COM Range call, eliminating per-paragraph COM overhead.
' Invariants: Never batch across headings, tables, Word Math, lists, or excluded.
' Highlight Parity: Handles mixed highlight (wdUndefined = 9999999) per paragraph.
' ==============================================================================

Public g_BatchCount As Long
Public g_BatchedParaCount As Long
Public g_SingleParaCount As Long

Public Sub ResetBatchCounters()
    g_BatchCount = 0
    g_BatchedParaCount = 0
    g_SingleParaCount = 0
End Sub

Public Function GetBatchTelemetry() As String
    Dim s As String
    s = "{"
    s = s & """batch_count"":" & g_BatchCount & ","
    s = s & """batched_para_count"":" & g_BatchedParaCount & ","
    s = s & """single_para_count"":" & g_SingleParaCount
    s = s & "}"
    GetBatchTelemetry = s
End Function

Public Function IsBatchableNormal(ByRef snap As ParagraphSnapshot) As Boolean
    IsBatchableNormal = False
    If snap.ElementType <> ElementTypeNormalParagraph Then
        Exit Function
    End If
    If snap.IsExcluded Then
        Exit Function
    End If
    If snap.HasSoftBreak Then
        Exit Function
    End If
    IsBatchableNormal = True
End Function

Public Sub FormatNormalBatch(ByVal doc As Document, ByVal rStart As Long, _
                            ByVal rEnd As Long, ByVal fontName As String, _
                            ByVal fontSize As Single, ByVal indentPts As Single, _
                            ByVal lineSpacingMode As String, ByVal paraCount As Long)
    Dim rng As Range
    Set rng = doc.Range(rStart, rEnd)
    
    ModFormatCore.ClearYellowAndRedBleed rng
    
    rng.Font.Name = fontName
    rng.Font.Size = fontSize
    rng.ParagraphFormat.Alignment = 3
    rng.ParagraphFormat.FirstLineIndent = indentPts
    rng.ParagraphFormat.LeftIndent = 0#
    rng.ParagraphFormat.KeepWithNext = False
    ModFormatSingle.SafeSetOutlineLevel rng, 10
    
    If lineSpacingMode <> "ls_keep" Then
        ModFormatCore.ApplyLineSpacing rng, lineSpacingMode
    End If
    
    If paraCount > 1 Then
        g_BatchCount = g_BatchCount + 1
        g_BatchedParaCount = g_BatchedParaCount + paraCount
    Else
        g_SingleParaCount = g_SingleParaCount + 1
    End If
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
