Attribute VB_Name = "ModTableIndex"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModTableIndex
' Purpose: Safe fail-closed Table Range Index for read-only snapshot scanning
' Compliance: M-05 (< 400 lines, compound = 0, Option Explicit, 7-bit ASCII)
' ==============================================================================

Private Type TableRangeDef
    StartPos As Long
    EndPos As Long
End Type

Private g_TblRanges() As TableRangeDef
Private g_TblCount As Long
Private g_IndexValid As Boolean

Public Sub BuildTableIndex(ByVal doc As Document)
    g_IndexValid = False
    g_TblCount = 0
    
    On Error GoTo ErrorHandler
    If doc Is Nothing Then
        Exit Sub
    End If
    
    g_TblCount = doc.Tables.Count
    If g_TblCount <= 0 Then
        ReDim g_TblRanges(0 To 0)
        g_IndexValid = True
        Exit Sub
    End If
    
    ReDim g_TblRanges(1 To g_TblCount)
    Dim i As Long
    For i = 1 To g_TblCount
        g_TblRanges(i).StartPos = doc.Tables(i).Range.Start
        g_TblRanges(i).EndPos = doc.Tables(i).Range.End
    Next i
    g_IndexValid = True
    Exit Sub
    
ErrorHandler:
    g_IndexValid = False
    g_TblCount = 0
End Sub

Public Function IsRangeInTable(ByVal rng As Range) As Boolean
    If Not g_IndexValid Then
        On Error Resume Next
        IsRangeInTable = CBool(rng.Information(12))
        On Error GoTo 0
        Exit Function
    End If
    
    If g_TblCount <= 0 Then
        IsRangeInTable = False
        Exit Function
    End If
    
    Dim s As Long
    s = rng.Start
    Dim i As Long
    For i = 1 To g_TblCount
        If s >= g_TblRanges(i).StartPos Then
            If s < g_TblRanges(i).EndPos Then
                IsRangeInTable = True
                Exit Function
            End If
        End If
    Next i
    IsRangeInTable = False
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
