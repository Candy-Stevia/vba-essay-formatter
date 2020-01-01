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
Private g_DocPtr As String

Public Sub InvalidateTableIndex()
    g_IndexValid = False
    g_TblCount = 0
    g_DocPtr = ""
End Sub

Public Sub BuildTableIndex(ByVal doc As Document)
    g_IndexValid = False
    g_TblCount = 0
    g_DocPtr = ""
    
    On Error GoTo ErrorHandler
    If doc Is Nothing Then
        Exit Sub
    End If
    
    Dim rawCount As Long
    rawCount = doc.Tables.Count
    g_DocPtr = CStr(ObjPtr(doc))
    If rawCount <= 0 Then
        ReDim g_TblRanges(0 To 0)
        g_IndexValid = True
        Exit Sub
    End If
    
    Dim rawRanges() As TableRangeDef
    ReDim rawRanges(1 To rawCount)
    Dim tbl As Table, i As Long, needSort As Boolean
    Dim tRng As Range
    i = 0
    needSort = False
    For Each tbl In doc.Tables
        i = i + 1
        Set tRng = tbl.Range
        rawRanges(i).StartPos = tRng.Start
        rawRanges(i).EndPos = tRng.End
        If i > 1 Then
            If rawRanges(i).StartPos < rawRanges(i - 1).StartPos Then
                needSort = True
            End If
        End If
    Next tbl
    
    If needSort Then
        Dim j As Long, keyItem As TableRangeDef
        For i = 2 To rawCount
            keyItem = rawRanges(i)
            j = i - 1
            Do While j >= 1
                If rawRanges(j).StartPos > keyItem.StartPos Then
                    rawRanges(j + 1) = rawRanges(j)
                    j = j - 1
                Else
                    Exit Do
                End If
            Loop
            rawRanges(j + 1) = keyItem
        Next i
    End If
    
    ' Merge nested/overlapping table ranges so intervals are strictly disjoint & sorted
    ReDim g_TblRanges(1 To rawCount)
    g_TblCount = 0
    For i = 1 To rawCount
        If g_TblCount = 0 Then
            g_TblCount = 1
            g_TblRanges(1) = rawRanges(i)
        ElseIf rawRanges(i).StartPos < g_TblRanges(g_TblCount).EndPos Then
            If rawRanges(i).EndPos > g_TblRanges(g_TblCount).EndPos Then
                g_TblRanges(g_TblCount).EndPos = rawRanges(i).EndPos
            End If
        Else
            g_TblCount = g_TblCount + 1
            g_TblRanges(g_TblCount) = rawRanges(i)
        End If
    Next i
    
    g_IndexValid = True
    Exit Sub
    
ErrorHandler:
    g_IndexValid = False
    g_TblCount = 0
    g_DocPtr = ""
End Sub

Public Function IsRangeInTable(ByVal rng As Range) As Boolean
    If rng Is Nothing Then
        IsRangeInTable = True
        Exit Function
    End If
    
    Dim docPtr As String
    On Error Resume Next
    docPtr = CStr(ObjPtr(rng.Document))
    If Err.Number <> 0 Then
        Err.Clear
        IsRangeInTable = True
        Exit Function
    End If
    On Error GoTo 0
    
    If (Not g_IndexValid) Or (docPtr <> g_DocPtr) Then
        On Error Resume Next
        IsRangeInTable = CBool(rng.Information(12))
        If Err.Number <> 0 Then
            Err.Clear
            IsRangeInTable = True
        End If
        On Error GoTo 0
        Exit Function
    End If
    
    If g_TblCount <= 0 Then
        IsRangeInTable = False
        Exit Function
    End If
    
    Dim s As Long, lo As Long, hi As Long, mid As Long
    On Error Resume Next
    s = rng.Start
    If Err.Number <> 0 Then
        Err.Clear
        IsRangeInTable = True
        Exit Function
    End If
    On Error GoTo 0
    
    lo = 1
    hi = g_TblCount
    Do While lo <= hi
        mid = lo + (hi - lo) \ 2
        If s < g_TblRanges(mid).StartPos Then
            hi = mid - 1
        ElseIf s >= g_TblRanges(mid).EndPos Then
            lo = mid + 1
        Else
            IsRangeInTable = True
            Exit Function
        End If
    Loop
    IsRangeInTable = False
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
