Attribute VB_Name = "ModTableIndex"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModTableIndex
' Purpose: Sparse in-RAM coordinate indexes (Tables, ListParagraphs, OMaths)
' Compliance: M-05 (< 600 lines, compound = 0, Option Explicit, 7-bit ASCII)
' ==============================================================================

Private Type SpanRangeDef
    StartPos As Long
    EndPos As Long
End Type

Private g_TblRanges() As SpanRangeDef
Private g_TblCount As Long
Private g_TblValid As Boolean
Private g_TblDocPtr As LongPtr

Private g_ListRanges() As SpanRangeDef
Private g_ListCount As Long
Private g_ListValid As Boolean
Private g_ListDocPtr As LongPtr

Private g_MathRanges() As SpanRangeDef
Private g_MathCount As Long
Private g_MathValid As Boolean
Private g_MathDocPtr As LongPtr

Public Sub InvalidateTableIndex()
    g_TblValid = False
    g_TblCount = 0
    g_TblDocPtr = 0
End Sub

Public Sub InvalidateAllIndexes()
    g_TblValid = False
    g_TblCount = 0
    g_TblDocPtr = 0
    g_ListValid = False
    g_ListCount = 0
    g_ListDocPtr = 0
    g_MathValid = False
    g_MathCount = 0
    g_MathDocPtr = 0
End Sub

Private Sub SortAndMergeSpans(ByRef rawRanges() As SpanRangeDef, ByVal rawCount As Long, _
                              ByRef outRanges() As SpanRangeDef, ByRef outCount As Long)
    Dim i As Long, j As Long, keyItem As SpanRangeDef, needSort As Boolean
    needSort = False
    For i = 2 To rawCount
        If rawRanges(i).StartPos < rawRanges(i - 1).StartPos Then
            needSort = True
            Exit For
        End If
    Next i
    If needSort Then
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
    ReDim outRanges(1 To rawCount)
    outCount = 0
    For i = 1 To rawCount
        If outCount = 0 Then
            outCount = 1
            outRanges(1) = rawRanges(i)
        ElseIf rawRanges(i).StartPos < outRanges(outCount).EndPos Then
            If rawRanges(i).EndPos > outRanges(outCount).EndPos Then
                outRanges(outCount).EndPos = rawRanges(i).EndPos
            End If
        Else
            outCount = outCount + 1
            outRanges(outCount) = rawRanges(i)
        End If
    Next i
End Sub

Private Function PosInSortedSpans(ByRef spans() As SpanRangeDef, ByVal cnt As Long, ByVal pos As Long) As Boolean
    If cnt <= 0 Then Exit Function
    Dim lo As Long, hi As Long, mid As Long
    lo = 1
    hi = cnt
    Do While lo <= hi
        mid = lo + (hi - lo) \ 2
        If pos < spans(mid).StartPos Then
            hi = mid - 1
        ElseIf pos >= spans(mid).EndPos Then
            lo = mid + 1
        Else
            PosInSortedSpans = True
            Exit Function
        End If
    Loop
End Function

Public Sub BuildTableIndex(ByVal doc As Document)
    InvalidateTableIndex
    On Error GoTo ErrorHandler
    If doc Is Nothing Then Exit Sub
    Dim rawCount As Long
    rawCount = doc.Tables.Count
    g_TblDocPtr = ObjPtr(doc)
    If rawCount <= 0 Then
        ReDim g_TblRanges(0 To 0)
        g_TblValid = True
        Exit Sub
    End If
    Dim rawRanges() As SpanRangeDef, tbl As Table, i As Long, tRng As Range
    ReDim rawRanges(1 To rawCount)
    i = 0
    For Each tbl In doc.Tables
        i = i + 1
        Set tRng = tbl.Range
        rawRanges(i).StartPos = tRng.Start
        rawRanges(i).EndPos = tRng.End
    Next tbl
    SortAndMergeSpans rawRanges, rawCount, g_TblRanges, g_TblCount
    g_TblValid = True
    Exit Sub
ErrorHandler:
    InvalidateTableIndex
End Sub

Public Sub EnsureTableIndex(ByVal doc As Document)
    If doc Is Nothing Then Exit Sub
    If g_TblValid And g_TblDocPtr = ObjPtr(doc) Then Exit Sub
    BuildTableIndex doc
End Sub

Public Function IsPosInTable(ByVal pos As Long) As Boolean
    If Not g_TblValid Or g_TblCount <= 0 Then Exit Function
    IsPosInTable = PosInSortedSpans(g_TblRanges, g_TblCount, pos)
End Function

Public Function IsRangeInTable(ByVal rng As Range) As Boolean
    If rng Is Nothing Then
        IsRangeInTable = True
        Exit Function
    End If
    Dim docPtr As LongPtr
    On Error Resume Next
    docPtr = ObjPtr(rng.Document)
    If Err.Number <> 0 Then
        Err.Clear
        IsRangeInTable = True
        Exit Function
    End If
    On Error GoTo 0
    If (Not g_TblValid) Or (docPtr <> g_TblDocPtr) Then
        On Error Resume Next
        IsRangeInTable = CBool(rng.Information(12))
        If Err.Number <> 0 Then
            Err.Clear
            IsRangeInTable = True
        End If
        On Error GoTo 0
        Exit Function
    End If
    If g_TblCount <= 0 Then Exit Function
    Dim s As Long
    On Error Resume Next
    s = rng.Start
    If Err.Number <> 0 Then
        Err.Clear
        IsRangeInTable = True
        Exit Function
    End If
    On Error GoTo 0
    IsRangeInTable = PosInSortedSpans(g_TblRanges, g_TblCount, s)
End Function

Public Sub BuildListIndex(ByVal doc As Document)
    g_ListValid = False
    g_ListCount = 0
    g_ListDocPtr = 0
    On Error GoTo ErrList
    If doc Is Nothing Then Exit Sub
    Dim rawCount As Long
    rawCount = doc.ListParagraphs.Count
    g_ListDocPtr = ObjPtr(doc)
    If rawCount <= 0 Then
        ReDim g_ListRanges(0 To 0)
        g_ListValid = True
        Exit Sub
    End If
    Dim rawRanges() As SpanRangeDef, lp As Paragraph, i As Long, r As Range
    ReDim rawRanges(1 To rawCount)
    i = 0
    For Each lp In doc.ListParagraphs
        i = i + 1
        If i > rawCount Then ReDim Preserve rawRanges(1 To i)
        Set r = lp.Range
        rawRanges(i).StartPos = r.Start
        rawRanges(i).EndPos = r.End
    Next lp
    If i > 0 Then SortAndMergeSpans rawRanges, i, g_ListRanges, g_ListCount
    g_ListValid = True
    Exit Sub
ErrList:
    g_ListValid = False
    g_ListCount = 0
    g_ListDocPtr = 0
End Sub

Public Function IsPosInList(ByVal pos As Long) As Boolean
    If Not g_ListValid Or g_ListCount <= 0 Then Exit Function
    IsPosInList = PosInSortedSpans(g_ListRanges, g_ListCount, pos)
End Function

Public Function IsListIndexValidForDoc(ByVal doc As Document) As Boolean
    If doc Is Nothing Then Exit Function
    IsListIndexValidForDoc = (g_ListValid And g_ListDocPtr = ObjPtr(doc))
End Function

Public Sub BuildOMathIndex(ByVal doc As Document)
    g_MathValid = False
    g_MathCount = 0
    g_MathDocPtr = 0
    On Error GoTo ErrMath
    If doc Is Nothing Then Exit Sub
    Dim rawCount As Long
    rawCount = doc.OMaths.Count
    g_MathDocPtr = ObjPtr(doc)
    If rawCount <= 0 Then
        ReDim g_MathRanges(0 To 0)
        g_MathValid = True
        Exit Sub
    End If
    Dim rawRanges() As SpanRangeDef, om As OMath, i As Long, r As Range
    ReDim rawRanges(1 To rawCount)
    i = 0
    For Each om In doc.OMaths
        i = i + 1
        If i > rawCount Then ReDim Preserve rawRanges(1 To i)
        Set r = om.Range
        rawRanges(i).StartPos = r.Start
        rawRanges(i).EndPos = r.End
    Next om
    If i > 0 Then SortAndMergeSpans rawRanges, i, g_MathRanges, g_MathCount
    g_MathValid = True
    Exit Sub
ErrMath:
    g_MathValid = False
    g_MathCount = 0
    g_MathDocPtr = 0
End Sub

Public Function HasOMathInSpan(ByVal rStart As Long, ByVal rEnd As Long) As Boolean
    If Not g_MathValid Or g_MathCount <= 0 Or rEnd <= rStart Then Exit Function
    Dim lo As Long, hi As Long, mid As Long
    lo = 1
    hi = g_MathCount
    Do While lo <= hi
        mid = lo + (hi - lo) \ 2
        If g_MathRanges(mid).EndPos <= rStart Then
            lo = mid + 1
        ElseIf g_MathRanges(mid).StartPos >= rEnd Then
            hi = mid - 1
        Else
            HasOMathInSpan = True
            Exit Function
        End If
    Loop
End Function

Public Function IsOMathIndexValidForDoc(ByVal doc As Document) As Boolean
    If doc Is Nothing Then Exit Function
    IsOMathIndexValidForDoc = (g_MathValid And g_MathDocPtr = ObjPtr(doc))
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
