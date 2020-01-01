Attribute VB_Name = "ModBookmark"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModBookmark
' Purpose: Bookmark identity generation, insertion, existence checking and clean hyperlinks
' Architecture: Academic Document Format Engine
' ==============================================================================

' --- Bookmark ID Formatting ---

Public Function FormatBookmarkId(ByVal prefix As String, ByVal idNum As Long) As String
    ' Format: bm_toc_0001, bm_tbl_0001, bm_fig_0001
    FormatBookmarkId = prefix & Format$(idNum, "0000")
End Function

' --- Bookmark Insertion & Checking ---

Public Function AddBookmark(ByVal doc As Document, _
                            ByVal bmName As String, _
                            ByVal rng As Range) As Boolean
    On Error GoTo ErrorHandler
    
    If doc.Bookmarks.Exists(bmName) Then
        doc.Bookmarks(bmName).Delete
    End If
    
    doc.Bookmarks.Add Name:=bmName, Range:=rng
    AddBookmark = True
    Exit Function

ErrorHandler:
    AddBookmark = False
End Function

Public Function CheckHasGeneratedBlock(ByVal doc As Document) As Boolean
    CheckHasGeneratedBlock = doc.Bookmarks.Exists(BM_SYS_GENERATED_BLOCK)
End Function

' --- Clean Hyperlink Insertion (Remove underline and blue color) ---

Public Sub InsertCleanHyperlink(ByVal doc As Document, _
                                ByVal rng As Range, _
                                ByVal targetBookmark As String)
    On Error Resume Next
    Dim hl As Hyperlink
    Set hl = doc.Hyperlinks.Add(Anchor:=rng, Address:="", SubAddress:=targetBookmark)
    
    ' Overwrite Word default hyperlink character styling
    With rng.Font
        .Underline = 0  ' wdUnderlineNone = 0
        .ColorIndex = 0 ' wdAuto = 0
    End With
    If Err.Number <> 0 Then
        ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
        Err.Clear
    End If
End Sub

Public Sub CleanExcessBookmarks(ByVal doc As Document, ByVal cntToc As Long, ByVal cntTbl As Long, ByVal cntFig As Long)
    Dim toDel() As String, delCount As Long, delCap As Long
    delCap = 32
    ReDim toDel(1 To delCap)
    Dim bm As Bookmark, nm As String, pfx As String, tailStr As String, bmId As Long
    For Each bm In doc.Bookmarks
        nm = bm.Name
        If Len(nm) >= 11 Then
            pfx = Left$(nm, 7)
            tailStr = Mid$(nm, 8)
            If IsNumeric(tailStr) Then
                bmId = CLng(Val(tailStr))
                If (pfx = BM_PREFIX_TOC And bmId > cntToc) Or (pfx = BM_PREFIX_TBL And bmId > cntTbl) Or (pfx = BM_PREFIX_FIG And bmId > cntFig) Then
                    delCount = delCount + 1
                    If delCount > delCap Then
                        delCap = delCap * 2
                        ReDim Preserve toDel(1 To delCap)
                    End If
                    toDel(delCount) = nm
                End If
            End If
        End If
    Next bm
    Dim i As Long
    For i = 1 To delCount
        ModZoneManager.SafeDelBm doc, toDel(i)
    Next i
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
