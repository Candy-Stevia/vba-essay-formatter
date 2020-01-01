Attribute VB_Name = "ModBookmark"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModBookmark
' Purpose: Bookmark identity generation, insertion, existence checking and clean hyperlinks
' Architecture: Academic Document Format Engine
' ==============================================================================

Private Const TAG_BM_PREFIX As String = "bm_af_"
Private mTagStarts() As Long
Private mTagEnds() As Long
Private mTagCodes() As Long
Private mTagCount As Long

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
    ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
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

Private Function IsAllAsciiDigits(ByVal s As String) As Boolean
    Dim i As Long, ch As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If ch < 48 Or ch > 57 Then Exit Function
    Next i
    IsAllAsciiDigits = True
End Function

Public Sub CleanExcessBookmarks(ByVal doc As Document, ByVal cntToc As Long, ByVal cntTbl As Long, ByVal cntFig As Long)
    Dim toDel() As String, delCount As Long, delCap As Long
    delCap = 32
    ReDim toDel(1 To delCap)
    Dim bm As Bookmark, nm As String, pfx As String, tailStr As String, bmId As Long
    For Each bm In doc.Bookmarks
        nm = bm.Name
        If Len(nm) >= 11 Then
            pfx = Left$(nm, 7)
            If pfx = BM_PREFIX_TOC Or pfx = BM_PREFIX_TBL Or pfx = BM_PREFIX_FIG Then
                tailStr = Mid$(nm, 8)
                If Len(tailStr) <= 9 Then
                    If IsAllAsciiDigits(tailStr) Then
                        bmId = CLng(tailStr)
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
            End If
        End If
    Next bm
    Dim i As Long
    For i = 1 To delCount
        ModZoneManager.SafeDelBm doc, toDel(i)
    Next i
End Sub

Public Sub ApplyAfTagBookmark(ByVal doc As Document, ByVal pStart As Long, ByVal pEnd As Long, ByVal tagType As String)
    Dim bm As Bookmark, toDel() As String, delCnt As Long, maxSeq As Long, suffix As String, k As Long
    For Each bm In doc.Bookmarks
        If Left$(bm.Name, Len(TAG_BM_PREFIX)) = TAG_BM_PREFIX Then
            If bm.Range.Start >= pStart And bm.Range.Start < pEnd Then
                delCnt = delCnt + 1
                ReDim Preserve toDel(1 To delCnt)
                toDel(delCnt) = bm.Name
            End If
            suffix = Mid$(bm.Name, InStrRev(bm.Name, "_") + 1)
            If Len(suffix) > 0 And Len(suffix) <= 9 And IsNumeric(suffix) Then
                If CLng(suffix) > maxSeq Then maxSeq = CLng(suffix)
            End If
        End If
    Next bm
    For k = 1 To delCnt
        ModZoneManager.SafeDelBm doc, toDel(k)
    Next k
    Dim uType As String, tEnd As Long, pRng As Range
    uType = UCase$(Trim$(tagType))
    If uType = "CLEAR" Then Exit Sub
    Set pRng = doc.Range(pStart, pEnd)
    If Right$(pRng.Text, 1) = Chr(7) Or CBool(pRng.Information(12)) Then
        tEnd = pEnd - 2
    Else
        tEnd = pEnd - 1
    End If
    If tEnd < pStart Then tEnd = pStart
    AddBookmark doc, TAG_BM_PREFIX & LCase$(uType) & "_" & Format$(maxSeq + 1, "0000"), doc.Range(pStart, tEnd)
End Sub

Public Sub ResetAfTagCache()
    mTagCount = 0
    Erase mTagStarts
    Erase mTagEnds
    Erase mTagCodes
End Sub

Public Sub LoadAfTagCache(ByVal doc As Document)
    ResetAfTagCache
    If doc Is Nothing Then Exit Sub
    Dim bm As Bookmark, rest As String, usPos As Long, code As Long, cap As Long
    cap = 16
    For Each bm In doc.Bookmarks
        If Left$(bm.Name, Len(TAG_BM_PREFIX)) = TAG_BM_PREFIX Then
            rest = Mid$(bm.Name, Len(TAG_BM_PREFIX) + 1)
            usPos = InStr(1, rest, "_", vbBinaryCompare)
            If usPos > 1 Then
                code = ModCustomProfile.TagTypeToCode(Left$(rest, usPos - 1))
                If code <> 0 Then
                    If mTagCount = 0 Then
                        ReDim mTagStarts(1 To cap)
                        ReDim mTagEnds(1 To cap)
                        ReDim mTagCodes(1 To cap)
                    ElseIf mTagCount >= cap Then
                        cap = cap * 2
                        ReDim Preserve mTagStarts(1 To cap)
                        ReDim Preserve mTagEnds(1 To cap)
                        ReDim Preserve mTagCodes(1 To cap)
                    End If
                    mTagCount = mTagCount + 1
                    mTagStarts(mTagCount) = bm.Range.Start
                    mTagEnds(mTagCount) = bm.Range.End
                    mTagCodes(mTagCount) = code
                End If
            End If
        End If
    Next bm
End Sub

Public Function HasActiveAfTags() As Boolean
    HasActiveAfTags = (mTagCount > 0)
End Function

Public Function LookupAfTagForRange(ByVal rStart As Long, ByVal rEnd As Long) As Long
    If mTagCount = 0 Then Exit Function
    Dim k As Long
    For k = 1 To mTagCount
        If mTagStarts(k) >= rStart And mTagStarts(k) < rEnd Then
            LookupAfTagForRange = mTagCodes(k)
            Exit Function
        End If
    Next k
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
