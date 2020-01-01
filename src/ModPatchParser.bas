Attribute VB_Name = "ModPatchParser"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModPatchParser
' Purpose: Fault-tolerant parser for AF-PATCH v1 flat plain-text patch format
' Architecture: Academic Document Format Engine (100% 7-bit ASCII, < 400 lines)
' ==============================================================================

Public Type PatchRefItem
    KeyId As String
    PrefixText As String
    LineNo As Long
End Type

Public Type PatchCiteItem
    RefKey As String
    CtxText As String
    FindText As String
    FixText As String
    HasFix As Boolean
    LineNo As Long
End Type

Public Type PatchFixItem
    CtxText As String
    FindText As String
    NewText As String
    LineNo As Long
End Type

Public Type PatchAcronymItem
    Abbr As String
    FullForm As String
    NoteText As String
    HasNote As Boolean
    LineNo As Long
End Type

Public Type PatchPayload
    Refs() As PatchRefItem
    RefCount As Long
    SkippedRefs As Long
    Cites() As PatchCiteItem
    CiteCount As Long
    SkippedCites As Long
    Fixes() As PatchFixItem
    FixCount As Long
    SkippedFixes As Long
    Acronyms() As PatchAcronymItem
    AcronymCount As Long
    SkippedAcronyms As Long
    HasHeader As Boolean
    HasFooter As Boolean
End Type

Private Sub EnsurePatchCapacities(ByRef p As PatchPayload)
    ReDim p.Refs(1 To 64)
    ReDim p.Cites(1 To 64)
    ReDim p.Fixes(1 To 64)
    ReDim p.Acronyms(1 To 64)
End Sub

Private Sub AppendRef(ByRef p As PatchPayload, ByVal keyId As String, ByVal prefixTxt As String, ByVal ln As Long)
    p.RefCount = p.RefCount + 1
    If p.RefCount > UBound(p.Refs) Then ReDim Preserve p.Refs(1 To p.RefCount + 64)
    p.Refs(p.RefCount).KeyId = keyId
    p.Refs(p.RefCount).PrefixText = prefixTxt
    p.Refs(p.RefCount).LineNo = ln
End Sub

Private Sub AppendCite(ByRef p As PatchPayload, ByRef item As PatchCiteItem)
    p.CiteCount = p.CiteCount + 1
    If p.CiteCount > UBound(p.Cites) Then ReDim Preserve p.Cites(1 To p.CiteCount + 64)
    p.Cites(p.CiteCount) = item
End Sub

Private Sub AppendFix(ByRef p As PatchPayload, ByRef item As PatchFixItem)
    p.FixCount = p.FixCount + 1
    If p.FixCount > UBound(p.Fixes) Then ReDim Preserve p.Fixes(1 To p.FixCount + 64)
    p.Fixes(p.FixCount) = item
End Sub

Private Sub AppendAcronym(ByRef p As PatchPayload, ByRef item As PatchAcronymItem)
    p.AcronymCount = p.AcronymCount + 1
    If p.AcronymCount > UBound(p.Acronyms) Then ReDim Preserve p.Acronyms(1 To p.AcronymCount + 64)
    p.Acronyms(p.AcronymCount) = item
End Sub

Private Function ExtractFieldValue(ByVal lineTxt As String, ByVal prefixColon As String) As Boolean
    Dim pLen As Long
    pLen = Len(prefixColon)
    If Len(lineTxt) >= pLen Then
        If UCase$(Left$(lineTxt, pLen)) = UCase$(prefixColon) Then ExtractFieldValue = True
    End If
End Function

Private Sub FlushOpenBlock(ByRef curMode As String, ByRef cItem As PatchCiteItem, ByRef fItem As PatchFixItem, _
                           ByRef hasOpenCite As Boolean, ByRef hasOpenFix As Boolean, ByRef outPatch As PatchPayload)
    If hasOpenCite Then
        If Len(cItem.RefKey) > 0 Or Len(cItem.CtxText) > 0 Or Len(cItem.FindText) > 0 Or cItem.HasFix Then
            If Len(cItem.RefKey) > 0 And Len(cItem.CtxText) > 0 And Len(cItem.FindText) > 0 Then
                AppendCite outPatch, cItem
            Else
                outPatch.SkippedCites = outPatch.SkippedCites + 1
            End If
        End If
        Dim emptyC As PatchCiteItem
        cItem = emptyC
        hasOpenCite = False
    ElseIf hasOpenFix Then
        If Len(fItem.CtxText) > 0 Or Len(fItem.FindText) > 0 Or Len(fItem.NewText) > 0 Then
            If Len(fItem.CtxText) > 0 And Len(fItem.FindText) > 0 And Len(fItem.NewText) > 0 Then
                AppendFix outPatch, fItem
            Else
                outPatch.SkippedFixes = outPatch.SkippedFixes + 1
            End If
        End If
        Dim emptyF As PatchFixItem
        fItem = emptyF
        hasOpenFix = False
    End If
End Sub

Private Sub ParseRefLine(ByVal trimmed As String, ByVal lineNo As Long, ByRef outPatch As PatchPayload)
    Dim parts() As String, k As String, v As String
    parts = Split(trimmed, "|", 2)
    If UBound(parts) = 1 Then
        k = Trim$(parts(0))
        v = Trim$(parts(1))
        If Len(k) > 0 And Len(v) > 0 Then
            AppendRef outPatch, k, v, lineNo
            Exit Sub
        End If
    End If
    outPatch.SkippedRefs = outPatch.SkippedRefs + 1
End Sub

Private Sub ParseAcronymLine(ByVal trimmed As String, ByVal lineNo As Long, ByRef outPatch As PatchPayload)
    Dim parts() As String, item As PatchAcronymItem
    parts = Split(trimmed, "|", 3)
    If UBound(parts) >= 1 Then
        item.Abbr = Trim$(parts(0))
        item.FullForm = Trim$(parts(1))
        item.LineNo = lineNo
        If UBound(parts) = 2 Then
            item.NoteText = Replace(Trim$(parts(2)), "\n", vbCr)
            item.HasNote = (Len(item.NoteText) > 0)
        End If
        If Len(item.Abbr) > 0 And Len(item.FullForm) > 0 Then
            AppendAcronym outPatch, item
            Exit Sub
        End If
    End If
    outPatch.SkippedAcronyms = outPatch.SkippedAcronyms + 1
End Sub

Private Sub ParseCiteField(ByVal trimmed As String, ByVal lineNo As Long, ByRef cItem As PatchCiteItem, ByRef hasOpen As Boolean)
    hasOpen = True
    If cItem.LineNo = 0 Then cItem.LineNo = lineNo
    If ExtractFieldValue(trimmed, "REF:") Then
        cItem.RefKey = Trim$(Mid$(trimmed, 5))
    ElseIf ExtractFieldValue(trimmed, "CTX:") Then
        cItem.CtxText = Trim$(Mid$(trimmed, 5))
    ElseIf ExtractFieldValue(trimmed, "FIND:") Then
        cItem.FindText = Trim$(Mid$(trimmed, 6))
    ElseIf ExtractFieldValue(trimmed, "FIX:") Or ExtractFieldValue(trimmed, "NEW:") Then
        cItem.FixText = Trim$(Mid$(trimmed, 5))
        cItem.HasFix = (Len(cItem.FixText) > 0)
    End If
End Sub

Private Sub ParseFixField(ByVal trimmed As String, ByVal lineNo As Long, ByRef fItem As PatchFixItem, ByRef hasOpen As Boolean)
    hasOpen = True
    If fItem.LineNo = 0 Then fItem.LineNo = lineNo
    If ExtractFieldValue(trimmed, "CTX:") Then
        fItem.CtxText = Trim$(Mid$(trimmed, 5))
    ElseIf ExtractFieldValue(trimmed, "FIND:") Then
        fItem.FindText = Trim$(Mid$(trimmed, 6))
    ElseIf ExtractFieldValue(trimmed, "NEW:") Or ExtractFieldValue(trimmed, "FIX:") Then
        fItem.NewText = Trim$(Mid$(trimmed, 5))
    End If
End Sub

Private Sub StripFenceAndNormalizeLines(ByVal rawText As String, ByRef outLines() As String)
    Dim norm As String, arr() As String, firstIdx As Long, lastIdx As Long, t As String
    norm = Replace(Replace(rawText, vbCrLf, vbLf), vbCr, vbLf)
    If Len(norm) > 0 Then
        If AscW(Left$(norm, 1)) = &HFEFF Then norm = Mid$(norm, 2)
    End If
    arr = Split(norm, vbLf)
    firstIdx = LBound(arr)
    lastIdx = UBound(arr)
    Do While firstIdx <= lastIdx
        t = Trim$(arr(firstIdx))
        If Len(t) = 0 Or Left$(t, 3) = "```" Or Left$(t, 3) = "~~~" Then
            If Left$(t, 3) = "```" Or Left$(t, 3) = "~~~" Then
                firstIdx = firstIdx + 1
                Exit Do
            End If
            firstIdx = firstIdx + 1
        Else
            Exit Do
        End If
    Loop
    Do While lastIdx >= firstIdx
        t = Trim$(arr(lastIdx))
        If Len(t) = 0 Or Left$(t, 3) = "```" Or Left$(t, 3) = "~~~" Then
            If Left$(t, 3) = "```" Or Left$(t, 3) = "~~~" Then
                lastIdx = lastIdx - 1
                Exit Do
            End If
            lastIdx = lastIdx - 1
        Else
            Exit Do
        End If
    Loop
    If lastIdx < firstIdx Then
        ReDim outLines(0 To 0)
        outLines(0) = ""
        Exit Sub
    End If
    ReDim outLines(0 To (lastIdx - firstIdx))
    Dim k As Long
    For k = firstIdx To lastIdx
        outLines(k - firstIdx) = arr(k)
    Next k
End Sub

Public Function ParseAiPatch(ByVal rawInput As String, ByRef outPatch As PatchPayload) As Boolean
    Dim emptyP As PatchPayload
    outPatch = emptyP
    EnsurePatchCapacities outPatch
    If Len(Trim$(rawInput)) = 0 Then Exit Function
    Dim lines() As String, idx As Long, trimmed As String, up As String
    Dim curMode As String, started As Boolean, hasOpenCite As Boolean, hasOpenFix As Boolean
    Dim curCite As PatchCiteItem, curFix As PatchFixItem
    StripFenceAndNormalizeLines rawInput, lines
    curMode = ""
    For idx = LBound(lines) To UBound(lines)
        trimmed = Trim$(lines(idx))
        up = UCase$(trimmed)
        If Left$(trimmed, 3) = "```" Or Left$(trimmed, 3) = "~~~" Then GoTo NextLn
        If Left$(up, 12) = "=== AF-PATCH" Then
            started = True
            outPatch.HasHeader = True
            GoTo NextLn
        End If
        If up = "=== END ===" Then
            FlushOpenBlock curMode, curCite, curFix, hasOpenCite, hasOpenFix, outPatch
            outPatch.HasFooter = True
            Exit For
        End If
        If up = "@REFS" Or up = "@CITE" Or up = "@FIX" Or up = "@ACRONYMS" Then
            FlushOpenBlock curMode, curCite, curFix, hasOpenCite, hasOpenFix, outPatch
            started = True
            curMode = up
            If curMode = "@CITE" Then
                hasOpenCite = True
                curCite.LineNo = idx + 1
            ElseIf curMode = "@FIX" Then
                hasOpenFix = True
                curFix.LineNo = idx + 1
            End If
            GoTo NextLn
        End If
        If Not started Then GoTo NextLn
        If Len(trimmed) = 0 Then
            FlushOpenBlock curMode, curCite, curFix, hasOpenCite, hasOpenFix, outPatch
            GoTo NextLn
        End If
        If curMode = "@REFS" Then
            ParseRefLine trimmed, idx + 1, outPatch
        ElseIf curMode = "@ACRONYMS" Then
            ParseAcronymLine trimmed, idx + 1, outPatch
        ElseIf curMode = "@CITE" Then
            ParseCiteField trimmed, idx + 1, curCite, hasOpenCite
        ElseIf curMode = "@FIX" Then
            ParseFixField trimmed, idx + 1, curFix, hasOpenFix
        End If
NextLn:
    Next idx
    FlushOpenBlock curMode, curCite, curFix, hasOpenCite, hasOpenFix, outPatch
    ParseAiPatch = (outPatch.HasHeader Or outPatch.RefCount > 0 Or outPatch.CiteCount > 0 Or _
                    outPatch.FixCount > 0 Or outPatch.AcronymCount > 0)
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
