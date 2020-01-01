Attribute VB_Name = "ModPatchDom"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModPatchDom
' Purpose: Word DOM Range.Find / Binary Search coordinate mapper and safety guards
' Architecture: Academic Document Format Engine (100% 7-bit ASCII, < 400 lines)
' ==============================================================================

Private Function FindUniqueViaWordFind(ByVal doc As Document, ByVal boundStart As Long, ByVal boundEnd As Long, _
                                       ByVal needle As String, ByRef outStart As Long, ByRef outEnd As Long) As Boolean
    If Len(needle) = 0 Or Len(needle) > 250 Then Exit Function
    Dim r As Range, cnt As Long, prevPos As Long, pos As Long
    pos = boundStart
    Do While pos < boundEnd
        Set r = doc.Range(pos, boundEnd)
        r.Find.ClearFormatting
        If Not r.Find.Execute(FindText:=needle, MatchCase:=True, MatchWholeWord:=False, _
                              MatchWildcards:=False, Forward:=True, Wrap:=0, Format:=False) Then Exit Do
        If r.End > boundEnd Or r.Start < pos Then Exit Do
        cnt = cnt + 1
        outStart = r.Start
        outEnd = r.End
        If cnt > 1 Then Exit Do
        prevPos = pos
        pos = r.End
        If pos <= prevPos Then pos = prevPos + 1
    Loop
    FindUniqueViaWordFind = (cnt = 1)
End Function

Private Function LocateDomViaBinary(ByVal doc As Document, ByRef p As ScannedPara, _
                                    ByVal rawS1 As Long, ByVal rawE1 As Long, _
                                    ByVal wantNorm As String, ByRef outStart As Long, ByRef outEnd As Long) As Boolean
    Dim targetS As Long, targetE As Long, lo As Long, hi As Long, midPos As Long, visLen As Long
    targetS = rawS1 - 1
    targetE = rawE1
    lo = p.RangeStart
    hi = p.RangeEnd
    Do While lo < hi
        midPos = (lo + hi + 1) \ 2
        visLen = Len(doc.Range(p.RangeStart, midPos).Text)
        If visLen <= targetS Then lo = midPos Else hi = midPos - 1
    Loop
    outStart = lo
    lo = outStart
    hi = p.RangeEnd
    Do While lo < hi
        midPos = (lo + hi) \ 2
        visLen = Len(doc.Range(p.RangeStart, midPos).Text)
        If visLen >= targetE Then hi = midPos Else lo = midPos + 1
    Loop
    outEnd = lo
    If outEnd > outStart Then
        If ModPatchMatcher.NormalizeStr(doc.Range(outStart, outEnd).Text) = wantNorm Then LocateDomViaBinary = True
    End If
End Function

Public Function ResolveDomCoordinates(ByVal doc As Document, ByRef p As ScannedPara, _
                                      ByVal normS As Long, ByVal normE As Long, _
                                      ByVal ctxS As Long, ByVal ctxE As Long, _
                                      ByRef outStart As Long, ByRef outEnd As Long) As Boolean
    Dim rawS1 As Long, rawE1 As Long, wantNorm As String, pRng As Range, canDirect As Boolean
    rawS1 = p.PosStart(normS)
    rawE1 = p.PosEnd(normE)
    wantNorm = Mid$(p.NormText, normS, normE - normS + 1)
    Set pRng = doc.Range(p.RangeStart, p.RangeEnd)
    canDirect = ((p.RangeEnd - p.RangeStart) = Len(p.RawText))
    If canDirect And doc.OMaths.Count > 0 Then
        If pRng.OMaths.Count > 0 Then canDirect = False
    End If
    If canDirect Then
        If pRng.Fields.Count > 0 Then canDirect = False
    End If
    If canDirect Then
        outStart = p.RangeStart + rawS1 - 1
        outEnd = p.RangeStart + rawE1
        ResolveDomCoordinates = True
        Exit Function
    End If
    Dim rawSlice As String, ctxDomS As Long, ctxDomE As Long
    rawSlice = Mid$(p.RawText, rawS1, rawE1 - rawS1 + 1)
    If ctxS > 0 And ctxE >= ctxS Then
        Dim rawCtx As String
        rawCtx = Mid$(p.RawText, p.PosStart(ctxS), p.PosEnd(ctxE) - p.PosStart(ctxS) + 1)
        If FindUniqueViaWordFind(doc, p.RangeStart, p.RangeEnd, rawCtx, ctxDomS, ctxDomE) Then
            If FindUniqueViaWordFind(doc, ctxDomS, ctxDomE, rawSlice, outStart, outEnd) Then
                ResolveDomCoordinates = True
                Exit Function
            End If
        End If
    Else
        If FindUniqueViaWordFind(doc, p.RangeStart, p.RangeEnd, rawSlice, outStart, outEnd) Then
            ResolveDomCoordinates = True
            Exit Function
        End If
    End If
    ResolveDomCoordinates = LocateDomViaBinary(doc, p, rawS1, rawE1, wantNorm, outStart, outEnd)
End Function

Public Function ValidateDomGuards(ByVal doc As Document, ByRef p As ScannedPara, ByVal opKind As Long, _
                                  ByVal sPos As Long, ByVal ePos As Long, ByVal targetBm As String, _
                                  ByVal repText As String, ByVal hasRep As Boolean, _
                                  ByRef outKeepLink As Boolean, ByRef outSynced As Boolean, _
                                  ByRef outReason As String) As Boolean
    outKeepLink = False
    outSynced = False
    If ePos <= sPos Then
        outReason = "Empty DOM span"
        Exit Function
    End If
    Dim mRng As Range, hl As Hyperlink
    Set mRng = doc.Range(sPos, ePos)
    If mRng.StoryType <> 1 Then
        outReason = "Not main story"
        Exit Function
    End If
    If mRng.OMaths.Count > 0 Then
        outReason = "Protected OMath span"
        Exit Function
    End If
    If p.IsInTable Then
        If mRng.Cells.Count <> 1 Or ePos > p.RangeEnd - 1 Then
            outReason = "Crosses table cell boundary"
            Exit Function
        End If
        If doc.Range(sPos, sPos).Cells(1).Range.Start <> doc.Range(ePos - 1, ePos - 1).Cells(1).Range.Start Then
            outReason = "Spans multiple table cells"
            Exit Function
        End If
    End If
    If opKind = OP_KIND_FIX Then
        If mRng.Hyperlinks.Count > 0 Or mRng.Fields.Count > 0 Then
            outReason = "Protected field/hyperlink in @FIX"
            Exit Function
        End If
    Else
        If mRng.Hyperlinks.Count = 0 Then
            If mRng.Fields.Count > 0 Then
                outReason = "Protected Word field in @CITE"
                Exit Function
            End If
        ElseIf mRng.Hyperlinks.Count = 1 Then
            Set hl = mRng.Hyperlinks(1)
            If hl.Type = 0 And hl.Address = "" And Left$(hl.SubAddress, 7) = BM_PREFIX_REF And _
               hl.Range.Start <= sPos And hl.Range.End >= ePos And hl.Range.Text = mRng.Text Then
                outKeepLink = True
                If hl.SubAddress = targetBm And (Not hasRep Or mRng.Text = repText) Then outSynced = True
            Else
                outReason = "Overlaps protected hyperlink"
                Exit Function
            End If
        Else
            outReason = "Overlaps multiple hyperlinks"
            Exit Function
        End If
    End If
    ValidateDomGuards = True
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
