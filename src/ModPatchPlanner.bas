Attribute VB_Name = "ModPatchPlanner"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModPatchPlanner
' Purpose: Dry-Run locator, DOM Range.Find/Binary fallback, Safety Guards & Overlap
' Architecture: Academic Document Format Engine (100% 7-bit ASCII, < 400 lines)
' ==============================================================================

Public Type PlanStats
    RefsTotal As Long
    RefsMatched As Long
    RefsSkipped As Long
    CitesTotal As Long
    CitesPlanned As Long
    CitesAlready As Long
    CitesSkipped As Long
    FixesTotal As Long
    FixesPlanned As Long
    FixesSkipped As Long
    StructsTotal As Long
    StructsPlanned As Long
    StructsSkipped As Long
    OpCount As Long
    WarnNotes As String
End Type

Private Sub AppendWarnNote(ByRef stats As PlanStats, ByVal noteLine As String)
    If Len(noteLine) = 0 Then Exit Sub
    If Len(stats.WarnNotes) > 900 Then Exit Sub
    If Len(stats.WarnNotes) > 0 Then stats.WarnNotes = stats.WarnNotes & vbCrLf
    stats.WarnNotes = stats.WarnNotes & "- " & noteLine
End Sub

Private Function IsParaBoundaryOk(ByVal normTxt As String, ByVal normLen As Long, _
                                  ByVal sPos As Long, ByVal ePos As Long, _
                                  ByVal needLeft As Boolean, ByVal needRight As Boolean) As Boolean
    If needLeft And sPos > 1 Then
        If ModPatchMatcher.IsIdentChar(Mid$(normTxt, sPos - 1, 1)) Then Exit Function
    End If
    If needRight And ePos < normLen Then
        If ModPatchMatcher.IsIdentChar(Mid$(normTxt, ePos + 1, 1)) Then Exit Function
    End If
    IsParaBoundaryOk = True
End Function

Private Function FindHitsInPara(ByRef p As ScannedPara, ByVal ctxNorm As String, ByVal findNorm As String, _
                                ByVal enforceBoundary As Boolean, ByRef outNormS As Long, _
                                ByRef outNormE As Long, ByRef outCtxS As Long, ByRef outCtxE As Long) As Long
    Dim fLen As Long, cLen As Long, needL As Boolean, needR As Boolean
    fLen = Len(findNorm)
    cLen = Len(ctxNorm)
    If fLen = 0 Or p.NormLen < fLen Then Exit Function
    If enforceBoundary Then
        needL = ModPatchMatcher.IsIdentChar(Left$(findNorm, 1))
        needR = ModPatchMatcher.IsIdentChar(Right$(findNorm, 1))
    End If
    Dim hits As Long, lastS As Long, fPos As Long, searchFrom As Long
    If cLen = 0 Then
        searchFrom = 1
        Do While searchFrom <= p.NormLen - fLen + 1
            fPos = InStr(searchFrom, p.NormText, findNorm, vbBinaryCompare)
            If fPos = 0 Then Exit Do
            If Not enforceBoundary Or IsParaBoundaryOk(p.NormText, p.NormLen, fPos, fPos + fLen - 1, needL, needR) Then
                hits = hits + 1
                If hits = 1 Then
                    outNormS = fPos
                    outNormE = fPos + fLen - 1
                    outCtxS = 0
                    outCtxE = 0
                End If
                If hits > 1 Then Exit Do
            End If
            searchFrom = fPos + 1
        Loop
    Else
        Dim ctxFrom As Long, ctxPos As Long, ctxSub As String, subFrom As Long, absS As Long
        ctxFrom = 1
        Do While ctxFrom <= p.NormLen - cLen + 1
            ctxPos = InStr(ctxFrom, p.NormText, ctxNorm, vbBinaryCompare)
            If ctxPos = 0 Then Exit Do
            ctxSub = Mid$(p.NormText, ctxPos, cLen)
            subFrom = 1
            Do While subFrom <= cLen - fLen + 1
                fPos = InStr(subFrom, ctxSub, findNorm, vbBinaryCompare)
                If fPos = 0 Then Exit Do
                absS = ctxPos + fPos - 1
                If absS <> lastS Then
                    If Not enforceBoundary Or IsParaBoundaryOk(p.NormText, p.NormLen, absS, absS + fLen - 1, needL, needR) Then
                        hits = hits + 1
                        lastS = absS
                        If hits = 1 Then
                            outNormS = absS
                            outNormE = absS + fLen - 1
                            outCtxS = ctxPos
                            outCtxE = ctxPos + cLen - 1
                        End If
                        If hits > 1 Then Exit Do
                    End If
                End If
                subFrom = fPos + 1
            Loop
            If hits > 1 Then Exit Do
            ctxFrom = ctxPos + 1
        Loop
    End If
    FindHitsInPara = hits
End Function

Private Function IsBalancedWrapped(ByVal s As String) As Boolean
    If Len(s) < 3 Or Left$(s, 1) <> "(" Or Right$(s, 1) <> ")" Then Exit Function
    Dim k As Long, depth As Long, ch As String
    For k = 2 To Len(s) - 1
        ch = Mid$(s, k, 1)
        If ch = "(" Then
            depth = depth + 1
        ElseIf ch = ")" Then
            depth = depth - 1
            If depth < 0 Then Exit Function
        End If
    Next k
    IsBalancedWrapped = (depth = 0)
End Function

Private Sub TrimOuterCitePunct(ByVal normTxt As String, ByRef normS As Long, ByRef normE As Long, _
                               ByRef repStr As String, ByVal hasRep As Boolean)
    Dim slice As String
    slice = Mid$(normTxt, normS, normE - normS + 1)
    If Right$(slice, 1) = ";" And (Not hasRep Or Right$(repStr, 1) <> ";") And normE > normS Then
        normE = normE - 1
        slice = Mid$(normTxt, normS, normE - normS + 1)
    End If
    If IsBalancedWrapped(slice) Then
        normS = normS + 1
        normE = normE - 1
        If hasRep Then
            If IsBalancedWrapped(repStr) Then repStr = Mid$(repStr, 2, Len(repStr) - 2)
        End If
    End If
End Sub

Private Function LocateSingleBodyOp(ByVal doc As Document, ByRef paras() As ScannedPara, ByVal pCount As Long, _
                                    ByVal opKind As Long, ByVal ctxRaw As String, ByVal findRaw As String, _
                                    ByVal repText As String, ByVal hasRep As Boolean, ByVal targetBm As String, _
                                    ByVal shortTip As String, ByVal srcLine As Long, _
                                    ByRef outOp As PatchOp, ByRef outSkipReason As String) As Boolean
    Dim ctxNorm As String, findNorm As String, j As Long, totalHits As Long, paraHits As Long
    Dim winParaIdx As Long, winS As Long, winE As Long, winCtxS As Long, winCtxE As Long
    Dim tmpS As Long, tmpE As Long, tmpCS As Long, tmpCE As Long
    ctxNorm = ModPatchMatcher.NormalizeStr(ctxRaw)
    findNorm = ModPatchMatcher.NormalizeStr(findRaw)
    If Len(findNorm) = 0 Then
        outSkipReason = "Empty FIND"
        Exit Function
    End If
    Dim isStruct As Boolean, inScope As Boolean
    isStruct = (opKind = OP_KIND_STRUCT)
    For j = 1 To pCount
        If isStruct Then
            inScope = (Not paras(j).IsInRefZone And Not paras(j).IsExcluded)
        Else
            inScope = (paras(j).InBodyScope And Not paras(j).IsInRefZone And Not paras(j).IsHeading And Not paras(j).IsExcluded)
        End If
        If inScope Then
            paraHits = FindHitsInPara(paras(j), ctxNorm, findNorm, (opKind = OP_KIND_FIX Or isStruct), tmpS, tmpE, tmpCS, tmpCE)
            If paraHits > 0 Then
                totalHits = totalHits + paraHits
                If totalHits = 1 Then
                    winParaIdx = j
                    winS = tmpS
                    winE = tmpE
                    winCtxS = tmpCS
                    winCtxE = tmpCE
                Else
                    outSkipReason = "Ambiguous (" & totalHits & "+ hits)"
                    Exit Function
                End If
            End If
        End If
    Next j
    If totalHits = 0 Then
        If isStruct Then outSkipReason = "Not found in document" Else outSkipReason = "Not found in body"
        Exit Function
    End If
    If opKind = OP_KIND_CITE Then TrimOuterCitePunct paras(winParaIdx).NormText, winS, winE, repText, hasRep
    Dim domS As Long, domE As Long, keepLnk As Boolean, synced As Boolean
    If Not ModPatchDom.ResolveDomCoordinates(doc, paras(winParaIdx), winS, winE, winCtxS, winCtxE, domS, domE) Then
        outSkipReason = "DOM mapping failed"
        Exit Function
    End If
    If Not ModPatchDom.ValidateDomGuards(doc, paras(winParaIdx), opKind, domS, domE, targetBm, repText, hasRep, keepLnk, synced, outSkipReason) Then
        Exit Function
    End If
    outOp.Kind = opKind
    outOp.ParaIdxInArr = winParaIdx
    outOp.ParaIndex = paras(winParaIdx).ParaIndex
    outOp.StartPos = domS
    outOp.EndPos = domE
    outOp.ReplaceText = repText
    outOp.HasReplace = hasRep
    outOp.BmName = targetBm
    outOp.ShortTip = shortTip
    outOp.SrcLine = srcLine
    outOp.KeepExistingLink = keepLnk
    outOp.IsAlreadySynced = synced
    LocateSingleBodyOp = True
End Function

Public Sub PlanAllPatchOps(ByVal doc As Document, ByRef payload As PatchPayload, ByRef paras() As ScannedPara, _
                           ByVal pCount As Long, ByVal maxRefNum As Long, ByRef outRefs() As RefTargetInfo, _
                           ByRef outOps() As PatchOp, ByRef outStats As PlanStats)
    Dim dictRefs As Object, i As Long, cap As Long, candOp As PatchOp, reason As String
    Set dictRefs = CreateObject("Scripting.Dictionary")
    outStats.RefsTotal = payload.RefCount + payload.SkippedRefs
    ModPatchMatcher.ResolveRefsTargets doc, payload.Refs, payload.RefCount, paras, pCount, maxRefNum, _
                                       outRefs, dictRefs, outStats.RefsMatched, outStats.RefsSkipped
    outStats.RefsSkipped = outStats.RefsSkipped + payload.SkippedRefs
    For i = 1 To payload.RefCount
        If Not outRefs(i).Found Then AppendWarnNote outStats, "@REFS [" & outRefs(i).KeyId & "]: " & outRefs(i).ErrReason
    Next i
    cap = payload.StructCount + payload.FixCount + payload.CiteCount
    If cap < 1 Then cap = 1
    ReDim outOps(1 To cap)
    outStats.StructsTotal = payload.StructCount + payload.SkippedStructs
    outStats.StructsSkipped = payload.SkippedStructs
    For i = 1 To payload.StructCount
        reason = ""
        If LocateSingleBodyOp(doc, paras, pCount, OP_KIND_STRUCT, payload.Structs(i).CtxText, payload.Structs(i).FindText, _
                              payload.Structs(i).NewText, True, "", "", payload.Structs(i).LineNo, candOp, reason) Then
            outStats.OpCount = outStats.OpCount + 1
            outOps(outStats.OpCount) = candOp
            outStats.StructsPlanned = outStats.StructsPlanned + 1
        Else
            outStats.StructsSkipped = outStats.StructsSkipped + 1
            AppendWarnNote outStats, "@STRUCT L" & payload.Structs(i).LineNo & ": " & reason
        End If
    Next i
    outStats.FixesTotal = payload.FixCount + payload.SkippedFixes
    outStats.FixesSkipped = payload.SkippedFixes
    For i = 1 To payload.FixCount
        reason = ""
        If LocateSingleBodyOp(doc, paras, pCount, OP_KIND_FIX, payload.Fixes(i).CtxText, payload.Fixes(i).FindText, _
                              payload.Fixes(i).NewText, True, "", "", payload.Fixes(i).LineNo, candOp, reason) Then
            outStats.OpCount = outStats.OpCount + 1
            outOps(outStats.OpCount) = candOp
            outStats.FixesPlanned = outStats.FixesPlanned + 1
        Else
            outStats.FixesSkipped = outStats.FixesSkipped + 1
            AppendWarnNote outStats, "@FIX L" & payload.Fixes(i).LineNo & ": " & reason
        End If
    Next i
    outStats.CitesTotal = payload.CiteCount + payload.SkippedCites
    outStats.CitesSkipped = payload.SkippedCites
    Dim tIdx As Long
    For i = 1 To payload.CiteCount
        If Not dictRefs.Exists(payload.Cites(i).RefKey) Then
            outStats.CitesSkipped = outStats.CitesSkipped + 1
            AppendWarnNote outStats, "@CITE L" & payload.Cites(i).LineNo & ": Unknown/unmatched key '" & payload.Cites(i).RefKey & "'"
        Else
            tIdx = CLng(dictRefs(payload.Cites(i).RefKey))
            reason = ""
            If LocateSingleBodyOp(doc, paras, pCount, OP_KIND_CITE, payload.Cites(i).CtxText, payload.Cites(i).FindText, _
                                  payload.Cites(i).FixText, payload.Cites(i).HasFix, outRefs(tIdx).BmName, _
                                  outRefs(tIdx).ShortTip, payload.Cites(i).LineNo, candOp, reason) Then
                If candOp.IsAlreadySynced Then
                    outStats.CitesAlready = outStats.CitesAlready + 1
                Else
                    outStats.OpCount = outStats.OpCount + 1
                    outOps(outStats.OpCount) = candOp
                    outStats.CitesPlanned = outStats.CitesPlanned + 1
                End If
            Else
                outStats.CitesSkipped = outStats.CitesSkipped + 1
                AppendWarnNote outStats, "@CITE L" & payload.Cites(i).LineNo & ": " & reason
            End If
        End If
    Next i
    SortAndFilterOverlaps outOps, outStats
End Sub

Private Sub SortAndFilterOverlaps(ByRef ops() As PatchOp, ByRef stats As PlanStats)
    If stats.OpCount <= 1 Then Exit Sub
    Dim i As Long, j As Long, tmp As PatchOp
    For i = 1 To stats.OpCount - 1
        For j = i + 1 To stats.OpCount
            If ops(j).StartPos < ops(i).StartPos Or _
               (ops(j).StartPos = ops(i).StartPos And ops(j).EndPos < ops(i).EndPos) Then
                tmp = ops(i)
                ops(i) = ops(j)
                ops(j) = tmp
            End If
        Next j
    Next i
    Dim writeIdx As Long, prevEnd As Long
    writeIdx = 0
    prevEnd = -1
    For i = 1 To stats.OpCount
        If ops(i).StartPos < prevEnd Then
            If ops(i).Kind = OP_KIND_FIX Then
                stats.FixesPlanned = stats.FixesPlanned - 1
                stats.FixesSkipped = stats.FixesSkipped + 1
            ElseIf ops(i).Kind = OP_KIND_STRUCT Then
                stats.StructsPlanned = stats.StructsPlanned - 1
                stats.StructsSkipped = stats.StructsSkipped + 1
            Else
                stats.CitesPlanned = stats.CitesPlanned - 1
                stats.CitesSkipped = stats.CitesSkipped + 1
            End If
            AppendWarnNote stats, "L" & ops(i).SrcLine & ": Overlapping span skipped"
        Else
            writeIdx = writeIdx + 1
            ops(writeIdx) = ops(i)
            prevEnd = ops(i).EndPos
        End If
    Next i
    stats.OpCount = writeIdx
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
