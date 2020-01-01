Attribute VB_Name = "ModPatchEngine"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModPatchEngine
' Purpose: Dry-Run UX, Protected Bottom-Up Execution for AF-PATCH v1
' Architecture: Academic Document Format Engine (100% 7-bit ASCII, < 400 lines)
' ==============================================================================

Private Function ResolveCitationApplyColor(ByRef outIsOffFallback As Boolean) As Long
    ModRibbonCallbacks.EnsureConfigInitialized
    Dim cid As String
    cid = g_Config.CrossRefColor
    outIsOffFallback = (cid = "xref_off")
    If Len(cid) = 0 Or outIsOffFallback Then
        ResolveCitationApplyColor = ModCrossRef.ResolveLinkColor("xref_blue")
    Else
        ResolveCitationApplyColor = ModCrossRef.ResolveLinkColor(cid)
    End If
End Function

Private Function BuildDryRunSummary(ByRef payload As PatchPayload, ByRef stats As PlanStats, _
                                    ByVal willBuildLoa As Boolean, ByVal canLoa As Boolean, _
                                    ByVal isOffFallback As Boolean) As String
    Dim s As String
    s = ModConstants.Uni("AI PATCH DRY-RUN (X\u00c1C NH\u1eacN TH\u1ef0C THI):") & vbCrLf & vbCrLf & _
        "- @REFS: " & stats.RefsMatched & "/" & stats.RefsTotal & _
        ModConstants.Uni(" matched (B\u1ecf qua: ") & stats.RefsSkipped & ")" & vbCrLf & _
        "- @CITE: " & stats.CitesPlanned & " new/update, " & stats.CitesAlready & " synced" & _
        ModConstants.Uni(" (B\u1ecf qua: ") & stats.CitesSkipped & ")" & vbCrLf & _
        "- @FIX: " & stats.FixesPlanned & " apply" & _
        ModConstants.Uni(" (B\u1ecf qua: ") & stats.FixesSkipped & ")" & vbCrLf & _
        "- @STRUCT: " & stats.StructsPlanned & " apply" & _
        ModConstants.Uni(" (B\u1ecf qua: ") & stats.StructsSkipped & ")" & vbCrLf
    If payload.AcronymCount <= 0 Then
        s = s & ModConstants.Uni("- @ACRONYMS: 0 (Gi\u1eef nguy\u00ean b\u1ea3ng t\u1eeb vi\u1ebft t\u1eaft hi\u1ec7n c\u00f3)")
    ElseIf canLoa Then
        s = s & "- @ACRONYMS: " & payload.AcronymCount & _
            ModConstants.Uni(" (D\u1ef1ng l\u1ea1i b\u1ea3ng A-Z)")
    Else
        s = s & ModConstants.Uni("- @ACRONYMS: C\u1ea3nh b\u00e1o - kh\u00f4ng t\u00ecm th\u1ea5y =listAcronyms() ho\u1eb7c ti\u00eau \u0111\u1ec1 DM Vi\u1ebft t\u1eaft (s\u1ebd b\u1ecf qua)")
    End If
    If isOffFallback Then
        s = s & vbCrLf & ModConstants.Uni("L\u01b0u \u00fd: Ref Link \u0111ang Off -> Link tr\u00edch d\u1eabn s\u1ebd d\u00f9ng m\u00e0u Xanh d\u01b0\u01a1ng (Blue).")
    End If
    If Len(stats.WarnNotes) > 0 Then
        s = s & vbCrLf & vbCrLf & ModConstants.Uni("Skipped details (Chi ti\u1ebft b\u1ecf qua):") & vbCrLf & stats.WarnNotes
    End If
    BuildDryRunSummary = s
End Function

Private Sub EnsureRefBookmarks(ByVal doc As Document, ByRef targets() As RefTargetInfo, _
                               ByVal targetCount As Long, ByRef paras() As ScannedPara)
    Dim i As Long, pIdx As Long, endPos As Long, bmRng As Range
    For i = 1 To targetCount
        If targets(i).Found And targets(i).RefArrIdx > 0 Then
            pIdx = targets(i).RefArrIdx
            endPos = paras(pIdx).RangeEnd - 1
            If endPos <= paras(pIdx).RangeStart Then endPos = paras(pIdx).RangeStart
            Set bmRng = doc.Range(paras(pIdx).RangeStart, endPos)
            ModBookmark.AddBookmark doc, targets(i).BmName, bmRng
        End If
    Next i
End Sub

Private Sub ApplySingleOpItem(ByVal doc As Document, ByRef op As PatchOp, ByVal linkColor As Long, _
                              ByRef outFixes As Long, ByRef outStructs As Long, ByRef outCites As Long)
    On Error GoTo ErrOp
    Dim mRng As Range, hl As Hyperlink
    Set mRng = doc.Range(op.StartPos, op.EndPos)
    If op.Kind = OP_KIND_FIX Or op.Kind = OP_KIND_STRUCT Then
        mRng.Text = op.ReplaceText
        If op.Kind = OP_KIND_STRUCT Then outStructs = outStructs + 1 Else outFixes = outFixes + 1
    Else
        If op.KeepExistingLink Then
            Set hl = mRng.Hyperlinks(1)
            If op.HasReplace Then
                If StrComp(hl.TextToDisplay, op.ReplaceText, vbBinaryCompare) <> 0 Then
                    hl.TextToDisplay = op.ReplaceText
                End If
            End If
            hl.SubAddress = op.BmName
            hl.ScreenTip = "Ref -> " & op.ShortTip
            hl.Range.Font.Color = linkColor
            hl.Range.Font.Underline = 0
        Else
            If op.HasReplace Then
                If StrComp(mRng.Text, op.ReplaceText, vbBinaryCompare) <> 0 Then
                    mRng.Text = op.ReplaceText
                End If
            End If
            Set hl = doc.Hyperlinks.Add(Anchor:=mRng, Address:="", SubAddress:=op.BmName, _
                                        ScreenTip:="Ref -> " & op.ShortTip)
            If Not hl Is Nothing Then
                hl.Range.Font.Color = linkColor
                hl.Range.Font.Underline = 0
            End If
        End If
        outCites = outCites + 1
    End If
    Exit Sub
ErrOp:
    ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
End Sub

Private Sub RestorePatchUiState(ByVal doc As Document, ByVal hasWinView As Boolean, ByVal prevView As Long, _
                                ByVal hasPag As Boolean, ByVal prevPag As Boolean, _
                                ByVal hasScreen As Boolean, ByVal prevScreen As Boolean)
    On Error Resume Next
    ModProgress.ProgressFinish
    If hasWinView Then
        If Not doc Is Nothing Then
            If doc.Windows.Count > 0 Then doc.ActiveWindow.View.Type = prevView
        End If
    End If
    If hasPag Then Application.Options.Pagination = prevPag
    If hasScreen Then Application.ScreenUpdating = prevScreen
    Err.Clear
End Sub

Public Function ExecuteAiPatchText(ByVal doc As Document, ByVal rawPatch As String, _
                                   ByVal showUiDialogs As Boolean) As String
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        ExecuteAiPatchText = "ERR|BUSY"
        Exit Function
    End If
    If doc Is Nothing Then
        ExecuteAiPatchText = "ERR|NODOC"
        Exit Function
    End If
    ModMain.g_IsEngineRunning = True
    ModRibbonCallbacks.g_LastError = ""
    ModConstants.g_SilentErrors = 0
    On Error GoTo ErrPatch

    Dim payload As PatchPayload
    If Not ModPatchParser.ParseAiPatch(rawPatch, payload) Then
        If showUiDialogs Then
            ModUI.MsgBoxW ModConstants.Uni("Invalid AF-PATCH v1 block. Ensure clipboard contains '=== AF-PATCH v1 ===' and '=== END ==='. (D\u1eef li\u1ec7u patch kh\u00f4ng h\u1ee3p l\u1ec7.)"), _
                          vbExclamation, ModConstants.Uni("AI Patch Error (L\u1ed7i AI Patch)")
        End If
        ExecuteAiPatchText = "ERR|PARSE"
        GoTo CleanExit
    End If

    Dim earlyPfErr As String
    If showUiDialogs Then
        If Not ModBackup.ValidatePreFlightConditions(doc, earlyPfErr) Then
            ModRibbonCallbacks.g_LastError = earlyPfErr
            ModUI.MsgBoxW earlyPfErr, vbExclamation, ModConstants.Uni("Pre-Flight Failed (Ki\u1ec3m tra th\u1ea5t b\u1ea1i)")
            ExecuteAiPatchText = "ERR|PREFLIGHT|" & earlyPfErr
            GoTo CleanExit
        End If
    End If

    Dim markers As MarkerLocations, paras() As ScannedPara, pCount As Long, maxRefNum As Long, scanErr As String
    If Not ModPatchMatcher.ScanDocForPatch(doc, markers, paras, pCount, maxRefNum, scanErr) Then
        If showUiDialogs Then
            ModUI.MsgBoxW scanErr, vbExclamation, ModConstants.Uni("AI Patch Error (L\u1ed7i AI Patch)")
        End If
        ExecuteAiPatchText = "ERR|SCAN|" & scanErr
        GoTo CleanExit
    End If

    Dim refTargets() As RefTargetInfo, ops() As PatchOp, stats As PlanStats
    ModPatchPlanner.PlanAllPatchOps doc, payload, paras, pCount, maxRefNum, refTargets, ops, stats
    If Len(stats.WarnNotes) > 0 Then ModRibbonCallbacks.g_LastError = stats.WarnNotes

    Dim willBuildLoa As Boolean, canLoa As Boolean, isOffFallback As Boolean, linkColor As Long
    willBuildLoa = (payload.AcronymCount >= 1)
    If willBuildLoa Then canLoa = ModAcronymBuilder.CanLocateAcronymTarget(doc)
    linkColor = ResolveCitationApplyColor(isOffFallback)

    Dim hasWork As Boolean
    hasWork = (stats.RefsMatched > 0 Or stats.OpCount > 0 Or (willBuildLoa And canLoa))
    If Not hasWork And stats.CitesAlready = 0 Then
        If showUiDialogs Then
            ModUI.MsgBoxW BuildDryRunSummary(payload, stats, willBuildLoa, canLoa, isOffFallback), _
                          vbInformation, ModConstants.Uni("AI Patch Dry-Run")
        End If
        ExecuteAiPatchText = "NOOP|" & stats.RefsSkipped & "|" & stats.CitesSkipped & "|" & stats.FixesSkipped & "|" & stats.StructsSkipped
        GoTo CleanExit
    End If

    If showUiDialogs Then
        Dim promptMsg As String, resp As Long
        promptMsg = BuildDryRunSummary(payload, stats, willBuildLoa, canLoa, isOffFallback) & vbCrLf & vbCrLf & _
                    ModConstants.Uni("Proceed to apply patch to document? (Ti\u1ebfn h\u00e0nh \u00e1p d\u1ee5ng v\u00e0o t\u00e0i li\u1ec7u?)")
        resp = ModUI.MsgBoxW(promptMsg, vbYesNo + vbQuestion + vbDefaultButton1, _
                             ModConstants.Uni("Confirm AI Patch (X\u00e1c nh\u1eadn AI Patch)"))
        If resp <> 6 Then
            ExecuteAiPatchText = "CANCELED"
            GoTo CleanExit
        End If
    End If

    Dim backupPath As String, wasTracking As Boolean, prevScreen As Boolean, hasScreenSaved As Boolean
    Dim prevViewType As Long, prevPagination As Boolean, hasWindowView As Boolean, hasPaginationState As Boolean
    Dim objUndo As Object, undoStarted As Boolean, protectedOk As Boolean
    If Not ModBackup.BeginProtectedSession(doc, backupPath, wasTracking) Then
        ExecuteAiPatchText = "ERR|PREFLIGHT|" & ModRibbonCallbacks.g_LastError
        GoTo CleanExit
    End If
    protectedOk = True
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    hasScreenSaved = True
    On Error Resume Next
    prevPagination = Application.Options.Pagination
    If Err.Number = 0 Then
        hasPaginationState = True
        Application.Options.Pagination = False
    End If
    Err.Clear
    If doc.Windows.Count > 0 Then
        prevViewType = doc.ActiveWindow.View.Type
        If Err.Number = 0 Then
            hasWindowView = True
            doc.ActiveWindow.View.Type = 1
        End If
    End If
    Err.Clear
    On Error GoTo ErrPatch
    ModProgress.ProgressInit "AI Patch Engine"
    ModProgress.ProgressSetPhase ModConstants.Uni("Applying Patches (\u0110ang \u00e1p d\u1ee5ng b\u1ea3n v\u00e1)"), 10#, 75#
    undoStarted = ModBackup.StartUndoSafe(ModConstants.Uni("Apply AI Patch (Ch\u1ea1y AI Patch)"), objUndo)

    Dim fixesDone As Long, structsDone As Long, citesDone As Long, acroBuilt As Long, i As Long
    EnsureRefBookmarks doc, refTargets, payload.RefCount, paras
    For i = stats.OpCount To 1 Step -1
        ModProgress.ProgressUpdate (stats.OpCount - i + 1), stats.OpCount
        ApplySingleOpItem doc, ops(i), linkColor, fixesDone, structsDone, citesDone
    Next i
    If stats.CitesAlready > 0 And Not isOffFallback Then
        ModPatchMatcher.SyncCitationLinkColors doc, linkColor
    End If
    If willBuildLoa And canLoa Then
        ModProgress.ProgressSetPhase ModConstants.Uni("Building Acronyms (D\u1ef1ng DM Vi\u1ebft t\u1eaft)"), 85#, 15#
        ModAcronymBuilder.BuildAcronymTable doc, payload.Acronyms, payload.AcronymCount, acroBuilt
    End If

    ModBackup.EndUndoSafe objUndo, undoStarted
    RestorePatchUiState doc, hasWindowView, prevViewType, hasPaginationState, prevPagination, hasScreenSaved, prevScreen
    ModBackup.EndProtectedSession doc, backupPath, wasTracking, (ModConstants.g_SilentErrors > 0)
    protectedOk = False

    ExecuteAiPatchText = "OK|" & stats.RefsMatched & "|" & citesDone & "|" & stats.CitesAlready & "|" & _
                         stats.CitesSkipped & "|" & fixesDone & "|" & stats.FixesSkipped & "|" & acroBuilt & "|" & _
                         structsDone & "|" & stats.StructsSkipped
    If showUiDialogs Then
        Dim doneMsg As String
        doneMsg = ModConstants.Uni("AI Patch applied (Ho\u00e0n t\u1ea5t \u00e1p d\u1ee5ng AI Patch):") & vbCrLf & _
                  "- @REFS bookmarks: " & stats.RefsMatched & vbCrLf & _
                  "- @CITE linked/updated: " & citesDone & " (" & stats.CitesAlready & " already synced, " & stats.CitesSkipped & " skipped)" & vbCrLf & _
                  "- @FIX replaced: " & fixesDone & " (" & stats.FixesSkipped & " skipped)" & vbCrLf & _
                  "- @STRUCT replaced: " & structsDone & " (" & stats.StructsSkipped & " skipped)" & vbCrLf & _
                  "- @ACRONYMS built: " & acroBuilt
        If ModConstants.g_SilentErrors > 0 Then
            doneMsg = doneMsg & vbCrLf & ModConstants.Uni("Note: minor warnings occurred; crash backup retained.")
        End If
        ModUI.MsgBoxW doneMsg, vbInformation, ModConstants.Uni("AI Patch Done (Ho\u00e0n t\u1ea5t)")
    End If
    GoTo CleanExit

ErrPatch:
    ModConstants.g_SilentErrors = ModConstants.g_SilentErrors + 1
    ModRibbonCallbacks.g_LastError = "AI Patch error (" & Err.Number & "): " & Err.Description
    On Error Resume Next
    If undoStarted Then ModBackup.EndUndoSafe objUndo, undoStarted
    RestorePatchUiState doc, hasWindowView, prevViewType, hasPaginationState, prevPagination, True, True
    If protectedOk Then ModBackup.EndProtectedSession doc, backupPath, wasTracking, True
    If showUiDialogs Then
        ModUI.MsgBoxW ModRibbonCallbacks.g_LastError, vbCritical, ModConstants.Uni("AI Patch Error (L\u1ed7i)")
    End If
    ExecuteAiPatchText = "ERR|RUNTIME|" & Err.Description

CleanExit:
    ModMain.g_IsEngineRunning = False
End Function

Public Sub RunAiPatchClipboard()
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then Exit Sub
    If Application.Documents.Count = 0 Then
        ModUI.MsgBoxW ModConstants.Uni("No open Word document found. (Kh\u00f4ng t\u00ecm th\u1ea5y t\u00e0i li\u1ec7u Word \u0111ang m\u1edf.)"), _
                      vbExclamation, ModConstants.Uni("Error (L\u1ed7i)")
        Exit Sub
    End If
    Dim clipOk As Boolean, rawTxt As String
    rawTxt = ModClipboard.GetClipboardTextUnicode(clipOk)
    If Not clipOk Or Len(Trim$(rawTxt)) = 0 Then
        ModUI.MsgBoxW ModConstants.Uni("Clipboard does not contain Unicode text. Copy the AI patch block first. (Clipboard tr\u1ed1ng ho\u1eb7c kh\u00f4ng c\u00f3 v\u0103n b\u1ea3n Unicode.)"), _
                      vbExclamation, ModConstants.Uni("Clipboard Empty")
        Exit Sub
    End If
    ExecuteAiPatchText ActiveDocument, rawTxt, True
End Sub

Public Sub RunAiPatchFile()
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then Exit Sub
    If Application.Documents.Count = 0 Then
        ModUI.MsgBoxW ModConstants.Uni("No open Word document found. (Kh\u00f4ng t\u00ecm th\u1ea5y t\u00e0i li\u1ec7u Word \u0111ang m\u1edf.)"), _
                      vbExclamation, ModConstants.Uni("Error (L\u1ed7i)")
        Exit Sub
    End If
    Dim errMsg As String, canceled As Boolean, rawTxt As String
    rawTxt = ModClipboard.PickAndReadPatchFile(errMsg, canceled)
    If canceled Then Exit Sub
    If Len(errMsg) > 0 Or Len(Trim$(rawTxt)) = 0 Then
        ModUI.MsgBoxW IIf(Len(errMsg) > 0, errMsg, "Empty patch file."), vbExclamation, ModConstants.Uni("File Error (L\u1ed7i File)")
        Exit Sub
    End If
    ExecuteAiPatchText ActiveDocument, rawTxt, True
End Sub

Public Function ExecuteAiPatchSilentForTest(ByVal doc As Document, ByVal rawPatch As String) As String
    ExecuteAiPatchSilentForTest = ExecuteAiPatchText(doc, rawPatch, False)
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
