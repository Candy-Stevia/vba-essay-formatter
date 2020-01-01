Attribute VB_Name = "ModMain"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModMain
' Purpose: Master Orchestrator connecting the complete 5-Phase Document Format Engine
' Architecture: Academic Document Format Engine (Win32 Unicode Integration)
' ==============================================================================

Public g_SimulateCrashError As Boolean
Public g_IsEngineRunning As Boolean
Public g_EnableDiag As Boolean

Public Sub SetSimulateCrashError(ByVal val As Boolean)
    g_SimulateCrashError = val
End Sub

Public Sub SetEnableDiag(ByVal val As Boolean)
    g_EnableDiag = val
End Sub

Public Sub RunEngine(Optional ByVal targetDoc As Document = Nothing, Optional ByVal runMode As Long = 1)
    If g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        Exit Sub
    End If
    If runMode <> EngineModeFull And runMode <> EngineModeListsOnly Then
        ModRibbonCallbacks.g_LastError = "Invalid runMode: " & CStr(runMode)
        ModUI.MsgBoxW ModConstants.Uni("Invalid run mode (Ch\u1ebf \u0111\u1ed9 ch\u1ea1y kh\u00f4ng h\u1ee3p l\u1ec7): ") & CStr(runMode), _
                      vbExclamation, ModConstants.Uni("Error (L\u1ed7i)")
        Exit Sub
    End If
    If targetDoc Is Nothing Then
        If Application.Documents.Count = 0 Then
            ModUI.MsgBoxW ModConstants.Uni("No open Word document found. (Kh\u00f4ng t\u00ecm th\u1ea5y t\u00e0i li\u1ec7u Word \u0111ang m\u1edf.)"), _
                          vbExclamation, ModConstants.Uni("Error (L\u1ed7i)")
            Exit Sub
        End If
    End If
    If runMode = ModConstants.EngineModeListsOnly Then
        ModQuickUpdate.RunQuickUpdate targetDoc
        Exit Sub
    End If
    g_IsEngineRunning = True
    ModRibbonCallbacks.InvalidateRibbonSafe
    
    On Error GoTo MainErrorHandler
    
    Dim isFullySuccess As Boolean, wasCancelled As Boolean
    isFullySuccess = False
    wasCancelled = False
    Dim diagStarted As Boolean
    diagStarted = False
    Dim pendingErrAlert As String
    pendingErrAlert = ""
    Dim backupPath As String, errMsg As String
    Dim wasTracking As Boolean
    wasTracking = False
    Dim prevScreen As Boolean
    prevScreen = Application.ScreenUpdating
    Dim screenUpdatingDisabled As Boolean
    screenUpdatingDisabled = False
    Dim objUndo As Object, isUndoStarted As Boolean
    isUndoStarted = False
    Dim doc As Document
    
    ModRibbonCallbacks.g_LastError = ""
    ModConstants.g_SilentErrors = 0
    
    ' Resolve Active Document
    If targetDoc Is Nothing Then Set doc = ActiveDocument Else Set doc = targetDoc
    If doc Is Nothing Then
        ModUI.MsgBoxW ModConstants.Uni("No open Word document found. (Kh\u00f4ng t\u00ecm th\u1ea5y t\u00e0i li\u1ec7u Word \u0111ang m\u1edf.)"), vbExclamation, ModConstants.Uni("Error (L\u1ed7i)")
        GoTo CleanExit
    End If
    
    ' Ensure Ribbon Config is Initialized and Snapshot Config for Run Duration
    ModRibbonCallbacks.EnsureConfigInitialized
    Dim cfg As EngineRibbonConfig
    cfg = ModRibbonCallbacks.g_Config
    ModCustomProfile.LoadStyleProfile doc
    If g_EnableDiag Then
        ModDiag.DiagStart doc, "baseline_run"
        diagStarted = True
    End If
    
    ' Step 1: Pre-Flight Document Safety Checks
    If Not ValidatePreFlightConditions(doc, errMsg) Then
        ModUI.MsgBoxW errMsg, vbCritical, ModConstants.Uni("Pre-Flight Failed (Ki\u1ec3m tra th\u1ea5t b\u1ea1i)")
        GoTo CleanExit
    End If
    If g_EnableDiag Then
        ModDiag.DiagLog "PreFlight Validated"
    End If
    
    ' Step 2: Boundary Marker Scan
    Dim markers As MarkerLocations
    If Not ScanMarkers(doc, markers, errMsg) Then
        ModUI.MsgBoxW errMsg, vbExclamation, ModConstants.Uni("Marker Warning (C\u1ea3nh b\u00e1o th\u1ebb)")
        GoTo CleanExit
    End If
    If g_EnableDiag Then
        ModDiag.DiagLog "Markers Scanned"
    End If
    
    ' Step 3: Physical Timestamped Backup
    If Not CreatePhysicalBackup(doc, backupPath, errMsg) Then
        ModUI.MsgBoxW errMsg, vbCritical, ModConstants.Uni("Backup Failed (Sao l\u01b0u th\u1ea5t b\u1ea1i)")
        GoTo CleanExit
    End If
    
    ' Step 4: Track Changes Management (Prompt user if active)
    If Not HandleTrackChanges(doc, wasTracking) Then
        ModBackup.DeleteFileUnicode backupPath
        GoTo CleanExit
    End If
    
    ' Disable Screen Updating for Performance and Visual Stability
    Application.EnableCancelKey = 0
    Application.ScreenUpdating = False
    screenUpdatingDisabled = True
    
    Dim prevViewType As Long, prevSeekView As Long, prevPagination As Boolean, hasWindowView As Boolean
    Dim hasPaginationState As Boolean, winCount As Long
    hasPaginationState = False
    hasWindowView = False
    prevSeekView = 0
    On Error Resume Next
    prevPagination = Application.Options.Pagination
    If Err.Number = 0 Then
        hasPaginationState = True
        Application.Options.Pagination = False
    Else
        Err.Clear
    End If
    winCount = doc.Windows.Count
    If Err.Number <> 0 Then
        winCount = 0
        Err.Clear
    End If
    If winCount > 0 Then
        If Not doc.ActiveWindow Is Nothing Then
            prevSeekView = doc.ActiveWindow.View.SeekView
            If Err.Number <> 0 Then
                prevSeekView = 0
                Err.Clear
            ElseIf prevSeekView <> 0 Then
                doc.ActiveWindow.View.SeekView = 0
                If Err.Number <> 0 Then
                    Err.Clear
                End If
            End If
            prevViewType = doc.ActiveWindow.View.Type
            If Err.Number = 0 Then
                hasWindowView = True
                doc.ActiveWindow.View.Type = 1 ' wdNormalView = 1
                If Err.Number <> 0 Then
                    Err.Clear
                End If
            Else
                Err.Clear
            End If
        End If
    End If
    On Error GoTo MainErrorHandler
    
    ' Simulated Crash Hook for Ephemeral Backup Verification
    If g_SimulateCrashError Then
        Err.Raise vbObjectError + 8888, "ModMain", "Simulated Crash for Ephemeral Backup Verification"
    End If
    
    ' Step 5: Master Undo Record for Full Execution (Encompassing Wipe & Normalize - R-06, R-02)
    isUndoStarted = ModBackup.StartUndoSafe("Academic Format Full Run", objUndo)
    If g_EnableDiag Then
        ModDiag.DiagLog "Master Undo Started"
    End If
    
    ModProgress.ProgressInit "Academic Formatter"
    
    ' Step 6: Wipe Previous Generated Zones (Zero-Drift Idempotency)
    ModProgress.ProgressSetPhase ModConstants.Uni("Wiping Old Zones (X\u00f3a v\u00f9ng c\u0169)"), 5#, 5#
    WipeAllZones doc
    If g_EnableDiag Then
        ModDiag.DiagLog "Step 6: WipeAllZones Completed"
    End If
    
    ' Step 7: Normalize Content Breaks (Full Run)
    ModProgress.ProgressSetPhase ModConstants.Uni("Normalizing Content (Chu\u1ea9n h\u00f3a d\u00f2ng)"), 10#, 15#
    NormalizeContentBreaks doc
    If g_EnableDiag Then
        ModDiag.DiagLog "Step 7: NormalizeContentBreaks Completed"
    End If
    
    ' R-08 Fix: Always re-scan markers after Wipe and Normalize
    If Not ScanMarkers(doc, markers, errMsg) Then
        Err.Raise vbObjectError + 1000, "RunEngine", errMsg
    End If
    
    ' Step 8: Phase 1 Read-Only Snapshot Scan (R-01 check)
    ModProgress.ProgressSetPhase ModConstants.Uni("Scanning Document (Qu\u00e9t t\u00e0i li\u1ec7u)"), 25#, 15#
    Dim snapshots() As ParagraphSnapshot, totalCount As Long, latexWarnings As Long
    If Not ScanDocumentSnapshot(doc, markers, snapshots, totalCount, latexWarnings) Then
        If ModProgress.g_CancelRequested Then
            wasCancelled = True
            GoTo CleanExit
        End If
        Err.Raise vbObjectError + 1001, "RunEngine", "ScanDocumentSnapshot failed"
    End If
    If g_EnableDiag Then
        ModDiag.DiagLog "Step 8: ScanDocumentSnapshot Completed: " & totalCount & " paras"
    End If
    
    ' Resolve Page Sync Mode
    Dim pageMode As Long
    If cfg.IsFooterPage Then pageMode = PageSyncRealFooter Else pageMode = PageSyncPhysicalPage
    
    ' Step 9: Phase 2 Record 1 (R-01 check)
    Dim stats As FormatStats
    ModProgress.ProgressSetPhase ModConstants.Uni("Formatting Content (\u0110\u1ecbnh d\u1ea1ng n\u1ed9i dung)"), 40#, 22#
    If Not ExecuteRecord1(doc, snapshots, totalCount, cfg.FontName, cfg.FontSize, _
                          cfg.LineSpacingMode, cfg.IndentPts, stats) Then
        If ModProgress.g_CancelRequested Then
            wasCancelled = True
            GoTo CleanExit
        End If
        Err.Raise vbObjectError + 1002, "RunEngine", "ExecuteRecord1 failed"
    End If
    If g_EnableDiag Then
        ModDiag.DiagLog "Step 9: ExecuteRecord1 Completed"
    End If
    If doc.Bookmarks.Exists(BM_LOA_HEADER) Then
        ModAcronymBuilder.FormatLoaHeadingPara doc.Bookmarks(BM_LOA_HEADER).Range, cfg.FontName, cfg.FontSize, True
    End If
    ModProgress.ProgressSetPhase ModConstants.Uni("Linking Cross-Refs (G\u1EAFn link tham chi\u1EBFu)"), 62#, 3#
    If markers.HasContentStart Then
        Dim freshCs As Long
        freshCs = ModMarkerLexer.FindMarkerParagraph(doc, "=contentstart()")
        If freshCs > 0 Then
            markers.ContentStartParaIndex = freshCs
        End If
    End If
    ModCrossRef.ApplyCrossReferences doc, markers, snapshots, totalCount, cfg.CrossRefColor
    If g_EnableDiag Then
        ModDiag.DiagLog "Step 9b: ApplyCrossReferences Completed"
    End If
    
    ' Step 10: Phase 3 Record 2 - List of Figures (LOF) (R-01 check)
    ModProgress.ProgressSetPhase ModConstants.Uni("Building Figure List (D\u1ef1ng DM H\u00ecnh)"), 65#, 5#
    If Not BuildListOfFigures(doc, markers, snapshots, totalCount, cfg.FontName, cfg.FontSize, pageMode, False) Then
        Err.Raise vbObjectError + 1003, "RunEngine", "BuildListOfFigures failed"
    End If
    If g_EnableDiag Then
        ModDiag.DiagLog "Step 10: BuildListOfFigures Completed"
    End If
    
    ' Step 11: Phase 4 Record 3 - List of Tables (LOT) (R-01 check)
    ModProgress.ProgressSetPhase ModConstants.Uni("Building Table List (D\u1ef1ng DM B\u1ea3ng)"), 70#, 5#
    If Not BuildListOfTables(doc, markers, snapshots, totalCount, cfg.FontName, cfg.FontSize, pageMode, False) Then
        Err.Raise vbObjectError + 1004, "RunEngine", "BuildListOfTables failed"
    End If
    If g_EnableDiag Then
        ModDiag.DiagLog "Step 11: BuildListOfTables Completed"
    End If
    
    ' Step 12: Phase 5 Record 4 - Table of Contents & 2-Phase Page Sync (R-01 check)
    ModProgress.ProgressSetPhase ModConstants.Uni("Building TOC (D\u1ef1ng M\u1ee5c l\u1ee5c)"), 75#, 5#
    If Not BuildTableOfContents(doc, markers, snapshots, totalCount, cfg.FontName, cfg.FontSize, _
                                pageMode, cfg.IsBoldTOC, cfg.TocDepth, False) Then
        Err.Raise vbObjectError + 1005, "RunEngine", "BuildTableOfContents failed"
    End If
    If g_EnableDiag Then
        ModDiag.DiagLog "Step 12: BuildTableOfContents Completed"
    End If
    
    ' Step 12b: Centralized Single-Pass Repaginate and Sync for all lists
    If markers.HasListFigure Or markers.HasListTable Or markers.HasListContent Then
        If hasWindowView Then
            If doc.Windows.Count > 0 Then
                doc.ActiveWindow.View.Type = 3 ' wdPrintView = 3
            End If
        End If
        ModProgress.ProgressSetPhase ModConstants.Uni("Repaginating (Ph\u00e2n trang)"), 80#, 5#
        ModProgress.ProgressPump
        doc.Repaginate
        If g_EnableDiag Then
            ModDiag.DiagLog "Step 12b: Centralized Repaginate Completed"
        End If
        ModProgress.ProgressSetPhase ModConstants.Uni("Syncing Pages (\u0110\u1ed3ng b\u1ed9 s\u1ed1 trang)"), 85#, 15#
        If Not SyncAllListPages(doc, pageMode) Then
            Err.Raise vbObjectError + 1006, "RunEngine", "SyncAllListPages failed"
        End If
        If g_EnableDiag Then
            ModDiag.DiagLog "Step 12b: SyncAllListPages Completed"
        End If
    End If
    
    ' Step 13: Final Marker Highlighting
    HighlightMarkers doc
    If diagStarted Then
        ModDiag.DiagEnd "SUCCESS"
        diagStarted = False
    End If
    isFullySuccess = True
    GoTo CleanExit

MainErrorHandler:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Dim prevErr As String
    prevErr = ModRibbonCallbacks.g_LastError
    If Len(prevErr) > 0 Then
        ModRibbonCallbacks.g_LastError = "Main Error " & CStr(errNum) & ": " & errDesc & " [" & prevErr & "]"
    Else
        ModRibbonCallbacks.g_LastError = "Main Error " & CStr(errNum) & ": " & errDesc
    End If
    Dim errAlert As String
    errAlert = ModUI.FormatFriendlyError(errNum, errDesc) & vbCrLf & "(" & CStr(errNum) & ")"
    If Len(prevErr) > 0 Then
        errAlert = errAlert & vbCrLf & "[" & prevErr & "]"
    End If
    If Len(backupPath) > 0 Then
        If ModBackup.FileExistsUnicode(backupPath) Then
            errAlert = errAlert & vbCrLf & vbCrLf & _
                       ModConstants.Uni("Safety backup preserved at (B\u1ea3n sao l\u01b0u an to\u00e0n \u0111\u01b0\u1ee3c gi\u1eef t\u1ea1i):") & _
                       vbCrLf & backupPath
        End If
    End If
    pendingErrAlert = errAlert
    Resume CleanExit

CleanExit:
    On Error Resume Next
    Application.EnableCancelKey = 1
    ModProgress.ProgressFinish
    If hasWindowView Then
        If Not doc Is Nothing Then
            If doc.Windows.Count > 0 Then
                If doc.ActiveWindow.View.Type <> prevViewType Then
                    doc.ActiveWindow.View.Type = prevViewType
                End If
                If prevSeekView <> 0 Then
                    doc.ActiveWindow.View.SeekView = prevSeekView
                End If
            End If
        End If
    End If
    If hasPaginationState Then
        Application.Options.Pagination = prevPagination
    End If
    Dim weStartedUndo As Boolean
    weStartedUndo = isUndoStarted
    ModBackup.EndUndoSafe objUndo, isUndoStarted
    If (Not isFullySuccess) And (Len(pendingErrAlert) > 0 Or wasCancelled) And weStartedUndo And (Not doc Is Nothing) Then
        If Not doc.Saved Then
            doc.Undo
        End If
    End If
    If (isFullySuccess And ModConstants.g_SilentErrors = 0) Or wasCancelled Then
        ModBackup.DeleteFileUnicode backupPath
    End If
    If Not doc Is Nothing Then
        RestoreTrackChanges doc, wasTracking
    End If
    If screenUpdatingDisabled Then
        Application.ScreenUpdating = prevScreen
        Application.ScreenRefresh
    End If
    If diagStarted Then
        ModDiag.DiagEnd "FAILED"
        diagStarted = False
    End If
    If wasCancelled Then
        If Application.DisplayAlerts <> 0 Then
            ModUI.MsgBoxW ModConstants.Uni("Operation cancelled safely. (\u0110\u00e3 h\u1ee7y thao t\u00e1c an to\u00e0n.)"), vbInformation, ModConstants.Uni("Cancelled (\u0110\u00e3 h\u1ee7y)")
        End If
    ElseIf Len(pendingErrAlert) > 0 Then
        ModUI.MsgBoxW pendingErrAlert, vbCritical, ModConstants.Uni("Error - Backup Safe (L\u1ed7i - \u0110\u00e3 b\u1ea3o to\u00e0n Backup)")
    ElseIf isFullySuccess Then
        ShowSummaryReport stats, latexWarnings, backupPath
    End If
    ModCustomProfile.ResetStyleProfileCache
    ModCustomProfile.ResetTagCache
    ModTableIndex.InvalidateAllIndexes
    g_IsEngineRunning = False
    ModRibbonCallbacks.InvalidateRibbonSafe
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
