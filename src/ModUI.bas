Attribute VB_Name = "ModUI"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModUI
' Purpose: Pure-VBA Unicode configuration dialog, Ribbon callback & summary report
' Architecture: Academic Document Format Engine (Win32 MessageBoxW Unicode Support)
' ==============================================================================

#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpText As LongPtr, _
        ByVal lpCaption As LongPtr, _
        ByVal uType As Long) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
#Else
    Private Declare Function MessageBoxW Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal lpText As Long, _
        ByVal lpCaption As Long, _
        ByVal uType As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
#End If

Private Const MB_TASKMODAL As Long = &H2000&

' --- Unicode Native Message Box ---

Public Function MsgBoxW(ByVal prompt As String, _
                        Optional ByVal buttons As Long = vbOKOnly, _
                        Optional ByVal title As String = "") As Long
    On Error Resume Next
    If Application.DisplayAlerts = 0 Then
        Select Case (buttons And &HF&)
            Case vbYesNo
                If (buttons And &H300&) = &H100& Then
                    MsgBoxW = 7 ' vbNo
                Else
                    MsgBoxW = 6 ' vbYes
                End If
            Case vbYesNoCancel
                MsgBoxW = 6 ' vbYes
            Case vbRetryCancel, 6
                MsgBoxW = 2 ' vbCancel
            Case Else
                MsgBoxW = 1 ' vbOK
        End Select
        Exit Function
    End If
    #If VBA7 Then
        Dim ownerHwnd As LongPtr
    #Else
        Dim ownerHwnd As Long
    #End If
    ownerHwnd = GetActiveWindow()
    If ownerHwnd = 0 Then
        If Application.Windows.Count > 0 Then
            ownerHwnd = Application.ActiveWindow.Hwnd
            Err.Clear
        End If
    End If
    Dim effectiveButtons As Long
    effectiveButtons = buttons
    If ownerHwnd = 0 Then
        effectiveButtons = effectiveButtons Or MB_TASKMODAL
    End If
    MsgBoxW = MessageBoxW(ownerHwnd, StrPtr(prompt), StrPtr(title), effectiveButtons)
End Function

' --- Ribbon OnAction Legacy / Fallback Callback ---

Public Sub Ribbon_RunEngine(Optional ByVal control As Object = Nothing)
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then
        Exit Sub
    End If
    If Application.Documents.Count = 0 Then
        MsgBoxW ModConstants.Uni("No open Word document found. (Kh\u00f4ng t\u00ecm th\u1ea5y t\u00e0i li\u1ec7u Word \u0111ang m\u1edf.)"), _
                vbExclamation, ModConstants.Uni("Error (L\u1ed7i)")
        Exit Sub
    End If
    ModRibbonCallbacks.EnsureConfigInitialized
    ModMain.RunEngine ActiveDocument, EngineModeFull
End Sub

' --- Interactive Pure-VBA Configuration Requester (Fallback) ---

Public Function RequestUserConfiguration(ByRef outFontName As String, _
                                         ByRef outFontSize As Single, _
                                         ByRef outPageMode As Long) As Boolean
    Dim promptMsg As String, dlgTitle As String
    promptMsg = ModConstants.Uni("ACADEMIC FORMATTER (CHU\u1ea8N H\u00d3A T\u00c0I LI\u1ec6U)") & vbCrLf & vbCrLf & _
                ModConstants.Uni("Default settings (C\u1ea5u h\u00ecnh m\u1eb7c \u0111\u1ecbnh):") & vbCrLf & _
                ModConstants.Uni("  \u2022 Font (Ph\u00f4ng): Times New Roman") & vbCrLf & _
                ModConstants.Uni("  \u2022 Size (C\u1ee1 ch\u1eef): 13 pt") & vbCrLf & _
                ModConstants.Uni("  \u2022 Page # (S\u1ed1 trang): Footer (Theo ch\u00e2n trang)") & vbCrLf & vbCrLf & _
                ModConstants.Uni("[Yes]: Run now (Ch\u1ea1y ngay)") & vbCrLf & _
                ModConstants.Uni("[No]: Customize font/size (T\u00f9y ch\u1ec9nh ph\u00f4ng/c\u1ee1)") & vbCrLf & _
                ModConstants.Uni("[Cancel]: Abort (H\u1ee7y b\u1ecf)")
                
    dlgTitle = ModConstants.Uni("Format Config (C\u1ea5u h\u00ecnh \u0111\u1ecbnh d\u1ea1ng)")
    
    Dim userChoice As Long
    userChoice = MsgBoxW(promptMsg, vbYesNoCancel + vbQuestion + vbDefaultButton1, dlgTitle)
    
    If userChoice = 2 Then ' vbCancel = 2
        RequestUserConfiguration = False
        Exit Function
    ElseIf userChoice = 6 Then ' vbYes = 6
        outFontName = DEFAULT_FONT_NAME
        outFontSize = DEFAULT_FONT_SIZE
        outPageMode = PageSyncRealFooter
        RequestUserConfiguration = True
        Exit Function
    Else ' vbNo = 7
        Dim customFont As String, customSize As String
        customFont = Trim$(InputBox("Enter font name (Nhap ten phong chu, e.g. Times New Roman, Arial):", "Font (Phong chu)", DEFAULT_FONT_NAME))
        If Len(customFont) = 0 Then customFont = DEFAULT_FONT_NAME
        
        customSize = Trim$(InputBox("Enter font size (Nhap co chu, e.g. 12, 13, 14):", "Font Size (Co chu)", CStr(DEFAULT_FONT_SIZE)))
        If IsNumeric(customSize) Then
            outFontSize = CSng(customSize)
        Else
            outFontSize = DEFAULT_FONT_SIZE
        End If
        outFontName = customFont
        
        Dim pagePrompt As String, pageTitle As String, pageChoice As Long
        pagePrompt = ModConstants.Uni("SELECT PAGE NUMBER MODE (CH\u1eccN KI\u1ec2U S\u1ed0 TRANG):") & vbCrLf & vbCrLf & _
                     ModConstants.Uni("[Yes]: Footer Page # (Theo ch\u00e2n trang - Khuy\u00ean d\u00f9ng).") & vbCrLf & _
                     ModConstants.Uni("[No]: Physical Page # (Theo trang v\u1eadt l\u00fd c\u1ee7a file).")
        pageTitle = ModConstants.Uni("Page Number Mode (Ki\u1ec3u s\u1ed1 trang)")
        pageChoice = MsgBoxW(pagePrompt, vbYesNo + vbQuestion + vbDefaultButton1, pageTitle)
        
        If pageChoice = 6 Then
            outPageMode = PageSyncRealFooter
        Else
            outPageMode = PageSyncPhysicalPage
        End If
        RequestUserConfiguration = True
    End If
End Function

' --- Friendly Error Mapper (#132) ---

Public Function FormatFriendlyError(ByVal errNum As Long, ByVal errDesc As String) As String
    Dim cleanDesc As String
    cleanDesc = Trim$(errDesc)
    Select Case errNum
        Case vbObjectError + 1000
            FormatFriendlyError = ModConstants.Uni("Marker scan failed. Check =contentStart() and list markers (=listContent(), =listTableHere(), =listFigureHere()). (Qu\u00e9t th\u1ebb c\u1ea5u tr\u00fac th\u1ea5t b\u1ea1i. H\u00e3y ki\u1ec3m tra =contentStart() v\u00e0 c\u00e1c th\u1ebb danh m\u1ee5c.)")
            If Len(cleanDesc) > 0 Then
                FormatFriendlyError = FormatFriendlyError & vbCrLf & cleanDesc
            End If
        Case vbObjectError + 1001
            FormatFriendlyError = ModConstants.Uni("Document snapshot scan failed. Save, reopen the document, and check for corrupted paragraphs or tables. (Qu\u00e9t t\u00e0i li\u1ec7u th\u1ea5t b\u1ea1i. H\u00e3y l\u01b0u, m\u1edf l\u1ea1i file v\u00e0 ki\u1ec3m tra \u0111o\u1ea1n/b\u1ea3ng l\u1ed7i.)")
        Case vbObjectError + 1002
            FormatFriendlyError = ModConstants.Uni("Body formatting failed. Check locked sections, Track Changes, or malformed tables. (\u0110\u1ecbnh d\u1ea1ng n\u1ed9i dung th\u1ea5t b\u1ea1i. Ki\u1ec3m tra ph\u1ea7n b\u1ecb kh\u00f3a, Track Changes ho\u1eb7c b\u1ea3ng l\u1ed7i.)")
        Case vbObjectError + 1003
            FormatFriendlyError = ModConstants.Uni("Failed to build List of Figures (LOF). Check the =listFigureHere() marker and figure captions. (D\u1ef1ng Danh m\u1ee5c H\u00ecnh th\u1ea5t b\u1ea1i. Ki\u1ec3m tra th\u1ebb =listFigureHere() v\u00e0 ch\u00fa th\u00edch h\u00ecnh.)")
        Case vbObjectError + 1004
            FormatFriendlyError = ModConstants.Uni("Failed to build List of Tables (LOT). Check the =listTableHere() marker and table captions. (D\u1ef1ng Danh m\u1ee5c B\u1ea3ng th\u1ea5t b\u1ea1i. Ki\u1ec3m tra th\u1ebb =listTableHere() v\u00e0 ti\u00eau \u0111\u1ec1 b\u1ea3ng.)")
        Case vbObjectError + 1005
            FormatFriendlyError = ModConstants.Uni("Failed to build Table of Contents (TOC). Check the =listContent() marker and heading markers. (D\u1ef1ng M\u1ee5c l\u1ee5c th\u1ea5t b\u1ea1i. Ki\u1ec3m tra th\u1ebb =listContent() v\u00e0 c\u00e1c th\u1ebb ti\u00eau \u0111\u1ec1.)")
        Case vbObjectError + 1006
            FormatFriendlyError = ModConstants.Uni("Page number sync failed. Switch to Print Layout, ensure the document is editable, and retry. (\u0110\u1ed3ng b\u1ed9 s\u1ed1 trang th\u1ea5t b\u1ea1i. Chuy\u1ec3n sang Print Layout, \u0111\u1ea3m b\u1ea3o file kh\u00f4ng ch\u1ec9 \u0111\u1ecdc r\u1ed3i th\u1eed l\u1ea1i.)")
        Case 4605
            FormatFriendlyError = ModConstants.Uni("Word command unavailable. Exit Protected View or unlock editing, then retry. (L\u1ec7nh Word kh\u00f4ng kh\u1ea3 d\u1ee5ng. Tho\u00e1t Protected View ho\u1eb7c m\u1edf kh\u00f3a ch\u1ec9nh s\u1eeda r\u1ed3i th\u1eed l\u1ea1i.)")
        Case 5991
            FormatFriendlyError = ModConstants.Uni("Table contains vertically merged cells that Word cannot access directly. Split merged cells or simplify the table. (B\u1ea3ng c\u00f3 \u00f4 g\u1ed9p d\u1ecdc khi\u1ebfn Word kh\u00f4ng truy c\u1eadp \u0111\u01b0\u1ee3c. H\u00e3y t\u00e1ch \u00f4 g\u1ed9p ho\u1eb7c \u0111\u01a1n gi\u1ea3n h\u00f3a b\u1ea3ng.)")
        Case 6124
            FormatFriendlyError = ModConstants.Uni("Document editing is restricted or protected. Disable Protect Document in Review tab, then retry. (T\u00e0i li\u1ec7u \u0111ang b\u1ecb h\u1ea1n ch\u1ebf ch\u1ec9nh s\u1eeda. H\u00e3y t\u1eaft b\u1ea3o v\u1ec7 t\u00e0i li\u1ec7u \u1edf th\u1ebb Review r\u1ed3i th\u1eed l\u1ea1i.)")
        Case Else
            If Len(cleanDesc) > 0 Then
                FormatFriendlyError = cleanDesc
            Else
                FormatFriendlyError = "Error " & CStr(errNum)
            End If
    End Select
End Function

' --- Summary Reports ---

Public Sub ShowSummaryReport(ByRef stats As FormatStats, ByVal latexWarningCount As Long, _
                             Optional ByVal keptBackupPath As String = "")
    Dim msg As String, title As String
    msg = ModConstants.Uni("Formatting complete (Ho\u00e0n t\u1ea5t \u0111\u1ecbnh d\u1ea1ng):") & vbCrLf & _
          "- " & stats.CountLevel1 & " " & ModConstants.Uni("Chapters / Level 1 (Ch\u01b0\u01a1ng / M\u1ee5c 1)") & vbCrLf & _
          "- " & stats.CountSubLevels & " " & ModConstants.Uni("Subheadings (M\u1ee5c con 2, 3, 4)") & vbCrLf & _
          "- " & stats.CountTables & " " & ModConstants.Uni("Tables (B\u1ea3ng bi\u1ec3u)") & vbCrLf & _
          "- " & stats.CountFigures & " " & ModConstants.Uni("Figures (H\u00ecnh \u1ea3nh)")
    Dim styleSum As String
    styleSum = ModCustomProfile.GetActiveStyleSummary()
    If Len(styleSum) > 0 Then
        msg = msg & vbCrLf & ModConstants.Uni("- Active @STYLE profile (H\u1ed3 s\u01a1 @STYLE): ") & styleSum & vbCrLf & _
              ModConstants.Uni("  Note: Saved @STYLE overrides Ribbon Font/Size; use Reset Custom Style to revert. (L\u01b0u \u00fd: @STYLE \u0111ang ghi \u0111\u00e8 Ribbon; d\u00f9ng Reset Custom Style \u0111\u1ec3 v\u1ec1 m\u1eb7c \u0111\u1ecbnh.)")
    End If
          
    If latexWarningCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              ModConstants.Uni("Note: ") & latexWarningCount & _
              ModConstants.Uni(" unclosed LaTeX block(s) safely skipped. (L\u01b0u \u00fd: ") & latexWarningCount & _
              ModConstants.Uni(" kh\u1ed1i LaTeX ch\u01b0a \u0111\u00f3ng \u0111\u00e3 \u0111\u01b0\u1ee3c b\u1ecf qua an to\u00e0n.)")
    End If
    
    If ModConstants.g_SilentErrors > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              ModConstants.Uni("Note: ") & ModConstants.g_SilentErrors & _
              ModConstants.Uni(" minor warning(s). Backup kept. (L\u01b0u \u00fd: ") & ModConstants.g_SilentErrors & _
              ModConstants.Uni(" c\u1ea3nh b\u00e1o ph\u1ee5, \u0111\u00e3 gi\u1eef b\u1ea3n sao.)")
        If Len(keptBackupPath) > 0 Then
            msg = msg & vbCrLf & ModConstants.Uni("  Backup (B\u1ea3n sao): ") & keptBackupPath
        End If
        If Len(ModRibbonCallbacks.g_LastError) > 0 Then
            msg = msg & vbCrLf & "  " & ModRibbonCallbacks.g_LastError
        End If
    End If
    
    title = ModConstants.Uni("Done (Ho\u00e0n t\u1ea5t)")
    MsgBoxW msg, vbInformation, title
End Sub

Public Sub ShowListsUpdatedReport()
    Dim msg As String, title As String
    msg = ModConstants.Uni("TOC, LOT, and LOF updated. Body formatting preserved. (\u0110\u00e3 l\u00e0m m\u1edbi M\u1ee5c l\u1ee5c & Danh m\u1ee5c, gi\u1eef nguy\u00ean th\u00e2n b\u00e0i.)")
    
    If ModConstants.g_SilentErrors > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              ModConstants.Uni("Note: ") & ModConstants.g_SilentErrors & _
              ModConstants.Uni(" minor warning(s). Backup kept. (L\u01b0u \u00fd: ") & ModConstants.g_SilentErrors & _
              ModConstants.Uni(" c\u1ea3nh b\u00e1o ph\u1ee5, \u0111\u00e3 gi\u1eef b\u1ea3n sao.)")
        If Len(ModRibbonCallbacks.g_LastError) > 0 Then
            msg = msg & vbCrLf & "  " & ModRibbonCallbacks.g_LastError
        End If
    End If
    
    title = ModConstants.Uni("Lists Updated (C\u1eadp nh\u1eadt xong)")
    MsgBoxW msg, vbInformation, title
End Sub

Private Function FormatQuickSlotStatus(ByVal slotFound As Boolean, ByVal itemCount As Long, ByVal unitWord As String) As String
    If Not slotFound Then
        FormatQuickSlotStatus = ModConstants.Uni("Not found (Kh\u00f4ng th\u1ea5y)")
    ElseIf itemCount = 0 Then
        FormatQuickSlotStatus = ModConstants.Uni("Found, 0 ") & unitWord & ModConstants.Uni(" (C\u00f3 v\u00f9ng, 0 m\u1ee5c)")
    Else
        FormatQuickSlotStatus = ModConstants.Uni("Updated: ") & itemCount & " " & unitWord & ModConstants.Uni(" (\u0110\u00e3 c\u1eadp nh\u1eadt)")
    End If
End Function

Public Sub ShowQuickUpdateReport(ByVal tocFound As Boolean, ByVal tocItems As Long, _
                                 ByVal lotFound As Boolean, ByVal lotItems As Long, _
                                 ByVal lofFound As Boolean, ByVal lofItems As Long, _
                                 ByVal cntHeadings As Long, ByVal cntTables As Long, _
                                 ByVal cntFigures As Long, ByVal cntNotes As Long)
    Dim msg As String, title As String
    msg = ModConstants.Uni("Quick Update Summary (K\u1ebft qu\u1ea3 c\u1eadp nh\u1eadt nhanh):") & vbCrLf & _
          "- TOC: " & FormatQuickSlotStatus(tocFound, tocItems, ModConstants.Uni("items")) & vbCrLf & _
          "- LOT: " & FormatQuickSlotStatus(lotFound, lotItems, ModConstants.Uni("tables")) & vbCrLf & _
          "- LOF: " & FormatQuickSlotStatus(lofFound, lofItems, ModConstants.Uni("figures")) & vbCrLf & vbCrLf & _
          ModConstants.Uni("Normalized: ") & cntHeadings & " headings, " & cntTables & " tables, " & cntFigures & " figures, " & cntNotes & " notes." & vbCrLf & _
          ModConstants.Uni("(\u0110\u00e3 chu\u1ea9n h\u00f3a: ") & cntHeadings & ModConstants.Uni(" m\u1ee5c, ") & cntTables & ModConstants.Uni(" b\u1ea3ng, ") & cntFigures & ModConstants.Uni(" h\u00ecnh, ") & cntNotes & ModConstants.Uni(" ngu\u1ed3n. Th\u00e2n b\u00e0i gi\u1eef nguy\u00ean 100%.)")
    Dim styleSum As String
    styleSum = ModCustomProfile.GetActiveStyleSummary()
    If Len(styleSum) > 0 Then
        msg = msg & vbCrLf & ModConstants.Uni("- Active @STYLE profile (H\u1ed3 s\u01a1 @STYLE): ") & styleSum & vbCrLf & _
              ModConstants.Uni("  Note: Saved @STYLE overrides Ribbon Font/Size; use Reset Custom Style to revert. (L\u01b0u \u00fd: @STYLE \u0111ang ghi \u0111\u00e8 Ribbon; d\u00f9ng Reset Custom Style \u0111\u1ec3 v\u1ec1 m\u1eb7c \u0111\u1ecbnh.)")
    End If
    If ModConstants.g_SilentErrors > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              ModConstants.Uni("Note: ") & ModConstants.g_SilentErrors & ModConstants.Uni(" minor warning(s) (") & ModConstants.g_SilentErrors & ModConstants.Uni(" c\u1ea3nh b\u00e1o ph\u1ee5).")
        If Len(ModRibbonCallbacks.g_LastError) > 0 Then
            msg = msg & vbCrLf & "  " & ModRibbonCallbacks.g_LastError
        End If
    End If
    title = ModConstants.Uni("Quick Update Done (C\u1eadp nh\u1eadt xong)")
    MsgBoxW msg, vbInformation, title
End Sub

Public Sub ShowMathSummaryReport(ByVal builtCount As Long, ByVal pinkCount As Long)
    Dim msg As String
    msg = ModConstants.Uni("LaTeX Math conversion complete (Ho\u00e0n t\u1ea5t d\u1ef1ng c\u00f4ng th\u1ee9c to\u00e1n):") & vbCrLf & _
          "- " & builtCount & " " & ModConstants.Uni("equation(s) converted (c\u00f4ng th\u1ee9c \u0111\u00e3 d\u1ef1ng)") & vbCrLf & _
          "- " & pinkCount & " " & ModConstants.Uni("Pink item(s) need review (m\u1ee5c H\u1ed3ng c\u1ea7n ki\u1ec3m tra l\u1ea1i)")
    If pinkCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & ModConstants.Uni("Tip: Click inside any Pink item and use Toggle / Edit to fix or force-convert. (G\u1ee3i \u00fd: \u0110\u1eb7t con tr\u1ecf v\u00e0o n\u1ec1n H\u1ed3ng v\u00e0 b\u1ea5m \u0110\u1ed5i / S\u1eeda.)")
    End If
    If ModConstants.g_SilentErrors > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              ModConstants.Uni("Note: ") & ModConstants.g_SilentErrors & _
              ModConstants.Uni(" minor warning(s). Safety backup was preserved. (L\u01b0u \u00fd: ") & ModConstants.g_SilentErrors & _
              ModConstants.Uni(" c\u1ea3nh b\u00e1o ph\u1ee5, \u0111\u00e3 gi\u1eef b\u1ea3n sao an to\u00e0n.)")
        If Len(ModRibbonCallbacks.g_LastError) > 0 Then
            msg = msg & vbCrLf & "  " & ModRibbonCallbacks.g_LastError
        End If
    End If
    MsgBoxW msg, vbInformation, ModConstants.Uni("Done (Ho\u00e0n t\u1ea5t)")
End Sub

Public Function HasPendingMathBookmarks(ByVal doc As Document) As Boolean
    Dim bm As Bookmark
    For Each bm In doc.Bookmarks
        If Left$(bm.Name, 8) = "bm_math_" Then
            HasPendingMathBookmarks = True
            Exit Function
        End If
    Next bm
End Function

Public Function ExportDocumentToPdfPathForTest(ByVal doc As Document, ByVal targetPath As String, _
                                               Optional ByVal openAfter As Boolean = False) As String
    On Error GoTo ErrPdf
    If doc Is Nothing Or Len(Trim$(targetPath)) = 0 Then
        ExportDocumentToPdfPathForTest = "ERR|INVALID_ARG"
        Exit Function
    End If
    Dim pdfPath As String
    pdfPath = Trim$(targetPath)
    If LCase$(Right$(pdfPath, 4)) <> ".pdf" Then pdfPath = pdfPath & ".pdf"
    doc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=17, OpenAfterExport:=openAfter, _
                            OptimizeFor:=0, Range:=0, Item:=0, IncludeDocProps:=True, KeepIRM:=True, _
                            CreateBookmarks:=1, DocStructureTags:=True, BitmapMissingFonts:=True, _
                            UseISO19005_1:=False
    ExportDocumentToPdfPathForTest = "OK|" & pdfPath
    Exit Function
ErrPdf:
    ExportDocumentToPdfPathForTest = "ERR|" & Err.Number & "|" & Err.Description
End Function

Public Sub ExportDocumentToPdf(ByVal doc As Document)
    If doc Is Nothing Then Exit Sub
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then Exit Sub
    ModMain.g_IsEngineRunning = True
    ModRibbonCallbacks.InvalidateRibbonSafe
    On Error GoTo ExportErr
    Dim hasMrk As Boolean, hasMath As Boolean
    hasMrk = ModMarkerLexer.HasAnySystemMarker(doc)
    hasMath = HasPendingMathBookmarks(doc)
    If hasMrk Or hasMath Then
        Dim warnMsg As String
        warnMsg = ModConstants.Uni("Warning: Document still has ")
        If hasMrk And hasMath Then
            warnMsg = warnMsg & ModConstants.Uni("system markers (=...()) and math shading. (C\u1ea3nh b\u00e1o: T\u00e0i li\u1ec7u v\u1eabn c\u00f2n th\u1ebb =...() v\u00e0 n\u1ec1n to\u00e1n h\u1ecdc.)")
        ElseIf hasMrk Then
            warnMsg = warnMsg & ModConstants.Uni("system markers (=...()). (C\u1ea3nh b\u00e1o: T\u00e0i li\u1ec7u v\u1eabn c\u00f2n th\u1ebb \u0111\u00e1nh d\u1ea5u =...().)")
        Else
            warnMsg = warnMsg & ModConstants.Uni("math shading/draft bookmarks. (C\u1ea3nh b\u00e1o: T\u00e0i li\u1ec7u v\u1eabn c\u00f2n n\u1ec1n to\u00e1n h\u1ecdc ch\u01b0a ch\u1ed1t.)")
        End If
        warnMsg = warnMsg & vbCrLf & vbCrLf & _
                  ModConstants.Uni("Continue exporting PDF anyway? Choose No to cancel and run Finalize Math / Clean Markers first. (Ti\u1ebfp t\u1ee5c xu\u1ea5t PDF? Ch\u1ecdn No \u0111\u1ec3 h\u1ee7y v\u00e0 d\u1ecdn tr\u01b0\u1edbc.)")
        If MsgBoxW(warnMsg, vbYesNo + vbExclamation + vbDefaultButton2, ModConstants.Uni("Export PDF Warning (C\u1ea3nh b\u00e1o xu\u1ea5t PDF)")) <> 6 Then
            GoTo ExportCleanup
        End If
    End If
    Dim baseDir As String, baseName As String, lowerName As String, defaultPdf As String
    If Len(Trim$(doc.Path)) = 0 Or ModBackup.IsCloudPath(doc.FullName) Then
        baseDir = ""
        On Error Resume Next
        baseDir = CStr(CreateObject("WScript.Shell").SpecialFolders("MyDocuments"))
        Err.Clear
        On Error GoTo ExportErr
        If Len(Trim$(baseDir)) = 0 Then baseDir = Environ$("USERPROFILE") & "\Documents"
        If Right$(baseDir, 1) <> "\" Then baseDir = baseDir & "\"
    Else
        baseDir = doc.Path
        If Right$(baseDir, 1) <> "\" Then baseDir = baseDir & "\"
    End If
    baseName = doc.Name
    lowerName = LCase$(baseName)
    If Right$(lowerName, 5) = ".docx" Or Right$(lowerName, 5) = ".docm" Then
        baseName = Left$(baseName, Len(baseName) - 5)
    ElseIf Right$(lowerName, 4) = ".doc" Then
        baseName = Left$(baseName, Len(baseName) - 4)
    End If
    defaultPdf = baseDir & baseName & ".pdf"
    Dim fd As Object, k As Long, chosenPath As String
    Set fd = Application.FileDialog(2)
    fd.Title = ModConstants.Uni("Export PDF (Xu\u1ea5t t\u1ec7p PDF)")
    fd.InitialFileName = defaultPdf
    On Error Resume Next
    For k = 1 To fd.Filters.Count
        If InStr(1, LCase$(fd.Filters(k).Extensions), "pdf", vbBinaryCompare) > 0 Then
            fd.FilterIndex = k
            Exit For
        End If
    Next k
    Err.Clear
    On Error GoTo ExportErr
    If fd.Show <> -1 Then GoTo ExportCleanup
    chosenPath = fd.SelectedItems(1)
    Dim res As String
    res = ExportDocumentToPdfPathForTest(doc, chosenPath, True)
    If Left$(res, 4) = "ERR|" Then
        MsgBoxW ModConstants.Uni("Could not export PDF. Check if the target PDF is currently open in Chrome/Edge/Adobe Reader. (Kh\u00f4ng th\u1ec3 xu\u1ea5t PDF. H\u00e3y \u0111\u00f3ng file PDF \u0111ang m\u1edf n\u1ebfu c\u00f3.)") & _
                vbCrLf & vbCrLf & Mid$(res, 5), vbCritical, ModConstants.Uni("Export PDF Failed (L\u1ed7i xu\u1ea5t PDF)")
    End If
ExportCleanup:
    ModMain.g_IsEngineRunning = False
    ModRibbonCallbacks.InvalidateRibbonSafe
    Exit Sub
ExportErr:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    ModMain.g_IsEngineRunning = False
    ModRibbonCallbacks.InvalidateRibbonSafe
    MsgBoxW ModConstants.Uni("Export PDF failed (L\u1ed7i xu\u1ea5t PDF): ") & errDesc, _
            vbCritical, ModConstants.Uni("Export PDF Failed (L\u1ed7i xu\u1ea5t PDF)")
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
