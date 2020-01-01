Attribute VB_Name = "ModUI"
Option Explicit

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
#Else
    Private Declare Function MessageBoxW Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal lpText As Long, _
        ByVal lpCaption As Long, _
        ByVal uType As Long) As Long
#End If

' --- Unicode Native Message Box ---

Public Function MsgBoxW(ByVal prompt As String, _
                        Optional ByVal buttons As Long = vbOKOnly, _
                        Optional ByVal title As String = "") As Long
    On Error Resume Next
    If Application.DisplayAlerts = 0 Then
        If (buttons And vbYesNo) = vbYesNo Then
            MsgBoxW = 6 ' vbYes
        ElseIf (buttons And vbYesNoCancel) = vbYesNoCancel Then
            MsgBoxW = 6 ' vbYes
        Else
            MsgBoxW = 1 ' vbOK
        End If
        Exit Function
    End If
    MsgBoxW = MessageBoxW(0, StrPtr(prompt), StrPtr(title), buttons)
End Function

' --- Ribbon OnAction Legacy / Fallback Callback ---

Public Sub Ribbon_RunEngine(Optional ByVal control As Object = Nothing)
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

' --- Summary Reports ---

Public Sub ShowSummaryReport(ByRef stats As FormatStats, ByVal latexWarningCount As Long)
    Dim msg As String, title As String
    msg = ModConstants.Uni("Formatting complete (Ho\u00e0n t\u1ea5t \u0111\u1ecbnh d\u1ea1ng):") & vbCrLf & _
          "- " & stats.CountLevel1 & " " & ModConstants.Uni("Chapters / Level 1 (Ch\u01b0\u01a1ng / M\u1ee5c 1)") & vbCrLf & _
          "- " & stats.CountSubLevels & " " & ModConstants.Uni("Subheadings (M\u1ee5c con 2, 3, 4)") & vbCrLf & _
          "- " & stats.CountTables & " " & ModConstants.Uni("Tables (B\u1ea3ng bi\u1ec3u)") & vbCrLf & _
          "- " & stats.CountFigures & " " & ModConstants.Uni("Figures (H\u00ecnh \u1ea3nh)")
          
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