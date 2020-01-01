Attribute VB_Name = "ModProgress"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModProgress
' Purpose: Pure-VBA Responsive Progress Controller & Message Pump
' Compliance: M-05 (< 400 lines, compound = 0, Option Explicit, 7-bit ASCII)
' ==============================================================================

#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
#End If

Public g_CancelRequested As Boolean

Private g_RunStartTimer As Double
Private g_LastPumpTime As Single
Private g_ActiveTitle As String
Private g_CurrentPhase As String
Private g_PhaseBasePct As Single
Private g_PhaseWeightPct As Single
Private g_CursorSet As Boolean
Private g_PrevDisplayStatusBar As Boolean
Private g_HasSavedStatusBar As Boolean

Private Const PUMP_INTERVAL_SEC As Single = 0.35
Private Const VK_ESCAPE As Long = &H1B

Public Function GetElapsedSec() As Double
    Dim nowT As Double
    nowT = Timer
    If nowT < g_RunStartTimer Then
        GetElapsedSec = (86400# - g_RunStartTimer) + nowT
    Else
        GetElapsedSec = nowT - g_RunStartTimer
    End If
End Function

Private Function FormatElapsedTag() As String
    FormatElapsedTag = " [" & Format$(GetElapsedSec(), "0.0") & "s]"
End Function

Private Sub PollEscapeCancel()
    On Error Resume Next
    If g_CancelRequested Then
        Exit Sub
    End If
    If Application.DisplayAlerts = 0 Then
        Exit Sub
    End If
    If Application.Windows.Count = 0 Then
        Exit Sub
    End If
    If GetForegroundWindow() = Application.ActiveWindow.Hwnd Then
        If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
            g_CancelRequested = True
        End If
    End If
    Err.Clear
End Sub

Public Sub ProgressInit(ByVal title As String)
    On Error Resume Next
    g_CancelRequested = False
    g_RunStartTimer = Timer
    g_ActiveTitle = title
    g_CurrentPhase = "Initializing"
    g_PhaseBasePct = 0#
    g_PhaseWeightPct = 10#
    g_LastPumpTime = Timer
    
    If Not g_HasSavedStatusBar Then
        g_PrevDisplayStatusBar = Application.DisplayStatusBar
        If Err.Number = 0 Then
            g_HasSavedStatusBar = True
        Else
            Err.Clear
        End If
    End If
    Application.DisplayStatusBar = True
    Application.System.Cursor = wdCursorWait
    g_CursorSet = True
    
    Application.StatusBar = "[>                   ] 0% | " & title & ": " & ModConstants.Uni("Starting... (\u0110ang kh\u1edfi t\u1ea1o...)") & FormatElapsedTag()
    DoEvents
End Sub

Public Sub ProgressSetPhase(ByVal phaseName As String, ByVal basePct As Single, ByVal weightPct As Single)
    On Error Resume Next
    g_CurrentPhase = phaseName
    g_PhaseBasePct = basePct
    g_PhaseWeightPct = weightPct
    
    Dim pctInt As Long
    pctInt = CLng(basePct)
    If pctInt > 100 Then
        pctInt = 100
    End If
    
    Dim barStr As String
    barStr = BuildProgressBar(pctInt)
    
    Application.StatusBar = barStr & " " & CStr(pctInt) & "% | " & phaseName & FormatElapsedTag()
    DoEvents
    g_LastPumpTime = Timer
End Sub

Public Sub ProgressUpdate(ByVal currentItem As Long, ByVal totalItems As Long, Optional ByVal subDetail As String = "")
    On Error Resume Next
    Dim nowT As Single
    nowT = Timer
    
    Dim shouldPump As Boolean
    shouldPump = False
    
    If (nowT - g_LastPumpTime >= PUMP_INTERVAL_SEC) Then
        shouldPump = True
    ElseIf (nowT < g_LastPumpTime) Then
        shouldPump = True
    End If
    
    If Not shouldPump Then
        Exit Sub
    End If
    
    PollEscapeCancel
    
    Dim localRatio As Single
    If totalItems > 0 Then
        localRatio = CSng(currentItem) / CSng(totalItems)
        If localRatio > 1# Then
            localRatio = 1#
        End If
    Else
        localRatio = 0#
    End If
    
    Dim totalPct As Single
    totalPct = g_PhaseBasePct + (localRatio * g_PhaseWeightPct)
    If totalPct > 100# Then
        totalPct = 100#
    End If
    
    Dim pctInt As Long
    pctInt = CLng(totalPct)
    
    Dim barStr As String
    barStr = BuildProgressBar(pctInt)
    
    Dim msg As String
    msg = barStr & " " & CStr(pctInt) & "% | " & g_CurrentPhase
    If Len(subDetail) > 0 Then
        msg = msg & " (" & subDetail & ")"
    ElseIf totalItems > 0 Then
        msg = msg & " (" & CStr(currentItem) & "/" & CStr(totalItems) & ")"
    End If
    msg = msg & FormatElapsedTag()
    
    Application.StatusBar = msg
    DoEvents
    g_LastPumpTime = Timer
End Sub

Public Sub ProgressPump()
    On Error Resume Next
    DoEvents
    g_LastPumpTime = Timer
End Sub

Public Sub ProgressFinish()
    On Error Resume Next
    g_CancelRequested = False
    If g_CursorSet Then
        Application.System.Cursor = wdCursorNormal
        g_CursorSet = False
    End If
    Application.StatusBar = ""
    If g_HasSavedStatusBar Then
        Application.DisplayStatusBar = g_PrevDisplayStatusBar
        g_HasSavedStatusBar = False
    End If
End Sub

Private Function BuildProgressBar(ByVal pct As Long) As String
    Dim filled As Long
    filled = pct \ 5
    If filled > 20 Then
        filled = 20
    ElseIf filled < 0 Then
        filled = 0
    End If
    
    Dim emptyCount As Long
    emptyCount = 20 - filled
    
    BuildProgressBar = "[" & String$(filled, "=") & String$(emptyCount, " ") & "]"
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
