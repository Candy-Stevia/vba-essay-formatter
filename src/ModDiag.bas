Attribute VB_Name = "ModDiag"
Option Explicit

' ==============================================================================
' Module: ModDiag
' Purpose: Low-overhead in-process diagnostic telemetry for Word VBA
' Architecture: Win32 API GetProcessMemoryInfo & QueryPerformanceCounter
' Compliance: M-05 (< 300 lines, compound = 0, Option Explicit, 7-bit ASCII)
' ==============================================================================

#If VBA7 Then
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
Private Declare PtrSafe Function GetProcessMemoryInfo Lib "psapi.dll" ( _
    ByVal hProcess As LongPtr, _
    ByRef pCounters As PROCESS_MEMORY_COUNTERS_EX, _
    ByVal cb As Long) As Long

Private Type PROCESS_MEMORY_COUNTERS_EX
    cb As Long
    PageFaultCount As Long
#If Win64 Then
    PeakWorkingSetSize As LongLong
    WorkingSetSize As LongLong
    QuotaPeakPagedPoolUsage As LongLong
    QuotaPagedPoolUsage As LongLong
    QuotaPeakNonPagedPoolUsage As LongLong
    QuotaNonPagedPoolUsage As LongLong
    PagefileUsage As LongLong
    PeakPagefileUsage As LongLong
    PrivateUsage As LongLong
#Else
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
    PrivateUsage As Long
#End If
End Type
#Else
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
Private Declare Function GetProcessMemoryInfo Lib "psapi.dll" ( _
    ByVal hProcess As Long, _
    ByRef pCounters As PROCESS_MEMORY_COUNTERS_EX, _
    ByVal cb As Long) As Long

Private Type PROCESS_MEMORY_COUNTERS_EX
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
    PrivateUsage As Long
End Type
#End If

Private g_T0 As Currency
Private g_Freq As Currency
Private g_UseQpc As Boolean
Private g_TimerT0 As Single
Private g_LogPath As String
Private g_RunId As String
Private g_Enabled As Boolean

Public Function GetBitness() As String
#If Win64 Then
    GetBitness = "VBA7-64bit"
#Else
    GetBitness = "VBA-32bit"
#End If
End Function

Public Sub DiagStart(ByVal doc As Document, Optional ByVal customRunId As String = "")
    On Error Resume Next
    g_Enabled = True
    g_UseQpc = False
    
    If QueryPerformanceFrequency(g_Freq) <> 0 Then
        If g_Freq > 0 Then
            If QueryPerformanceCounter(g_T0) <> 0 Then
                g_UseQpc = True
            End If
        End If
    End If
    If Not g_UseQpc Then
        g_TimerT0 = Timer
    End If
    
    If Len(customRunId) > 0 Then
        g_RunId = customRunId
    Else
        g_RunId = Format$(Now, "yyyymmdd_hhnnss")
    End If
    
    g_LogPath = Environ$("TEMP") & "\academic_perf_" & g_RunId & ".log"
    
    Dim pCount As Long
    Dim tCount As Long
    Dim bCount As Long
    Dim vCount As Long
    pCount = 0
    tCount = 0
    bCount = 0
    vCount = 0
    If Not doc Is Nothing Then
        pCount = doc.Paragraphs.Count
        tCount = doc.Tables.Count
        bCount = doc.Bookmarks.Count
        vCount = doc.Variables.Count
    End If
    Dim initLine As String
    initLine = "START | runId=" & g_RunId & " | " & GetBitness() & _
               " | Word " & Application.Version & _
               " | paras=" & pCount & " | tables=" & tCount & _
               " | bms=" & bCount & " | vars=" & vCount
    DiagLog initLine
    On Error GoTo 0
End Sub

Public Sub DiagLog(ByVal phaseName As String)
    If Not g_Enabled Then
        Exit Sub
    End If
    On Error Resume Next
    
    Dim elapsedSec As Double
    Dim timeOk As Boolean
    elapsedSec = 0#
    timeOk = False
    
    If g_UseQpc Then
        Dim tNow As Currency
        If QueryPerformanceCounter(tNow) <> 0 Then
            elapsedSec = CDbl(tNow - g_T0) / CDbl(g_Freq)
            timeOk = True
        End If
    End If
    
    If Not timeOk Then
        Dim curTimer As Single
        curTimer = Timer
        If curTimer < g_TimerT0 Then
            elapsedSec = CDbl((curTimer + 86400!) - g_TimerT0)
        Else
            elapsedSec = CDbl(curTimer - g_TimerT0)
        End If
    End If
    
    Dim wsStr As String
    Dim privStr As String
    GetProcessMemoryStrings wsStr, privStr
    
    Dim logLine As String
    logLine = Format$(elapsedSec, "0.000") & "s | ws=" & _
              wsStr & " | priv=" & privStr & " | " & phaseName
              
    Dim fNum As Integer
    fNum = FreeFile
    Open g_LogPath For Append As #fNum
    Print #fNum, logLine
    Close #fNum
    On Error GoTo 0
End Sub

Public Sub DiagEnd(ByVal finalStatus As String)
    If Not g_Enabled Then
        Exit Sub
    End If
    DiagLog "FINISH | status=" & finalStatus
    g_Enabled = False
End Sub

Public Function GetCurrentLogPath() As String
    GetCurrentLogPath = g_LogPath
End Function

Private Sub GetProcessMemoryStrings(ByRef outWsStr As String, ByRef outPrivStr As String)
    outWsStr = "n/a"
    outPrivStr = "n/a"
    On Error Resume Next
    
    Dim pmc As PROCESS_MEMORY_COUNTERS_EX
    pmc.cb = LenB(pmc)
    
    #If VBA7 Then
    Dim hProc As LongPtr
    #Else
    Dim hProc As Long
    #End If
    hProc = GetCurrentProcess()
    
    If GetProcessMemoryInfo(hProc, pmc, pmc.cb) <> 0 Then
        outWsStr = Format$(CDbl(pmc.WorkingSetSize) / 1048576#, "0.0") & "MB"
        outPrivStr = Format$(CDbl(pmc.PrivateUsage) / 1048576#, "0.0") & "MB"
    End If
    On Error GoTo 0
End Sub
