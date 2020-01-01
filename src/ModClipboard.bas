Attribute VB_Name = "ModClipboard"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModClipboard
' Purpose: Win32 32/64-bit Unicode Clipboard R/W & UTF-8 text file reader
' Architecture: Academic Document Format Engine (Zero UserForm, 100% 7-bit ASCII)
' ==============================================================================

Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE As Long = &H2

#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub WinSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
    Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
    Private Declare Sub WinSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

Private Function TryOpenClipboardWithRetry() As Boolean
    Dim attempt As Long
    For attempt = 1 To 4
        If OpenClipboard(0) <> 0 Then
            TryOpenClipboardWithRetry = True
            Exit Function
        End If
        WinSleep 30
    Next attempt
End Function

Public Function GetClipboardTextUnicode(Optional ByRef outOk As Boolean = False) As String
    outOk = False
    If Not TryOpenClipboardWithRetry() Then Exit Function
    On Error GoTo CleanClose
    #If VBA7 Then
        Dim hMem As LongPtr, lpStr As LongPtr
    #Else
        Dim hMem As Long, lpStr As Long
    #End If
    Dim cLen As Long, buf As String, isLocked As Boolean
    isLocked = False
    If IsClipboardFormatAvailable(CF_UNICODETEXT) <> 0 Then
        hMem = GetClipboardData(CF_UNICODETEXT)
        If hMem <> 0 Then
            lpStr = GlobalLock(hMem)
            If lpStr <> 0 Then
                isLocked = True
                cLen = lstrlenW(lpStr)
                If cLen > 0 Then
                    buf = String$(cLen, vbNullChar)
                    CopyMem StrPtr(buf), lpStr, CLng(cLen) * 2
                    GetClipboardTextUnicode = buf
                End If
                GlobalUnlock hMem
                isLocked = False
                outOk = True
            End If
        End If
    End If
CleanClose:
    On Error Resume Next
    If isLocked And hMem <> 0 Then
        GlobalUnlock hMem
        isLocked = False
    End If
    CloseClipboard
    Err.Clear
End Function

Public Function SetClipboardTextUnicode(ByVal textVal As String) As Boolean
    On Error GoTo CleanFail
    #If VBA7 Then
        Dim hMem As LongPtr, lpDst As LongPtr, hRes As LongPtr
    #Else
        Dim hMem As Long, lpDst As Long, hRes As Long
    #End If
    Dim withNull As String, byteLen As Long
    Dim isLocked As Boolean, isClipOpen As Boolean, ownsHMem As Boolean
    isLocked = False
    isClipOpen = False
    ownsHMem = False
    withNull = textVal & vbNullChar
    byteLen = LenB(withNull)
    hMem = GlobalAlloc(GMEM_MOVEABLE, byteLen)
    If hMem = 0 Then Exit Function
    ownsHMem = True
    lpDst = GlobalLock(hMem)
    If lpDst = 0 Then GoTo CleanFail
    isLocked = True
    CopyMem lpDst, StrPtr(withNull), byteLen
    GlobalUnlock hMem
    isLocked = False
    If Not TryOpenClipboardWithRetry() Then GoTo CleanFail
    isClipOpen = True
    If EmptyClipboard() = 0 Then GoTo CleanFail
    hRes = SetClipboardData(CF_UNICODETEXT, hMem)
    If hRes = 0 Then GoTo CleanFail
    ownsHMem = False
    SetClipboardTextUnicode = True
CleanFail:
    On Error Resume Next
    If isLocked And hMem <> 0 Then
        GlobalUnlock hMem
        isLocked = False
    End If
    If ownsHMem And hMem <> 0 Then
        GlobalFree hMem
        ownsHMem = False
    End If
    If isClipOpen Then
        CloseClipboard
        isClipOpen = False
    End If
    Err.Clear
End Function

Public Function ReadUtf8TextFile(ByVal filePath As String, ByRef outErr As String) As String
    On Error GoTo ErrRead
    outErr = ""
    Dim stm As Object, content As String
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    content = stm.ReadText(-1)
    stm.Close
    If Len(content) > 0 Then
        If (AscW(Left$(content, 1)) And &HFFFF&) = &HFEFF& Then content = Mid$(content, 2)
    End If
    ReadUtf8TextFile = content
    Exit Function
ErrRead:
    outErr = ModConstants.Uni("File read error (L\u1ed7i \u0111\u1ecdc file): ") & Err.Description
    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
    Err.Clear
End Function

Public Function PickAndReadPatchFile(ByRef outErr As String, ByRef outCanceled As Boolean) As String
    outCanceled = False
    outErr = ""
    Dim fd As Object
    On Error GoTo ErrDlg
    Set fd = Application.FileDialog(3)
    fd.Title = ModConstants.Uni("Select AI Patch File (Ch\u1ecdn file Patch AI)")
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "AI Patch Text", "*.txt; *.patch"
    fd.Filters.Add "All Files", "*.*"
    If fd.Show <> -1 Then
        outCanceled = True
        Exit Function
    End If
    PickAndReadPatchFile = ReadUtf8TextFile(fd.SelectedItems(1), outErr)
    Exit Function
ErrDlg:
    outErr = ModConstants.Uni("Cannot open file dialog (Kh\u00f4ng m\u1edf \u0111\u01b0\u1ee3c h\u1ed9p tho\u1ea1i file): ") & Err.Description
    Err.Clear
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
