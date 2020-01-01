Attribute VB_Name = "ModBackup"
Option Explicit

' ==============================================================================
' Module: ModBackup
' Purpose: Pre-flight document validation, Unicode safe physical backup & cleanup
' Architecture: Academic Document Format Engine (Win32 Unicode Safe)
' ==============================================================================

' --- Cloud Path Detection ---

Public Function IsCloudPath(ByVal fullPath As String) As Boolean
    Dim p As String
    p = LCase$(Trim$(fullPath))
    If Left$(p, 7) = "http://" Or Left$(p, 8) = "https://" Then
        IsCloudPath = True
    Else
        IsCloudPath = False
    End If
End Function

' --- Pre-Flight Conditions Validation ---

Public Function ValidatePreFlightConditions(ByVal doc As Document, ByRef outErrorMessage As String) As Boolean
    outErrorMessage = ""
    
    ' 1. Check whether document was ever saved to disk
    If Len(Trim$(doc.Path)) = 0 Then
        outErrorMessage = ModConstants.Uni("Document is not saved to disk. Press Ctrl+S first. (T\u00e0i li\u1ec7u ch\u01b0a \u0111\u01b0\u1ee3c l\u01b0u tr\u00ean \u1ed5 \u0111\u0129a, h\u00e3y l\u01b0u tr\u01b0\u1edbc.)")
        ValidatePreFlightConditions = False
        Exit Function
    End If
    
    ' 2. Check for OneDrive / SharePoint cloud monikers
    If IsCloudPath(doc.FullName) Then
        outErrorMessage = ModConstants.Uni("Cloud file (OneDrive/SharePoint) is not supported. Save a local copy first. (Kh\u00f4ng h\u1ed7 tr\u1ee3 file tr\u00ean m\u1ea1ng, h\u00e3y l\u01b0u v\u1ec1 m\u00e1y tr\u01b0\u1edbc.)")
        ValidatePreFlightConditions = False
        Exit Function
    End If
    
    ' 3. Check Read-Only state
    If doc.ReadOnly Then
        outErrorMessage = ModConstants.Uni("Document is Read-Only. Unlock it first. (T\u00e0i li\u1ec7u \u0111ang \u1edf ch\u1ebf \u0111\u1ed9 ch\u1ec9 \u0111\u1ecdc, h\u00e3y m\u1edf kh\u00f3a tr\u01b0\u1edbc.)")
        ValidatePreFlightConditions = False
        Exit Function
    End If
    
    ' 4. Check Document Protection (wdNoProtection = -1)
    If doc.ProtectionType <> -1 Then
        outErrorMessage = ModConstants.Uni("Document is protected. Unprotect it first. (T\u00e0i li\u1ec7u \u0111ang b\u1ecb kh\u00f3a b\u1ea3o v\u1ec7, h\u00e3y b\u1ecf kh\u00f3a tr\u01b0\u1edbc.)")
        ValidatePreFlightConditions = False
        Exit Function
    End If
    
    ' 5. Check unsaved changes
    If Not doc.Saved Then
        outErrorMessage = ModConstants.Uni("Unsaved changes detected. Press Ctrl+S first. (T\u00e0i li\u1ec7u c\u00f3 thay \u0111\u1ed5i ch\u01b0a l\u01b0u, h\u00e3y nh\u1ea5n Ctrl+S tr\u01b0\u1edbc.)")
        ValidatePreFlightConditions = False
        Exit Function
    End If
    
    ValidatePreFlightConditions = True
End Function

' --- Backup Filepath Generation ---

Public Function GenerateBackupFilePath(ByVal doc As Document) As String
    Dim folderPath As String, fileName As String, baseName As String, ext As String
    Dim dotPos As Long, timeStamp As String
    
    folderPath = doc.Path
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    fileName = doc.Name
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        baseName = Left$(fileName, dotPos - 1)
        ext = Mid$(fileName, dotPos)
    Else
        baseName = fileName
        ext = ".docx"
    End If
    timeStamp = Format$(Now, "yyyymmdd_hhnnss")
    GenerateBackupFilePath = folderPath & baseName & "_crash_backup_" & timeStamp & ext
End Function

' --- Unicode Native File Helpers (Scripting.FileSystemObject) ---

Public Function FileExistsUnicode(ByVal filePath As String) As Boolean
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If fso.FileExists(filePath) Then
            If fso.GetFile(filePath).Size > 0 Then FileExistsUnicode = True
        End If
    End If
End Function

Public Sub DeleteFileUnicode(ByVal filePath As String)
    On Error Resume Next
    If Len(filePath) = 0 Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
    End If
End Sub

' --- Physical Backup Execution ---

Public Function CreatePhysicalBackup(ByVal doc As Document, ByRef outBackupPath As String, ByRef outErrorMessage As String) As Boolean
    On Error GoTo ErrorHandler
    outErrorMessage = ""
    outBackupPath = GenerateBackupFilePath(doc)
    
    ' 1. Attempt WordBasic.CopyFileA first
    On Error Resume Next
    WordBasic.CopyFileA doc.FullName, outBackupPath
    On Error GoTo ErrorHandler
    
    ' 2. Fallback to FSO CopyFile if destination does not exist (Unicode path)
    If Not FileExistsUnicode(outBackupPath) Then
        On Error Resume Next
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso Is Nothing Then fso.CopyFile doc.FullName, outBackupPath, True
        On Error GoTo ErrorHandler
    End If
    
    ' 3. Deterministic verification that destination file physically exists with content
    If FileExistsUnicode(outBackupPath) Then
        CreatePhysicalBackup = True
    Else
        outErrorMessage = ModConstants.Uni("Backup failed: file not found at (Sao l\u01b0u th\u1ea5t b\u1ea1i, kh\u00f4ng th\u1ea5y file t\u1ea1i): ") & outBackupPath
        CreatePhysicalBackup = False
    End If
    Exit Function

ErrorHandler:
    outErrorMessage = ModConstants.Uni("Backup error (L\u1ed7i sao l\u01b0u) (") & Err.Number & "): " & Err.Description
    CreatePhysicalBackup = False
End Function

' --- Track Changes Management ---

Public Function HandleTrackChanges(ByVal doc As Document, ByRef outWasTracking As Boolean) As Boolean
    outWasTracking = doc.TrackRevisions
    If outWasTracking Then
        Dim resp As Long
        resp = ModUI.MsgBoxW(ModConstants.Uni("Track Changes is ON. Temporarily pause it while formatting and restore afterward? (Ch\u1ebf \u0111\u1ed9 Track Changes \u0111ang b\u1eadt. T\u1ea1m t\u1eaft trong l\u00fac \u0111\u1ecbnh d\u1ea1ng r\u1ed3i b\u1eadt l\u1ea1i sau nh\u00e9?)"), _
                             vbYesNo + vbQuestion, ModConstants.Uni("Track Changes (Theo d\u00f5i thay \u0111\u1ed5i)"))
        If resp <> 6 Then ' vbYes = 6
            HandleTrackChanges = False
            Exit Function
        End If
        doc.TrackRevisions = False
    End If
    HandleTrackChanges = True
End Function

Public Sub RestoreTrackChanges(ByVal doc As Document, ByVal wasTracking As Boolean)
    If wasTracking Then
        doc.TrackRevisions = True
    End If
End Sub
