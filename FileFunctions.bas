Attribute VB_Name = "FileFunctions"
Option Explicit

Public Const vbffSaveDialog = True
Public Const vbffLoadDialog = False

Private Const cdlOFNOverwritePrompt = &H2
Private Const cdlOFNExplorer = &H80000
Private Const cdlOFNFileMustExist = &H1000

Public Function IsWriteable(ByVal filepath As String) As Boolean
    Debug.Assert False
    IsWriteable = False
End Function

Public Function RemoveExtension(ByVal filepath As String) As String
    Dim Filename As String
    Dim strLen As Long
    Dim pos As Long

    Debug.Assert False
    ' caution: this function cannot handle a name with no extension but a directory has an extension
    Filename = filepath
    ' find last extension delimiter
    pos = InStrRev(Filename, ".")
    ' if no extension delimiter found
    If pos = 0 Then
        ' return entire filepath
        Filename = filepath
    Else
        ' return filepath before last extension delimiter
        Filename = Left$(filepath, pos - 1)
    End If

    RemoveExtension = Filename
End Function

Public Function RemoveFilename(ByVal filepath As String) As String
    Dim path As String
    Dim pos As Integer
    Dim strLen As Integer

    ' this function needs to be debugged and validated
    Debug.Assert False
    strLen = Len(filepath)
    pos = InStrRev(filepath, "\")

    If pos = 0 Then
        pos = strLen
    End If

    path = Left$(filepath, pos)
    RemoveFilename = path
End Function

Public Function RemovePathname(ByVal filepath As String) As String
    Dim Filename As String
    Dim pos As Integer

    Debug.Assert False
    ' find last directory delimiter
    pos = InStrRev(filepath, "\")

    ' if no directory delimiter found
    If pos = 0 Then
        ' filepath just contains a filename, return the entire filepath
        Filename = filepath
    Else
        ' return filepath after last directory delimiter
        Filename = Mid$(filepath, pos + 1)
    End If

    ' set return code to the filename
    RemovePathname = Filename
End Function

Public Function DoesFileExist(ByVal filepath As String) As Boolean
    Dim Filename As String
    
    On Local Error GoTo ErrorHandler
    Filename = Dir(filepath)
    DoesFileExist = Len(Filename) > 0
    Exit Function
ErrorHandler:
    DoesFileExist = False
End Function

Public Function OverwriteIfFileExists(ByVal Filename As String) As Boolean
    Dim s As String
    Dim result As Integer
    
    On Local Error GoTo ErrorHandler
    If DoesFileExist(Filename) Then
        s = Trim(Filename) & " already exists" & vbCrLf
        s = s & "Do you want to replace it?"
        result = MsgBox(s, vbYesNo + vbExclamation + vbDefaultButton2, "Save Level As")
        If result = vbYes Then
            OverwriteIfFileExists = True
        Else
            OverwriteIfFileExists = False
        End If
    End If
    
    Exit Function
ErrorHandler:
    ReportError "OverwriteIfFileExists"
    OverwriteIfFileExists = False
End Function


Public Function InputFilename(ByVal DefaultExt As String, ByVal Filter As String, ByVal title As String, ByVal isSave As Boolean) As String
    Dim fileDlg As New CCommonDialog
    
    On Error GoTo cancelled
    fileDlg.Filename = ""
    fileDlg.DefaultExt = DefaultExt
    fileDlg.Filter = Filter
    fileDlg.FilterIndex = 0
    fileDlg.DialogTitle = title
    fileDlg.CancelError = True
    If isSave Then
        fileDlg.Flags = fileDlg.Flags Or cdlOFNOverwritePrompt
        fileDlg.ShowSave
    Else
        fileDlg.Flags = fileDlg.Flags Or cdlOFNExplorer Or cdlOFNFileMustExist
        fileDlg.ShowOpen
    End If
    
    InputFilename = fileDlg.Filename
    Exit Function
    
cancelled:
    InputFilename = ""
End Function


