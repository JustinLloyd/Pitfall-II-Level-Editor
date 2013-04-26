VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilepath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_drive As String
Private m_path As String
Private m_filespec As String
Private m_extension As String

Private Function GetDrive(ByVal s As String) As String
    Dim colonPos As Integer
    
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    colonPos = InStr(s, ":")
    If colonPos = 2 Then
        GetDrive = Left$(s, 2)
    Else
        GetDrive = ""
    End If
    
    Exit Function
ErrorHandler:
    ReportError "GetDrive"
    GetDrive = ""
End Function

Private Function GetPath(ByVal s As String) As String
    Dim firstSlashPos As Integer
    Dim lastSlashPos As Integer
    
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    firstSlashPos = InStr(s, "\")
    If firstSlashPos = 0 Then firstSlashPos = InStr(s, "/")
    If firstSlashPos = 0 Then
        GetPath = ""
    Else
        lastSlashPos = InStrRev(s, "\")
        If lastSlashPos = 0 Then lastSlashPos = InStrRev(s, "/")
        If lastSlashPos = firstSlashPos Then
            GetPath = Mid$(s, firstSlashPos, 1)
        Else
            GetPath = Mid$(s, firstSlashPos, lastSlashPos - firstSlashPos + 1)
        End If
    End If
    Exit Function
ErrorHandler:
    ReportError "GetPath"
End Function

Private Function GetFilespec(ByVal s As String) As String
    Dim lastSlashPos As Integer
    Dim lastDotPos As Integer
    
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    lastSlashPos = InStrRev(s, "\")
    If lastSlashPos = 0 Then lastSlashPos = InStrRev(s, "/")
    lastDotPos = InStrRev(s, ".")
    If lastDotPos > 0 Then
        If lastDotPos < lastSlashPos Then
            GetFilespec = Mid$(s, lastSlashPos + 1)
        Else
            GetFilespec = Mid$(s, lastSlashPos + 1, lastDotPos - 1)
        End If
    Else
        GetFilespec = Mid$(s, lastSlashPos + 1)
    End If
    
    Exit Function
ErrorHandler:
    ReportError "GetFilespec"
    GetFilespec = ""
End Function

Private Function GetExtension(ByVal s As String) As String
    Dim ext As String
    Dim dotPos As Integer
    Dim lastSlashPos As Integer
    
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    dotPos = InStrRev(s, ".")
    lastSlashPos = InStrRev(s, "\")
    If dotPos < lastSlashPos Then
        GetExtension = ""
    Else
        GetExtension = Mid$(s, dotPos)
    End If
    Exit Function
ErrorHandler:
    ReportError "GetExtension"
End Function

Public Property Get filepath() As String
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    filepath = m_drive & m_path & m_filespec & m_extension
    Exit Sub
ErrorHandler:
    ReportError "FilePath"
End Property

Public Property Let filepath(ByVal filePathString As String)
    Dim s As String
    Dim pos As Integer
    
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    m_drive = GetDrive(filePathString)
    m_path = GetPath(filePathString)
    m_filespec = GetFilespec(filePathString)
    m_extension = GetExtension(filePathString)
    Exit Property
ErrorHandler:
    ReportError "FilePath"
End Property

Public Property Get Drive() As String
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    Drive = m_drive
    Exit Sub
ErrorHandler:
    ReportError "Drive"
End Property

Public Property Let Drive(ByVal s As String)
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    m_drive = s
    Exit Property
ErrorHandler:
    ReportError "Drive"
End Property

Public Property Get path() As String
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    path = m_path
    Exit Sub
ErrorHandler:
    ReportError "Path"
End Property

Public Property Let path(ByVal s As String)
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    m_path = s
    Exit Property
ErrorHandler:
    ReportError "Path"
End Property

Public Property Get Filespec() As String
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    Filespec = m_filespec
    Exit Sub
ErrorHandler:
    ReportError "Filespec"
End Property

Public Property Let Filespec(ByVal s As String)
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    m_filespec = s
    Exit Property
ErrorHandler:
    ReportError "Filespec"
End Property

Public Property Get Extension() As String
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    Extension = m_extension
    Exit Sub
ErrorHandler:
    ReportError "Extension"
End Property

Public Property Let Extension(ByVal s As String)
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    m_extension = s
    Exit Property
ErrorHandler:
    ReportError "Extension"
End Property

Private Sub Class_Initialize()
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    m_drive = ""
    m_path = ""
    m_filespec = ""
    m_extension = ""
    Exit Sub
ErrorHandler:
    ReportError "Class_Initialize"
End Sub

Public Sub ReplaceDrive(ByVal s As String)
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    m_drive = GetDrive(s)
    Exit Sub
ErrorHandler:
    ReportError "ReplaceDrive"
End Sub

Public Sub ReplacePath(ByVal s As String)
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    m_path = GetPath(s)
    Exit Sub
ErrorHandler:
    ReportError "ReplacePath"
End Sub

Public Sub ReplaceFilespec(ByVal s As String)
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    m_filespec = GetFilespec(s)
    m_extension = GetExtension(s)
    Exit Sub
ErrorHandler:
    ReportError "ReplaceFilespec"
End Sub

Public Sub ReplaceExtension(ByVal s As String)
    On Local Error GoTo ErrorHandler
    Debug.Assert False ' debug me
    m_extension = GetExtension(s)
    Exit Sub
ErrorHandler:
    ReportError "ReplaceExtension"
End Sub
