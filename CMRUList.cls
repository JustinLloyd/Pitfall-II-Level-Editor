VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMRUList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const k_MRU_BASENAME = "MRU"
Private Const k_MRU_SECTION = "MRU"
Private Const k_MRU_EMPTY_KEY = "Empty"
Private Const k_MRU_MAX_ITEMS = 4

Private m_mruList As Collection
Private m_menuObject(0 To k_MRU_MAX_ITEMS) As Menu
Private m_menuDefined As Boolean

'************************************************************************************
'                                                                                   *
' Name: BuildMRUMenu                                                                *
' Purpose: Updates an application's MRU menu options.                               *
'                                                                                   *
'                                                                                   *
'************************************************************************************

Public Sub BuildMRUMenu()
    Dim filename As String
    Dim mruEntry As String
    Dim index As Integer
    
    If m_menuDefined = False Then
        Exit Sub
    End If
    
    m_menuObject(0).Visible = False
    m_menuObject(0).Visible = False
    If m_mruList.Count = 0 Then
        Exit Sub
    End If
    
    m_menuObject(0).Visible = True
    m_menuObject(0).Caption = "-"
    For index = 1 To m_mruList.Count
        filename = m_mruList.Item(index)
        mruEntry = "&" & Trim(index) & " " & Trim(filename)
        m_menuObject(index).Caption = mruEntry
        m_menuObject(index).Visible = True
    Next
    
End Sub

Public Sub Init(Optional ByVal menuObject As Object)
    Dim index As Integer
    
    On Error GoTo InitError
    For index = 0 To k_MRU_MAX_ITEMS
        On Error Resume Next
        Load menuObject(index)
        On Error GoTo InitError
        Set m_menuObject(index) = menuObject.Item(index)
        m_menuObject(index).Visible = False
    Next
    
    m_menuDefined = True
    ReadList
    BuildMRUMenu
    On Error GoTo 0
    Exit Sub
    
InitError:
    Debug.Print "Did not set menu"
    On Error GoTo 0
End Sub

'************************************************************************************
'                                                                                   *
' Name: ClearList                                                                   *
' Purpose: Clears the MRU list from the Windows registry.                           *
'                                                                                   *
'                                                                                   *
'************************************************************************************

Public Sub ClearList()
    Dim index As Integer
    Dim keyName As String
    
    ' debug this function
    Debug.Assert False
    If m_mruList.Count = 0 Then
        Exit Sub
    End If
    
    For index = 1 To m_mruList.Count
        keyName = k_MRU_BASENAME & Trim(index)
        DeleteSetting App.title, k_MRU_SECTION, keyName
    Next
    
End Sub

'************************************************************************************
'                                                                                   *
' Name: ReadList                                                                    *
' Purpose: Reads the MRU list from the Windows registry.                            *
'                                                                                   *
'                                                                                   *
'************************************************************************************

Public Sub ReadList()
    Dim filename As String
    Dim keyName As String
    Dim index As Integer
    
    Set m_mruList = New Collection
    For index = 1 To k_MRU_MAX_ITEMS
        keyName = k_MRU_BASENAME & Trim(index)
        filename = GetSetting(App.title, k_MRU_SECTION, keyName, k_MRU_EMPTY_KEY)
        If filename <> k_MRU_EMPTY_KEY Then
            m_mruList.Add filename, Trim(index)
        End If
        
    Next
    
End Sub

'************************************************************************************
'                                                                                   *
' Name: WriteList                                                                   *
' Purpose: Writes the MRU list to the Windows registry.                             *
'                                                                                   *
'                                                                                   *
'************************************************************************************

Public Sub WriteList()
    Dim keyName As String
    Dim index As Integer
    
    On Error GoTo errorWriting
    
    If m_mruList.Count = 0 Then
        Exit Sub
    End If
    
    For index = 1 To m_mruList.Count
        keyName = k_MRU_BASENAME & Trim(index)
        SaveSetting App.title, k_MRU_SECTION, keyName, Trim(m_mruList.Item(index))
    Next
    
    Exit Sub
errorWriting:
    Stop
End Sub

'************************************************************************************
'                                                                                   *
' Name: AddItem                                                                     *
' Purpose: Adds an item to the MRU list.                                            *
'                                                                                   *
'                                                                                   *
'************************************************************************************

Public Sub AddItem(ByVal filename As String)
    Dim index As Integer
    Dim keyName As String
    Dim mruEntry As String
    Dim filenameExists As Boolean
    
    ' look for filename in mru list
    filenameExists = False
    If m_mruList.Count > 0 Then
        For index = 1 To m_mruList.Count
            mruEntry = m_mruList.Item(index)
            If UCase(mruEntry) = UCase(filename) Then
                filenameExists = True
                Exit For
            End If
        Next
        
        ' if filename exists in mru list
        If filenameExists Then
            ' remove filename from mru list
            m_mruList.Remove index
        End If
    End If
    
    
    ' add filename at top of mru list
    If m_mruList.Count > 0 Then
        m_mruList.Add filename, , 1
    Else
        m_mruList.Add filename
    End If
    
    ' if mru list > maximum allowed items
    If m_mruList.Count > k_MRU_MAX_ITEMS Then
        ' remove last item from list
        m_mruList.Remove m_mruList.Count
    End If
    
    ' write mru list to registry
    WriteList
    ' rebuild mru menu
    BuildMRUMenu
End Sub

Public Function Item(ByVal index As Integer) As String
    Debug.Assert index >= 1 And index <= k_MRU_MAX_ITEMS
    If index > m_mruList.Count Then
        Item = ""
    Else
        Item = m_mruList.Item(index)
    End If
    
End Function

Private Sub Class_Initialize()
    Dim index As Integer
    
    m_menuDefined = False
    For index = 0 To k_MRU_MAX_ITEMS
        Set m_menuObject(index) = Nothing
    Next
End Sub
