VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const vbffSaveDialog = 2
Private Const vbffOpenDialog = 1
 
Private Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'constants for API memory functions
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFilename) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OpenFilename) As Long
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PrintDlg) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

'constants for color dialog
Private Const CDERR_DIALOGFAILURE = &HFFFF
Private Const CDERR_FINDRESFAILURE = &H6
Private Const CDERR_GENERALCODES = &H0
Private Const CDERR_INITIALIZATION = &H2
Private Const CDERR_LOADRESFAILURE = &H7
Private Const CDERR_LOADSTRFAILURE = &H5
Private Const CDERR_LOCKRESFAILURE = &H8
Private Const CDERR_MEMALLOCFAILURE = &H9
Private Const CDERR_MEMLOCKFAILURE = &HA
Private Const CDERR_NOHINSTANCE = &H4
Private Const CDERR_NOHOOK = &HB
Private Const CDERR_NOTEMPLATE = &H3
Private Const CDERR_REGISTERMSGFAIL = &HC
Private Const CDERR_STRUCTSIZE = &H1


'constants for file dialog
Private Const FNERR_BUFFERTOOSMALL = &H3003
Private Const FNERR_FILENAMECODES = &H3000
Private Const FNERR_INVALIDFILENAME = &H3002
Private Const FNERR_SUBCLASSFAILURE = &H3001

'data buffer for the ChooseColor function
Private Type ChooseColor
        m_lStructSize As Long
        m_hwndOwner As Long
        m_hInstance As Long
        m_rgbResult As Long
        m_lpCustColors As Long
        m_flags As Long
        m_lCustData As Long
        m_lpfnHook As Long
        m_lpTemplateName As String
End Type

'constants for LOGFONT
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64
Private Const FW_BOLD = 700

'data buffer for the ChooseFont function
Private Type LOGFONT
        m_lfHeight As Long
        m_lfWidth As Long
        m_lfEscapement As Long
        m_lfOrientation As Long
        m_lfWeight As Long
        m_lfItalic As Byte
        m_lfUnderline As Byte
        m_lfStrikeOut As Byte
        m_lfCharSet As Byte
        m_lfOutPrecision As Byte
        m_lfClipPrecision As Byte
        m_lfQuality As Byte
        m_lfPitchAndFamily As Byte
        m_lfFaceName(LF_FACESIZE) As Byte
End Type

'data buffer for the ChooseFont function
Private Type ChooseFont
        m_lStructSize As Long
        m_hwndOwner As Long
        m_hdc As Long
        m_lpLogFont As Long
        m_iPointSize As Long
        m_flags As Long
        m_rgbColors As Long
        m_lCustData As Long
        m_lpfnHook As Long
        m_lpTemplateName As String
        m_hInstance As Long
        m_lpszStyle As String
        m_nFontType As Integer
        m_MISSING_ALIGNMENT As Integer
        m_nSizeMin As Long
        m_nSizeMax As Long
End Type


'data buffer for the GetOpenFileName and GetSaveFileName functions
Private Type OpenFilename
        m_lStructSize As Long
        m_hwndOwner As Long
        m_hInstance As Long
        m_lpstrFilter As String
        m_lpstrCustomFilter As String
        m_nMaxCustFilter As Long
        m_iFilterIndex As Long
        m_lpstrFile As String
        m_nMaxFile As Long
        m_lpstrFileTitle As String
        m_nMaxFileTitle As Long
        m_lpstrInitialDir As String
        m_lpstrTitle As String
        m_flags As Long
        m_nFileOffset As Integer
        m_nFileExtension As Integer
        m_lpstrDefExt As String
        m_lCustData As Long
        m_lpfnHook As Long
        m_lpTemplateName As String
End Type


'data buffer for the PrintDlg function
Private Type PrintDlg
        m_lStructSize As Long
        m_hwndOwner As Long
        m_hDevMode As Long
        m_hDevNames As Long
        m_hdc As Long
        m_flags As Long
        m_nFromPage As Integer
        m_nToPage As Integer
        m_nMinPage As Integer
        m_nMaxPage As Integer
        m_nCopies As Integer
        m_hInstance As Long
        m_lCustData As Long
        m_lpfnPrintHook As Long
        m_lpfnSetupHook As Long
        m_lpPrintTemplateName As String
        m_lpSetupTemplateName As String
        m_hPrintTemplate As Long
        m_hSetupTemplate As Long
End Type


'internal property buffers

Private m_iAction As Integer         'internal buffer for Action property
Private m_bCancelError As Boolean    'internal buffer for CancelError property
Private m_lColor As Long             'internal buffer for Color property
Private m_lCopies As Long            'internal buffer for lCopies property
Private m_sDefaultExt As String      'internal buffer for sDefaultExt property
Private m_sDialogTitle As String     'internal buffer for DialogTitle property
Private m_sFileName As String        'internal buffer for FileName property
Private m_sFileTitle As String       'internal buffer for FileTitle property
Private m_sFilter As String          'internal buffer for Filter property
Private m_iFilterIndex As Integer    'internal buffer for FilterIndex property
Private m_lFlags As Long             'internal buffer for Flags property
Private m_bFontBold As Boolean       'internal buffer for FontBold property
Private m_bFontItalic As Boolean     'internal buffer for FontItalic property
Private m_sFontName As String        'internal buffer for FontName property
Private m_lFontSize As Long          'internal buffer for FontSize property
Private m_bFontStrikethru As Boolean 'internal buffer for FontStrikethru property
Private m_bFontUnderline As Boolean  'internal buffer for FontUnderline property
Private m_lFromPage As Long          'internal buffer for FromPage property
Private m_lhdc As Long               'internal buffer for hdc property
Private m_lHelpCommand As Long       'internal buffer for HelpCommand property
Private m_sHelpContext As String     'internal buffer for HelpContext property
Private m_sHelpFile As String        'internal buffer for HelpFile property
Private m_sHelpKey As String         'internal buffer for HelpKey property
Private m_sInitDir As String         'internal buffer for InitDir property
Private m_lMax As Long               'internal buffer for Max property
Private m_lMaxFileSize As Long       'internal buffer for MaxFileSize property
Private m_lMin As Long               'internal buffer for Min property
Private m_objObject As Object        'internal buffer for Object property
Private m_iPrinterDefault As Integer 'internal buffer for PrinterDefault property
Private m_lToPage As Long            'internal buffer for ToPage property

Private m_lApiReturn As Long          'internal buffer for APIReturn property
Private m_lExtendedError As Long      'internal buffer for ExtendedError property



Public Property Get Filter() As String
    'return object's Filter property
    Filter = m_sFilter
End Property

Public Sub ShowColor()
    'display the color dialog box
    
    Dim tChooseColor As ChooseColor
    Dim alCustomColors(15) As Long
    Dim lCustomColorSize As Long
    Dim lCustomColorAddress As Long
    Dim lMemHandle As Long
    
    Dim n As Integer
        
    On Error GoTo ShowColorError
    
    
    '***    init property buffers
    
    m_iAction = 3  'Action property - ShowColor
    m_lApiReturn = 0  'APIReturn property
    m_lExtendedError = 0  'ExtendedError property
    
    
    '***    prepare tChooseColor data
    
    'lStructSize As Long
    tChooseColor.m_lStructSize = Len(tChooseColor)
    
    'hwndOwner As Long
    tChooseColor.m_hwndOwner = m_lhdc

    'hInstance As Long
    
    'rgbResult As Long
    tChooseColor.m_rgbResult = m_lColor
    
    'lpCustColors As Long
    ' Fill custom colors array with all white
    For n = 0 To UBound(alCustomColors)
        alCustomColors(n) = &HFFFFFF
    Next
    ' Get size of memory needed for custom colors
    lCustomColorSize = Len(alCustomColors(0)) * 16
    ' Get a global memory block to hold a copy of the custom colors
    lMemHandle = GlobalAlloc(GHND, lCustomColorSize)
    
    If lMemHandle = 0 Then
        Exit Sub
    End If
    ' Lock the custom color's global memory block
    lCustomColorAddress = GlobalLock(lMemHandle)
    If lCustomColorAddress = 0 Then
        Exit Sub
    End If
    ' Copy custom colors to the global memory block
    Call CopyMemory(ByVal lCustomColorAddress, alCustomColors(0), lCustomColorSize)
 
    tChooseColor.m_lpCustColors = lCustomColorAddress
    
    'flags As Long
    tChooseColor.m_flags = m_lFlags
        
    'lCustData As Long
    'lpfnHook As Long
    'lpTemplateName As String
    
    
    '***    call the ChooseColor API function
    m_lApiReturn = ChooseColor(tChooseColor)
    
    
    '***    handle return from ChooseColor API function
    Select Case m_lApiReturn
        
        Case 0  'user canceled
        If m_bCancelError = True Then
            'generate an error
            On Error GoTo 0
            Err.Raise Number:=vbObjectError + 894, _
                Description:="Cancel Pressed"
            Exit Sub
        End If
        
        Case 1  'user selected a color
            'update property buffer
            m_lColor = tChooseColor.m_rgbResult
        
        Case Else   'an error occured
            'call CommDlgExtendedError
            m_lExtendedError = CommDlgExtendedError
        
    End Select

Exit Sub

ShowColorError:
    Exit Sub
End Sub

Public Sub ShowFont()
    'display the font dialog box
    
    Dim tLogFont As LOGFONT
    Dim tChooseFont As ChooseFont
    
    Dim lLogFontSize As Long
    Dim lLogFontAddress As Long
    Dim lMemHandle As Long
    
    Dim lReturn As Long
    Dim sFont As String
    Dim lBytePoint As Long
    On Error GoTo ShowFontError
    
    '***    init property buffers
    
    m_iAction = 4  'Action property - ShowFont
    m_lApiReturn = 0  'APIReturn property
    m_lExtendedError = 0  'ExtendedError property

    
    '***    prepare tChooseFont data
        
    'tLogFont.m_lfHeight As Long
    'tLogFont.m_lfWidth As Long
    'tLogFont.m_lfEscapement As Long
    'tLogFont.m_lfOrientation As Long
    
    'tLogFont.m_lfWeight As Long - init from FontBold property
    If m_bFontBold = True Then
        tLogFont.m_lfWeight = FW_BOLD
    End If
    
    'tLogFont.m_lfItalic As Byte - init from FontItalic property
    If m_bFontItalic = True Then
        tLogFont.m_lfItalic = 1
    End If
    
    'tLogFont.m_lfUnderline As Byte - init from FontUnderline property
    If m_bFontUnderline = True Then
        tLogFont.m_lfUnderline = 1
    End If

    'tLogFont.m_lfStrikeOut As Byte - init from FontStrikethru property
    If m_bFontStrikethru = True Then
        tLogFont.m_lfStrikeOut = 1
    End If

    'tLogFont.m_lfCharSet As Byte
    'tLogFont.m_lfOutPrecision As Byte
    'tLogFont.m_lfClipPrecision As Byte
    'tLogFont.m_lfQuality As Byte
    'tLogFont.m_lfPitchAndFamily As Byte
    'tLogFont.m_lfFaceName(LF_FACESIZE) As Byte
    
    'tChooseFont.m_lStructSize As Long
    tChooseFont.m_lStructSize = Len(tChooseFont)
    
    'tChooseFont.m_hwndOwner As Long
    'tChooseFont.m_hdc As Long
    
    'tChooseFont.m_lpLogFont As Long
    lLogFontSize = Len(tLogFont)
    
    ' Get a global memory block to hold a copy of tLogFont - exit on failure
    lMemHandle = GlobalAlloc(GHND, lLogFontSize)
    If lMemHandle = 0 Then
        Exit Sub
    End If
    
    ' Lock tLogFont's global memory block - exit on failure
    lLogFontAddress = GlobalLock(lMemHandle)
    If lLogFontAddress = 0 Then
        Exit Sub
    End If
    
    ' Copy tLogFont to the global memory block
    Call CopyMemory(ByVal lLogFontAddress, tLogFont, lLogFontSize)
 
    tChooseFont.m_lpLogFont = lLogFontAddress
    
    'tChooseFont.m_iPointSize As Long - init from FontSize property
    tChooseFont.m_iPointSize = m_lFontSize * 10
    
    'tChooseFont.flags As Long - init from Flags property
    tChooseFont.m_flags = m_lFlags

    'tChooseFont.rgbColors As Long
    'tChooseFont.m_lCustData As Long
    'tChooseFont.m_lpfnHook As Long
    'tChooseFont.m_lpTemplateName As String
    'tChooseFont.m_hInstance As Long
    
    'tChooseFont.m_lpszStyle As String
    'sFont = Chr$(0) & Space$(20) & Chr$(0)
    'tChooseFont.m_lpszStyle = sFont
    
    'tChooseFont.m_nFontType As Integer
    'tChooseFont.MISSING_ALIGNMENT As Integer
    'tChooseFont.m_nSizeMin As Long
    'tChooseFont.m_nSizeMax As Long
                    
    
    '***    call the CHOOSEFONT API function
    m_lApiReturn = ChooseFont(tChooseFont)    'store to APIReturn property
    
    
    '***    handle return from CHOOSEFONT API function
    Select Case m_lApiReturn
        
        Case 0  'user canceled
        If m_bCancelError = True Then
            'generate an error
            Err.Raise (2001)
            Exit Sub
        End If
        
        Case 1  'user selected a font
            ' Copy global memory block to tLogFont
            Call CopyMemory(tLogFont, ByVal lLogFontAddress, lLogFontSize)
            
            'tLogFont.m_lfWeight As Long  - store to FontBold property
            If tLogFont.m_lfWeight >= FW_BOLD Then
                m_bFontBold = True
            Else
                m_bFontBold = False
            End If
                        
            'tLogFont.m_lfItalic As Byte - store to FontItalic property
            If tLogFont.m_lfItalic = 1 Then
                m_bFontItalic = True
            Else
                m_bFontItalic = False
            End If
            
            'tLogFont.m_lfUnderline As Byte - store to FontUnderline property
            If tLogFont.m_lfUnderline = 1 Then
                m_bFontUnderline = True
            Else
                m_bFontUnderline = False
            End If
        
            'tLogFont.m_lfStrikeOut As Byte - store to FontStrikethru property
            If tLogFont.m_lfStrikeOut = 1 Then
                m_bFontStrikethru = True
            Else
                m_bFontStrikethru = False
            End If
            
            'tLogFont.m_lfFaceName(LF_FACESIZE) As Byte - store to FontName property
            FontName = ByteArrayToString(tLogFont.m_lfFaceName())
            
            'tChooseFont.m_iPointSize As Long - store to FontSize property
            m_lFontSize = CLng(tChooseFont.m_iPointSize / 10)
        
        Case Else   'an error occured
            'call CommDlgExtendedError
            m_lExtendedError = CommDlgExtendedError   'store to ExtendedError property
        
    End Select
Exit Sub

ShowFontError:
    Exit Sub
End Sub

Public Sub ShowHelp()
    'run winhelp.exe with the specified help file
    Dim sHelpFileBuff As String
    Dim lData As Long
    
    On Error GoTo ShowHelpError
    
    '***    init Private properties
    m_iAction = 6  'Action property - ShowHelp
    m_lApiReturn = 0  'APIReturn property
    m_lExtendedError = 0  'ExtendedError property

    '***    prepare the buffers and parameters for the API function
    'sHelpFile is a null terminated string
    sHelpFileBuff = m_sHelpFile & Chr$(0)
    
    'sData is dependent on lHelpCommand
    Select Case m_lHelpCommand
        Case 0
            lData = 0
        Case Else
            lData = 0
    End Select
    
    '***    call the API function
    m_lApiReturn = WinHelp(m_lhdc, m_sHelpFile, m_lHelpCommand, lData)     ' - Store to APIReturn property
    
    Select Case m_lApiReturn
        
        Case 0  '
            'call CommDlgExtendedError
            m_lExtendedError = CommDlgExtendedError   ' - store to ExtendedError property
        
        Case Else   '
            'call CommDlgExtendedError
            m_lExtendedError = CommDlgExtendedError
        
    End Select
        
Exit Sub

ShowHelpError:
    Exit Sub
End Sub


Public Sub ShowOpen()
    'display the file open dialog box
    ShowFileDialog vbffOpenDialog
End Sub

Public Sub ShowPrinter()
    'display the print dialog
    Dim tPrintDlg As PrintDlg
    
    On Error GoTo ShowPrinterError
    
    '***    init public properties
    m_iAction = 5  'Action property - ShowPrint
    m_lApiReturn = 0   'APIReturn property
    m_lExtendedError = 0  'ExtendedError property

    '***    prepare tPrintDlg data
    
    'lStructSize As Long
    tPrintDlg.m_lStructSize = Len(tPrintDlg)
    
    'hwndOwner As Long
    
    'hDevMode As Long
    
    'hDevNames As Long
    
    'hdc As Long - init from hDC property
    tPrintDlg.m_hdc = m_lhdc
    
    'flags As Long - init from Flags property
    tPrintDlg.m_flags = m_lFlags
    
    'nFromPage As Integer - init from FromPage property
    tPrintDlg.m_nFromPage = m_lFromPage
    
    'nToPage As Integer - init from ToPage property
    tPrintDlg.m_nToPage = m_lToPage
    
    'nMinPage As Integer - init from Min property
    tPrintDlg.m_nMinPage = m_lMin
    
    'nMaxPage As Integer - init from Max property
    tPrintDlg.m_nMaxPage = m_lMax
    
    'nCopies As Integer - init from Copies property
    tPrintDlg.m_nCopies = m_lCopies
    
    'hInstance As Long
    
    'lCustData As Long
    
    
    '***    Call the PrintDlg API function
    m_lApiReturn = PrintDlg(tPrintDlg)
    
    '***    handle return from PrintDlg API function
    Select Case m_lApiReturn
        
        Case 0  'user canceled
            If m_bCancelError = True Then
                'generate an error
                Err.Raise (2001)
                Exit Sub
            End If
        
        Case 1  'user selected OK
            'nFromPage As Integer - store to FromPage property
            m_lFromPage = tPrintDlg.m_nFromPage
            
            'nToPage As Integer - store to ToPage property
            m_lToPage = tPrintDlg.m_nToPage
            
            'nMinPage As Integer - store to Min property
            m_lMin = tPrintDlg.m_nMinPage
            
            'nMaxPage As Integer - store to Max property
            m_lMax = tPrintDlg.m_nMaxPage
            
            'nCopies As Integer - store to Copies property
            m_lCopies = tPrintDlg.m_nCopies
    
        Case Else   'an error occured
            'call CommDlgExtendedError
            m_lExtendedError = CommDlgExtendedError   'store to ExtendedError property
    
    End Select

Exit Sub

ShowPrinterError:
    
    Exit Sub
    
End Sub


Public Sub ShowSave()
    ShowFileDialog vbffSaveDialog
End Sub


Public Property Get Filename() As String
    Filename = m_sFileName
End Property

Public Property Let Filename(ByVal newFilename As String)
    m_sFileName = newFilename
End Property


Public Property Let Filter(ByVal newFilter As String)
    m_sFilter = newFilter
End Property


Private Function LeftOfNull(ByVal sIn As String)
    Dim lNullPos As Long
    LeftOfNull = sIn
    'get position of first Chr$(0) in sIn
    lNullPos = InStr(sIn, Chr$(0))
    'return part of sIn to left of first Chr$(0) if found
    If lNullPos > 0 Then
        LeftOfNull = Mid$(sIn, 1, lNullPos - 1)
    Else
        LeftOfNull = sIn
    End If
   
End Function


Public Property Get Action() As Integer
    Action = m_iAction
End Property

Private Function APIFilter(sIn)
    Dim sOneChr As String
    Dim sOutStr As String
    Dim pos As Long
    
    'convert any | characters to nulls
    sOutStr = sIn
    pos = InStr(sOutStr, "|")
    While pos > 0
        Mid$(sOutStr, pos, 1) = Chr$(0)
        pos = InStr(sOutStr, "|")
    Wend

    'add a null to the end
    sOutStr = sOutStr & Chr$(0)
    
    APIFilter = sOutStr
End Function

Public Property Get FilterIndex() As Integer
    FilterIndex = m_iFilterIndex
End Property

Public Property Let FilterIndex(vNewValue As Integer)
    m_iFilterIndex = vNewValue
End Property

Public Property Get CancelError() As Boolean
    CancelError = m_bCancelError
End Property

Public Property Let CancelError(vNewValue As Boolean)
    m_bCancelError = vNewValue
End Property

Public Property Get Color() As Long
    Color = m_lColor
End Property

Public Property Let Color(vNewValue As Long)
    m_lColor = vNewValue
End Property

Public Property Get Copies() As Long
    Copies = m_lCopies
End Property

Public Property Let Copies(vNewValue As Long)
    m_lCopies = vNewValue
End Property

Public Property Get DefaultExt() As String
    DefaultExt = m_sDefaultExt
End Property

Public Property Let DefaultExt(vNewValue As String)
    m_sDefaultExt = vNewValue
End Property

Public Property Get DialogTitle() As String
    DialogTitle = m_sDialogTitle
End Property

Public Property Let DialogTitle(vNewValue As String)
    m_sDialogTitle = vNewValue
End Property

Public Property Get Flags() As Long
    Flags = m_lFlags
End Property

Public Property Let Flags(vNewValue As Long)
    m_lFlags = vNewValue
End Property

Public Property Get FontBold() As Boolean
    FontBold = m_bFontBold
End Property

Public Property Let FontBold(vNewValue As Boolean)
    m_bFontBold = vNewValue
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = m_bFontItalic
End Property

Public Property Let FontItalic(vNewValue As Boolean)
    m_bFontItalic = vNewValue
End Property

Public Property Get FontName() As String
    FontName = m_sFontName
End Property

Public Property Let FontName(vNewValue As String)
    m_sFontName = vNewValue
End Property

Public Property Get FontSize() As Long
    FontSize = m_lFontSize
End Property

Public Property Let FontSize(vNewValue As Long)
    m_lFontSize = vNewValue
End Property

Public Property Get FontStrikethru() As Boolean
    FontStrikethru = m_bFontStrikethru
End Property

Public Property Let FontStrikethru(vNewValue As Boolean)
    m_bFontStrikethru = vNewValue
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = m_bFontUnderline
End Property

Public Property Let FontUnderline(ByVal newFontUnderline As Boolean)
    m_bFontUnderline = newFontUnderline
End Property

Public Property Get FromPage() As Long
    FromPage = m_lFromPage
End Property

Public Property Let FromPage(ByVal newFromPage As Long)
    m_lFromPage = newFromPage
End Property

Public Property Get hdc() As Long
    hdc = m_lhdc
End Property

Public Property Let hdc(ByVal newHDC As Long)
    m_lhdc = newHDC
End Property

Public Property Get HelpCommand() As Long
    HelpCommand = m_lHelpCommand
End Property

Public Property Let HelpCommand(ByVal newHelpCommand As Long)
    m_lHelpCommand = newHelpCommand
End Property

Public Property Get HelpContext() As String
    HelpContext = m_sHelpContext
End Property

Public Property Let HelpContext(ByVal newHelpContext As String)
    m_sHelpContext = newHelpContext
End Property

Public Property Get HelpFile() As String
    HelpFile = m_sHelpFile
End Property

Public Property Let HelpFile(ByVal newHelpFile As String)
    m_sHelpFile = newHelpFile
End Property

Public Property Get HelpKey() As String
    HelpKey = m_sHelpKey
End Property

Public Property Let HelpKey(ByVal newHelpKey As String)
    m_sHelpKey = newHelpKey
End Property

Public Property Get InitDir() As String
    InitDir = m_sInitDir
End Property

Public Property Let InitDir(ByVal newInitDir As String)
    m_sInitDir = newInitDir
End Property

Public Property Get Max() As Long
    Max = m_lMax
End Property

Public Property Let Max(ByVal newMax As Long)
    m_lMax = newMax
End Property

Public Property Get MaxFileSize() As Long
    MaxFileSize = m_lMaxFileSize
End Property

Public Property Let MaxFileSize(ByVal newMaxFileSize As Long)
    m_lMaxFileSize = newMaxFileSize
End Property

Public Property Get Min() As Long
    Min = m_lMin
End Property

Public Property Let Min(ByVal newMinimum As Long)
    m_lMin = newMinimum
End Property

Public Property Get Object() As Object
    Object = m_objObject
End Property

Public Property Let Object(ByVal newObject As Object)
    m_objObject = newObject
End Property

Public Property Get PrinterDefault() As Integer
    PrinterDefault = m_iPrinterDefault
End Property

Public Property Let PrinterDefault(ByVal newPrinterDefault As Integer)
    m_iPrinterDefault = newPrinterDefault
End Property

Public Property Get ToPage() As Long
    ToPage = m_lToPage
End Property

Public Property Let ToPage(ByVal newToPage As Long)
    m_lToPage = newToPage
End Property

Public Property Get FileTitle() As String
    FileTitle = m_sFileTitle
End Property

Public Property Let FileTitle(ByVal newFileTitle As String)
    m_sFileTitle = newFileTitle
End Property

Public Property Get APIReturn() As Long
    APIReturn = m_lApiReturn
End Property

Public Property Get ExtendedError() As Long
    ExtendedError = m_lExtendedError
End Property


Private Function ByteArrayToString(ByRef abBytes() As Byte) As String
    Dim lBytePoint As Long
    Dim lByteVal As Long
    Dim sOut As String
    
    'init array pointer
    lBytePoint = LBound(abBytes)
    
    'fill sOut with characters in array
    While lBytePoint <= UBound(abBytes)
        
        lByteVal = abBytes(lBytePoint)
        
        'return sOut and stop if Chr$(0) is encountered
        If lByteVal = 0 Then
            ByteArrayToString = sOut
            Exit Function
        Else
            sOut = sOut & Chr$(lByteVal)
        End If
        
        lBytePoint = lBytePoint + 1
    
    Wend
    
    'return sOut if Chr$(0) wasn't encountered
    ByteArrayToString = sOut
End Function

Private Sub ShowFileDialog(ByVal dlgAction As Integer)
    Dim openFile As OpenFilename
    Dim maxSize As Long
    Dim fileNameBuffer As String
    Dim fileTitleBuffer As String
    
   On Local Error GoTo ShowFileDialogError
    m_iAction = dlgAction
    m_lApiReturn = 0
    m_lExtendedError = 0
        
    openFile.m_lStructSize = Len(openFile)
    openFile.m_hwndOwner = m_lhdc
    openFile.m_lpstrFilter = APIFilter(m_sFilter)
    openFile.m_iFilterIndex = m_iFilterIndex
    'determine size of buffer from MaxFileSize property
    If m_lMaxFileSize > 0 Then
        maxSize = m_lMaxFileSize
    Else
        maxSize = 255
    End If
    
    fileNameBuffer = m_sFileName
    'pad with spaces & null terminate
    fileNameBuffer = fileNameBuffer & Space$(maxSize - 1 - Len(fileNameBuffer)) & Chr$(0)
    openFile.m_lpstrFile = fileNameBuffer
    
    openFile.m_nMaxFile = maxSize
            
    fileTitleBuffer = m_sFileTitle
    'pad with spaces & null terminate
    fileTitleBuffer = fileTitleBuffer & Space$(maxSize - 1 - Len(fileTitleBuffer)) & Chr$(0)
    
    openFile.m_lpstrFileTitle = fileTitleBuffer
    openFile.m_lpstrInitialDir = m_sInitDir
    openFile.m_lpstrTitle = m_sDialogTitle
    openFile.m_flags = m_lFlags
    openFile.m_lpstrDefExt = m_sDefaultExt
    
    Select Case Action
        Case vbffOpenDialog  'ShowOpen
            m_lApiReturn = GetOpenFileName(openFile)
        Case vbffSaveDialog 'ShowSave
            m_lApiReturn = GetSaveFileName(openFile)
        Case Else   'unknown action
            Exit Sub
    End Select
    
    Select Case m_lApiReturn
        Case 0  'user canceled
            m_lExtendedError = CommDlgExtendedError
            If m_bCancelError = True Then
                'generate an error
                Err.Raise (2001)
                Exit Sub
            End If
        
        Case 1  'user selected or entered a file
            m_sFileName = LeftOfNull(openFile.m_lpstrFile)
            m_sFileTitle = LeftOfNull(openFile.m_lpstrFileTitle)
        
        Case Else   'an error occured
            'call CommDlgExtendedError
            m_lExtendedError = CommDlgExtendedError
    End Select
    
    Exit Sub

ShowFileDialogError:
    Exit Sub
'    Debug.Assert False 'uh oh
End Sub


