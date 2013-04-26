VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDirect3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Flexible vertex format the describes transformed and lit vertices.
Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR


'Module level variable to store the caps of the device.
Private m_d3dCaps As D3DCAPS8

'Module level variable to store the display mode.
Private m_d3dDisplayMode As D3DDISPLAYMODE

'Variables to store the render surface width and height.
'Private m_lClientWidth As Long
'Private m_lClientHeight As Long

Private m_hwnd As Long
'********************************************************************
'*                                                                  *
'* DirectX8 References                                              *
'*                                                                  *
'********************************************************************

Private m_dx As DirectX8
Private m_d3d As Direct3D8
Private m_d3dDevice As Direct3DDevice8
Private m_d3dX As D3DX8
Private m_d3dPresentParams As D3DPRESENT_PARAMETERS


Public Function GetDirect3DDevice() As Direct3DDevice8
    Set GetDirect3DDevice = m_d3dDevice
End Function

Public Function GetDirect3DX() As D3DX8
    Set GetDirect3DX = m_d3dX
End Function

Public Function GetDX() As DirectX8
    Set GetDX = m_dx
End Function

Public Function Init(ByVal hwnd As Long) As Boolean
    Dim DevType As CONST_D3DDEVTYPE
    Dim i As Long
    Dim lCount As Long
    Dim lErrNum As Long
    Dim format As Long
    Dim bFoundMode As Boolean
        
    m_hwnd = hwnd
    'Turn off error checking. The app will check for errors and handle them.
    On Local Error Resume Next
    
    'Initiazlize the DirectX8 object
    Set m_dx = New DirectX8
        
    'Check to make sure that the dx object was created successfully.
    If Err.Number Then
        'There were problems creating the dx object. Return the error number.
        Init = False
        Exit Function
    End If
    
    'Create the Direct3D object
    Set m_d3d = m_dx.Direct3DCreate
    
    'Check to make sure that the d3d object was created successfully.
    If Err.Number Then
        'There were problems creating the d3d object. Return the error number,
        Init = False
        Exit Function
    End If
    
    'We'll start by attempting to create a HAL device. This variable
    'will hold the final type of device that we create after we check
    'some capabilities.
    DevType = D3DDEVTYPE_HAL
    
    'Get the capabilities of the Direct3D device that we specify. In this case,
    'we'll be using the adapter default (the primiary card on the system).
    Call m_d3d.GetDeviceCaps(D3DADAPTER_DEFAULT, DevType, m_d3dCaps)
    
    'Check for errors. If there is an error, the card more than likely doesn't support at least DX7,
    'so get the caps of the reference device instead.
    If Err.Number Then
        Err.Clear
        DevType = D3DDEVTYPE_REF
        Call m_d3d.GetDeviceCaps(D3DADAPTER_DEFAULT, DevType, m_d3dCaps)
        
        'If there is *still* an error, then the driver has problems. We'll
        'have to exit at this point, because there isn't anything else we can
        'do.
        If Err.Number Then
            Init = False
            Exit Function
        End If
        
    End If
    
    'Grab some information about the current display mode.
    Call m_d3d.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, m_d3dDisplayMode)
    
    'Now we'll go ahead and fill the D3DPRESENT_PARAMETERS type.
    With m_d3dPresentParams
        'Make sure that the adapter is in a color bit-depth greater than 8 bits per pixel.
        If m_d3dDisplayMode.format = D3DFMT_P8 Or m_d3dDisplayMode.format = D3DFMT_A8P8 Then
            'Device is running in some variation of an 8 bit format
            MsgBox " For this application to run, the primary display needs to be in 16 bit or higher color depth.", vbCritical
            Init = False
            Exit Function
        Else
            'Device is greater than 8 bit. Set the format variable to the current display format.
            format = m_d3dDisplayMode.format
        End If
        
        'For windowed mode, we just discard any information instead of flipping it.
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    
        'Set windowed mode to true.
        .Windowed = 1
    
        'Set the backbuffer format
        .BackBufferFormat = format
        
        .BackBufferWidth = Screen.Width / Screen.TwipsPerPixelX
        .BackBufferHeight = Screen.Height / Screen.TwipsPerPixelY
    End With
    
    'Try to create the device now that we have everything set.
    Set m_d3dDevice = m_d3d.CreateDevice(D3DADAPTER_DEFAULT, DevType, m_hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, m_d3dPresentParams)
    
    'If the creation above failed, try to create a REF device instead.
    If Err.Number Then
        Err.Clear
        Set m_d3dDevice = m_d3d.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_REF, m_hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, m_d3dPresentParams)
        If Err.Number Then
            'The app still hit an error. Both HAL and REF devices weren't created. The app will have to exit at this point.
            Init = False
            Exit Function
        End If
    
    End If
    
    If InitDevice(m_d3dDevice) Then
        MsgBox "Unable to initialize the device"
        Init = False
        Exit Function
    End If

    Set m_d3dX = New D3DX8
    Init = True
End Function

Private Function InitDevice(dev As Direct3DDevice8) As Long

    '***********************************************************************
    '
    ' This function initializes the device with some renderstates, and also
    ' sets up the viewport, camera, and world.
    '
    ' Parameters:
    '
    ' [IN]
    '       dev:    An existing Direct3DDevice8 object
    '       m_d3dpp:  A filled D3DPRESENT_PARAMETERS type
    '       hwnd:   Handle to the target window
    '
    '
    ' Return value:
    '     If an error occurs, it returns D3DERR_INVALIDCALL.
    '
    '***********************************************************************

    Call dev.Reset(m_d3dPresentParams)

    'On Local Error Resume Next
    With dev
        'Set the vertex shader to an FVF that contains texture coords,
        'and transformed and lit vertex coords.
        Call .SetVertexShader(FVF)
        
        'Turn off lighting
        Call .SetRenderState(D3DRS_LIGHTING, 0)
                                
        'Set the render state that uses the alpha component as the source for blending.
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        
        'Set the render state that uses the inverse alpha component as the destination blend.
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
    End With
        
    If Err.Number Then InitDevice = D3DERR_INVALIDCALL
End Function

Public Sub EndFrame()
    With m_d3dDevice
        Call .EndScene
        'Draw the graphics to the front buffer.
        Call .Present(ByVal 0&, ByVal 0&, 0, ByVal 0&)
    End With
End Sub


Public Sub BeginFrame()
    With m_d3dDevice
        'Clear the back buffer
        Call .Clear(0, ByVal 0&, D3DCLEAR_TARGET, &HFF, 0, 0)
        Call .BeginScene
    End With
    
End Sub

Public Function LoadTexture(ByVal fp As String, ByVal transparentColour As Long, ByRef scaling As D3DVECTOR2) As Direct3DTexture8
    Dim tex As Direct3DTexture8
    Dim inf As D3DXIMAGE_INFO
    Dim surfDesc As D3DSURFACE_DESC
    
    'Check to make sure the media was found
    If Dir(fp) = vbNullString Then
        Err.Raise 65000, , "Unable to locate specified graphic file."
        Stop
    End If
    
    'Load the Sprite texture. We need to get alpha information embedded into this
    'surface, so we'll call the more complex CreateTextureFromFileEx() method instead.
    'The main thing we need to do is just let it know we want to use black as the
    'alpha channel. We do this by passing &HFF000000 to the method, and it fills in
    'the high order byte of any pixel that contains black with full alpha so that it
    'becomes transparent when rendered with alpha blending enabled.
'    Set tex = m_d3dX.CreateTextureFromFileEx(m_d3dDevice, _
'                                            fp, _
'                                            D3DX_DEFAULT, _
'                                            D3DX_DEFAULT, _
'                                            1, _
'                                            0, _
'                                            D3DFMT_UNKNOWN, _
'                                            D3DPOOL_MANAGED, _
'                                            D3DX_FILTER_POINT, _
'                                            D3DX_FILTER_POINT, _
'                                            &HFF000000 Or transparentColour, _
'                                            inf, _
'                                            ByVal 0)
 
    Set tex = m_d3dX.CreateTextureFromFileEx(m_d3dDevice, _
                                            fp, _
                                            D3DX_DEFAULT, _
                                            D3DX_DEFAULT, _
                                            1&, _
                                            0&, _
                                            D3DFMT_UNKNOWN, _
                                            D3DPOOL_MANAGED, _
                                            D3DX_FILTER_NONE, _
                                            D3DX_FILTER_NONE, _
                                            &HFF000000 Or transparentColour, _
                                            ByVal 0, _
                                            ByVal 0&)
 
'    inf.Width = 160
'    inf.Height = 48
'    tex.GetLevelDesc 0, surfDesc
'    scaling.x = inf.Width / surfDesc.Width
'    scaling.y = inf.Height / surfDesc.Height
    scaling.X = 1
    scaling.Y = 1
    If Err.Number Then
        'Something happened while loading the texture.
'        MsgBox "Error loading texture. Error number: " & Err.Number
        Set LoadTexture = Nothing
        Err.Raise 65000, , "Error loading texture. Error number : " & Err.Number
        Stop
    End If
    
    Set LoadTexture = tex
End Function


