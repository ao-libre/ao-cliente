Attribute VB_Name = "mDx8_Engine"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' No matter what you do with DirectX8, you will need to start with
' the DirectX8 object. You will need to create a new instance of
' the object, using the New keyword, rather than just getting a
' pointer to it, since there's nowhere to get a pointer from yet (duh!).

Public DirectX As New DirectX8

' The D3DX8 object contains lots of helper functions, mostly math
' to make Direct3D alot easier to use. Notice we create a new
' instance of the object using the New keyword.
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8

' The Direct3DDevice8 represents our rendering device, which could
' be a hardware or a software device. The great thing is we still
' use the same object no matter what it is
Public DirectDevice As Direct3DDevice8

Public SurfaceDB As New clsTextureManager
Public SpriteBatch As New clsBatch

Private Viewport As D3DVIEWPORT8
Private Projection As D3DMATRIX
Private View As D3DMATRIX

Public Engine_BaseSpeed As Single

Public ScreenWidth As Long
Public ScreenHeight As Long

Public Const HeadOffsetAltos As Integer = -8
Public Const HeadOffsetBajos As Integer = 2

Public MainScreenRect As RECT

Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Private EndTime As Long

Public Function Engine_DirectX8_Init() As Boolean
    
    'Establecemos cual va a ser el tamano del render.
    ScreenWidth = frmMain.MainViewPic.ScaleWidth
    ScreenHeight = frmMain.MainViewPic.ScaleHeight
    
    ' The D3DDISPLAYMODE type structure that holds
    ' the information about your current display adapter.
    Dim DispMode  As D3DDISPLAYMODE
    
    ' The D3DPRESENT_PARAMETERS type holds a description of the way
    ' in which DirectX will display it's rendering.
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    
    ' Initialize all DirectX objects.
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8

    ' Retrieve the information about your current display adapter.
    Call DirectD3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispMode)
    
    ' Fill the D3DPRESENT_PARAMETERS type, describing how DirectX should
    ' display it's renders.
    With D3DWindow
        .Windowed = True
        
        ' The swap effect determines how the graphics get from the backbuffer to the screen.
        ' D3DSWAPEFFECT_DISCARD:
        '   Means that every time the render is presented, the backbuffer
        '   image is destroyed, so everything must be rendered again.
        .SwapEffect = IIf((ClientSetup.vSync) = True, D3DSWAPEFFECT_COPY_VSYNC, D3DSWAPEFFECT_DISCARD)
        
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = ScreenWidth
        .BackBufferHeight = ScreenHeight
        .hDeviceWindow = frmMain.MainViewPic.hWnd
    End With
    
    ' Create the rendering device.
    ' Here we request a Hardware or Mixed rasterization.
    ' If your computer does not have this, the request may fail, so use
    ' D3DDEVTYPE_REF instead of D3DDEVTYPE_HAL if this happens. A real
    ' program would be able to detect an error and automatically switch device.
    ' We also request software vertex processing, which means the CPU has to
    ' transform and light our geometry.
    Select Case ClientSetup.Aceleracion

        Case 0 '   Hardware
            Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, _
                                                      D3DDEVTYPE_HAL, _
                                                      D3DWindow.hDeviceWindow, _
                                                      D3DCREATE_HARDWARE_VERTEXPROCESSING, _
                                                      D3DWindow)

        Case 1 '   Mixed
            Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, _
                                                      D3DDEVTYPE_HAL, _
                                                      D3DWindow.hDeviceWindow, _
                                                      D3DCREATE_MIXED_VERTEXPROCESSING, _
                                                      D3DWindow)

        Case Else 'Si no hay opcion entramos en Hardware para asegurarnos que funcione el cliente.
            Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, _
                                                      D3DDEVTYPE_HAL, _
                                                      D3DWindow.hDeviceWindow, _
                                                      D3DCREATE_HARDWARE_VERTEXPROCESSING, _
                                                      D3DWindow)
            
    End Select
    
    'Seteamos la matriz de proyeccion.
    Call D3DXMatrixOrthoOffCenterLH(Projection, 0, ScreenWidth, ScreenHeight, 0, -1#, 1#)
    Call D3DXMatrixIdentity(View)
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)
    Call DirectDevice.SetTransform(D3DTS_VIEW, View)

    ' Set rendering options
    Call Engine_Init_RenderStates
    
    'Carga dinamica de texturas por defecto.
    Set SurfaceDB = New clsTextureManager
    
    'Sprite batching.
    Set SpriteBatch = New clsBatch
    Call SpriteBatch.Initialise(2000)
    
    EndTime = timeGetTime

    If Err Then
        MsgBox JsonLanguage.item("ERROR_DIRECTX_INIT").item("TEXTO")
        Engine_DirectX8_Init = False
        Exit Function
    End If
    
    If DirectDevice Is Nothing Then
        MsgBox JsonLanguage.item("ERROR_DIRECTDEVICE_INIT").item("TEXTO")
        Engine_DirectX8_Init = False
        Exit Function
    End If
    
    Engine_DirectX8_Init = True
    
End Function

Private Sub Engine_Init_RenderStates()

    'Set the render states
    With DirectDevice
    
        Call .SetVertexShader(D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
        Call .SetRenderState(D3DRS_LIGHTING, False)
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, True)
        Call .SetRenderState(D3DRS_FILLMODE, D3DFILL_SOLID)
        Call .SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
        Call .SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
        
    End With
    
End Sub

Public Sub Engine_DirectX8_End()
'***************************************************
'Author: Standelf
'Last Modification: 26/05/2010
'Destroys all DX objects
'***************************************************
On Error Resume Next
    
    '   DeInit Lights
    Call DeInit_LightEngine
    
    '   DeInit Auras
    Call DeInit_Auras
    
    '   Clean Particles
    Call Particle_Group_Remove_All
    
    '   Clean Texture
    Call DirectDevice.SetTexture(0, Nothing)

    '   Erase Data
    Erase MapData()
    Erase charlist()
    
    Set DirectD3D8 = Nothing
    Set DirectD3D = Nothing
    Set DirectX = Nothing
    Set DirectDevice = Nothing
    Set SpriteBatch = Nothing
    Set Audio = Nothing
    
End Sub

Public Sub Engine_DirectX8_Aditional_Init()
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/12/2010
'**************************************************************

    FPS = 101
    FramesPerSecCounter = 101
    
    ColorTecho = 250
    colorRender = 240

    TileBufferSize = 9
    Engine_BaseSpeed = 0.018
    
    With MainScreenRect
        .Bottom = ScreenHeight
        .Right = ScreenWidth
    End With
    
    ' Seteamos algunos colores por adelantado y unica vez.
    Call Engine_Long_To_RGB_List(Normal_RGBList(), -1)
    Call Engine_Long_To_RGB_List(Color_Shadow(), D3DColorARGB(50, 0, 0, 0))
    Call Engine_Long_To_RGB_List(Color_Arbol(), D3DColorARGB(100, 100, 100, 100))
    
    ' Inicializamos otros sistemas.
    Call mDx8_Text.Engine_Init_FontTextures
    Call mDx8_Text.Engine_Init_FontSettings
    Call mDx8_Auras.Load_Auras
    Call mDx8_Clima.Init_MeteoEngine
    Call mDx8_Dibujado.Damage_Initialize
    
End Sub

Public Sub Engine_BeginScene(Optional ByVal Color As Long = 0)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | DD Clear & BeginScene
'***************************************************

    Call DirectDevice.BeginScene
    Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, Color, 1#, 0)
    Call SpriteBatch.Begin
    
End Sub

Public Sub Engine_EndScene(ByRef DestRect As RECT, Optional ByVal hWndDest As Long = 0)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | DD EndScene & Present
'***************************************************
    
    Call SpriteBatch.Flush
    
    Call DirectDevice.EndScene
        
    If hWndDest = 0 Then
        Call DirectDevice.Present(DestRect, ByVal 0&, ByVal 0&, ByVal 0&)
    Else
        Call DirectDevice.Present(DestRect, ByVal 0, hWndDest, ByVal 0)
    End If
    
End Sub

Public Function Engine_ElapsedTime() As Long
'**************************************************************
'Gets the time that past since the last call
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_ElapsedTime
'**************************************************************
Dim Start_Time As Long

    'Get current time
    Start_Time = timeGetTime

    'Calculate elapsed time
    Engine_ElapsedTime = Start_Time - EndTime

    'Get next end time
    EndTime = Start_Time

End Function

Public Sub Engine_Update_FPS()
    '***************************************************
    'Author: Standelf
    'Last Modification: 09/09/2019
    'Calculate FPS
    '***************************************************

    If FPSLastCheck + 1000 < GetTickCount Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 1
        FPSLastCheck = GetTickCount
    Else
        FramesPerSecCounter = FramesPerSecCounter + 1

    End If

End Sub
