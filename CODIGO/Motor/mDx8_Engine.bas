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

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamano muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Public ScreenWidth As Long
Public ScreenHeight As Long

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
    Dim i As Byte
    
    '   DeInit Lights
    Call DeInit_LightEngine
    
    '   DeInit Auras
    Call DeInit_Auras
    
    '   Clean Particles
    Call Particle_Group_Remove_All
    
    '   Clean Texture
    Call DirectDevice.SetTexture(0, Nothing)
    
    '   Borrar DBI Surface
    Call CleanDrawBuffer
    
    '   Erase Data
    Erase MapData()
    Erase charlist()
    
    Set DirectD3D8 = Nothing
    Set DirectD3D = Nothing
    Set DirectX = Nothing
    Set DirectDevice = Nothing
    Set SpriteBatch = Nothing
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

    TileBufferSize = Areas.TilesBuffer
    
    Call CalcularAreas(HalfWindowTileWidth, HalfWindowTileHeight)
    
    Engine_BaseSpeed = 0.018
    
    With MainScreenRect
        .Bottom = ScreenHeight
        .Right = ScreenWidth
    End With
    
    ' Seteamos algunos colores por adelantado y unica vez.
    Call Engine_Long_To_RGB_List(Normal_RGBList(), -1)
    Call Engine_Long_To_RGB_List(Color_Shadow(), D3DColorARGB(50, 0, 0, 0))
    Call Engine_Long_To_RGB_List(Color_Arbol(), D3DColorARGB(100, 100, 100, 100))
    Color_Paralisis = D3DColorARGB(180, 230, 230, 250)
    Color_Invisibilidad = D3DColorARGB(180, 236, 136, 66)
    Color_Montura = D3DColorARGB(180, 15, 230, 40)
    
    ' Inicializamos otros sistemas.
    Call mDx8_Text.Engine_Init_FontTextures
    Call mDx8_Text.Engine_Init_FontSettings
    Call mDx8_Auras.Load_Auras
    Call mDx8_Clima.Init_MeteoEngine
    Call mDx8_Dibujado.Damage_Initialize
    
    ' Inicializa DIB surface, un buffer usado para dejar imagenes estaticas en PictureBox
    Call PrepareDrawBuffer
    
End Sub

Public Sub Engine_Draw_Line(x1 As Single, y1 As Single, x2 As Single, y2 As Single, Optional Color As Long = -1, Optional Color2 As Long = -1)
On Error GoTo Error
    
    Call Engine_Long_To_RGB_List(temp_rgb(), Color)
    
    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.Draw(x1, y1, x2, y2, temp_rgb())
    
Exit Sub

Error:
    'Call Log_Engine("Error in Engine_Draw_Line, " & Err.Description & " (" & Err.number & ")")
End Sub

Public Sub Engine_Draw_Point(x1 As Single, y1 As Single, Optional Color As Long = -1)
On Error GoTo Error
    
    Call Engine_Long_To_RGB_List(temp_rgb(), Color)
    
    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.Draw(x1, y1, 0, 1, temp_rgb(), 0, 0)
    
Exit Sub

Error:
    'Call Log_Engine("Error in Engine_Draw_Point, " & Err.Description & " (" & Err.number & ")")
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

Public Function Engine_PixelPosX(ByVal X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_PixelPosX
'*****************************************************************

    Engine_PixelPosX = (X - 1) * 32
    
End Function

Public Function Engine_PixelPosY(ByVal Y As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_PixelPosY
'*****************************************************************

    Engine_PixelPosY = (Y - 1) * 32
    
End Function

Public Function Engine_TPtoSPX(ByVal X As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_TPtoSPX
'************************************************************

    Engine_TPtoSPX = Engine_PixelPosX(X - ((UserPos.X - HalfWindowTileWidth) - TileBufferSize)) + OffsetCounterX - 272 + ((10 - TileBufferSize) * 32)
    
End Function

Public Function Engine_TPtoSPY(ByVal Y As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_TPtoSPY
'************************************************************

    Engine_TPtoSPY = Engine_PixelPosY(Y - ((UserPos.Y - HalfWindowTileHeight) - TileBufferSize)) + OffsetCounterY - 272 + ((10 - TileBufferSize) * 32)
    
End Function

Public Sub Engine_Draw_Box(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As Long)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | Render Box
'***************************************************

    Call Engine_Long_To_RGB_List(temp_rgb(), Color)

    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.Draw(X, Y, Width, ByVal Height, temp_rgb())
    
End Sub

Public Sub Engine_D3DColor_To_RGB_List(rgb_list() As Long, Color As D3DCOLORVALUE)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 14/05/10
'Blisse-AO | Set a D3DColorValue to a RGB List
'***************************************************
    rgb_list(0) = D3DColorARGB(Color.a, Color.r, Color.g, Color.b)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Public Sub Engine_Long_To_RGB_List(rgb_list() As Long, long_color As Long)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 16/05/10
'Blisse-AO | Set a Long Color to a RGB List
'***************************************************
    rgb_list(0) = long_color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Public Function SetARGB_Alpha(rgb_list() As Long, Alpha As Byte) As Long()

    '***************************************************
    'Author: Juan Manuel Couso (Cucsifae)
    'Last Modification: 29/08/18
    'Obtiene un ARGB list le modifica el alpha y devuelve una copia
    '***************************************************
    Dim TempColor        As D3DCOLORVALUE
    Dim tempARGB(0 To 3) As Long

    'convertimos el valor del rgb list a D3DCOLOR
    Call ARGBtoD3DCOLORVALUE(rgb_list(1), TempColor)

    'comprobamos ue no se salga del rango permitido
    If Alpha > 255 Then Alpha = 255
    If Alpha < 0 Then Alpha = 0
    
    'seteamos el alpha
    TempColor.a = Alpha
    
    'generamos el nuevo RGB_List
    Call Engine_D3DColor_To_RGB_List(tempARGB(), TempColor)

    SetARGB_Alpha = tempARGB()

End Function

Private Function Engine_Collision_Between(ByVal Value As Single, ByVal Bound1 As Single, ByVal Bound2 As Single) As Byte
'*****************************************************************
'Find if a value is between two other values (used for line collision)
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Between
'*****************************************************************

    'Checks if a value lies between two bounds
    If Bound1 > Bound2 Then
        If Value >= Bound2 Then
            If Value <= Bound1 Then Engine_Collision_Between = 1
        End If
    Else
        If Value >= Bound1 Then
            If Value <= Bound2 Then Engine_Collision_Between = 1
        End If
    End If
    
End Function

Public Function Engine_Collision_Line(ByVal L1X1 As Long, ByVal L1Y1 As Long, ByVal L1X2 As Long, ByVal L1Y2 As Long, ByVal L2X1 As Long, ByVal L2Y1 As Long, ByVal L2X2 As Long, ByVal L2Y2 As Long) As Byte
'*****************************************************************
'Check if two lines intersect (return 1 if true)
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Line
'*****************************************************************
Dim m1 As Single
Dim M2 As Single
Dim b1 As Single
Dim b2 As Single
Dim IX As Single

    'This will fix problems with vertical lines
    If L1X1 = L1X2 Then L1X1 = L1X1 + 1
    If L2X1 = L2X2 Then L2X1 = L2X1 + 1

    'Find the first slope
    m1 = (L1Y2 - L1Y1) / (L1X2 - L1X1)
    b1 = L1Y2 - m1 * L1X2

    'Find the second slope
    M2 = (L2Y2 - L2Y1) / (L2X2 - L2X1)
    b2 = L2Y2 - M2 * L2X2
    
    'Check if the slopes are the same
    If M2 - m1 = 0 Then
    
        If b2 = b1 Then
            'The lines are the same
            Engine_Collision_Line = 1
        Else
            'The lines are parallel (can never intersect)
            Engine_Collision_Line = 0
        End If
        
    Else
        
        'An intersection is a point that lies on both lines. To find this, we set the Y equations equal and solve for X.
        'M1X+B1 = M2X+B2 -> M1X-M2X = -B1+B2 -> X = B1+B2/(M1-M2)
        IX = ((b2 - b1) / (m1 - M2))
        
        'Check for the collision
        If Engine_Collision_Between(IX, L1X1, L1X2) Then
            If Engine_Collision_Between(IX, L2X1, L2X2) Then Engine_Collision_Line = 1
        End If
        
    End If
    
End Function

Public Function Engine_Collision_LineRect(ByVal sX As Long, ByVal sY As Long, ByVal SW As Long, ByVal SH As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Byte
'*****************************************************************
'Check if a line intersects with a rectangle (returns 1 if true)
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_LineRect
'*****************************************************************

    'Top line
    If Engine_Collision_Line(sX, sY, sX + SW, sY, x1, y1, x2, y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If
    
    'Right line
    If Engine_Collision_Line(sX + SW, sY, sX + SW, sY + SH, x1, y1, x2, y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    'Bottom line
    If Engine_Collision_Line(sX, sY + SH, sX + SW, sY + SH, x1, y1, x2, y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    'Left line
    If Engine_Collision_Line(sX, sY, sX, sY + SW, x1, y1, x2, y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

End Function

Function Engine_Collision_Rect(ByVal x1 As Integer, ByVal y1 As Integer, ByVal Width1 As Integer, ByVal Height1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal Width2 As Integer, ByVal Height2 As Integer) As Boolean
'*****************************************************************
'Check for collision between two rectangles
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Rect
'*****************************************************************

    If x1 + Width1 >= x2 Then
        If x1 <= x2 + Width2 Then
            If y1 + Height1 >= y2 Then
                If y1 <= y2 + Height2 Then
                    Engine_Collision_Rect = True
                End If
            End If
        End If
    End If

End Function

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

Public Sub Engine_ZoomIn()
'**************************************************************
'Author: Standelf
'Last Modify Date: 29/12/2010
'**************************************************************

    With MainScreenRect
        .Top = 0
        .Left = 0
        .Bottom = IIf(.Bottom - 1 <= 367, .Bottom, .Bottom - 1)
        .Right = IIf(.Right - 1 <= 491, .Right, .Right - 1)
    End With
    
End Sub

Public Sub Engine_ZoomOut()
'**************************************************************
'Author: Standelf
'Last Modify Date: 29/12/2010
'**************************************************************

    With MainScreenRect
        .Top = 0
        .Left = 0
        .Bottom = IIf(.Bottom + 1 >= 459, .Bottom, .Bottom + 1)
        .Right = IIf(.Right + 1 >= 583, .Right, .Right + 1)
    End With
    
End Sub

Public Sub Engine_ZoomNormal()
'**************************************************************
'Author: Standelf
'Last Modify Date: 29/12/2010
'**************************************************************

    With MainScreenRect
        .Top = 0
        .Left = 0
        .Bottom = ScreenHeight
        .Right = ScreenWidth
    End With
    
End Sub

Public Function ZoomOffset(ByVal Offset As Byte) As Single
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/01/2011
'**************************************************************

    ZoomOffset = IIf((Offset = 1), (ScreenHeight - MainScreenRect.Bottom) / 2, (ScreenWidth - MainScreenRect.Right) / 2)
    
End Function

Function Engine_Distance(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer) As Long
'***************************************************
'Author: Standelf
'Last Modification: -
'***************************************************

    Engine_Distance = Abs(x1 - x2) + Abs(y1 - y2)
    
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

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetAngle
'************************************************************
Dim SideA As Single
Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then

        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90

            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If

        'Exit the function
        Exit Function

    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then

        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function
        Exit Function

    End If

    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle

    'Exit function

Exit Function

    'Check for error
ErrOut:

    'Return a 0 saying there was an error
    Engine_GetAngle = 0

Exit Function

End Function
