Attribute VB_Name = "mDx8_Engine"
'DX8 Objects
Public DirectX As New DirectX8
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8
Public DirectDevice As Direct3DDevice8

Public SurfaceDB As New clsSurfaceManager

Public Engine_BaseSpeed As Single
Public TileBufferSize As Integer

Public Const ScreenWidth As Long = 536
Public Const ScreenHeight As Long = 412

Public Const HeadOffsetAltos As Integer = -8
Public Const HeadOffsetBajos As Integer = 2

Public MainScreenRect As RECT
Public ConnectScreenRect As RECT
'
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

    Dim DispMode As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8

    DirectD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = IIf((ClientSetup.vSync) = True, D3DSWAPEFFECT_COPY_VSYNC, D3DSWAPEFFECT_DISCARD)
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = frmMain.MainViewPic.ScaleWidth
        .BackBufferHeight = frmMain.MainViewPic.ScaleHeight
        .hDeviceWindow = frmMain.MainViewPic.hwnd
    End With

    Select Case ClientSetup.Aceleracion
        Case 0 '   Software
            Set DirectDevice = DirectD3D.CreateDevice( _
                                D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                                frmMain.MainViewPic.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                D3DWindow)
        Case 1 '   Hardware
            Set DirectDevice = DirectD3D.CreateDevice( _
                                D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                                frmMain.MainViewPic.hwnd, _
                                D3DCREATE_HARDWARE_VERTEXPROCESSING, _
                                D3DWindow)
        Case 2 '   Mixed
            Set DirectDevice = DirectD3D.CreateDevice( _
                                D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                                frmMain.MainViewPic.hwnd, _
                                D3DCREATE_MIXED_VERTEXPROCESSING, _
                                D3DWindow)
        Case Else '   Si no hay opcion entramos en Software para asegurarnos que funcione el cliente
            Set DirectDevice = DirectD3D.CreateDevice( _
                                D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                                frmMain.MainViewPic.hwnd, _
                                D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                D3DWindow)
    End Select

    Engine_Init_FontTextures
    Engine_Init_FontSettings
    
    DirectDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1 Or D3DFVF_SPECULAR
    DirectDevice.SetRenderState D3DRS_LIGHTING, False
    DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DirectDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    DirectDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    DirectDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    DirectDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    DirectDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
    EndTime = GetTickCount
    
    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la última versión correctamente instalada. Puede descargarla desde: " & Client_Web & "support/directx.zip"
        Engine_DirectX8_Init = False
        Exit Function
    End If
    
    If Err Then
        MsgBox "No se puede iniciar DirectD3D. Por favor asegurese de tener la última versión correctamente instalada. Puede descargarla desde: " & Client_Web & "support/directx.zip"
        Engine_DirectX8_Init = False
        Exit Function
    End If
    
    If DirectDevice Is Nothing Then
        MsgBox "No se puede inicializar DirectDevice. Por favor asegurese de tener la última versión correctamente instalada. Puede descargarla desde: " & Client_Web & "support/directx.zip"
        Engine_DirectX8_Init = False
        Exit Function
    End If
    
    Engine_DirectX8_Init = True
End Function

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
    For i = 1 To UBound(ParticleTexture)
        If Not ParticleTexture(i) Is Nothing Then Set ParticleTexture(i) = Nothing
    Next i
    
    '   Clean Texture
    DirectDevice.SetTexture 0, Nothing

    '   Erase Data
    Erase MapData()
    Erase charlist()
    
    Set DirectD3D8 = Nothing
    Set DirectD3D = Nothing
    Set DirectX = Nothing
    Set DirectDevice = Nothing
End Sub

Public Sub Engine_DirectX8_Aditional_Init()
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/12/2010
'**************************************************************

    FPS = 101
    FramesPerSecCounter = 101

    Engine_Set_TileBuffer 9
    
    Engine_Set_BaseSpeed 0.018
    
    With MainScreenRect
        .bottom = frmMain.MainViewPic.ScaleHeight
        .Right = frmMain.MainViewPic.ScaleWidth
    End With


    Call Engine_Long_To_RGB_List(Normal_RGBList(), -1)

    Load_Auras
    Init_MeteoEngine
    Engine_Init_ParticleEngine
    
End Sub

Public Sub Engine_Draw_Line(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, Optional Color As Long = -1, Optional Color2 As Long = -1)
On Error GoTo error
Dim Vertex(1) As TLVERTEX

    Vertex(0) = Geometry_Create_TLVertex(X1, Y1, 0, 1, Color, 0, 0)
    Vertex(1) = Geometry_Create_TLVertex(X2, Y2, 0, 1, Color2, 0, 0)

    DirectDevice.SetTexture 0, Nothing
    DirectDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, Vertex(0), Len(Vertex(0))
Exit Sub

error:
    'Call Log_Engine("Error in Engine_Draw_Line, " & Err.Description & " (" & Err.number & ")")
End Sub

Public Sub Engine_Draw_Point(X1 As Single, Y1 As Single, Optional Color As Long = -1)
On Error GoTo error
Dim Vertex(0) As TLVERTEX

    Vertex(0) = Geometry_Create_TLVertex(X1, Y1, 0, 1, Color, 0, 0)

    DirectDevice.SetTexture 0, Nothing
    DirectDevice.DrawPrimitiveUP D3DPT_POINTLIST, 1, Vertex(0), Len(Vertex(0))
Exit Sub

error:
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

    Engine_TPtoSPX = Engine_PixelPosX(X - ((UserPos.X - HalfWindowTileWidth) - Engine_Get_TileBuffer)) + OffsetCounterX - 272 + ((10 - TileBufferSize) * 32)
    
End Function

Public Function Engine_TPtoSPY(ByVal Y As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_TPtoSPY
'************************************************************

    Engine_TPtoSPY = Engine_PixelPosY(Y - ((UserPos.Y - HalfWindowTileHeight) - Engine_Get_TileBuffer)) + OffsetCounterY - 272 + ((10 - TileBufferSize) * 32)
    
End Function

Private Function Engine_FToDW(f As Single) As Long
'*****************************************************************
'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_FToDW
'*****************************************************************
Dim buf As D3DXBuffer

    'Converts a single into a long (Float to DWORD)
    Set buf = DirectD3D8.CreateBuffer(4)
    DirectD3D8.BufferSetData buf, 0, 4, 1, f
    DirectD3D8.BufferGetData buf, 0, 4, 1, Effect_FToDW

End Function

Public Sub Engine_Draw_Box(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As Long)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | Render Box
'***************************************************
    Dim b_Rect As RECT
    Dim b_Color(0 To 3) As Long
    Dim b_Vertex(0 To 3) As TLVERTEX
    
    Engine_Long_To_RGB_List b_Color(), Color

    With b_Rect
        .bottom = Y + Height
        .Left = X
        .Right = X + Width
        .Top = Y
    End With

    Geometry_Create_Box b_Vertex(), b_Rect, b_Rect, b_Color(), 0, 0
    
    DirectDevice.SetTexture 0, Nothing
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, b_Vertex(0), Len(b_Vertex(0))
End Sub

Public Sub Engine_D3DColor_To_RGB_List(RGB_List() As Long, Color As D3DCOLORVALUE)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 14/05/10
'Blisse-AO | Set a D3DColorValue to a RGB List
'***************************************************
    RGB_List(0) = D3DColorARGB(Color.a, Color.r, Color.g, Color.b)
    RGB_List(1) = RGB_List(0)
    RGB_List(2) = RGB_List(0)
    RGB_List(3) = RGB_List(0)
End Sub

Public Sub Engine_Long_To_RGB_List(RGB_List() As Long, long_color As Long)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 16/05/10
'Blisse-AO | Set a Long Color to a RGB List
'***************************************************
    RGB_List(0) = long_color
    RGB_List(1) = RGB_List(0)
    RGB_List(2) = RGB_List(0)
    RGB_List(3) = RGB_List(0)
End Sub
Public Function SetARGB_Alpha(RGB_List() As Long, alpha As Byte) As Long()
'***************************************************
'Author: Juan Manuel Couso (Cucsifae)
'Last Modification: 29/08/18
'Obtiene un ARGB list le modifica el alpha y devuelve una copia
'***************************************************
Dim TempColor As D3DCOLORVALUE
Dim tempARGB(0 To 3) As Long
'convertimos el valor del rgb list a D3DCOLOR
Call ARGBtoD3DCOLORVALUE(RGB_List(1), TempColor)
'comprobamos ue no se salga del rango permitido
If alpha > 255 Then alpha = 255
If alpha < 0 Then alpha = 0
'seteamos el alpha
TempColor.a = alpha
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

Public Function Engine_Collision_LineRect(ByVal SX As Long, ByVal SY As Long, ByVal SW As Long, ByVal SH As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Byte
'*****************************************************************
'Check if a line intersects with a rectangle (returns 1 if true)
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_LineRect
'*****************************************************************

    'Top line
    If Engine_Collision_Line(SX, SY, SX + SW, SY, X1, Y1, X2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If
    
    'Right line
    If Engine_Collision_Line(SX + SW, SY, SX + SW, SY + SH, X1, Y1, X2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    'Bottom line
    If Engine_Collision_Line(SX, SY + SH, SX + SW, SY + SH, X1, Y1, X2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    'Left line
    If Engine_Collision_Line(SX, SY, SX, SY + SW, X1, Y1, X2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

End Function

Function Engine_Collision_Rect(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal Width1 As Integer, ByVal Height1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal Width2 As Integer, ByVal Height2 As Integer) As Boolean
'*****************************************************************
'Check for collision between two rectangles
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Rect
'*****************************************************************

    If X1 + Width1 >= X2 Then
        If X1 <= X2 + Width2 Then
            If Y1 + Height1 >= Y2 Then
                If Y1 <= Y2 + Height2 Then
                    Engine_Collision_Rect = True
                End If
            End If
        End If
    End If

End Function

Public Sub Engine_BeginScene(Optional ByVal Color As Long = 0)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | DD Clear & BeginScene
'***************************************************

    DirectDevice.BeginScene
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, Color, 1#, 0
    
End Sub

Public Sub Engine_EndScene(ByRef destRect As RECT, Optional ByVal hWndDest As Long = 0)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | DD EndScene & Present
'***************************************************
    
    DirectDevice.EndScene
        
    If hWndDest = 0 Then
        DirectDevice.Present destRect, ByVal 0&, ByVal 0&, ByVal 0&
    Else
        DirectDevice.Present destRect, ByVal 0, hWndDest, ByVal 0
    End If
    
End Sub

Public Sub Geometry_Create_Box(ByRef Verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef RGB_List() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal Angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'**************************************************************

    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim Temp As Single
    
    If Angle > 0 Then
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.bottom - dest.Top) / 2
        
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        Temp = (dest.Right - x_center) / radius
        right_point = Atn(Temp / Sqr(-Temp * Temp + 1))
        left_point = 3.1459 - right_point
    End If
    
    If Angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - Angle) * radius
        y_Cor = y_center - Sin(-left_point - Angle) * radius
    End If

    If Textures_Width And Textures_Height Then
        Verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(0), src.Left / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        Verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(0), 0, 0)
    End If

    If Angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - Angle) * radius
        y_Cor = y_center - Sin(left_point - Angle) * radius
    End If
    
    If Textures_Width And Textures_Height Then
        Verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(1), src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        Verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(1), 0, 1)
    End If

    If Angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - Angle) * radius
        y_Cor = y_center - Sin(-right_point - Angle) * radius
    End If

    If Textures_Width And Textures_Height Then
        Verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(2), (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        Verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(2), 1, 0)
    End If

    If Angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - Angle) * radius
        y_Cor = y_center - Sin(right_point - Angle) * radius
    End If

    If Textures_Width And Textures_Height Then
        Verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(3), (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        Verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(3), 1, 1)
    End If

End Sub

Public Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal Color As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.Color = Color
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Public Sub Engine_ZoomIn()
'**************************************************************
'Author: Standelf
'Last Modify Date: 29/12/2010
'**************************************************************

    With MainScreenRect
        .Top = 0
        .Left = 0
        .bottom = IIf(.bottom - 1 <= 367, .bottom, .bottom - 1)
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
        .bottom = IIf(.bottom + 1 >= 459, .bottom, .bottom + 1)
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
        .bottom = ScreenHeight
        .Right = ScreenWidth
    End With
    
End Sub

Public Function ZoomOffset(ByVal offset As Byte) As Single
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/01/2011
'**************************************************************

    ZoomOffset = IIf((offset = 1), (ScreenHeight - MainScreenRect.bottom) / 2, (ScreenWidth - MainScreenRect.Right) / 2)
    
End Function

Public Sub Engine_Set_BaseSpeed(ByVal BaseSpeed As Single)
'**************************************************************
'Author: Standelf
'Last Modify Date: 29/12/2010
'**************************************************************

    Engine_BaseSpeed = BaseSpeed
    
End Sub

Public Function Engine_Get_BaseSpeed() As Single
'**************************************************************
'Author: Standelf
'Last Modify Date: 29/12/2010
'**************************************************************

    Engine_Get_BaseSpeed = Engine_BaseSpeed
    
End Function

Public Sub Engine_Set_TileBuffer(ByVal setTileBufferSize As Single)
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/12/2010
'**************************************************************

    TileBufferSize = setTileBufferSize
    
End Sub

Public Function Engine_Get_TileBuffer() As Single
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/12/2010
'**************************************************************

    Engine_Get_TileBuffer = TileBufferSize
    
End Function

Function Engine_Distance(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Long
'***************************************************
'Author: Standelf
'Last Modification: -
'***************************************************

    Engine_Distance = Abs(X1 - X2) + Abs(Y1 - Y2)
    
End Function

Public Sub Engine_Update_FPS()
'***************************************************
'Author: Standelf
'Last Modification: 10/01/2011
'Limit FPS & Calculate later
'***************************************************

        If ClientSetup.LimiteFPS And Not ClientSetup.vSync Then
            While (GetTickCount - FPSLastCheck) \ 10 < FramesPerSecCounter
                Sleep 5
            Wend
        End If
        
        If FPSLastCheck + 1000 < GetTickCount Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            FPSLastCheck = GetTickCount
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If

        'If Settings.MostrarFPS = True Then
            'Fonts_Render_String FPS, 2, 2, -1, Settings.Engine_Font
            'DrawText 2, 2, FPS, -1
       ' End If
End Sub

Public Sub DrawPJ(ByVal Index As Byte)

    If LenB(cPJ(Index).Nombre) = 0 Then Exit Sub

    Dim cColor As Long

    frmPanelAccount.lblAccData(Index).Caption = cPJ(Index).Nombre

    If cPJ(Index).GameMaster Then
        cColor = RGB(57, 255, 20)
    Else
        If cPJ(Index).Criminal Then
            cColor = RGB(255, 128, 128)
        Else
            cColor = RGB(0, 192, 192)
        End If
    End If

    frmPanelAccount.lblAccData(Index).ForeColor = cColor

    Dim i As Integer

    Dim init_x As Integer
    Dim init_y As Integer
    Dim head_offset As Integer
    Dim grhtemp As Grh
    Static re As RECT
   
    re.Left = 0
    re.Top = 0
    re.bottom = 80
    re.Right = 76

    init_x = 25
    init_y = 20

    Dim Light(3) As Long
    Light(0) = D3DColorXRGB(255, 255, 255)
    Light(1) = D3DColorXRGB(255, 255, 255)
    Light(2) = D3DColorXRGB(255, 255, 255)
    Light(3) = D3DColorXRGB(255, 255, 255)

    If cPJ(Index).Race = eRaza.Humano Or cPJ(Index).Race = eRaza.Elfo Or cPJ(Index).Race = eRaza.ElfoOscuro Then
        head_offset = HeadOffsetAltos
    Else
        head_offset = HeadOffsetBajos
    End If

    Call Engine_BeginScene

    If cPJ(Index).Body <> 0 Then
        Call DDrawTransGrhtoSurface(BodyData(cPJ(Index).Body).Walk(3), PixelOffsetX + init_x, PixelOffsetY + init_y, 0, Light(), 0, init_x, init_y)
    End If

    If cPJ(Index).Dead Then
        Call DDrawTransGrhtoSurface(HeadData(CASPER_HEAD).Head(3), PixelOffsetX + init_x + 4, PixelOffsetY + init_y + head_offset, 0, Light(), 0, init_x, init_y)
    Else
        If cPJ(Index).Head <> 0 Then
            Call DDrawTransGrhtoSurface(HeadData(cPJ(Index).Head).Head(3), PixelOffsetX + init_x + 4, PixelOffsetY + init_y + head_offset, 0, Light(), 0, init_x, init_y)
        End If
    End If

    If cPJ(Index).helmet <> 0 Then
        Call DDrawTransGrhtoSurface(CascoAnimData(cPJ(Index).helmet).Head(3), PixelOffsetX + init_x + 4, PixelOffsetY + init_y + head_offset, 0, Light(), 0, init_x, init_y)
    End If
     
    If cPJ(Index).weapon <> 0 Then
        Call DDrawTransGrhtoSurface(WeaponAnimData(cPJ(Index).weapon).WeaponWalk(3), PixelOffsetX + init_x, PixelOffsetY + init_y, 0, Light(), 0, init_x, init_y)
    End If
     
    If cPJ(Index).shield <> 0 Then
        Call DDrawTransGrhtoSurface(ShieldAnimData(cPJ(Index).shield).ShieldWalk(3), PixelOffsetX + init_x, PixelOffsetY + init_y, 0, Light(), 0, init_x, init_y)
    End If

    Engine_EndScene re, frmPanelAccount.picChar(Index - 1).hwnd
End Sub
