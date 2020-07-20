Attribute VB_Name = "mDx8_Utilities"
Option Explicit

Public Sub Engine_Draw_Line(x1 As Single, _
                            y1 As Single, _
                            x2 As Single, _
                            y2 As Single, _
                            Optional Color As Long = -1, _
                            Optional Color2 As Long = -1)

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

Public Sub Engine_Draw_Box(ByVal X As Integer, _
                           ByVal Y As Integer, _
                           ByVal Width As Integer, _
                           ByVal Height As Integer, _
                           Color As Long)
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

Private Function Engine_Collision_Between(ByVal Value As Single, _
                                          ByVal Bound1 As Single, _
                                          ByVal Bound2 As Single) As Byte
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

Public Function Engine_Collision_Line(ByVal L1X1 As Long, _
                                      ByVal L1Y1 As Long, _
                                      ByVal L1X2 As Long, _
                                      ByVal L1Y2 As Long, _
                                      ByVal L2X1 As Long, _
                                      ByVal L2Y1 As Long, _
                                      ByVal L2X2 As Long, _
                                      ByVal L2Y2 As Long) As Byte

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

Public Function Engine_Collision_LineRect(ByVal sX As Long, _
                                          ByVal sY As Long, _
                                          ByVal SW As Long, _
                                          ByVal SH As Long, _
                                          ByVal x1 As Long, _
                                          ByVal y1 As Long, _
                                          ByVal x2 As Long, _
                                          ByVal y2 As Long) As Byte
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

Function Engine_Collision_Rect(ByVal x1 As Integer, _
                               ByVal y1 As Integer, _
                               ByVal Width1 As Integer, _
                               ByVal Height1 As Integer, _
                               ByVal x2 As Integer, _
                               ByVal y2 As Integer, _
                               ByVal Width2 As Integer, _
                               ByVal Height2 As Integer) As Boolean
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

Function Engine_Distance(ByVal x1 As Integer, _
                         ByVal y1 As Integer, _
                         ByVal x2 As Integer, _
                         ByVal y2 As Integer) As Long
    '***************************************************
    'Author: Standelf
    'Last Modification: -
    '***************************************************

    Engine_Distance = Abs(x1 - x2) + Abs(y1 - y2)
    
End Function

Public Function Engine_GetAngle(ByVal CenterX As Integer, _
                                ByVal CenterY As Integer, _
                                ByVal TargetX As Integer, _
                                ByVal TargetY As Integer) As Single

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

