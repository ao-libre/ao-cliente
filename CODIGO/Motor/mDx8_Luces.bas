Attribute VB_Name = "mDx8_Luces"
Option Explicit

'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 14/05/10
'Blisse-AO | Light Engine, Read the _
 #LightEngine to Set the type of Lights
'***************************************************

Option Base 0

Private Type tLight
    RGBcolor As D3DCOLORVALUE
    active As Boolean
    map_x As Byte
    map_y As Byte
    range As Byte
End Type
 
Private Light_List() As tLight
Private NumLights    As Byte

Public Function Create_Light_To_Map(ByVal map_x As Byte, _
                                    ByVal map_y As Byte, _
                                    Optional range As Byte = 3, _
                                    Optional ByVal Red As Byte = 255, _
                                    Optional ByVal Green As Byte = 255, _
                                    Optional ByVal Blue As Byte = 255)
    NumLights = NumLights + 1
   
    ReDim Preserve Light_List(1 To NumLights) As tLight
    
    With Light_List(NumLights)
        
        With .RGBcolor
            .r = Red
            .g = Green
            .b = Blue
            .a = 255
        End With
        
        .range = range
        .active = True
        .map_x = map_x
        .map_y = map_y
    
    End With
   
    Call LightRender(NumLights)
    
End Function

Public Function Delete_Light_To_Map(ByVal X As Byte, ByVal Y As Byte)
   
    Dim i As Long
   
    For i = 1 To NumLights

        If Light_List(i).map_x = X And Light_List(i).map_y = Y Then
            Call Delete_Light_To_Index(i)
            Exit Function
        End If
    Next i
 
End Function

#If LightEngine = 1 Then '   Luces Radiales

    Public Function Delete_Light_To_Index(ByVal light_index As Integer)
   
        Dim min_x As Integer
        Dim min_y As Integer
        Dim max_x As Integer
        Dim max_y As Integer
        Dim Ya    As Integer
        Dim Xa    As Integer
        
        Light_List(light_index).active = False
        
        min_x = Light_List(light_index).map_x - Light_List(light_index).range
        max_x = Light_List(light_index).map_x + Light_List(light_index).range
        min_y = Light_List(light_index).map_y - Light_List(light_index).range
        max_y = Light_List(light_index).map_y + Light_List(light_index).range
       
        For Ya = min_y To max_y
            For Xa = min_x To max_x

                If InMapBounds(Xa, Ya) Then
                    Call Engine_D3DColor_To_RGB_List(MapData(Xa, Ya).Engine_Light(), Estado_Actual)
                End If
                
            Next Xa
        Next Ya
   
    End Function

Private Sub LightRender(ByVal light_index As Integer)
 
    On Local Error Resume Next
 
    If light_index = 0 Then Exit Sub
    If Light_List(light_index).active = False Then Exit Sub
   
    Dim min_x        As Integer
    Dim min_y        As Integer
    Dim max_x        As Integer
    Dim max_y        As Integer
    Dim Ya           As Integer
    Dim Xa           As Integer
   
    Dim AmbientColor As D3DCOLORVALUE
    Dim LightColor   As D3DCOLORVALUE
   
    Dim XCoord       As Integer
    Dim YCoord       As Integer
   
    AmbientColor.r = Estado_Actual.r
    AmbientColor.g = Estado_Actual.g
    AmbientColor.b = Estado_Actual.b

    LightColor = Light_List(light_index).RGBcolor
       
    min_x = Light_List(light_index).map_x - Light_List(light_index).range
    max_x = Light_List(light_index).map_x + Light_List(light_index).range
    min_y = Light_List(light_index).map_y - Light_List(light_index).range
    max_y = Light_List(light_index).map_y + Light_List(light_index).range
       
    For Ya = min_y To max_y
        For Xa = min_x To max_x

            If InMapBounds(Xa, Ya) Then
                
                XCoord = Xa * 32
                YCoord = Ya * 32
                MapData(Xa, Ya).Engine_Light(0) = LightCalculate(Light_List(light_index).range, _
                                                                 Light_List(light_index).map_x * 32, _
                                                                 Light_List(light_index).map_y * 32, _
                                                                 XCoord, _
                                                                 YCoord, _
                                                                 MapData(Xa, Ya).Engine_Light(0), _
                                                                 LightColor, _
                                                                 AmbientColor)
 
                XCoord = Xa * 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).Engine_Light(1) = LightCalculate(Light_List(light_index).range, _
                                                                 Light_List(light_index).map_x * 32, _
                                                                 Light_List(light_index).map_y * 32, _
                                                                 XCoord, _
                                                                 YCoord, _
                                                                 MapData(Xa, Ya).Engine_Light(1), _
                                                                 LightColor, _
                                                                 AmbientColor)
                       
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).Engine_Light(2) = LightCalculate(Light_List(light_index).range, _
                                                                 Light_List(light_index).map_x * 32, _
                                                                 Light_List(light_index).map_y * 32, _
                                                                 XCoord, _
                                                                 YCoord, _
                                                                 MapData(Xa, Ya).Engine_Light(2), _
                                                                 LightColor, _
                                                                 AmbientColor)
   
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32
                MapData(Xa, Ya).Engine_Light(3) = LightCalculate(Light_List(light_index).range, _
                                                                 Light_List(light_index).map_x * 32, _
                                                                 Light_List(light_index).map_y * 32, _
                                                                 XCoord, _
                                                                 YCoord, _
                                                                 MapData(Xa, Ya).Engine_Light(3), _
                                                                 LightColor, _
                                                                 AmbientColor)
               
            End If
        Next Xa
    Next Ya

End Sub

Private Function LightCalculate(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim XDist As Single
    Dim YDist As Single
    Dim VertexDist As Single
    Dim pRadio As Integer
   
    Dim CurrentColor As D3DCOLORVALUE
   
    pRadio = cRadio * 32
   
    XDist = LightX + 16 - XCoord
    YDist = LightY + 16 - YCoord
   
    VertexDist = Sqr(XDist * XDist + YDist * YDist)
   
    If VertexDist <= pRadio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, VertexDist / pRadio) 'aca hay algo mal ;) Ambient color ;)
        LightCalculate = D3DColorXRGB(Round(CurrentColor.r), Round(CurrentColor.g), Round(CurrentColor.b))
    Else
        LightCalculate = TileLight
    End If
    
End Function

#Else 'Luces Normales


Private Sub LightRender(ByVal light_index As Integer)

    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim ia As Single
    Dim i As Integer
    Dim Color(3) As Long
    Dim Ya As Integer
    Dim Xa As Integer

    Dim XCoord As Integer
    Dim YCoord As Integer
    
    With Light_List(light_index)
    
        Color(0) = D3DColorARGB(255, .RGBcolor.r, .RGBcolor.g, .RGBcolor.b)
        Color(1) = Color(0)
        Color(2) = Color(0)
        Color(3) = Color(0)
    
        'Set up light borders
        min_x = .map_x - .range
        min_y = .map_y - .range
        max_x = .map_x + .range
        max_y = .map_y + .range
    
    End With
    
    'Arrange corners
    
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).Engine_Light(2) = Color(2)
    End If
    
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).Engine_Light(1) = Color(1)
    End If
    
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).Engine_Light(0) = Color(0)
    End If
    
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).Engine_Light(3) = Color(3)
    End If
    
    'Arrange borders
    
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).Engine_Light(1) = Color(1)
            MapData(X, min_y).Engine_Light(2) = Color(2)
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).Engine_Light(0) = Color(0)
            MapData(X, max_y).Engine_Light(3) = Color(3)
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).Engine_Light(2) = Color(2)
            MapData(min_x, Y).Engine_Light(3) = Color(3)
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).Engine_Light(0) = Color(0)
            MapData(max_x, Y).Engine_Light(1) = Color(1)
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).Engine_Light(0) = Color(0)
                MapData(X, Y).Engine_Light(1) = Color(1)
                MapData(X, Y).Engine_Light(2) = Color(2)
                MapData(X, Y).Engine_Light(3) = Color(3)
            End If
        Next Y
    Next X
    
    
End Sub

Private Sub Delete_Light_To_Index(ByVal light_index As Integer)
'***************************************'
'Author: Juan Martin Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim colorz As Long

    colorz = D3DColorARGB(Estado_Actual.a, Estado_Actual.r, Estado_Actual.g, Estado_Actual.b)
    
    With Light_List(light_index)
    
        'Set up light borders
        min_x = .map_x - .range
        min_y = .map_y - .range
        max_x = .map_x + .range
        max_y = .map_y + .range
    
    End With
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).Engine_Light(2) = colorz
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).Engine_Light(0) = colorz
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).Engine_Light(1) = colorz
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).Engine_Light(3) = colorz
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).Engine_Light(0) = colorz
            MapData(X, min_y).Engine_Light(2) = colorz
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).Engine_Light(1) = colorz
            MapData(X, max_y).Engine_Light(3) = colorz
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).Engine_Light(2) = colorz
            MapData(min_x, Y).Engine_Light(3) = colorz
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).Engine_Light(0) = colorz
            MapData(max_x, Y).Engine_Light(1) = colorz
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).Engine_Light(0) = colorz
                MapData(X, Y).Engine_Light(1) = colorz
                MapData(X, Y).Engine_Light(2) = colorz
                MapData(X, Y).Engine_Light(3) = colorz
            End If
        Next Y
    Next X
    
End Sub

#End If 'Terminamos de Seleccionar las luces

Public Sub DeInit_LightEngine()
    
    'Kill Font's
    Erase Light_List()
    
    'Exit, The works is done.
    Exit Sub
    
End Sub

Public Function LightRenderAll() As Boolean
    On Error GoTo handle

    If Not ArrayInitialized(Not Light_List) Then Exit Function
    
    Dim i As Long

    For i = 1 To UBound(Light_List)
        Call LightRender(i)
    Next i
    
    LightRenderAll = True

handle:
    LightRenderAll = False
    Exit Function
    
End Function

Public Function LightRemoveAll() As Boolean

    On Error GoTo handle
    
    If Not ArrayInitialized(Not Light_List) Then Exit Function
    
    Dim i As Long

    For i = 1 To UBound(Light_List)
        Call Delete_Light_To_Index(i)
    Next i
    
    LightRemoveAll = True
    
handle:
    LightRemoveAll = False
    Exit Function

End Function

