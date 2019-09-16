Attribute VB_Name = "Mod_TileEngine"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Marquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matias Fernando Pequeno
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez



Option Explicit

Public indexList(0 To 5) As Integer
Public ibQuad As DxVBLibA.Direct3DIndexBuffer8
Public vbQuadIdx As DxVBLibA.Direct3DVertexBuffer8
Dim temp_verts(3) As TLVERTEX

Public OffsetCounterX As Single
Public OffsetCounterY As Single
    
Public WeatherFogX1 As Single
Public WeatherFogY1 As Single
Public WeatherFogX2 As Single
Public WeatherFogY2 As Single
Public WeatherFogCount As Byte

Public ParticleOffsetX As Long
Public ParticleOffsetY As Long
Public LastOffsetX As Integer
Public LastOffsetY As Integer

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

Private Const GrhFogata As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180

'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamano y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    speed As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type


'Apariencia del personaje
Public Type Char
    Movement As Boolean
    active As Byte
    Heading As E_Heading
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    fX As Grh
    FxIndex As Integer
    
    Criminal As Byte
    Atacable As Byte
    
    Nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    attacking As Boolean
    
    Aura(1 To 4) As Aura
    
    ParticleIndex As Integer
    Particle_Count As Long
    Particle_Group() As Long
End Type

'Info de un objeto
Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    Damage As DList
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    Engine_Light(0 To 3) As Long 'Standelf, Light Engine.
    Particle_Group_Index As Long 'Particle Engine
    
    fX As Grh
    FxIndex As Integer
End Type

'Info de cada mapa
Public Type mapInfo
    Music As String
    name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

Public IniPath As String

'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Public FPSLastCheck As Long

'Tamano del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamano muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

'Tamano de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Dim timerElapsedTime As Single
Public timerTicksPerFrame As Single

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte




'?????????Graficos???????????
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'?????????????????????????

'?????????Mapa????????????
Public MapData() As MapBlock ' Mapa
Public mapInfo As mapInfo ' Info acerca del mapa en uso
'?????????????????????????

Public Normal_RGBList(3) As Long
Public Color_Shadow(3) As Long
Public Color_Arbol(3) As Long

'   Control de Lluvia
Public bRain As Boolean
Public bTecho       As Boolean 'hay techo?
Public bFogata       As Boolean

Public charlist(1 To 10000) As Char

' Used by GetTextExtentPoint32
Private Type Size
    cx As Long
    cy As Long
End Type

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'?????????????????????????


'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'attack state
        .attacking = False
        
        'Make active
        .active = 1
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
End Sub




Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.speed = GrhData(Grh.GrhIndex).speed
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addx As Integer
    Dim addy As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
        End Select
        
        nX = X + addx
        nY = Y + addy
        
        If nX <= 0 Then nX = 1
        If nY <= 0 Then nY = 1
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        
        If (X Or Y) = 0 Then Exit Sub
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy
    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        If CharIndex <> UserCharIndex Then
            Call Char_Erase(CharIndex)
        End If
    End If
End Sub

Public Sub DoFogataFx()
    Dim Location As Position
    
    If bFogata Then
        bFogata = HayFogata(Location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(Location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", Location.X, Location.Y, LoopStyle.Enabled)
    End If
End Sub

Public Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 09/21/2010
' 09/21/2010: C4b3z0n - Changed from Private Funtion tu Public Function.
'***************************************************
    
    With charlist(CharIndex).Pos
    
        EstaPCarea = .X > UserPos.X - MinXBorder And _
                     .X < UserPos.X + MinXBorder And _
                     .Y > UserPos.Y - MinYBorder And _
                     .Y < UserPos.Y + MinYBorder
    End With
    
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
    If Not UserNavegando Then
        With charlist(CharIndex)
            If Not .muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5) Then
                .pie = Not .pie
                
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS1, .Pos.X, .Pos.Y)
                Else
                    Call Audio.PlayWave(SND_PASOS2, .Pos.X, .Pos.Y)
                End If
            End If
        End With
    Else
' TODO : Actually we would have to check if the CharIndex char is in the water or not....
        Call Audio.PlayWave(SND_NAVEGANDO, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y)
    End If
End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim X As Integer
    Dim Y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addx = nX - X
        addy = nY - Y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
            .FxIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call Char_Erase(CharIndex)
    End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
    End If
End Sub

Private Function HayFogata(ByRef Location As Position) As Boolean
    Dim J As Long
    Dim k As Long
    
    For J = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(J, k) Then
                If MapData(J, k).ObjGrh.GrhIndex = GrhFogata Then
                    Location.X = J
                    Location.Y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next J
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim loopc As Long
    Dim Dale As Boolean
    
    loopc = 1
    Do While charlist(loopc).active And Dale
        loopc = loopc + 1
        Dale = (loopc <= UBound(charlist))
    Loop
    
    NextOpenChar = loopc
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Public Function LoadGrhData() As Boolean
On Error Resume Next
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    'Open files
    handle = FreeFile()
    Open IniPath & "Graficos.ind" For Binary Access Read As handle
    Get handle, , fileVersion
    
    Get handle, , grhCount
    
    ReDim GrhData(0 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
        
        With GrhData(Grh)
           ' GrhData(Grh).active = True
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then Resume Next
            
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        Resume Next
                    End If
                Next Frame
                
                Get handle, , .speed
                
                If .speed <= 0 Then Resume Next
                
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then Resume Next
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then Resume Next
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then Resume Next
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then Resume Next
            Else
                Get handle, , .FileNum
                If .FileNum <= 0 Then Resume Next
                
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then Resume Next
                
                Get handle, , .sY
                If .sY < 0 Then Resume Next
                
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then Resume Next
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then Resume Next
                
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
    Wend
    
    Close handle

    
    LoadGrhData = True
Exit Function

ErrorHandler:
    LoadGrhData = False
End Function


Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    'Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    If UserEvento Then
        Exit Function
    End If
    
    LegalPos = True
End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 01/08/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
'*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(X, Y).CharIndex
    'Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.X, UserPos.Y).Blocked = 1 Then
            Exit Function
        End If
        
        With charlist(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> eCabezas.CASPER_HEAD And .iBody <> eCabezas.FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.X, UserPos.Y) Then
                    If Not HayAgua(X, Y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(X, Y) Then Exit Function
                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 6 Then
                    If charlist(UserCharIndex).invisible = True Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Function GetBitmapDimensions(ByVal BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
    Dim BMHeader As BITMAPFILEHEADER
    Dim BINFOHeader As BITMAPINFOHEADER
    
    Open BmpFile For Binary Access Read As #1
    
    Get #1, , BMHeader
    Get #1, , BINFOHeader
    
    Close #1
    
    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight
End Function

Public Sub DrawTransparentGrhtoHdc(ByVal dsthdc As Long, ByVal srchdc As Long, ByRef SourceRect As RECT, ByRef destRect As RECT, ByVal TransparentColor As Long)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 27/07/2012 - ^[GS]^
'*************************************************************
    Dim Color As Long
    Dim X As Long
    Dim Y As Long
    
    For X = SourceRect.Left To SourceRect.Right
        For Y = SourceRect.Top To SourceRect.Bottom
            Color = GetPixel(srchdc, X, Y)
            
            If Color <> TransparentColor Then
                Call SetPixel(dsthdc, destRect.Left + (X - SourceRect.Left), destRect.Top + (Y - SourceRect.Top), Color)
            End If
        Next Y
    Next X
End Sub

Public Sub DrawImageInPicture(ByRef PictureBox As PictureBox, ByRef Picture As StdPicture, ByVal x1 As Single, ByVal y1 As Single, Optional Width1, Optional Height1, Optional x2, Optional y2, Optional Width2, Optional Height2)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 12/28/2009
'Draw Picture in the PictureBox
'*************************************************************

Call PictureBox.PaintPicture(Picture, x1, y1, Width1, Height1, x2, y2, Width2, Height2)

End Sub

Sub RenderScreen(ByVal tilex As Integer, _
                 ByVal tiley As Integer, _
                 ByVal PixelOffsetX As Integer, _
                 ByVal PixelOffsetY As Integer)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Renders everything to the viewport
    '**************************************************************
    Dim Y                As Long     'Keeps track of where on map we are
    Dim X                As Long     'Keeps track of where on map we are
    
    Dim screenminY       As Integer  'Start Y pos on current screen
    Dim screenmaxY       As Integer  'End Y pos on current screen
    Dim screenminX       As Integer  'Start X pos on current screen
    Dim screenmaxX       As Integer  'End X pos on current screen
    
    Dim minY             As Integer  'Start Y pos on current map
    Dim maxY             As Integer  'End Y pos on current map
    Dim minX             As Integer  'Start X pos on current map
    Dim maxX             As Integer  'End X pos on current map
    
    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen
    
    Dim minXOffset       As Integer
    Dim minYOffset       As Integer
    
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    
    Dim ElapsedTime      As Single
    
    ElapsedTime = Engine_ElapsedTime()
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - Engine_Get_TileBuffer
    maxY = screenmaxY + Engine_Get_TileBuffer
    minX = screenminX - Engine_Get_TileBuffer
    maxX = screenmaxX + Engine_Get_TileBuffer
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    ParticleOffsetX = (Engine_PixelPosX(screenminX) - PixelOffsetX)
    ParticleOffsetY = (Engine_PixelPosY(screenminY) - PixelOffsetY)

    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            
            PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY
            
            'Layer 1 **********************************
            If MapData(X, Y).Graphic(1).GrhIndex <> 0 Then
                Call Draw_Grh(MapData(X, Y).Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1)
            End If
            '******************************************

            'Layer 2 **********************************
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call Draw_Grh(MapData(X, Y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1)
            End If
            '******************************************
            
            ScreenX = ScreenX + 1
        Next
    
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next
   
    
    '<----- Layer Obj, Char, 3 ----->
    ScreenY = minYOffset - Engine_Get_TileBuffer

    For Y = minY To maxY
        
        ScreenX = minXOffset - Engine_Get_TileBuffer

        For X = minX To maxX
            If Map_InBounds(X, Y) Then
            
                PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
                PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
                
                With MapData(X, Y)
                    'Object Layer **********************************
                    If .ObjGrh.GrhIndex <> 0 Then
                        Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                    End If
                    '***********************************************
        
        
                    'Char layer********************************
                    If .CharIndex <> 0 Then
                        Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                    End If
                    '*************************************************
        
                    
                    
                    'Layer 3 *****************************************
                    If .Graphic(3).GrhIndex <> 0 Then
                        If .Graphic(3).GrhIndex = 735 Or .Graphic(3).GrhIndex >= 6994 And .Graphic(3).GrhIndex <= 7002 Then
                            If Abs(UserPos.X - X) < 3 And (Abs(UserPos.Y - Y)) < 8 And (Abs(UserPos.Y) < Y) Then
                                Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, Color_Arbol(), 1)
                            Else 'NORMAL
                                Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                            End If
                        Else 'NORMAL
                            Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                        End If
                    End If
                    '************************************************
                    
                    'Dibujamos los danos.
                    If .Damage.Activated Then _
                        Call mDx8_Dibujado.Damage_Draw(X, Y, PixelOffsetXTemp, PixelOffsetYTemp - 20)
                
                    'Particulas ORE
                    If .Particle_Group_Index Then
                        If Abs(UserPos.X - X) < Engine_Get_TileBuffer + 3 And (Abs(UserPos.Y - Y)) < Engine_Get_TileBuffer + 3 Then _
                            Call Particle_Group_Render(.Particle_Group_Index, PixelOffsetXTemp + 16, PixelOffsetYTemp + 16)
                    End If
                    
                    'Particulas vbGORE
                    Call Effect_UpdateAll
                    
                    If Not .FxIndex = 0 Then
                        Call Draw_Grh(.fX, PixelOffsetXTemp + FxData(MapData(X, Y).FxIndex).OffsetX, PixelOffsetYTemp + FxData(.FxIndex).OffsetY, 1, .Engine_Light(), 1, True)
                        If .fX.Started = 0 Then .FxIndex = 0
                    End If
                    
                    
                End With
                
            End If
            
            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y
    
    '<----- Layer 4 ----->
    ScreenY = minYOffset - Engine_Get_TileBuffer

    For Y = minY To maxY
        ScreenX = minXOffset - Engine_Get_TileBuffer

        For X = minX To maxX
   
            'Layer 4
            If MapData(X, Y).Graphic(4).GrhIndex Then
                If bTecho Then
                    Call Draw_Grh(MapData(X, Y).Graphic(4), ScreenX * TilePixelWidth + PixelOffsetX, ScreenY * TilePixelHeight + PixelOffsetY, 1, temp_rgb(), 1)
                Else
                    If ColorTecho = 250 Then
                        Call Draw_Grh(MapData(X, Y).Graphic(4), ScreenX * TilePixelWidth + PixelOffsetX, ScreenY * TilePixelHeight + PixelOffsetY, 1, MapData(X, Y).Engine_Light(), 1)
                    Else
                        Call Draw_Grh(MapData(X, Y).Graphic(4), ScreenX * TilePixelWidth + PixelOffsetX, ScreenY * TilePixelHeight + PixelOffsetY, 1, temp_rgb(), 1)
                    End If
                End If
            End If
            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y

    If ClientSetup.ProyectileEngine Then
                            
        If LastProjectile > 0 Then
            Dim J As Long ' Long siempre en los bucles es mucho mas rapido
                                
            For J = 1 To LastProjectile

                If ProjectileList(J).Grh.GrhIndex Then
                    Dim angle As Single
                    
                    'Update the position
                    angle = DegreeToRadian * Engine_GetAngle(ProjectileList(J).X, ProjectileList(J).Y, ProjectileList(J).tX, ProjectileList(J).tY)
                    ProjectileList(J).X = ProjectileList(J).X + (Sin(angle) * ElapsedTime * 0.63)
                    ProjectileList(J).Y = ProjectileList(J).Y - (Cos(angle) * ElapsedTime * 0.63)
                    
                    'Update the rotation
                    If ProjectileList(J).RotateSpeed > 0 Then
                        ProjectileList(J).Rotate = ProjectileList(J).Rotate + (ProjectileList(J).RotateSpeed * ElapsedTime * 0.01)

                        Do While ProjectileList(J).Rotate > 360
                            ProjectileList(J).Rotate = ProjectileList(J).Rotate - 360
                        Loop
                    End If
    
                    'Draw if within range
                    X = ((-minX - 1) * 32) + ProjectileList(J).X + PixelOffsetX + ((10 - TileBufferSize) * 32) - 288 + ProjectileList(J).OffsetX
                    Y = ((-minY - 1) * 32) + ProjectileList(J).Y + PixelOffsetY + ((10 - TileBufferSize) * 32) - 288 + ProjectileList(J).OffsetY

                    If Y >= -32 Then
                        If Y <= (ScreenHeight + 32) Then
                            If X >= -32 Then
                                If X <= (ScreenWidth + 32) Then
                                    If ProjectileList(J).Rotate = 0 Then
                                        Call Draw_Grh(ProjectileList(J).Grh, X, Y, 0, MapData(50, 50).Engine_Light(), 0, True, 0)
                                    Else
                                        Call Draw_Grh(ProjectileList(J).Grh, X, Y, 0, MapData(50, 50).Engine_Light(), 0, True, ProjectileList(J).Rotate)
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                End If
            Next J
            
            'Check if it is close enough to the target to remove
            For J = 1 To LastProjectile

                If ProjectileList(J).Grh.GrhIndex Then
                    If Abs(ProjectileList(J).X - ProjectileList(J).tX) < 20 Then
                        If Abs(ProjectileList(J).Y - ProjectileList(J).tY) < 20 Then
                            Call Engine_Projectile_Erase(J)
                        End If
                    End If
                End If
            Next J
            
        End If
    End If
    If colorRender <> 240 Then _
    Call DrawText(272, 50, renderText, render_msg(0), True, 2)
    
    '   Set Offsets
    LastOffsetX = ParticleOffsetX
    LastOffsetY = ParticleOffsetY
    
    If ClientSetup.PartyMembers Then Call Draw_Party_Members

    Call RenderCount
End Sub

Public Function RenderSounds()
'**************************************************************
'Author: Juan Martin Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
Dim Location As Position

        If bRain And bLluvia(UserMap) Then
                If bTecho Then
                    If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                        If RainBufferIndex Then _
                            Call Audio.StopWave(RainBufferIndex)
                        RainBufferIndex = Audio.PlayWave("lluviain.wav", 0, 0, LoopStyle.Enabled)
                        frmMain.IsPlaying = PlayLoop.plLluviain
                    End If
                Else
                    If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                        If RainBufferIndex Then _
                            Call Audio.StopWave(RainBufferIndex)
                        RainBufferIndex = Audio.PlayWave("lluviaout.wav", 0, 0, LoopStyle.Enabled)
                        frmMain.IsPlaying = PlayLoop.plLluviaout
                    End If
                End If
        End If

        If bFogata Then
                bFogata = Map_CheckBonfire(Location)

                If Not bFogata Then
                        Call Audio.StopWave(FogataBufferIndex)
                        FogataBufferIndex = 0
                End If

        Else
                bFogata = Map_CheckBonfire(Location)

                If bFogata And FogataBufferIndex = 0 Then
                        FogataBufferIndex = Audio.PlayWave("fuego.wav", Location.X, Location.Y, LoopStyle.Enabled)
                End If
        End If
End Function

Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Long) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.Y <= Y
    End If
End Function

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer) As Boolean
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
'Configures the engine to start running.
'***************************************************
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = Round(frmMain.MainViewPic.Height / 32, 0)
    WindowTileWidth = Round(frmMain.MainViewPic.Width / 32, 0)
    
    IniPath = Game.path(INIT)
    HalfWindowTileHeight = WindowTileHeight \ 2
    HalfWindowTileWidth = WindowTileWidth \ 2

    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    

    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY

On Error GoTo 0

    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    Call LoadGraphics
    Call CargarParticulas
    
    indexList(0) = 0: indexList(1) = 1: indexList(2) = 2
    indexList(3) = 3: indexList(4) = 4: indexList(5) = 5
    
    Set ibQuad = DirectDevice.CreateIndexBuffer(Len(indexList(0)) * 4, 0, D3DFMT_INDEX16, D3DPOOL_MANAGED)
    
    Call D3DIndexBuffer8SetData(ibQuad, 0, Len(indexList(0)) * 4, 0, indexList(0))
    
    Set vbQuadIdx = DirectDevice.CreateVertexBuffer(Len(temp_verts(0)) * 4, 0, D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1, D3DPOOL_MANAGED)
    
    InitTileEngine = True
End Function

Public Sub LoadGraphics()
    Call SurfaceDB.Initialize(DirectD3D8, ClientSetup.byMemory)
End Sub

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, _
                  ByVal DisplayFormLeft As Integer, _
                  ByVal MouseViewX As Integer, _
                  ByVal MouseViewY As Integer)

    If EngineRun Then
        Call Engine_BeginScene
        
        Call DesvanecimientoTechos
        Call DesvanecimientoMsg
        
        If UserMoving Then
        
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame

                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False

                End If
                
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame

                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                    
                End If

            End If

        End If
        
        'Update mouse position within view area
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
        
        '****** Update screen ******
        If UserCiego Then
            DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        Else
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)

        End If
        
        Call Dialogos.Render
        Call DibujarCartel

        If ClientSetup.bGuildNews Then Call DialogosClanes.Draw

        ' Calculamos los FPS y los mostramos
        Call Engine_Update_FPS

        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * Engine_Get_BaseSpeed
        
        Call Engine_EndScene(MainScreenRect, 0)
        
        '//Banco
        If frmBancoObj.Visible Then
            If frmBancoObj.PicBancoInv.Visible Then Call InvBanco(0).DrawInv
            If frmBancoObj.PicInv.Visible Then Call InvBanco(1).DrawInv
        End If
    
        '//Comercio
        If frmComerciar.Visible Then
            If frmComerciar.picInvNpc.Visible Then Call InvComNpc.DrawInv
            If frmComerciar.picInvUser.Visible Then Call InvComUsu.DrawInv
        End If
    
        '//Comercio entre usuarios
        If frmComerciarUsu.Visible Then
            If frmComerciarUsu.picInvComercio.Visible Then InvComUsu.DrawInv (1)
            If frmComerciarUsu.picInvOfertaProp.Visible Then InvOfferComUsu(0).DrawInv (1)
            If frmComerciarUsu.picInvOfertaOtro Then InvOfferComUsu(1).DrawInv (1)
            If frmComerciarUsu.picInvOroProp.Visible Then InvOroComUsu(0).DrawInv (1)
            If frmComerciarUsu.picInvOroOfertaProp.Visible Then InvOroComUsu(1).DrawInv (1)
            If frmComerciarUsu.picInvOroOfertaOtro.Visible Then InvOroComUsu(2).DrawInv (1)
        End If
    
        '//Herrero
        If frmHerrero.Visible Then
            If frmHerrero.picLingotes0.Visible Or frmHerrero.picMejorar0.Visible Then InvLingosHerreria(1).DrawInv (1)
            If frmHerrero.picLingotes1.Visible Or frmHerrero.picMejorar1.Visible Then InvLingosHerreria(2).DrawInv (1)
            If frmHerrero.picLingotes2.Visible Or frmHerrero.picMejorar2.Visible Then InvLingosHerreria(3).DrawInv (1)
            If frmHerrero.picLingotes3.Visible Or frmHerrero.picMejorar3.Visible Then InvLingosHerreria(4).DrawInv (1)
        End If
    
        '//Carpintero
        If frmCarp.Visible Then
            If frmCarp.picMaderas0.Visible Or frmCarp.imgMejorar0.Visible Then InvMaderasCarpinteria(1).DrawInv (1)
            If frmCarp.picMaderas1.Visible Or frmCarp.imgMejorar1.Visible Then InvMaderasCarpinteria(2).DrawInv (1)
            If frmCarp.picMaderas2.Visible Or frmCarp.imgMejorar2.Visible Then InvMaderasCarpinteria(3).DrawInv (1)
            If frmCarp.picMaderas3.Visible Or frmCarp.imgMejorar3.Visible Then InvMaderasCarpinteria(4).DrawInv (1)
        End If
    
        '//Inventario
        If frmMain.Visible Then
            If frmMain.PicInv.Visible Then
                Call Inventario.DrawInv
            End If
        End If
        

    End If
  
End Sub

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim Start_Time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Private Sub CharRender(ByVal CharIndex As Long, _
                       ByVal PixelOffsetX As Integer, _
                       ByVal PixelOffsetY As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 16/09/2010 (Zama)
    'Draw char's to screen without offcentering them
    '16/09/2010: ZaMa - Ya no se dibujan los bodies cuando estan invisibles.
    '***************************************************
    Dim moved As Boolean
    
    With charlist(CharIndex)

        If .Moving Then

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0

                End If

            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0

                End If

            End If

        End If
        
        'if is attacking we set the attack anim
        If .attacking And .Arma.WeaponWalk(.Heading).Started = 0 Then
            .Arma.WeaponWalk(.Heading).Started = 1
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
        
            'if the anim has ended or we are no longer attacking end the animation
        ElseIf .Arma.WeaponWalk(.Heading).FrameCounter > 4 And .attacking Then
            .attacking = False 'this is just for testing, it shouldnt be done here

        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            
            '//Evito runtime
            If Not .Heading <> 0 Then .Heading = EAST
            
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            '//Movimiento del arma y el escudo
            If Not .Movement And Not .attacking Then
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 1
                
                .Escudo.ShieldWalk(.Heading).Started = 0
                .Escudo.ShieldWalk(.Heading).FrameCounter = 1

            End If
            
            .Moving = False

        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        Dim ColorFinal(0 To 3) As Long
        Dim RenderSpell        As Boolean
        
        If Not .muerto Then
        
            If Abs(MouseTileX - .Pos.X) < 1 And (Abs(MouseTileY - .Pos.Y)) < 1 And CharIndex <> UserCharIndex And ClientSetup.TonalidadPJ Then
                
                If Len(.Nombre) > 0 Then
                    
                    If .Criminal Then
                        Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorXRGB(204, 100, 100))
                    Else
                        Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorXRGB(100, 100, 255))

                    End If
                    
                Else
                    ColorFinal(0) = MapData(.Pos.X, .Pos.Y).Engine_Light(0)
                    ColorFinal(1) = MapData(.Pos.X, .Pos.Y).Engine_Light(1)
                    ColorFinal(2) = MapData(.Pos.X, .Pos.Y).Engine_Light(2)
                    ColorFinal(3) = MapData(.Pos.X, .Pos.Y).Engine_Light(3)

                End If

                RenderSpell = True
            
            Else
                ColorFinal(0) = MapData(.Pos.X, .Pos.Y).Engine_Light(0)
                ColorFinal(1) = MapData(.Pos.X, .Pos.Y).Engine_Light(1)
                ColorFinal(2) = MapData(.Pos.X, .Pos.Y).Engine_Light(2)
                ColorFinal(3) = MapData(.Pos.X, .Pos.Y).Engine_Light(3)

            End If

        Else

            If esGM(Val(CharIndex)) Then
                Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorARGB(150, 200, 200, 0))
            Else

                If .Criminal Then
                    Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorARGB(100, 255, 100, 100))
                Else
                    Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorARGB(100, 128, 255, 255))
                End If

            End If

        End If
        
        If Not .invisible Then
            Movement_Speed = 0.5
            
            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then
                Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1)
            End If
            
            'Draw name when navigating
            If Len(.Nombre) > 0 Then
                If Nombres Then
                    If .iHead = 0 And .iBody > 0 Then
                        Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY)
                    End If
                End If
            End If
            
            'Draw Head
            If .Head.Head(.Heading).GrhIndex Then
                Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, ColorFinal(), 0)
                
                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then
                    Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, ColorFinal(), 0)
                End If
                
                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then
                    Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1)
                End If
                
                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then
                    Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1)
                End If
            
                'Draw name over head
                If LenB(.Nombre) > 0 Then
                    If Nombres Then
                        Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY)
                    End If
                End If
                
                '************Particulas************
                Dim i As Integer
                If .Particle_Count > 0 Then

                    For i = 1 To .Particle_Count
                    
                        If .Particle_Group(i) > 0 Then
                            Call Particle_Group_Render(.Particle_Group(i), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY)
                        End If
                        
                    Next i

                End If
            
            Else 'Usuario invisible
        
                If CharIndex = UserCharIndex Or mid$(charlist(CharIndex).Nombre, getTagPosition(.Nombre)) = mid$(charlist(UserCharIndex).Nombre, getTagPosition(charlist(UserCharIndex).Nombre)) And Len(mid$(charlist(CharIndex).Nombre, getTagPosition(.Nombre))) > 0 Then
                
                    Movement_Speed = 0.5
                
                    'Draw Body
                    If .Body.Walk(.Heading).GrhIndex Then
                        Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, True)
                    End If
                
                    'Draw Head
                    If .Head.Head(.Heading).GrhIndex Then
                        Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, ColorFinal(), 0, True)
                        
                        'Draw Helmet
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, ColorFinal(), 0, True)
                    
                        'Draw Weapon
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, True)
                    
                        'Draw Shield
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, True)
                
                        'Draw name over head
                        If LenB(.Nombre) > 0 Then
                            If Nombres Then
                                Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY, True)
                            End If

                        End If

                    End If

                End If
            
            End If

        End If
        
        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, CharIndex) '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo
        
        Movement_Speed = 1
        
        'Draw FX
        If .FxIndex <> 0 Then
            Call Draw_Grh(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, SetARGB_Alpha(MapData(.Pos.X, .Pos.Y).Engine_Light(), 180), 1, True)
            
            'Check if animation is over
            If .fX.Started = 0 Then .FxIndex = 0
            
        End If
        
    End With
    
End Sub

Private Sub RenderName(ByVal CharIndex As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal Invi As Boolean = False)
    Dim Pos As Integer
    Dim line As String
    Dim Color As Long
   
    With charlist(CharIndex)
            Pos = getTagPosition(.Nombre)
    
            If .priv = 0 Then
                    If .muerto Then
                        Color = D3DColorARGB(255, 220, 220, 255)
                    Else
                        If .Criminal Then
                            Color = ColoresPJ(50)
                        Else
                            Color = ColoresPJ(49)
                        End If
                    End If
            Else
                Color = ColoresPJ(.priv)
            End If
    
            If Invi Then
                Color = D3DColorARGB(180, 150, 180, 220)
            End If

            'Nick
            line = Left$(.Nombre, Pos - 2)
            Call DrawText(X + 16, Y + 30, line, Color, True)
            
            'Clan
            line = mid$(.Nombre, Pos)
            Call DrawText(X + 16, Y + 45, line, Color, True)

    End With
End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    
    With charlist(CharIndex)
        .FxIndex = fX
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
            .fX.Loops = Loops
        End If
        
    End With
    
End Sub

Public Sub Device_Textured_Render(ByVal X As Single, ByVal Y As Single, _
                                  ByVal Width As Integer, ByVal Height As Integer, _
                                  ByVal sX As Integer, ByVal sY As Integer, _
                                  ByVal tex As Long, _
                                  ByRef Color() As Long, _
                                  Optional ByVal Alpha As Boolean = False, _
                                  Optional ByVal angle As Single = 0)

        Dim Texture As Direct3DTexture8
        
        Dim TextureWidth As Long, TextureHeight As Long
        Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
        
        With SpriteBatch

                Call .SetTexture(Texture)
                    
                If TextureWidth <> 0 And TextureHeight <> 0 Then
                    Call .Draw(X, Y, Width, Height, Color, sX / TextureWidth, sY / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight, Alpha, angle)
                Else
                    Call .Draw(X, Y, TextureWidth, TextureHeight, Color, , , , , Alpha, angle)
                End If
                
        End With
        
End Sub

Public Sub RenderItem(ByVal hWndDest As Long, ByVal GrhIndex As Long)
    Dim DR As RECT
    
    With DR
        .Left = 0
        .Top = 0
        .Right = 32
        .Bottom = 32
    End With
    
    Call Engine_BeginScene

    Call Draw_GrhIndex(GrhIndex, 0, 0, 0, Normal_RGBList(), 0, False)
        
    Call Engine_EndScene(DR, hWndDest)
    
End Sub

Sub Draw_GrhIndex(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, Optional ByVal angle As Single = 0, Optional ByVal Alpha As Boolean = False)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If

        'Draw
        Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List())
    End With
    
End Sub

Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, ByVal Animate As Byte, Optional ByVal Alpha As Boolean = False, Optional ByVal angle As Single = 0)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    
    If Grh.GrhIndex = 0 Then Exit Sub
    
On Error GoTo Error
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.speed) * Movement_Speed
            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If

        Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha, angle)
        
    End With
    
Exit Sub

Error:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        'Call Log_Engine("Error in Draw_Grh, " & Err.Description & ", (" & Err.number & ")")
        MsgBox "Error en el Engine Grafico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
        Call CloseClient
    End If
End Sub

Public Function GrhCheck(ByVal GrhIndex As Long) As Boolean
        '**************************************************************
        'Author: Aaron Perkins - Modified by Juan Martin Sotuyo Dodero
        'Last Modify Date: 1/04/2003
        '
        '**************************************************************
        'check grh_index

        If GrhIndex > 0 And GrhIndex <= UBound(GrhData()) Then
                GrhCheck = GrhData(GrhIndex).NumFrames
        End If

End Function

Public Sub GrhUninitialize(Grh As Grh)
        '*****************************************************************
        'Author: Aaron Perkins
        'Last Modify Date: 1/04/2003
        'Resets a Grh
        '*****************************************************************

        With Grh
        
                'Copy of parameters
                .GrhIndex = 0
                .Started = False
                .Loops = 0
        
                'Set frame counters
                .FrameCounter = 0
                .speed = 0
                
        End With

End Sub

Public Sub DesvanecimientoTechos()
 
    If bTecho Then
        If Not Val(ColorTecho) = 150 Then ColorTecho = ColorTecho - 1
    Else
        If Not Val(ColorTecho) = 250 Then ColorTecho = ColorTecho + 1
    End If
    
    If Not Val(ColorTecho) = 250 Then
        Call Engine_Long_To_RGB_List(temp_rgb(), D3DColorARGB(ColorTecho, ColorTecho, ColorTecho, ColorTecho))
    End If
    
End Sub

Public Sub DesvanecimientoMsg()
'*****************************************************************
'Author: FrankoH
'Last Modify Date: 04/09/2019
'DESVANECIMIENTO DE LOS TEXTOS DEL RENDER
'*****************************************************************
    Static lastmovement As Long
    
    If GetTickCount - lastmovement > 1 Then
        lastmovement = GetTickCount
    Else
        Exit Sub
    End If

    If LenB(renderText) Then
        If Not Val(colorRender) = 0 Then colorRender = colorRender - 1
    ElseIf LenB(renderText) = 0 Then
        Exit Sub
    Else
        If Not Val(colorRender) = 240 Then colorRender = colorRender + 1
    End If
    
    If Not Val(colorRender) = 240 Then
        Call Engine_Long_To_RGB_List(render_msg(), ARGB(255, 255, 255, colorRender))
    End If
    
    If colorRender = 0 Then renderMsgReset
    
End Sub

Public Sub renderMsgReset()

    renderFont = 1
    renderText = vbNullString
    nameMap = vbNullString

End Sub
