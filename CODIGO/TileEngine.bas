Attribute VB_Name = "Mod_TileEngine"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez



Option Explicit

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

Private Const GrhFogata As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1


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
    x As Long
    y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    x As Integer
    y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
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
    
    Speed As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
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
    Active As Byte
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
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

'DX7 Objects
Public DirectX As New DirectX7
Public DirectDraw As DirectDraw7
Private PrimarySurface As DirectDrawSurface7
Private PrimaryClipper As DirectDrawClipper
Private BackBufferSurface As DirectDrawSurface7

Public IniPath As String
Public MapPath As String


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
Private fpsLastCheck As Long

'Tamaño del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer

'Offset del desde 0,0 del main view
Private MainViewTop As Integer
Private MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Dim timerElapsedTime As Single
Dim timerTicksPerFrame As Single
Dim engineBaseSpeed As Single


Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer


Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte




'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer

Public charlist(1 To 10000) As Char

#If SeguridadAlkon Then

Public MI(1 To 1233) As clsManagerInvisibles
Public CualMI As Integer

#End If

'estados internos del surface (read only)
Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

#If ConAlfaB Then

Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long

#End If

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open App.path & "\init\Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
End Sub

Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    N = FreeFile()
    Open App.path & "\init\Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
End Sub

Sub CargarTips()
    Dim N As Integer
    Dim i As Long
    Dim NumTips As Integer
    
    N = FreeFile
    Open App.path & "\init\Tips.ayu" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumTips
    
    'Resize array
    ReDim Tips(1 To NumTips) As String * 255
    
    For i = 1 To NumTips
        Get #N, , Tips(i)
    Next i
    
    Close #N
End Sub

Sub CargarArrayLluvia()
    Dim N As Integer
    Dim i As Long
    Dim Nu As Integer
    
    N = FreeFile()
    Open App.path & "\init\fk.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Nu
    
    'Resize array
    ReDim bLluvia(1 To Nu) As Byte
    
    For i = 1 To Nu
        Get #N, , bLluvia(i)
    Next i
    
    Close #N
End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.x + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .Active = 0 Then _
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
        .Pos.x = x
        .Pos.y = y
        
        'Make active
        .Active = 1
    End With
    
    'Plot on map
    MapData(x, y).CharIndex = CharIndex
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        .Active = 0
        .Criminal = 0
        .FxIndex = 0
        .invisible = False
    
#If SeguridadAlkon Then
        Call MI(CualMI).ResetInvisible(CharIndex)
#End If
    
        .Moving = 0
        .muerto = False
        .Nombre = ""
        .pie = False
        .Pos.x = 0
        .Pos.y = 0
        .UsandoArma = False
    End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    charlist(CharIndex).Active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).Active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y).CharIndex = 0
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
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
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addX As Integer
    Dim addY As Integer
    Dim x As Integer
    Dim y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        x = .Pos.x
        y = .Pos.y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addY = -1
        
            Case E_Heading.EAST
                addX = 1
        
            Case E_Heading.SOUTH
                addY = 1
            
            Case E_Heading.WEST
                addX = -1
        End Select
        
        nX = x + addX
        nY = y + addY
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.x = nX
        .Pos.y = nY
        MapData(x, y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addX
        .scrollDirectionY = addY
    End With
    
    If UserEstado <> 1 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.x, location.y, LoopStyle.Enabled)
    End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .x > UserPos.x - MinXBorder And .x < UserPos.x + MinXBorder And .y > UserPos.y - MinYBorder And .y < UserPos.y + MinYBorder
    End With
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
    If Not UserNavegando Then
        With charlist(CharIndex)
            If Not .muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5) Then
                .pie = Not .pie
                
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS1, .Pos.x, .Pos.y)
                Else
                    Call Audio.PlayWave(SND_PASOS2, .Pos.x, .Pos.y)
                End If
            End If
        End With
    Else
' TODO : Actually we would have to check if the CharIndex char is in the water or not....
        Call Audio.PlayWave(SND_NAVEGANDO, charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y)
    End If
End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim x As Integer
    Dim y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        x = .Pos.x
        y = .Pos.y
        
        MapData(x, y).CharIndex = 0
        
        addX = nX - x
        addY = nY - y
        
        If Sgn(addX) = 1 Then
            nHeading = E_Heading.EAST
        End If
        
        If Sgn(addX) = -1 Then
            nHeading = E_Heading.WEST
        End If
        
        If Sgn(addY) = -1 Then
            nHeading = E_Heading.NORTH
        End If
        
        If Sgn(addY) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        
        .Pos.x = nX
        .Pos.y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addX)
        .scrollDirectionY = Sgn(addY)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Then
            .FxIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim x As Integer
    Dim y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            y = -1
        
        Case E_Heading.EAST
            x = 1
        
        Case E_Heading.SOUTH
            y = 1
        
        Case E_Heading.WEST
            x = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.y + y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.y = y
        UserPos.y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
    End If
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.x - 8 To UserPos.x + 8
        For k = UserPos.y - 6 To UserPos.y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    location.x = j
                    location.y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim loopc As Long
    Dim Dale As Boolean
    
    loopc = 1
    Do While charlist(loopc).Active And Dale
        loopc = loopc + 1
        Dale = (loopc <= UBound(charlist))
    Loop
    
    NextOpenChar = loopc
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhData() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    Open IniPath & "Graficos.ind" For Binary Access Read As handle
    Seek #1, 1
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
        
        With GrhData(Grh)
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next Frame
                
                Get handle, , .Speed
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
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

Function LegalPos(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(x, y).Blocked = 1 Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(x, y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> HayAgua(x, y) Then
        Exit Function
    End If
    
    LegalPos = True
End Function

Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Private Sub DDrawGrhtoSurface(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Byte)
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
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
    
    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
        End If
        
        If GrhData(CurrentGrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If
    
    With SourceRect
        .Left = GrhData(CurrentGrhIndex).sX
        .Top = GrhData(CurrentGrhIndex).sY
        .Right = .Left + GrhData(CurrentGrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrhIndex).pixelHeight
    End With
    
    'Draw
    Call BackBufferSurface.BltFast(x, y, SurfaceDB.Surface(GrhData(CurrentGrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT)
End Sub

Sub DDrawTransGrhIndextoSurface(ByVal GrhIndex As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte)
    Dim SourceRect As RECT
    
    'Center Grh over X,Y pos
    If Center Then
        If GrhData(GrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(GrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
        End If
        
        If GrhData(GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(GrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If
    
    With SourceRect
        .Left = GrhData(GrhIndex).sX
        .Top = GrhData(GrhIndex).sY
        .Right = .Left + GrhData(GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(GrhIndex).pixelHeight
    End With
    
    'Draw
    Call BackBufferSurface.BltFast(x, y, SurfaceDB.Surface(GrhData(GrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
End Sub

Sub DDrawTransGrhtoSurface(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Byte)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            
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
    
    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
        End If
        
        If GrhData(CurrentGrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If
    
    With SourceRect
        .Left = GrhData(CurrentGrhIndex).sX
        .Top = GrhData(CurrentGrhIndex).sY
        .Right = .Left + GrhData(CurrentGrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrhIndex).pixelHeight
    End With
    
    'Draw
    Call BackBufferSurface.BltFast(x, y, SurfaceDB.Surface(GrhData(CurrentGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
End Sub

#If ConAlfaB = 1 Then

Sub DDrawTransGrhtoSurfaceAlpha(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Byte)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.Speed)
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
    
    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
        End If
        If GrhData(CurrentGrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If
    
    With SourceRect
        .Left = GrhData(CurrentGrhIndex).sX
        .Top = GrhData(CurrentGrhIndex).sY
        .Right = .Left + GrhData(CurrentGrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrhIndex).pixelHeight
    End With
    
    Dim Src As DirectDrawSurface7
    Dim rDest As RECT
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    
    Set Src = SurfaceDB.Surface(GrhData(CurrentGrhIndex).FileNum)
    
    Src.GetSurfaceDesc ddsdSrc
    Surface.GetSurfaceDesc ddsdDest
    
    With rDest
        .Left = x
        .Top = y
        .Right = x + GrhData(CurrentGrhIndex).pixelWidth
        .Bottom = y + GrhData(iGrhIndex).pixelHeight
        
        If .Right > ddsdDest.lWidth Then
            .Right = ddsdDest.lWidth
        End If
        If .Bottom > ddsdDest.lHeight Then
            .Bottom = ddsdDest.lHeight
        End If
    End With
    
    ' 0 -> 16 bits 555
    ' 1 -> 16 bits 565
    ' 2 -> 16 bits raro (Sin implementar)
    ' 3 -> 24 bits
    ' 4 -> 32 bits
    
    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0 Then
        Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
        Modo = 1
'TODO : Revisar las máscaras de 24 y 32 bits!!
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
        Modo = 3
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &HFF00 And ddsdSrc.ddpfPixelFormat.lGBitMask = &HFF00 Then
        Modo = 4
    Else
        'Modo = 2 '16 bits raro ?
        Call BackBufferSurface.BltFast(x, y, Src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        Exit Sub
    End If
    
    Dim SrcLock As Boolean
    Dim DstLock As Boolean
    
    SrcLock = False
    DstLock = False
    
On Local Error GoTo HayErrorAlpha
    
    Call Src.Lock(SourceRect, ddsdSrc, DDLOCK_WAIT, 0)
    SrcLock = True
    Call BackBufferSurface.Lock(rDest, ddsdDest, DDLOCK_WAIT, 0)
    DstLock = True
    
    Call BackBufferSurface.GetLockedArray(dArray())
    Call Src.GetLockedArray(sArray())
    
    Call BltAlphaFast(ByVal VarPtr(dArray(x + x, y)), ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, Modo)
    
    Surface.Unlock rDest
    DstLock = False
    Src.Unlock SourceRect
    SrcLock = False
Exit Sub

HayErrorAlpha:
    If SrcLock Then Src.Unlock SourceRect
    If DstLock Then Surface.Unlock rDest
End Sub

#End If 'ConAlfaB = 1

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

Sub DrawGrhtoHdc(ByVal hdc As Long, ByVal GrhIndex As Integer, ByRef SourceRect As RECT, ByRef destRect As RECT)
'*****************************************************************
'Draws a Grh's portion to the given area of any Device Context
'*****************************************************************
    Call SurfaceDB.Surface(GrhData(GrhIndex).FileNum).BltToDC(hdc, SourceRect, destRect)
End Sub

Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim y           As Long     'Keeps track of where on map we are
    Dim x           As Long     'Keeps track of where on map we are
    Dim screenminY  As Integer  'Start Y pos on current screen
    Dim screenmaxY  As Integer  'End Y pos on current screen
    Dim screenminX  As Integer  'Start X pos on current screen
    Dim screenmaxX  As Integer  'End X pos on current screen
    Dim minY        As Integer  'Start Y pos on current map
    Dim maxY        As Integer  'End Y pos on current map
    Dim minX        As Integer  'Start X pos on current map
    Dim maxX        As Integer  'End X pos on current map
    Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
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
    
    
    'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            
            'Layer 1 **********************************
            Call DDrawGrhtoSurface(MapData(x, y).Graphic(1), _
                (ScreenX - 1) * TilePixelWidth + PixelOffsetX + TileBufferPixelOffsetX, _
                (ScreenY - 1) * TilePixelHeight + PixelOffsetY + TileBufferPixelOffsetY, _
                0, 1)
            '******************************************
            
            ScreenX = ScreenX + 1
        Next x
        
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next y
    
    'Draw floor layer 2
    ScreenY = minYOffset
    For y = minY To maxY
        ScreenX = minXOffset
        For x = minX To maxX
            
            'Layer 2 **********************************
            If MapData(x, y).Graphic(2).GrhIndex <> 0 Then
                Call DDrawTransGrhtoSurface(MapData(x, y).Graphic(2), _
                        (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                        (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                        1, 1)
            End If
            '******************************************
            
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    
    'Draw Transparent Layers
    ScreenY = minYOffset
    For y = minY To maxY
        ScreenX = minXOffset
        For x = minX To maxX
            PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY
            
            With MapData(x, y)
                'Object Layer **********************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(.ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If
                '***********************************************
                
                
                'Char layer ************************************
                If .CharIndex <> 0 Then
                    Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                End If
                '*************************************************
                
                
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    'Draw
                    Call DDrawTransGrhtoSurface(.Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If
                '************************************************
            End With
            
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    
    If Not bTecho Then
        'Draw blocked tiles and grid
        ScreenY = minYOffset
        For y = minY To maxY
            ScreenX = minXOffset
            For x = minX To maxX
                
                'Layer 4 **********************************
                If MapData(x, y).Graphic(4).GrhIndex Then
                    'Draw
                    Call DDrawTransGrhtoSurface(MapData(x, y).Graphic(4), _
                        (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                        (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                        1, 1)
                End If
                '**********************************
                
                ScreenX = ScreenX + 1
            Next x
            ScreenY = ScreenY + 1
        Next y
    End If
    
'TODO : Check this!!
    If bLluvia(UserMap) = 1 Then
        If bRain Then
            'Figure out what frame to draw
            If llTick < DirectX.TickCount - 50 Then
                iFrameIndex = iFrameIndex + 1
                If iFrameIndex > 7 Then iFrameIndex = 0
                llTick = DirectX.TickCount
            End If

            For y = 0 To 4
                For x = 0 To 4
                    Call BackBufferSurface.BltFast(LTLluvia(y), LTLluvia(x), SurfaceDB.Surface(5556), RLluvia(iFrameIndex), DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
                Next x
            Next y
        End If
    End If
End Sub

Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
    If bLluvia(UserMap) = 1 Then
        If bRain Then
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
    End If
    
    DoFogataFx
End Function

Function HayUserAbajo(ByVal x As Integer, ByVal y As Integer, ByVal GrhIndex As Integer) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.x >= x - (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.x <= x + (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.y >= y - (GrhData(GrhIndex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.y <= y
    End If
End Function

Sub LoadGraphics()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero - complete rewrite
'Last Modify Date: 11/03/2006
'Initializes the SurfaceDB and sets up the rain rects
'**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(DirectDraw, ClientSetup.bUseVideo, DirGraficos, ClientSetup.byMemory)
    
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128
    
    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
    
    'We are done!
    'Saco esto porque el texto del cargar queda horrible
    'AddtoRichTextBox frmCargando.status, "Hecho.", , , , 1, , False
End Sub

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single) As Boolean
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************
    Dim SurfaceDesc As DDSURFACEDESC2
    Dim ddck As DDCOLORKEY
    
    IniPath = App.path & "\Init\"
    
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = setWindowTileHeight \ 2
    HalfWindowTileWidth = setWindowTileWidth \ 2
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = engineSpeed
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.x = MinXBorder
    UserPos.y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    'Set the view rect
    With MainViewRect
        .Left = MainViewLeft
        .Top = MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    'Set the dest rect
    With MainDestRect
        .Left = TilePixelWidth * TileBufferSize - TilePixelWidth
        .Top = TilePixelHeight * TileBufferSize - TilePixelHeight
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
On Error Resume Next
    Set DirectX = New DirectX7
    
    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If

    
    '****** INIT DirectDraw ******
    ' Create the root DirectDraw object
    Set DirectDraw = DirectX.DirectDrawCreate("")
    
    If Err Then
        MsgBox "No se puede iniciar DirectDraw. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If
    
On Error GoTo 0
    Call DirectDraw.SetCooperativeLevel(setDisplayFormhWnd, DDSCL_NORMAL)
    
    'Primary Surface
    ' Fill the surface description structure
    With SurfaceDesc
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
    ' Create the surface
    Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)
    
    'Create Primary Clipper
    Set PrimaryClipper = DirectDraw.CreateClipper(0)
    Call PrimaryClipper.SetHWnd(frmMain.hWnd)
    Call PrimarySurface.SetClipper(PrimaryClipper)
    
    With BackBufferRect
        .Left = 0
        .Top = 0
        .Right = TilePixelWidth * (WindowTileWidth + 2 * TileBufferSize)
        .Bottom = TilePixelHeight * (WindowTileHeight + 2 * TileBufferSize)
    End With
    
    With SurfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        If ClientSetup.bUseVideo Then
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
        Else
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        End If
        .lHeight = BackBufferRect.Bottom
        .lWidth = BackBufferRect.Right
    End With
    
    ' Create surface
    Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)
    
    'Set color key
    ddck.low = 0
    ddck.high = 0
    Call BackBufferSurface.SetColorKey(DDCKEY_SRCBLT, ddck)
    
    'Set font transparency
    Call BackBufferSurface.SetFontTransparency(D_TRUE)
    
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    
    
    LTLluvia(0) = 224
    LTLluvia(1) = 352
    LTLluvia(2) = 480
    LTLluvia(3) = 608
    LTLluvia(4) = 736
    
    'Saco esto porque el texto del cargar queda horrible
    'AddtoRichTextBox frmCargando.status, "Cargando Gráficos....", 0, 0, 0, , , True
    
    Call LoadGraphics
    
    InitTileEngine = True
End Function

Public Sub DeinitTileEngine()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Destroys all DX objects
'***************************************************
On Error Resume Next
    Set PrimarySurface = Nothing
    Set PrimaryClipper = Nothing
    Set BackBufferSurface = Nothing
    
    Set DirectDraw = Nothing
    
    Set DirectX = Nothing
End Sub

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
'***************************************************
'Author: Arron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Updates the game's model and renders everything.
'***************************************************
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
    '****** Set main view rectangle ******
    With MainViewRect
        .Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
        .Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    If EngineRun Then
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.x <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.x * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.x) Then
                    OffsetCounterX = 0
                    AddtoUserPos.x = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.y = 0
                    UserMoving = False
                End If
            End If
        End If
        
        'Update mouse position within view area
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
        
        '****** Update screen ******
        If UserCiego Then
            Call CleanViewPort
        Else
            Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
        End If
        
        If IScombate Then Call RenderText(260, 260, "MODO COMBATE", vbRed, frmMain.font)
        
        Call Dialogos.Render
        Call DibujarCartel
        
        Call DialogosClanes.Draw
        
        'Display front-buffer!
        Call PrimarySurface.Blt(MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT)
        
        'Limit FPS to 100 (an easy number higher than monitor's vertical refresh rates)
        While (DirectX.TickCount - fpsLastCheck) \ 10 < FramesPerSecCounter
            Sleep 5
        Wend
        
        'FPS update
        If fpsLastCheck + 1000 < DirectX.TickCount Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            fpsLastCheck = DirectX.TickCount
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
        
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    End If
End Sub

#If ConAlfaB Then

Public Sub EfectoNoche(ByRef Surface As DirectDrawSurface7)
    Dim dArray() As Byte
    Dim ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    Dim rRect As RECT
    
    Surface.GetSurfaceDesc ddsdDest
    
    With rRect
        .Left = 0
        .Top = 0
        .Right = ddsdDest.lWidth
        .Bottom = ddsdDest.lHeight
    End With
    
    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
        Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
        Modo = 1
    Else
        Modo = 2
    End If
    
    Dim DstLock As Boolean
    DstLock = False
    
    On Local Error GoTo HayErrorAlpha
    
    Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
    DstLock = True
    
    Surface.GetLockedArray dArray()
    Call BltEfectoNoche(ByVal VarPtr(dArray(0, 0)), _
        ddsdDest.lWidth, ddsdDest.lHeight, ddsdDest.lPitch, _
        Modo)
    
HayErrorAlpha:
    If DstLock = True Then
        Surface.Unlock rRect
        DstLock = False
    End If
End Sub

#End If

Public Sub RenderText(ByVal lngXPos As Integer, ByVal lngYPos As Integer, ByRef strText As String, ByVal lngColor As Long, ByRef font As StdFont)
   If strText <> "" Then
        Call BackBufferSurface.SetForeColor(vbBlack)
        Call BackBufferSurface.SetFont(font)
        
        Call BackBufferSurface.DrawText(lngXPos - 2, lngYPos - 1, strText, False)
        
        
        Call BackBufferSurface.SetForeColor(lngColor)
        Call BackBufferSurface.SetFont(font)
        
        Call BackBufferSurface.DrawText(lngXPos, lngYPos, strText, False)
   End If
End Sub

Public Sub RenderTextBig(ByVal lngXPos As Integer, ByVal lngYPos As Integer, ByRef strText As String, ByVal lngColor As Long, ByRef backFont As StdFont, ByRef frontFont As StdFont)
    If strText <> "" Then
        Call BackBufferSurface.SetForeColor(vbBlack)
        Call BackBufferSurface.SetFont(backFont)
        Call BackBufferSurface.DrawText(lngXPos - 2, lngYPos - 1, strText, False)
        
        Call BackBufferSurface.SetForeColor(lngColor)
        Call BackBufferSurface.SetFont(frontFont)
        Call BackBufferSurface.DrawText(lngXPos, lngYPos, strText, False)
    End If
End Sub

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
    
    With charlist(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
                .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
                .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            .Arma.WeaponWalk(.Heading).Started = 0
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
            
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            
            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        If Not .invisible Then
            'Draw Body
            Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
            
            'Draw Head
            If .Head.Head(.Heading).GrhIndex Then
                Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0)
                
                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0)
                
                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
                
                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
            End If
            
            'Draw name over head
            If LenB(.Nombre) > 0 Then
                If Nombres And Abs(MouseTileX - .Pos.x) < 2 And (Abs(MouseTileY - .Pos.y)) < 2 Then
                    Pos = InStr(.Nombre, "<")
                    If Pos = 0 Then Pos = Len(.Nombre) + 2
                    
                    If .priv = 0 Then
                        If .Criminal Then
                            color = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                        Else
                            color = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                        End If
                    Else
                        color = RGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                    End If
                    
                    'Nick
                    line = Left$(.Nombre, Pos - 2)
                    Call RenderText(PixelOffsetX + 4 - Len(line) * 2, PixelOffsetY + 30, line, color, frmMain.font)
                    
                    'Clan
                    line = mid$(.Nombre, Pos)
                    Call RenderText(PixelOffsetX + 15 - Len(line) * 3, PixelOffsetY + 45, line, color, frmMain.font)
                End If
            End If
        End If
        
        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, CharIndex)
        
        'Draw FX
        If .FxIndex <> 0 Then
#If (ConAlfaB = 1) Then
            Call DDrawTransGrhtoSurfaceAlpha(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1)
#Else
            Call DDrawTransGrhtoSurface(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1)
#End If
            
            'Check if animation is over
            If .fX.Started = 0 Then _
                .FxIndex = 0
        End If
    End With
End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
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

Private Sub CleanViewPort()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Fills the viewport with black.
'***************************************************
    Dim r As RECT
    Call BackBufferSurface.BltColorFill(r, vbBlack)
End Sub
