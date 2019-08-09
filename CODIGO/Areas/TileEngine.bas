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

Private OffsetCounterX As Single
Private OffsetCounterY As Single
    
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

'Contiene info acerca de donde se puede encontrar un grh tamano y animacion
Public Type GrhData
    SX As Integer
    SY As Integer
    
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

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer

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

Public Normal_RGBList(0 To 3) As Long

'   Control de Lluvia
Public bRain As Boolean
Public bTecho       As Boolean 'hay techo?
Public bFogata       As Boolean

Public charlist(1 To 10000) As Char

' Used by GetTextExtentPoint32
Private Type size
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

Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

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
    ReDim FxData(0 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
        'MsgBox FxData(i).Animacion & FxData(i).OffsetX
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
        .Pos.x = x
        .Pos.y = y
        
        'attack state
        .attacking = False
        
        'Make active
        .active = 1
    End With
    
    'Plot on map
    MapData(x, y).CharIndex = CharIndex
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
    Dim addx As Integer
    Dim addy As Integer
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
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
        End Select
        
        nX = x + addx
        nY = y + addy
        
        If nX <= 0 Then nX = 1
        If nY <= 0 Then nY = 1
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.x = nX
        .Pos.y = nY
        
        If (x Or y) = 0 Then Exit Sub
        MapData(x, y).CharIndex = 0
        
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
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", Location.x, Location.y, LoopStyle.Enabled)
    End If
End Sub

Public Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 09/21/2010
' 09/21/2010: C4b3z0n - Changed from Private Funtion tu Public Function.
'***************************************************
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
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        x = .Pos.x
        y = .Pos.y
        
        MapData(x, y).CharIndex = 0
        
        addx = nX - x
        addy = nY - y
        
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
        
        .Pos.x = nX
        .Pos.y = nY
        
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

Private Function HayFogata(ByRef Location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.x - 8 To UserPos.x + 8
        For k = UserPos.y - 6 To UserPos.y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    Location.x = j
                    Location.y = k
                    
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
    Dim LoopC As Long
    Dim Dale As Boolean
    
    LoopC = 1
    Do While charlist(LoopC).active And Dale
        LoopC = LoopC + 1
        Dale = (LoopC <= UBound(charlist))
    Loop
    
    NextOpenChar = LoopC
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
                
                Get handle, , .Speed
                
                If .Speed <= 0 Then Resume Next
                
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
                
                Get handle, , GrhData(Grh).SX
                If .SX < 0 Then Resume Next
                
                Get handle, , .SY
                If .SY < 0 Then Resume Next
                
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

errorHandler:
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
    
    'Hay un personaje?
    If MapData(x, y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> HayAgua(x, y) Then
        Exit Function
    End If
    
    LegalPos = True
End Function

Function MoveToLegalPos(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 01/08/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
'*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(x, y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(x, y).CharIndex
    'Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.x, UserPos.y).Blocked = 1 Then
            Exit Function
        End If
        
        With charlist(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> eCabezas.CASPER_HEAD And .iBody <> eCabezas.FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.x, UserPos.y) Then
                    If Not HayAgua(x, y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(x, y) Then Exit Function
                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 6 Then
                    If charlist(UserCharIndex).invisible = True Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> HayAgua(x, y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
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
    Dim color As Long
    Dim x As Long
    Dim y As Long
    
    For x = SourceRect.Left To SourceRect.Right
        For y = SourceRect.Top To SourceRect.Bottom
            color = GetPixel(srchdc, x, y)
            
            If color <> TransparentColor Then
                Call SetPixel(dsthdc, destRect.Left + (x - SourceRect.Left), destRect.Top + (y - SourceRect.Top), color)
            End If
        Next y
    Next x
End Sub
Public Sub DrawImageInPicture(ByRef PictureBox As PictureBox, ByRef Picture As StdPicture, ByVal X1 As Single, ByVal Y1 As Single, Optional Width1, Optional Height1, Optional X2, Optional Y2, Optional Width2, Optional Height2)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 12/28/2009
'Draw Picture in the PictureBox
'*************************************************************

Call PictureBox.PaintPicture(Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2)
End Sub


Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
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
    Dim ElapsedTime As Single
    
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

    '<----- Layer 1, 2 ----->
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
        
            If Map_InBounds(x, y) Then
                'Layer 1
                Call DDrawGrhtoSurface(MapData(x, y).Graphic(1), _
                    (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                    (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                    0, 1, x, y)
            End If
            
            ScreenX = ScreenX + 1
        Next x
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next y
    
    '<----- Layer 2 ----->
    ScreenY = minYOffset - Engine_Get_TileBuffer
    For y = minY To maxY
        ScreenX = minXOffset - Engine_Get_TileBuffer
        For x = minX To maxX
        If Map_InBounds(x, y) Then
            PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
            'Layer 2
            If MapData(x, y).Graphic(2).GrhIndex <> 0 Then
                Call DDrawTransGrhtoSurface(MapData(x, y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(x, y).Engine_Light(), 1, x, y)
            End If
        End If
            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y
    
    '<----- Layer Obj, Char, 3 ----->
    ScreenY = minYOffset - Engine_Get_TileBuffer
    For y = minY To maxY
        ScreenX = minXOffset - Engine_Get_TileBuffer
        For x = minX To maxX
            PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
            
            If Map_InBounds(x, y) Then
                With MapData(x, y)
                    'Object Layer
                    If .ObjGrh.GrhIndex <> 0 Then
                        Call DDrawTransGrhtoSurface(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(x, y).Engine_Light(), 1, x, y)
                    End If
                    
                    'Char layer
                    If .CharIndex <> 0 Then
                        Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                    End If
                    
                    'Dibujamos los daños.
                    If .Damage.Activated Then
                        mDx8_Dibujado.Damage_Draw x, y, PixelOffsetXTemp, PixelOffsetYTemp - 20
                    End If
                    
                    'Layer 3
                    If .Graphic(3).GrhIndex <> 0 Then
                    
                        If .Graphic(3).GrhIndex = 735 Or .Graphic(3).GrhIndex >= 6994 And .Graphic(3).GrhIndex <= 7002 Then
                            If Abs(UserPos.x - x) < 4 And (Abs(UserPos.y - y)) < 4 Then
                                Call DDrawTransGrhtoSurface(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, SetARGB_Alpha(MapData(x, y).Engine_Light(), 150), 1, x, y, True)
                            Else 'NORMAL
                                Call DDrawTransGrhtoSurface(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(x, y).Engine_Light(), 1, x, y)
    
                            End If
                        Else 'NORMAL
                            Call DDrawTransGrhtoSurface(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(x, y).Engine_Light(), 1, x, y)
                        End If
                    End If
                    
                    If MapData(x, y).FxIndex <> 0 Then
                        
                        Call DDrawTransGrhtoSurface(MapData(x, y).fX, PixelOffsetXTemp + FxData(.FxIndex).OffsetX, PixelOffsetYTemp + FxData(.FxIndex).OffsetY, 1, MapData(x, y).Engine_Light(), 1, x, y, True)
                        
                        If MapData(x, y).fX.Started = 0 Then _
                            MapData(x, y).FxIndex = 0
                    End If
                End With
            End If
            
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    
        '<----- Layer 4 ----->
        ScreenY = minYOffset - Engine_Get_TileBuffer
        For y = minY To maxY
            ScreenX = minXOffset - Engine_Get_TileBuffer
            For x = minX To maxX
                
                'Layer 4
                If MapData(x, y).Graphic(4).GrhIndex Then
                    If bTecho Then
                    
                        Call DDrawTransGrhtoSurface(MapData(x, y).Graphic(4), _
                            ScreenX * TilePixelWidth + PixelOffsetX, _
                            ScreenY * TilePixelHeight + PixelOffsetY, _
                            1, temp_rgb(), 1, x, y)
                    
                    Else
                        If ColorTecho = 250 Then
                            Call DDrawTransGrhtoSurface(MapData(x, y).Graphic(4), _
                                ScreenX * TilePixelWidth + PixelOffsetX, _
                                ScreenY * TilePixelHeight + PixelOffsetY, _
                                1, MapData(x, y).Engine_Light(), 1, x, y)
                        Else
                            Call DDrawTransGrhtoSurface(MapData(x, y).Graphic(4), _
                                ScreenX * TilePixelWidth + PixelOffsetX, _
                                ScreenY * TilePixelHeight + PixelOffsetY, _
                                1, temp_rgb(), 1, x, y)
                        End If
                    End If
                End If
                ScreenX = ScreenX + 1
            Next x
            ScreenY = ScreenY + 1
        Next y
 

    'Weather Update & Render
    Call Engine_Weather_Update
    
    'Effects Update
    Call Effect_UpdateAll
    
    If ClientSetup.ProyectileEngine Then
                            
        If LastProjectile > 0 Then
            Dim j As Long ' Long siempre en los bucles es mucho mas rapido
                                
            For j = 1 To LastProjectile
                If ProjectileList(j).Grh.GrhIndex Then
                    Dim Angle As Single
                    'Update the position
                    Angle = DegreeToRadian * Engine_GetAngle(ProjectileList(j).x, ProjectileList(j).y, ProjectileList(j).tX, ProjectileList(j).tY)
                    ProjectileList(j).x = ProjectileList(j).x + (Sin(Angle) * ElapsedTime * 0.63)
                    ProjectileList(j).y = ProjectileList(j).y - (Cos(Angle) * ElapsedTime * 0.63)
                    
                    'Update the rotation
                    If ProjectileList(j).RotateSpeed > 0 Then
                        ProjectileList(j).Rotate = ProjectileList(j).Rotate + (ProjectileList(j).RotateSpeed * ElapsedTime * 0.01)
                        Do While ProjectileList(j).Rotate > 360
                            ProjectileList(j).Rotate = ProjectileList(j).Rotate - 360
                        Loop
                    End If
    
                    'Draw if within range
                    x = ((-minX - 1) * 32) + ProjectileList(j).x + PixelOffsetX + ((10 - TileBufferSize) * 32) - 288 + ProjectileList(j).OffsetX
                    y = ((-minY - 1) * 32) + ProjectileList(j).y + PixelOffsetY + ((10 - TileBufferSize) * 32) - 288 + ProjectileList(j).OffsetY
                    If y >= -32 Then
                        If y <= (ScreenHeight + 32) Then
                            If x >= -32 Then
                                If x <= (ScreenWidth + 32) Then
                                    If ProjectileList(j).Rotate = 0 Then
                                        DDrawTransGrhtoSurface ProjectileList(j).Grh, x, y, 0, MapData(50, 50).Engine_Light(), 0, 50, 50, True, 0
                                    Else
                                        DDrawTransGrhtoSurface ProjectileList(j).Grh, x, y, 0, MapData(50, 50).Engine_Light(), 0, 50, 50, True, ProjectileList(j).Rotate
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                End If
            Next j
            
            'Check if it is close enough to the target to remove
            For j = 1 To LastProjectile
                If ProjectileList(j).Grh.GrhIndex Then
                    If Abs(ProjectileList(j).x - ProjectileList(j).tX) < 20 Then
                        If Abs(ProjectileList(j).y - ProjectileList(j).tY) < 20 Then
                            Engine_Projectile_Erase j
                        End If
                    End If
                End If
            Next j
            
        End If
    End If
    
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
                        FogataBufferIndex = Audio.PlayWave("fuego.wav", Location.x, Location.y, LoopStyle.Enabled)
                End If
        End If
End Function

Function HayUserAbajo(ByVal x As Integer, ByVal y As Integer, ByVal GrhIndex As Long) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.x >= x - (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.x <= x + (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.y >= y - (GrhData(GrhIndex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.y <= y
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
    
    IniPath = App.path & "\Init\"
    HalfWindowTileHeight = WindowTileHeight \ 2
    HalfWindowTileWidth = WindowTileWidth \ 2

    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    

    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.x = MinXBorder
    UserPos.y = MinYBorder
    
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

    InitTileEngine = True
End Function
Public Sub LoadGraphics()
    Call SurfaceDB.Initialize(DirectD3D8, App.path & "\graficos\", ClientSetup.byMemory)
End Sub


Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)

    If EngineRun Then
        Call Engine_BeginScene
        
        Call DesvanecimientoTechos
        
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
            DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        Else
            Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
        End If
        
        Call Dialogos.Render
        Call DibujarCartel
        
        Call DialogosClanes.Draw
      
        ' Calculamos los FPS y los mostramos
        Call Engine_Update_FPS

        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * Engine_Get_BaseSpeed
        
        Call Engine_EndScene(MainScreenRect, 0)
    End If
    
    
          '//Banco
    If frmBancoObj.PicBancoInv.Visible Then _
        Call InvBanco(0).DrawInv
         
    If frmBancoObj.PicInv.Visible Then _
        Call InvBanco(1).DrawInv
    
    
    '//Comercio
    If frmComerciar.picInvNpc.Visible Then _
        Call InvComNpc.DrawInv
        
    If frmComerciar.picInvUser.Visible Then _
        Call InvComUsu.DrawInv
        
    
    '//Comercio entre usuarios
    If frmComerciarUsu.picInvComercio.Visible Then _
        InvComUsu.DrawInv (1)
            
    If frmComerciarUsu.picInvOfertaProp.Visible Then _
        InvOfferComUsu(0).DrawInv (1)
            
    If frmComerciarUsu.picInvOfertaOtro Then _
        InvOfferComUsu(1).DrawInv (1)
            
    If frmComerciarUsu.picInvOroProp.Visible Then _
        InvOroComUsu(0).DrawInv (1)
            
    If frmComerciarUsu.picInvOroOfertaProp.Visible Then _
        InvOroComUsu(1).DrawInv (1)
                
    If frmComerciarUsu.picInvOroOfertaOtro.Visible Then _
        InvOroComUsu(2).DrawInv (1)
        
        
        
    '//Herrero
    If frmHerrero.Visible Then
        With frmHerrero
        If .picLingotes0.Visible Or .picMejorar0.Visible Then _
            InvLingosHerreria(1).DrawInv (1)
            
        If .picLingotes1.Visible Or .picMejorar1.Visible Then _
            InvLingosHerreria(2).DrawInv (1)
            
        If .picLingotes2.Visible Or .picMejorar2.Visible Then _
            InvLingosHerreria(3).DrawInv (1)
            
        If .picLingotes3.Visible Or .picMejorar3.Visible Then _
            InvLingosHerreria(4).DrawInv (1)
        End With
    End If
        
    '//FIN HERRERO
    
    '//Carpintero
    If frmCarp.Visible Then
        With frmCarp
            If .picMaderas0.Visible Or .imgMejorar0.Visible Then _
                InvMaderasCarpinteria(1).DrawInv (1)
                
            If .picMaderas1.Visible Or .imgMejorar1.Visible Then _
            InvMaderasCarpinteria(2).DrawInv (1)
            
            If .picMaderas2.Visible Or .imgMejorar2.Visible Then _
            InvMaderasCarpinteria(3).DrawInv (1)
            
            If .picMaderas3.Visible Or .imgMejorar3.Visible Then _
            InvMaderasCarpinteria(4).DrawInv (1)
        End With
    End If

    '//Inventario
    If frmMain.Visible Then Call Inventario.DrawInv

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

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
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
                If .Body.Walk(.Heading).Speed > 0 Then _
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
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
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
        Dim RenderSpell As Boolean
        
        If Not .muerto Then
            If Abs(MouseTileX - .Pos.x) < 1 And (Abs(MouseTileY - .Pos.y)) < 1 And CharIndex <> UserCharIndex And ClientSetup.TonalidadPJ Then
                If .Nombre <> "" Then
                    If .Criminal Then
                        Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorXRGB(204, 100, 100))
                    Else
                        Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorXRGB(100, 100, 255))
                    End If
                    
                Else
                    ColorFinal(0) = MapData(.Pos.x, .Pos.y).Engine_Light(0)
                    ColorFinal(1) = MapData(.Pos.x, .Pos.y).Engine_Light(1)
                    ColorFinal(2) = MapData(.Pos.x, .Pos.y).Engine_Light(2)
                    ColorFinal(3) = MapData(.Pos.x, .Pos.y).Engine_Light(3)
                End If
                RenderSpell = True
            Else
                ColorFinal(0) = MapData(.Pos.x, .Pos.y).Engine_Light(0)
                ColorFinal(1) = MapData(.Pos.x, .Pos.y).Engine_Light(1)
                ColorFinal(2) = MapData(.Pos.x, .Pos.y).Engine_Light(2)
                ColorFinal(3) = MapData(.Pos.x, .Pos.y).Engine_Light(3)
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
                Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, .Pos.x, .Pos.y, False, 0)
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
                Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, ColorFinal(), 0, .Pos.x, .Pos.y)
                
                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y + OFFSET_HEAD, 1, ColorFinal(), 0, .Pos.x, .Pos.y)
                
                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then
                    Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, .Pos.x, .Pos.y)
                End If
                
                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then
                    Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, .Pos.x, .Pos.y)
                End If
            
                'Draw name over head
                If LenB(.Nombre) > 0 Then
                    If Nombres Then
                        Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY)
                    End If
            End If
            
        Else 'Usuario invisible
        
            If CharIndex = UserCharIndex Or mid$(charlist(CharIndex).Nombre, _
                getTagPosition(.Nombre)) = mid$(charlist(UserCharIndex).Nombre, getTagPosition(charlist(UserCharIndex).Nombre)) And _
                    Len(mid$(charlist(CharIndex).Nombre, getTagPosition(.Nombre))) > 0 Then
                
                Movement_Speed = 0.5
                
                'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, .Pos.x, .Pos.y, True, 0)
                
                'Draw Head
                If .Head.Head(.Heading).GrhIndex Then
                    Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, ColorFinal(), 0, .Pos.x, .Pos.y, True)
                    
                    'Draw Helmet
                    If .Casco.Head(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y + OFFSET_HEAD, 1, ColorFinal(), 0, .Pos.x, .Pos.y, True)
                    
                    'Draw Weapon
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, .Pos.x, .Pos.y, True)
                    
                    'Draw Shield
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, .Pos.x, .Pos.y, True)
                
                
                    'Draw name over head
                    If LenB(.Nombre) > 0 Then
                        If Nombres Then
                             Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY, True)
                        End If
                    End If
                    
                    'OffSetName = 35

                End If
            End If
        End If
        End If
        
        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y + OFFSET_HEAD, CharIndex) '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo
        
        Movement_Speed = 1
        'Draw FX
        If .FxIndex <> 0 Then
            Call DDrawTransGrhtoSurface(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, SetARGB_Alpha(MapData(.Pos.x, .Pos.y).Engine_Light(), 180), 1, .Pos.x, .Pos.y, True)
            
            'Check if animation is over
            If .fX.Started = 0 Then _
                .FxIndex = 0
        End If
        
    End With
End Sub
Private Sub RenderName(ByVal CharIndex As Long, ByVal x As Integer, ByVal y As Integer, Optional ByVal Invi As Boolean = False)
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
   
    With charlist(CharIndex)
            Pos = getTagPosition(.Nombre)
    
            If .priv = 0 Then
                    If .muerto Then
                        color = D3DColorARGB(255, 220, 220, 255)
                    Else
                        If .Criminal Then
                            color = ColoresPJ(50)
                        Else
                            color = ColoresPJ(49)
                        End If
                    End If
            Else
                color = ColoresPJ(.priv)
            End If
    
            If Invi Then
                color = D3DColorARGB(180, 150, 180, 220)
            End If

            'Nick
            line = Left$(.Nombre, Pos - 2)
            Call DrawText(x + 16, y + 30, line, color, True)
            
            'Clan
            line = mid$(.Nombre, Pos)
            Call DrawText(x + 16, y + 45, line, color, True)
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


Public Sub Device_Textured_Render(ByVal x As Integer, ByVal y As Integer, ByVal Texture As Direct3DTexture8, ByRef src_rect As RECT, ByRef Color_List() As Long, Optional Alpha As Boolean = False, Optional ByVal Angle As Single = 0, Optional ByVal Shadow As Boolean = False)
    If Shadow And ClientSetup.UsarSombras = False Then Exit Sub
    
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim SRDesc As D3DSURFACE_DESC

    With dest_rect
        .Bottom = y + (src_rect.Bottom - src_rect.Top)
        .Left = x
        .Right = x + (src_rect.Right - src_rect.Left)
        .Top = y
    End With
    
    Dim texwidth As Long, texheight As Long
    Texture.GetLevelDesc 0, SRDesc

    texwidth = SRDesc.Width
    texheight = SRDesc.Height
    
    If Shadow Then
        Dim Color_Shadow(3) As Long
        Engine_Long_To_RGB_List Color_Shadow(), D3DColorARGB(50, 0, 0, 0)
        Geometry_Create_Box temp_verts(), dest_rect, src_rect, Color_Shadow(), texwidth, texheight, Angle
    Else
        Geometry_Create_Box temp_verts(), dest_rect, src_rect, Color_List(), texwidth, texheight, Angle
    End If
    
    DirectDevice.SetTexture 0, Texture

    If Shadow Then
        temp_verts(1).x = temp_verts(1).x + (src_rect.Bottom - src_rect.Top) * 0.5
        temp_verts(1).y = temp_verts(1).y - (src_rect.Right - src_rect.Left) * 0.5
       
        temp_verts(3).x = temp_verts(3).x + (src_rect.Right - src_rect.Left)
        temp_verts(3).y = temp_verts(3).y - (src_rect.Right - src_rect.Left) * 0.5
    End If
    
    If Alpha Then
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
    
    ' Medium load.
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))

    If Alpha Then
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub

Public Sub Device_Textured_Render_Scale(ByVal x As Integer, ByVal y As Integer, ByVal Texture As Direct3DTexture8, ByRef src_rect As RECT, ByRef Color_List() As Long, Optional Alpha As Boolean = False, Optional ByVal Angle As Single = 0, Optional ByVal Shadow As Boolean = False)
    If Shadow And ClientSetup.UsarSombras = False Then Exit Sub
    
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim SRDesc As D3DSURFACE_DESC

    With dest_rect
        .Bottom = y + 2 '(src_rect.bottom - src_rect.Top)
        .Left = x
        .Right = x + 2 '(src_rect.Right - src_rect.Left)
        .Top = y
    End With
    
    Dim texwidth As Long, texheight As Long
    Texture.GetLevelDesc 0, SRDesc

    texwidth = SRDesc.Width
    texheight = SRDesc.Height
    
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, Color_List(), texwidth, texheight, Angle
    
    DirectDevice.SetTexture 0, Texture
    
    If Alpha Then
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
        
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
        
    If Alpha Then
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
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

    Call DDrawTransGrhIndextoSurface(GrhIndex, 0, 0, 0, Normal_RGBList(), 0, False)
        
    Call Engine_EndScene(DR, hWndDest)
    
End Sub

Private Sub DDrawGrhtoSurface(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Byte, ByVal posX As Byte, ByVal posY As Byte)
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    
    If Grh.GrhIndex = 0 Then Exit Sub
On Error GoTo error
        
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
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render(x, y, SurfaceDB.Surface(.FileNum), SourceRect, MapData(posX, posY).Engine_Light(), False, 0, False)
    End With
Exit Sub

error:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        'Call Log_Engine("Error in DDrawGrhtoSurface, " & Err.Description & ", (" & Err.number & ")")
        MsgBox "Error en el Engine Grafico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
        Call CloseClient
    End If
End Sub

Sub DDrawTransGrhIndextoSurface(ByVal GrhIndex As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, Optional ByVal Angle As Single = 0, Optional ByVal Alpha As Boolean = False)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render(x, y, SurfaceDB.Surface(.FileNum), SourceRect, Color_List(), Alpha, Angle, False)
    End With
End Sub

Public Sub DDrawGrhtoSurfaceScale(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Byte, ByVal posX As Byte, ByVal posY As Byte)
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    
    If Grh.GrhIndex = 0 Then Exit Sub
On Error GoTo error
        
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
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render_Scale(x, y, SurfaceDB.Surface(.FileNum), SourceRect, MapData(posX, posY).Engine_Light(), False, 0, False)
    End With
Exit Sub

error:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        'Call Log_Engine("Error in DDrawGrhtoSurface, " & Err.Description & ", (" & Err.number & ")")
        MsgBox "Error en el Engine Grafico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
        Call CloseClient
    End If
End Sub


Sub DDrawTransGrhIndextoSurfaceScale(ByVal GrhIndex As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, Optional ByVal Angle As Single = 0, Optional ByVal Alpha As Boolean = False)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render_Scale(x, y, SurfaceDB.Surface(.FileNum), SourceRect, Color_List(), Alpha, Angle, False)
    End With
End Sub

Sub DDrawTransGrhtoSurface(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, ByVal Animate As Byte, ByVal posX As Byte, ByVal posY As Byte, Optional ByVal Alpha As Boolean = False, Optional ByVal Angle As Single = 0, Optional ByVal Shadow As Boolean = False)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    
    If Grh.GrhIndex = 0 Then Exit Sub
    
On Error GoTo error
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * Movement_Speed
            
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
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        Call Device_Textured_Render(x, y, SurfaceDB.Surface(.FileNum), SourceRect, Color_List(), Alpha, Angle, False)
    End With
Exit Sub

error:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        'Call Log_Engine("Error in DDrawTransGrhtoSurface, " & Err.Description & ", (" & Err.number & ")")
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
                .Speed = 0
                
        End With

End Sub

Public Sub DesvanecimientoTechos()
 
    If bTecho Then
        If Not Val(ColorTecho) = 150 Then ColorTecho = ColorTecho - 1
    Else

        If Not Val(ColorTecho) = 250 Then ColorTecho = ColorTecho + 1

    End If
 
    temp_rgb(0) = D3DColorARGB(ColorTecho, ColorTecho, ColorTecho, ColorTecho)
    temp_rgb(1) = D3DColorARGB(ColorTecho, ColorTecho, ColorTecho, ColorTecho)
    temp_rgb(2) = D3DColorARGB(ColorTecho, ColorTecho, ColorTecho, ColorTecho)
    temp_rgb(3) = D3DColorARGB(ColorTecho, ColorTecho, ColorTecho, ColorTecho)
 
End Sub
