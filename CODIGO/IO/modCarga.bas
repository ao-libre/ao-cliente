Attribute VB_Name = "Carga"
Option Explicit

'********************************
'Load Map with .CSM format
'********************************
Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
End Type

Private Type tDatosBloqueados
    X As Integer
    Y As Integer
End Type

Private Type tDatosGrh
    X As Integer
    Y As Integer
    GrhIndex As Long
End Type

Private Type tDatosTrigger
    X As Integer
    Y As Integer
    Trigger As Integer
End Type

Private Type tDatosLuces
    X As Integer
    Y As Integer
    light_value(3) As Long
    base_light(0 To 3) As Boolean 'Indica si el tile tiene luz propia.
End Type

Private Type tDatosParticulas
    X As Integer
    Y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    X As Integer
    Y As Integer
    NPCIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    Y As Integer
    objindex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSize

    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer

End Type

Private Type tMapDat
    map_name As String
    battle_mode As Boolean
    backup_mode As Boolean
    restrict_mode As String
    midi_number As String
    mp3_number As String
    zone As String
    terrain As String
    ambient As String
    lvlMinimo As String
    SePuedeDomar As Boolean
    ResuSinEfecto As Boolean
    MagiaSinEfecto As Boolean
    InviSinEfecto As Boolean
    NoEncriptarMP As Boolean
    version As Long
End Type

Private MapSize As tMapSize
Private MapDat As tMapDat
'********************************
'END - Load Map with .CSM format
'********************************

Private FileManager As clsIniManager

''
' Loads grh data using the Graficos.INI
'

Public Sub LoadGrhIni()
    On Error GoTo hErr

    Dim FileHandle     As Integer
    Dim Grh            As Long
    Dim Frame          As Long
    Dim SeparadorClave As String
    Dim SeparadorGrh   As String
    Dim CurrentLine    As String
    Dim Fields()       As String
    
    ' Guardo el separador en una variable asi no lo busco en cada bucle.
    SeparadorClave = "="
    SeparadorGrh = "-"
    
    ' Abrimos el archivo. No uso FileManager porque obliga a cargar todo el archivo en memoria
    ' y es demasiado grande. En cambio leo linea por linea y procesamos de a una.
    FileHandle = FreeFile()
    Open Game.path(INIT) & "Graficos.ini" For Input As FileHandle

    ' Leemos el total de Grhs
    Do While Not EOF(FileHandle)
        ' Leemos la linea actual
        Line Input #FileHandle, CurrentLine

        Fields = Split(CurrentLine, SeparadorClave)
            
        ' Buscamos la clave "NumGrh"
        If Fields(0) = "NumGrh" Then
            ' Asignamos el tamano al array de Grhs
            ReDim GrhData(1 To Val(Fields(1))) As GrhData
                
            Exit Do
        End If
    Loop
        
    ' Chequeamos si pudimos leer la cantidad de Grhs
    If UBound(GrhData) <= 0 Then GoTo hErr
        
    ' Buscamos la posicion del primer Grh
    Do While Not EOF(FileHandle)
        ' Leemos la linea actual
        Line Input #FileHandle, CurrentLine
            
        ' Buscamos el nodo "[Graphics]"
        If UCase$(CurrentLine) = "[GRAPHICS]" Then
            ' Ya lo tenemos, salimos
            Exit Do
        End If
    Loop
        
    ' Recorremos todos los Grhs
    Do While Not EOF(FileHandle)
        ' Leemos la linea actual
        Line Input #FileHandle, CurrentLine
            
        ' Ignoramos lineas vacias
        If CurrentLine <> vbNullString Then
            
            ' Divimos por el "="
            Fields = Split(CurrentLine, SeparadorClave)
                
            ' Leemos el numero de Grh (el numero a la derecha de la palabra "Grh")
            Grh = Right(Fields(0), Len(Fields(0)) - 3)
            
            ' Leemos los campos de datos del Grh
            Fields = Split(Fields(1), SeparadorGrh)
                
            With GrhData(Grh)
                    
                ' Primer lugar: cantidad de frames.
                .NumFrames = Val(Fields(0))
    
                ReDim .Frames(1 To .NumFrames)
                    
                ' Tiene mas de un frame entonces es una animacion
                If .NumFrames > 1 Then
                    
                    ' Segundo lugar: Leemos los numeros de grh de la animacion
                    For Frame = 1 To .NumFrames
                        .Frames(Frame) = Val(Fields(Frame))
                        If .Frames(Frame) <= LBound(GrhData) Or .Frames(Frame) > UBound(GrhData) Then GoTo hErr
                    Next
                        
                    ' Tercer lugar: leemos la velocidad de la animacion
                    .speed = Val(Fields(Frame))
                    If .speed <= 0 Then GoTo hErr
                        
                    ' Por ultimo, copiamos las dimensiones del primer frame
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo hErr
                        
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo hErr
                        
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo hErr
                        
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo hErr
        
                ElseIf .NumFrames = 1 Then
                    
                    ' Si es un solo frame lo asignamos a si mismo
                    .Frames(1) = Grh
                        
                    ' Segundo lugar: NumeroDelGrafico.bmp, pero sin el ".bmp"
                    .FileNum = Val(Fields(1))
                    If .FileNum <= 0 Then GoTo hErr
                            
                    ' Tercer Lugar: La coordenada X del grafico
                    .sX = Val(Fields(2))
                    If .sX < 0 Then GoTo hErr
                            
                    ' Cuarto Lugar: La coordenada Y del grafico
                    .sY = Val(Fields(3))
                    If .sY < 0 Then GoTo hErr
                            
                    ' Quinto lugar: El ancho del grafico
                    .pixelWidth = Val(Fields(4))
                    If .pixelWidth <= 0 Then GoTo hErr
                            
                    ' Sexto lugar: La altura del grafico
                    .pixelHeight = Val(Fields(5))
                    If .pixelHeight <= 0 Then GoTo hErr
                        
                    ' Calculamos el ancho y alto en tiles
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                        
                Else
                    ' 0 frames o negativo? Error
                    GoTo hErr
                End If
        
            End With
        End If
    Loop
    
hErr:
    Close FileHandle
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Graficos.ini no existe. Por favor, reinstale el juego.", , "Argentum Online")
        
        ElseIf Grh > 0 Then
            Call MsgBox("Hay un error en Graficos.ini con el Grh" & Grh & ".", , "Argentum Online")
        
        Else
            Call MsgBox("Hay un error en Graficos.ini. Por favor, reinstale el juego.", , "Argentum Online")
        End If
        
        Call CloseClient
        
    End If
    
    Exit Sub

End Sub

''
' Loads grh data using the new file format.
'

Public Sub LoadGrhInd()
On Error GoTo ErrorHandler:

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
            
                '.active = True
                Get handle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ReDim .Frames(1 To .NumFrames)
                
                If .NumFrames > 1 Then
                
                    For Frame = 1 To .NumFrames
                        Get handle, , .Frames(Frame)
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then GoTo ErrorHandler
                    Next Frame
                    
                    Get handle, , .speed
                    If .speed <= 0 Then GoTo ErrorHandler
                    
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                    
                Else
                    
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
                    
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = Grh
                    
                End If
                
            End With
            
        Wend
    
    Close handle
    
Exit Sub

ErrorHandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Graficos.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
        
    End If
    
End Sub

Public Sub CargarCabezas()
On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open Game.path(INIT) & "Cabezas.ind" For Binary Access Read As #N
    
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
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Cabezas.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
        
    End If
    
End Sub

Sub CargarCascos()
On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open Game.path(INIT) & "Cascos.ind" For Binary Access Read As #N
    
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
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Cascos.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
        
    End If
    
End Sub

Sub CargarCuerpos()
On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open Game.path(INIT) & "Personajes.ind" For Binary Access Read As #N
    
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
            Call InitGrh(BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0)
            Call InitGrh(BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0)
            Call InitGrh(BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0)
            Call InitGrh(BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0)
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Personajes.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
        
    End If
    
End Sub

Sub CargarFxs()
On Error GoTo errhandler:

    Dim i As Long
    
    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Game.path(INIT) & "Fxs.ini")
    
    'Resize array
    ReDim FxData(0 To FileManager.GetValue("INIT", "NumFxs")) As tIndiceFx
    
    For i = 1 To UBound(FxData())
        
        With FxData(i)
            .Animacion = Val(FileManager.GetValue("FX" & CStr(i), "Animacion"))
            .OffsetX = Val(FileManager.GetValue("FX" & CStr(i), "OffsetX"))
            .OffsetY = Val(FileManager.GetValue("FX" & CStr(i), "OffsetY"))
        End With
    
    Next
    
    Set FileManager = Nothing
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Fxs.ini no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
        
    End If

End Sub

Public Sub CargarTips()
'************************************************************************************.
' Carga el JSON con los tips del juego en un objeto para su uso a lo largo del proyecto
'************************************************************************************
On Error GoTo errhandler:
    
    Dim TipFile As String
        TipFile = FileToString(Game.path(INIT) & "tips_" & Language & ".json")
    
    Set JsonTips = JSON.parse(TipFile)

errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo" & "tips_" & Language & ".json no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
        
    End If
End Sub

Sub CargarArrayLluvia()
On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Long
    Dim Nu As Integer
    
    N = FreeFile()
    Open Game.path(INIT) & "fk.ind" For Binary Access Read As #N
    
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
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo fk.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
        
    End If
    
End Sub

Sub CargarAnimArmas()
On Error GoTo errhandler:

    Dim LoopC As Long

    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Game.path(INIT) & "armas.dat")
    
    NumWeaponAnims = Val(FileManager.GetValue("INIT", "NumArmas"))
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For LoopC = 1 To NumWeaponAnims
        Call InitGrh(WeaponAnimData(LoopC).WeaponWalk(1), Val(FileManager.GetValue("ARMA" & LoopC, "Dir1")), 0)
        Call InitGrh(WeaponAnimData(LoopC).WeaponWalk(2), Val(FileManager.GetValue("ARMA" & LoopC, "Dir2")), 0)
        Call InitGrh(WeaponAnimData(LoopC).WeaponWalk(3), Val(FileManager.GetValue("ARMA" & LoopC, "Dir3")), 0)
        Call InitGrh(WeaponAnimData(LoopC).WeaponWalk(4), Val(FileManager.GetValue("ARMA" & LoopC, "Dir4")), 0)
    Next LoopC
    
    Set FileManager = Nothing
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo armas.dat no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
        
    End If

End Sub


Public Sub CargarColores()
On Error GoTo errhandler:

    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Game.path(INIT) & "colores.dat")
    
    Dim i As Long
    
    For i = 0 To 47 '48, 49 y 50 reservados para atacables, ciudadano y criminal
        ColoresPJ(i) = D3DColorXRGB(FileManager.GetValue(CStr(i), "R"), FileManager.GetValue(CStr(i), "G"), FileManager.GetValue(CStr(i), "B"))
    Next i
    
    '   Crimi
    ColoresPJ(50) = D3DColorXRGB(FileManager.GetValue("CR", "R"), FileManager.GetValue("CR", "G"), FileManager.GetValue("CR", "B"))

    '   Ciuda
    ColoresPJ(49) = D3DColorXRGB(FileManager.GetValue("CI", "R"), FileManager.GetValue("CI", "G"), FileManager.GetValue("CI", "B"))
    
    '   Atacable TODO: hay que implementar un color para los atacables y hacer que funcione.
    'ColoresPJ(48) = D3DColorXRGB(FileManager.GetValue("AT", "R"), FileManager.GetValue("AT", "G"), FileManager.GetValue("AT", "B"))
    
    For i = 51 To 56 'Colores reservados para la renderizacion de dano
        ColoresDano(i) = D3DColorXRGB(FileManager.GetValue(CStr(i), "R"), FileManager.GetValue(CStr(i), "G"), FileManager.GetValue(CStr(i), "B"))
    Next i
    
    Set FileManager = Nothing
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo colores.dat no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
        
    End If
    
End Sub

Sub CargarAnimEscudos()
On Error GoTo errhandler:

    Dim LoopC As Long
    Dim NumEscudosAnims As Integer
    
    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Game.path(INIT) & "escudos.dat")
    
    NumEscudosAnims = Val(FileManager.GetValue("INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For LoopC = 1 To NumEscudosAnims
        Call InitGrh(ShieldAnimData(LoopC).ShieldWalk(1), Val(FileManager.GetValue("ESC" & LoopC, "Dir1")), 0)
        Call InitGrh(ShieldAnimData(LoopC).ShieldWalk(2), Val(FileManager.GetValue("ESC" & LoopC, "Dir2")), 0)
        Call InitGrh(ShieldAnimData(LoopC).ShieldWalk(3), Val(FileManager.GetValue("ESC" & LoopC, "Dir3")), 0)
        Call InitGrh(ShieldAnimData(LoopC).ShieldWalk(4), Val(FileManager.GetValue("ESC" & LoopC, "Dir4")), 0)
    Next LoopC
    
    Set FileManager = Nothing
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo escudos.dat no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
        
    End If
    
End Sub

Public Sub CargarHechizos()
'********************************
'Author: Shak
'Last Modification:
'Cargamos los hechizos del juego. [Solo datos necesarios]
'********************************
On Error GoTo errorH

    Dim J As Long
    
    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Game.path(INIT) & "Hechizos.dat")

    NumHechizos = Val(FileManager.GetValue("INIT", "NumHechizos"))
 
    ReDim Hechizos(1 To NumHechizos) As tHechizos
    
    For J = 1 To NumHechizos
        
        With Hechizos(J)
            .Desc = FileManager.GetValue("HECHIZO" & J, "Desc")
            .PalabrasMagicas = FileManager.GetValue("HECHIZO" & J, "PalabrasMagicas")
            .Nombre = FileManager.GetValue("HECHIZO" & J, "Nombre")
            .SkillRequerido = Val(FileManager.GetValue("HECHIZO" & J, "MinSkill"))
         
            If J <> 38 And J <> 39 Then
                
                .EnergiaRequerida = Val(FileManager.GetValue("HECHIZO" & J, "StaRequerido"))
                 
                .HechiceroMsg = FileManager.GetValue("HECHIZO" & J, "HechizeroMsg")
                .ManaRequerida = Val(FileManager.GetValue("HECHIZO" & J, "ManaRequerido"))
             
                .PropioMsg = FileManager.GetValue("HECHIZO" & J, "PropioMsg")
                .TargetMsg = FileManager.GetValue("HECHIZO" & J, "TargetMsg")
                
            End If
            
        End With
        
    Next J
    
    Set FileManager = Nothing
    
Exit Sub
 
errorH:

    If Err.number <> 0 Then
        
        Select Case Err.number
            
            Case 9
                Call MsgBox("Error cargando el archivo Hechizos.dat (Hechizo " & J & "). Por favor, avise a los administradores enviandoles el archivo Errores.log que se encuentra en la carpeta del cliente.", , "Argentum Online Libre")
                Call LogError(Err.number, Err.Description, "CargarHechizos")
            
            Case 53
                Call MsgBox("El archivo Hechizos.dat no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
        
        End Select
        
        Call CloseClient

    End If

End Sub

Sub CargarMapa(ByVal Map As Integer)

    On Error GoTo ErrorHandler

    Dim fh           As Integer
    Dim File         As Integer
    
    Dim MH           As tMapHeader
    Dim Blqs()       As tDatosBloqueados

    Dim L1()         As Long
    Dim L2()         As tDatosGrh
    Dim L3()         As tDatosGrh
    Dim L4()         As tDatosGrh

    Dim Triggers()   As tDatosTrigger
    Dim Luces()      As tDatosLuces
    Dim Particulas() As tDatosParticulas
    Dim Objetos()    As tDatosObjs
    Dim NPCs()       As tDatosNPC
    Dim TEs()        As tDatosTE

    Dim i            As Long
    Dim J            As Long

    DoEvents

    fh = FreeFile
    Open Game.path(Mapas) & "Mapa" & CStr(Map) & ".csm" For Binary Access Read As fh
    Get #fh, , MH
    Get #fh, , MapSize
    
    '¿Queremos cargar un mapa de IAO 1.4?
    Get #fh, , MapDat
    
    With MapSize
        ReDim MapData(.XMin To .XMax, .YMin To .YMax)
        ReDim L1(.XMin To .XMax, .YMin To .YMax)
    End With
    
    Get #fh, , L1
    
    With MH

        If .NumeroBloqueados > 0 Then
            ReDim Blqs(1 To .NumeroBloqueados)
            Get #fh, , Blqs

            For i = 1 To .NumeroBloqueados
                MapData(Blqs(i).X, Blqs(i).Y).Blocked = 1
            Next i

        End If
        
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2

            For i = 1 To .NumeroLayers(2)
                Call InitGrh(MapData(L2(i).X, L2(i).Y).Graphic(2), L2(i).GrhIndex)
            Next i

        End If
        
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3

            For i = 1 To .NumeroLayers(3)
                Call InitGrh(MapData(L3(i).X, L3(i).Y).Graphic(3), L3(i).GrhIndex)
            Next i

        End If
        
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4

            For i = 1 To .NumeroLayers(4)
                Call InitGrh(MapData(L4(i).X, L4(i).Y).Graphic(4), L4(i).GrhIndex)
            Next i

        End If
        
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers

            For i = 1 To .NumeroTriggers
                
                With Triggers(i)
                    MapData(.X, .Y).Trigger = .Trigger
                End With
                
            Next i

        End If
        
        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas
            
            For i = 1 To .NumeroParticulas

                With Particulas(i)
                    MapData(.X, .Y).Particle_Group_Index = General_Particle_Create(.Particula, .X, .Y)
                End With

            Next i

        End If
            
        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)
            Get #fh, , Luces
            
            'Aca metes la carga de las luces...
        End If
            
        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos
            
            For i = 1 To .NumeroOBJs
                'Erase OBJs
                MapData(Objetos(i).X, Objetos(i).Y).ObjGrh.GrhIndex = 0
            Next i
            
        End If
            
        If .NumeroNPCs > 0 Then
            ReDim NPCs(1 To .NumeroNPCs)
            Get #fh, , NPCs
            
            For i = 1 To .NumeroNPCs
                MapData(NPCs(i).X, NPCs(i).Y).NPCIndex = NPCs(i).NPCIndex
            Next
                
        End If

        If .NumeroTE > 0 Then
            ReDim TEs(1 To .NumeroTE)
            Get #fh, , TEs

            For i = 1 To .NumeroTE
                
                With TEs(i)
                
                    MapData(.X, .Y).TileExit.Map = .DestM
                    MapData(.X, .Y).TileExit.X = .DestX
                    MapData(.X, .Y).TileExit.Y = .DestY
                
                End With
                
            Next i

        End If
        
    End With

    Close fh

    For J = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax

            If L1(i, J) > 0 Then
                Call InitGrh(MapData(i, J).Graphic(1), L1(i, J))
            End If

        Next i
    Next J

ErrorHandler:
    
    If fh <> 0 Then Close fh
    
    If Err.number <> 0 Then
        'Call LogError(Err.number, Err.Description, "modCarga.CargarMapa")
        Call MsgBox("err: " & Err.number, "desc: " & Err.Description)
    End If

End Sub
