Attribute VB_Name = "Carga"
Option Explicit

Private GrhIndex As Long

Private FileManager As clsIniManager
Private Extension As String

''
' Cargar indices de graficos.
'
Public Sub LoadGraphicsIndex()

    On Error GoTo ErrorHandler:
    
    Dim FileHandle As Integer: FileHandle = FreeFile()
    
    #If IndicesBinarios = 1 Then

        Dim Grh         As Long
        Dim Frame       As Long
        Dim grhCount    As Long
        Dim handle      As Integer
        Dim fileVersion As Long
    
        'Open files
        handle = FreeFile()
        Open Game.path(INIT) & "Graficos.ind" For Binary Access Read As FileHandle
        
        'Obtenemos la version del .ind
        Get FileHandle, , fileVersion
        
        'Obtenemos la cantidad de indices.
        Get FileHandle, , grhCount
        ReDim GrhData(0 To grhCount) As GrhData
        
        While Not EOF(FileHandle)

            Get FileHandle, , Grh
            
            With GrhData(Grh)
            
                '.active = True
                Get FileHandle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ReDim .Frames(1 To .NumFrames)
                
                If .NumFrames > 1 Then
                
                    For Frame = 1 To .NumFrames
                        Get FileHandle, , .Frames(Frame)
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then GoTo ErrorHandler
                    Next Frame
                    
                    Get FileHandle, , .speed
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
                    
                    Get FileHandle, , .FileNum
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    Get FileHandle, , GrhData(Grh).sX
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    Get FileHandle, , .sY
                    If .sY < 0 Then GoTo ErrorHandler
                    
                    Get FileHandle, , .pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get FileHandle, , .pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = Grh
                    
                End If
                
            End With
            
        Wend

    #Else

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
        If UBound(GrhData) <= 0 Then GoTo ErrorHandler
        
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
            If LenB(CurrentLine) <> 0 Then
            
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
                            If .Frames(Frame) <= LBound(GrhData) Or .Frames(Frame) > UBound(GrhData) Then GoTo ErrorHandler
                        Next
                        
                        ' Tercer lugar: leemos la velocidad de la animacion
                        .speed = Val(Fields(Frame))
                        If .speed <= 0 Then GoTo ErrorHandler
                        
                        ' Por ultimo, copiamos las dimensiones del primer frame
                        .pixelHeight = GrhData(.Frames(1)).pixelHeight
                        If .pixelHeight <= 0 Then GoTo ErrorHandler
                        
                        .pixelWidth = GrhData(.Frames(1)).pixelWidth
                        If .pixelWidth <= 0 Then GoTo ErrorHandler
                        
                        .TileWidth = GrhData(.Frames(1)).TileWidth
                        If .TileWidth <= 0 Then GoTo ErrorHandler
                        
                        .TileHeight = GrhData(.Frames(1)).TileHeight
                        If .TileHeight <= 0 Then GoTo ErrorHandler
        
                    ElseIf .NumFrames = 1 Then
                    
                        ' Si es un solo frame lo asignamos a si mismo
                        .Frames(1) = Grh
                        
                        ' Segundo lugar: NumeroDelGrafico.bmp, pero sin el ".bmp"
                        .FileNum = Val(Fields(1))

                        If .FileNum <= 0 Then GoTo ErrorHandler
                            
                        ' Tercer Lugar: La coordenada X del grafico
                        .sX = Val(Fields(2))
                        If .sX < 0 Then GoTo ErrorHandler
                            
                        ' Cuarto Lugar: La coordenada Y del grafico
                        .sY = Val(Fields(3))
                        If .sY < 0 Then GoTo ErrorHandler
                            
                        ' Quinto lugar: El ancho del grafico
                        .pixelWidth = Val(Fields(4))
                        If .pixelWidth <= 0 Then GoTo ErrorHandler
                            
                        ' Sexto lugar: La altura del grafico
                        .pixelHeight = Val(Fields(5))
                        If .pixelHeight <= 0 Then GoTo ErrorHandler
                        
                        ' Calculamos el ancho y alto en tiles
                        .TileWidth = .pixelWidth / TilePixelHeight
                        .TileHeight = .pixelHeight / TilePixelWidth
                        
                    Else
                        ' 0 frames o negativo? Error
                        GoTo ErrorHandler

                    End If
        
                End With

            End If

        Loop
        
    #End If
    
    Close FileHandle
    
    Exit Sub
    
ErrorHandler:
    
    Close FileHandle
    
    If Err.number <> 0 Then

        #If IndicesBinarios = 1 Then
            Extension = ".ind"
        #Else
            Extension = ".ini"
        #End If
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Graficos" & Extension & " no existe. Por favor, reinstale el juego.", , "Argentum Online")
        
        ElseIf Grh > 0 Then
            Call MsgBox("Hay un error en Graficos" & Extension & " con el Grh" & Grh & ".", , "Argentum Online")
        
        Else
            Call MsgBox("Hay un error [" & Err.number & " - " & Err.Description & "] en Graficos" & Extension & ". Por favor, reinstale el juego.", , "Argentum Online")
        
        End If
        
        Call CloseClient
        
    End If
    
    Exit Sub
    
End Sub

Public Sub CargarCabezas()

    On Error GoTo errhandler:

    Dim i          As Long
    Dim j          As Long
    Dim NumHeads   As Integer
    Dim MisCabezas As tIndiceCabeza
    
    #If IndicesBinarios = 1 Then

        Dim N As Integer: N = FreeFile()
        Open Game.path(INIT) & "Cabezas.ind" For Binary Access Read As #N
            
        'cabecera
        Get #N, , MiCabecera
            
        'num de cabezas
        Get #N, , NumHeads
            
        'Resize array
        ReDim HeadData(0 To NumHeads) As HeadData
            
        For i = 1 To NumHeads
            Get #N, , MisCabezas
                
            If MisCabezas.Head(1) Then
                
                For j = 1 To 4
                    Call InitGrh(HeadData(i).Head(j), MisCabezas.Head(j), 0)
                Next
                
            End If

        Next i
            
        Close #N
            
    #Else

        Set FileManager = New clsIniManager
        Call FileManager.Initialize(Game.path(INIT) & "Cabezas.ini")
            
        'Obtenemos la cantidad de indices de las cabezas.
        NumHeads = Val(FileManager.GetValue("INIT", "NumHeads"))
            
        'Resize array
        ReDim HeadData(0 To NumHeads) As HeadData
            
        For i = 1 To NumHeads
            For j = 1 To 4
            
                GrhIndex = Val(FileManager.GetValue("HEAD" & i, "HEAD" & j))

                If GrhIndex > 0 Then
                    Call InitGrh(HeadData(i).Head(j), GrhIndex, 0)
                End If
                
            Next j
        Next i
            
        Set FileManager = Nothing
            
    #End If
    
errhandler:
    
    #If IndicesBinarios = 1 Then
        Extension = ".ind"
    #Else
        Extension = ".ini"
    #End If
        
    Select Case Err.number
    
        Case 0
            Exit Sub
            
        Case 53
            Call MsgBox("El archivo Cabezas" & Extension & " no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
            
        Case Else
            Call MsgBox("Hay un error [" & Err.number & " - " & Err.Description & "] en Cabezas" & Extension & ". Por favor, reinstale el juego.", , "Argentum Online")
            Call CloseClient
            
    End Select

End Sub

Sub CargarCascos()
On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim MisCabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open Game.path(INIT) & "Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim MisCabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , MisCabezas(i)
        
        If MisCabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), MisCabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), MisCabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), MisCabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), MisCabezas(i).Head(4), 0)
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

    Dim i     As Long
    Dim j     As Long
    
    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Game.path(INIT) & "armas.dat")
    
    NumWeaponAnims = Val(FileManager.GetValue("INIT", "NumArmas"))
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For i = 1 To NumWeaponAnims
        For j = 1 To 4
            
            GrhIndex = Val(FileManager.GetValue("ARMA" & i, "Dir" & j))

            If GrhIndex > 0 Then
                Call InitGrh(WeaponAnimData(i).WeaponWalk(j), GrhIndex, 0)
            End If
                
        Next j
    Next i
    
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
    Dim R As Long, G As Long, B As Long
    
    For i = 0 To 47 '48, 49 y 50 reservados para atacables, ciudadano y criminal
        R = Val(FileManager.GetValue(CStr(i), "R"))
        G = Val(FileManager.GetValue(CStr(i), "G"))
        B = Val(FileManager.GetValue(CStr(i), "B"))
        ColoresPJ(i) = D3DColorXRGB(R, G, B)
    Next i
    
    '   Atacable TODO: hay que implementar un color para los atacables y hacer que funcione.
    'R = Val(FileManager.GetValue("AT", "R"))
    'G = Val(FileManager.GetValue("AT", "G"))
    'B = Val(FileManager.GetValue("AT", "B"))
    'ColoresPJ(48) = D3DColorXRGB(R, G, B)
    
    '   Ciuda
    R = Val(FileManager.GetValue("CI", "R"))
    G = Val(FileManager.GetValue("CI", "G"))
    B = Val(FileManager.GetValue("CI", "B"))
    ColoresPJ(49) = D3DColorXRGB(R, G, B)
    
    '   Crimi
    R = Val(FileManager.GetValue("CR", "R"))
    G = Val(FileManager.GetValue("CR", "G"))
    B = Val(FileManager.GetValue("CR", "B"))
    ColoresPJ(50) = D3DColorXRGB(R, G, B)
    
    For i = 51 To 56 'Colores reservados para la renderizacion de dano
        R = Val(FileManager.GetValue(CStr(i), "R"))
        G = Val(FileManager.GetValue(CStr(i), "G"))
        B = Val(FileManager.GetValue(CStr(i), "B"))
        ColoresDano(i) = D3DColorXRGB(R, G, B)
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

    Dim i           As Long
    Dim j           As Long
    Dim NumEscudosAnims As Long
    
    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Game.path(INIT) & "escudos.dat")
    
    NumEscudosAnims = Val(FileManager.GetValue("INIT", "NumEscudos"))
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For i = 1 To NumEscudosAnims
        For j = 1 To 4
            
            GrhIndex = Val(FileManager.GetValue("ESC" & i, "Dir" & j))

            If GrhIndex > 0 Then
                Call InitGrh(ShieldAnimData(i).ShieldWalk(j), GrhIndex, 0)
            End If
                
        Next j
    Next i
    
    Set FileManager = Nothing
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo escudos.dat no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient

        End If
        
    End If
    
End Sub

'************************************************************************************.
' De aca en adelante solo se cargan archivos en texto plano.
'************************************************************************************

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
