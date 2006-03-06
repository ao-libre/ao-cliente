Attribute VB_Name = "Mod_General"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
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

Public bO As Integer
Public bK As Long
Public bRK As Long


Public iplst As String
Public banners As String

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Private lFrameLimiter As Long
Public lFrameTimer As Long
Public sHKeys() As String

Public Function DirGraficos() As String
    DirGraficos = App.Path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
    DirSound = App.Path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.Path & "\" & Config_Inicio.DirMusica & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.Path & "\" & Config_Inicio.DirMapas & "\"
End Function

Public Function SumaDigitos(ByVal numero As Integer) As Integer
'Suma digitos
    Dim suma As Integer
    
    While (numero <> 0)
        suma = suma + (numero Mod 10)
        numero = numero \ 10
    Wend
    
    SumaDigitos = suma
End Function

Public Function SumaDigitosMenos(ByVal numero As Integer) As Integer
    'Suma digitos, y resta el total de dígitos
    Dim suma As Integer
    
    Do While (numero <> 0)
        suma = suma + (numero Mod 10) - 1
        
        numero = numero \ 10
    Loop
    
    SumaDigitosMenos = suma
End Function

Public Function Complex(ByVal numero As Integer) As Integer
    If numero Mod 2 <> 0 Then
        Complex = numero * SumaDigitos(numero)
    Else
        Complex = numero * SumaDigitosMenos(numero)
    End If
End Function

Public Function ValidarLoginMSG(ByVal numero As Integer) As Integer
    Dim AuxInteger As Integer
    Dim AuxInteger2 As Integer
    
    AuxInteger = SumaDigitos(numero)
    AuxInteger2 = SumaDigitosMenos(numero)
    ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.Path & "\init\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub CargarVersiones()
On Error GoTo errorH:

    Versiones(1) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Graficos", "Val"))
    Versiones(2) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Wavs", "Val"))
    Versiones(3) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Midis", "Val"))
    Versiones(4) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Init", "Val"))
    Versiones(5) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Mapas", "Val"))
    Versiones(6) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "E", "Val"))
    Versiones(7) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "O", "Val"))
Exit Sub

errorH:
    Call MsgBox("Error cargando versiones")
End Sub

Sub CargarColores()
    Dim archivoC As String
    
    archivoC = App.Path & "\init\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim I As Long
    
    For I = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(I).r = Val(GetVar(archivoC, Str(I), "R"))
        ColoresPJ(I).G = Val(GetVar(archivoC, Str(I), "G"))
        ColoresPJ(I).B = Val(GetVar(archivoC, Str(I), "B"))
    Next I
        
    ColoresPJ(50).r = Val(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).G = Val(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).B = Val(GetVar(archivoC, "CR", "B"))
    ColoresPJ(49).r = Val(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).G = Val(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).B = Val(GetVar(archivoC, "CI", "B"))
End Sub

#If SeguridadAlkon Then
Sub InitMI()
    Dim alternativos As Integer
    Dim CualMITemp As Integer
    
    alternativos = RandomNumber(1, 7368)
    CualMITemp = RandomNumber(1, 1233)
    

    Set MI(CualMITemp) = New clsManagerInvisibles
    Call MI(CualMITemp).Inicializar(alternativos, 10000)
    
    If CualMI <> 0 Then
        Call MI(CualMITemp).CopyFrom(MI(CualMI))
        Set MI(CualMI) = Nothing
    End If
    CualMI = CualMITemp
End Sub
#End If

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.Path & "\init\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal Red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************
    With RichTextBox
        If (Len(.Text)) > 10000 Then .Text = ""
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        
        .SelBold = Bold
        .SelItalic = Italic
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        RichTextBox.Refresh
    End With
End Sub

'TODO : Never was sure this is really necessary....
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).Active = 1 Then
            MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).CharIndex = loopc
        End If
    Next loopc
End Sub

Sub SaveGameini()
    'Grabamos los datos del usuario en el Game.ini
    Config_Inicio.Name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort
    
    Call EscribirGameIni(Config_Inicio)
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim I As Long
    
    cad = LCase$(cad)
    
    For I = 1 To Len(cad)
        car = Asc(Mid$(cad, I, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next I
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(Mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(Mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    Call SaveGameini

    'Unload the connect form
    Unload frmConnect
    
    frmMain.Label8.Caption = UserName
    'Load main form
    frmMain.Visible = True
End Sub

Sub CargarTip()
    Dim N As Integer
    N = RandomNumber(1, UBound(Tips))
    
    frmtip.tip.Caption = Tips(N)
End Sub

'TODO : Hay formas más limpias, mantenibles y eficientes....
'Sub MoveNorth()
'    If Cartel Then Cartel = False
'
'    If LegalPos(UserPos.X, UserPos.Y - 1) Then
'        Call SendData("M" & E_Heading.NORTH)
'        If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
'            Call MoveCharbyHead(UserCharIndex, E_Heading.NORTH)
'            Call MoveScreen(E_Heading.NORTH)
'            DoFogataFx
'        End If
'    Else
'        If charlist(UserCharIndex).Heading <> E_Heading.NORTH Then
'            Call SendData("CHEA" & E_Heading.NORTH)
'        End If
'    End If
'End Sub
'
''TODO : Hay formas más limpias, mantenibles y eficientes....
'Sub MoveEast()
'    If Cartel Then Cartel = False
'
'    If LegalPos(UserPos.X + 1, UserPos.Y) Then
'        Call SendData("M" & E_Heading.EAST)
'        If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
'            Call MoveCharbyHead(UserCharIndex, E_Heading.EAST)
'            Call MoveScreen(E_Heading.EAST)
'            Call DoFogataFx
'        End If
'    Else
'        If charlist(UserCharIndex).Heading <> E_Heading.EAST Then
'            Call SendData("CHEA" & E_Heading.EAST)
'        End If
'    End If
'End Sub
'
''TODO : Hay formas más limpias, mantenibles y eficientes....
'Sub MoveSouth()
'    If Cartel Then Cartel = False
'
'    If LegalPos(UserPos.X, UserPos.Y + 1) Then
'        Call SendData("M" & E_Heading.SOUTH)
'        If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
'            MoveCharbyHead UserCharIndex, E_Heading.SOUTH
'            MoveScreen E_Heading.SOUTH
'            DoFogataFx
'        End If
'    Else
'        If charlist(UserCharIndex).Heading <> E_Heading.SOUTH Then
'            Call SendData("CHEA" & E_Heading.SOUTH)
'        End If
'    End If
'End Sub
'
''TODO : Hay formas más limpias, mantenibles y eficientes....
'Sub MoveWest()
'    If Cartel Then Cartel = False
'
'    If LegalPos(UserPos.X - 1, UserPos.Y) Then
'        Call SendData("M" & E_Heading.WEST)
'        If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
'            MoveCharbyHead UserCharIndex, E_Heading.WEST
'            MoveScreen E_Heading.WEST
'            DoFogataFx
'        End If
'    Else
'        If charlist(UserCharIndex).Heading <> E_Heading.WEST Then
'            Call SendData("CHEA" & E_Heading.WEST)
'        End If
'    End If
'End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    Select Case Direccion
    Case E_Heading.NORTH
        LegalOk = LegalPos(UserPos.X, UserPos.Y - 1)
    Case E_Heading.EAST
        LegalOk = LegalPos(UserPos.X + 1, UserPos.Y)
    Case E_Heading.SOUTH
        LegalOk = LegalPos(UserPos.X, UserPos.Y + 1)
    Case E_Heading.WEST
        LegalOk = LegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    If LegalOk Then
        Call SendData("M" & Direccion)
        If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
            DoFogataFx
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call SendData("CHEA" & Direccion)
        End If
    End If
End Sub

'TODO : Hay formas más limpias, mantenibles y eficientes....
Sub RandomMove()
'    Dim j As Integer
'
'    j = RandomNumber(1, 4)
'
'    Select Case j
'        Case 1
'            Call MoveEast
'        Case 2
'            Call MoveNorth
'        Case 3
'            Call MoveWest
'        Case 4
'            Call MoveSouth
'    End Select

    MoveTo RandomNumber(1, 4)
    
End Sub

Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
On Error Resume Next
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(vbKeyUp) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(NORTH)
                frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
        
            'Move Right
            If GetKeyState(vbKeyRight) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(EAST)
                frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(vbKeyDown) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(SOUTH)
                frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(vbKeyLeft) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(WEST)
                frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
        Else
            Dim kp As Boolean
            kp = (GetKeyState(vbKeyUp) < 0) Or _
                GetKeyState(vbKeyRight) < 0 Or _
                GetKeyState(vbKeyDown) < 0 Or _
                GetKeyState(vbKeyLeft) < 0
            If kp Then Call RandomMove
            If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
            frmMain.Coord.Caption = "(" & UserPos.X & "," & UserPos.Y & ")"
        End If
    End If
End Sub

'TODO : esto no es del tileengine??
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

    If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
        Exit Sub
    End If
End Sub

'TODO : esto no es del tileengine??
Function NextOpenChar()
'******************************************
'Finds next open Char
'******************************************
    Dim loopc As Long
    
    loopc = 1
    Do While charlist(loopc).Active And loopc < UBound(charlist)
        loopc = loopc + 1
    Loop
    
    NextOpenChar = loopc
End Function

'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!
Sub SwitchMap(ByVal Map As Integer)
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
    Dim loopc As Long
    Dim Y As Long
    Dim X As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    
    Open DirMapas & "Mapa" & Map & ".map" For Binary As #1
    Seek #1, 1
            
    'map Header
    Get #1, , MapInfo.MapVersion
    Get #1, , MiCabecera
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            Get #1, , ByFlags
            
            MapData(X, Y).Blocked = (ByFlags And 1)
            
            Get #1, , MapData(X, Y).Graphic(1).GrhIndex
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get #1, , MapData(X, Y).Graphic(2).GrhIndex
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
            Else
                MapData(X, Y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get #1, , MapData(X, Y).Graphic(3).GrhIndex
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
            Else
                MapData(X, Y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get #1, , MapData(X, Y).Graphic(4).GrhIndex
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
            Else
                MapData(X, Y).Graphic(4).GrhIndex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get #1, , MapData(X, Y).Trigger
            Else
                MapData(X, Y).Trigger = 0
            End If
            
            'Erase NPCs
            If MapData(X, Y).CharIndex > 0 Then
                Call EraseChar(MapData(X, Y).CharIndex)
            End If
            
            'Erase OBJs
            MapData(X, Y).ObjGrh.GrhIndex = 0
        Next X
    Next Y
    
    Close #1
    
    MapInfo.Name = ""
    MapInfo.Music = ""
    
    CurMap = Map
End Sub

'TODO : Reemplazar por la nueva versión, esta apesta!!!
Public Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************
    Dim I As Integer
    Dim LastPos As Integer
    Dim CurChar As String * 1
    Dim FieldNum As Integer
    Dim Seperator As String
    
    Seperator = Chr$(SepASCII)
    LastPos = 0
    FieldNum = 0
    
    For I = 1 To Len(Text)
        CurChar = Mid$(Text, I, 1)
        If CurChar = Seperator Then
            FieldNum = FieldNum + 1
            If FieldNum = Pos Then
                ReadField = Mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = I
        End If
    Next I
    FieldNum = FieldNum + 1
    
    If FieldNum = Pos Then
        ReadField = Mid$(Text, LastPos + 1)
    End If
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Sub WriteClientVer()
    Dim hFile As Integer
        
    hFile = FreeFile()
    Open App.Path & "\init\Ver.bin" For Binary Access Write Lock Read As #hFile
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    
    Put #hFile, , CInt(App.Major)
    Put #hFile, , CInt(App.Minor)
    Put #hFile, , CInt(App.Revision)
    
    Close #hFile
End Sub

Public Function IsIp(ByVal Ip As String) As Boolean
    Dim I As Long
    
    For I = 1 To UBound(ServersLst)
        If ServersLst(I).Ip = Ip Then
            IsIp = True
            Exit Function
        End If
    Next I
End Function

Public Sub CargarServidores()
On Error GoTo errorH
    Dim f As String
    Dim C As Integer
    Dim I As Long
    
    f = App.Path & "\init\sinfo.dat"
    C = Val(GetVar(f, "INIT", "Cant"))
    
    ReDim ServersLst(1 To C) As tServerInfo
    For I = 1 To C
        ServersLst(I).desc = GetVar(f, "S" & I, "Desc")
        ServersLst(I).Ip = Trim$(GetVar(f, "S" & I, "Ip"))
        ServersLst(I).PassRecPort = CInt(GetVar(f, "S" & I, "P2"))
        ServersLst(I).Puerto = CInt(GetVar(f, "S" & I, "PJ"))
    Next I
    CurServer = 1
Exit Sub

errorH:
    Call MsgBox("Error cargando los servidores, actualicelos de la web", vbCritical + vbOKOnly, "Argentum Online")
    End
End Sub

Public Sub InitServersList(ByVal Lst As String)
On Error Resume Next
    Dim NumServers As Integer
    Dim I As Integer
    Dim Cont As Integer
    
    I = 1
    
    Do While (ReadField(I, RawServersList, Asc(";")) <> "")
        I = I + 1
        Cont = Cont + 1
    Loop
    
    ReDim ServersLst(1 To Cont) As tServerInfo
    
    For I = 1 To Cont
        Dim cur$
        cur$ = ReadField(I, RawServersList, Asc(";"))
        ServersLst(I).Ip = ReadField(1, cur$, Asc(":"))
        ServersLst(I).Puerto = ReadField(2, cur$, Asc(":"))
        ServersLst(I).desc = ReadField(4, cur$, Asc(":"))
        ServersLst(I).PassRecPort = ReadField(3, cur$, Asc(":"))
    Next I
    
    CurServer = 1
End Sub

Public Function CurServerPasRecPort() As Integer
    If CurServer <> 0 Then
        CurServerPasRecPort = 7667
    Else
        CurServerPasRecPort = CInt(frmConnect.PortTxt)
    End If
End Function

Public Function CurServerIp() As String
    If CurServer <> 0 Then
        CurServerIp = ServersLst(CurServer).Ip
    Else
        CurServerIp = frmConnect.IPTxt
    End If
End Function

Public Function CurServerPort() As Integer
    If CurServer <> 0 Then
        CurServerPort = ServersLst(CurServer).Puerto
    Else
        CurServerPort = CInt(frmConnect.PortTxt)
    End If
End Function


Sub Main()
'TODO : Cambiar esto cuando se corrija el bug de los timers
'On Error GoTo ManejadorErrores
On Error Resume Next
    Call WriteClientVer
    
    Call LeerLineaComandos

    If App.PrevInstance Then
        Call MsgBox("Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If

Dim f As Boolean
Dim ulttick As Long, esttick As Long
Dim timers(1 To 2) As Integer

    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.Path & "\")
    
    ChDrive App.Path
    ChDir App.Path

'Obtengo mi MD5 hash
    'Obtener el HushMD5
    Dim fMD5HushYo As String * 32
    fMD5HushYo = MD5File(App.Path & "\" & App.EXEName & ".exe")
    MD5HushYo = txtOffset(hexMd52Asc(fMD5HushYo), 53)
    
    
    'Cargamos el archivo de configuracion inicial
    If FileExist(App.Path & "\init\Inicio.con", vbNormal) Then
        Config_Inicio = LeerGameIni()
    End If


If FileExist(App.Path & "\init\ao.dat", vbArchive) Then
        Call LoadRenderMode
    End If
    
    
    tipf = Config_Inicio.tip
    
    frmCargando.Show
    frmCargando.Refresh
    
    frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    AddtoRichTextBox frmCargando.status, "Buscando servidores....", 0, 0, 0, 0, 0, 1

#If UsarWrench = 1 Then
    frmMain.Socket1.Startup
#End If

    Call CargarServidores
    'TODO : esto de ServerRecibidos no se podría sacar???
    ServersRecibidos = True
    
    AddtoRichTextBox frmCargando.status, "Encontrado", , , , 1
    AddtoRichTextBox frmCargando.status, "Iniciando constantes...", 0, 0, 0, 0, 0, 1
    
    Call InicializarNombres
    
    frmOldPersonaje.NameTxt.Text = Config_Inicio.Name
    frmOldPersonaje.PasswordTxt.Text = ""
    
    AddtoRichTextBox frmCargando.status, "Hecho", , , , 1
    
    IniciarObjetosDirectX
    
    AddtoRichTextBox frmCargando.status, "Cargando Sonidos....", 0, 0, 0, 0, 0, 1
    AddtoRichTextBox frmCargando.status, "Hecho", , , , 1

Dim loopc As Integer

LastTime = GetTickCount

Call InitTileEngine(frmMain.hWnd, 152, 7, 32, 32, 13, 17, 9)
    
    Call AddtoRichTextBox(frmCargando.status, "Creando animaciones extra....")
    
    Call CargarAnimsExtra
    Call CargarTips

UserMap = 1

    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarVersiones
    Call CargarColores
    
#If SeguridadAlkon Then
    CualMI = 0
    Call InitMI
#End If

    AddtoRichTextBox frmCargando.status, "                    ¡Bienvenido a Argentum Online!", , , , 1
    
    Unload frmCargando
    
    'Inicializamos el sonido
    Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectSound....", 0, 0, 0, 0, 0, True)
    Call Audio.Initialize(DirectX, frmMain.hWnd, App.Path & "\" & Config_Inicio.DirSonidos & "\", App.Path & "\" & Config_Inicio.DirMusica & "\")
    Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1, , False)
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(DirectDraw, frmMain.picInv)
    
    If Musica = 0 Then
        Call Audio.PlayMIDI(MIdi_Inicio & ".mid")
    End If

    frmPres.Picture = LoadPicture(App.Path & "\Graficos\bosquefinal.jpg")
    frmPres.Show vbModal    'Es modal, así que se detiene la ejecución de Main hasta que se desaparece
    
    frmConnect.Visible = True

    MainViewRect.Left = MainViewLeft + 32 * RenderMod.iImageSize
    MainViewRect.Top = MainViewTop + 32 * RenderMod.iImageSize
    MainViewRect.Right = (MainViewRect.Left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
    MainViewRect.Bottom = (MainViewRect.Top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)
    
'TODO : Esto va en Engine Initialization
    MainDestRect.Left = ((TilePixelWidth * TileBufferSize) - TilePixelWidth) + 32 * RenderMod.iImageSize
    MainDestRect.Top = ((TilePixelHeight * TileBufferSize) - TilePixelHeight) + 32 * RenderMod.iImageSize
    MainDestRect.Right = (MainDestRect.Left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
    MainDestRect.Bottom = (MainDestRect.Top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)
    
    'Inicialización de variables globales
    PrimeraVez = True
    prgRun = True
    pausa = False
    lFrameLimiter = GetTickCount
    
    Do While prgRun
        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 Then
            Call ShowNextFrame
        End If
    
'TODO : Porque el pausado de 20 ms???
        If GetTickCount - LastTime > 20 Then
            If Not pausa And frmMain.Visible And Not frmForo.Visible And Not frmComerciar.Visible And Not frmComerciarUsu.Visible And Not frmBancoObj.Visible Then
                CheckKeys
                LastTime = GetTickCount
            End If
        End If

        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            FramesPerSec = FramesPerSecCounter
            
            If FPSFLAG Then frmMain.Caption = FramesPerSec
            
            FramesPerSecCounter = 0
            lFrameTimer = GetTickCount
        End If
        ''Limitamos las FPS a 20 (con el nuevo engine 60 es un número mucho mejor)
        If 50 - GetTickCount + lFrameLimiter > 0 Then
            Sleep 50 - GetTickCount + lFrameLimiter
        End If
        
        'Esto nos permite calcular cuanto tarda en dibujar un cuadro
        lFrameLimiter = GetTickCount
        
'TODO : Sería mejor comparar el tiempo desde la última vez que se hizo hasta el actual SOLO cuando se precisa. Además evitás el corte de intervalos con 2 golpes seguidos.
        'Sistema de timers renovado:
        esttick = GetTickCount
        For loopc = 1 To UBound(timers)
            timers(loopc) = timers(loopc) + (esttick - ulttick)
            'Timer de trabajo
            If timers(1) >= tUs Then
                timers(1) = 0
                NoPuedeUsar = False
            End If
            'timer de attaque (77)
            If timers(2) >= tAt Then
                timers(2) = 0
                UserCanAttack = 1
                UserPuedeRefrescar = True
            End If
        Next loopc
        ulttick = GetTickCount
        
        DoEvents
    Loop

    EngineRun = False
    frmCargando.Show
    AddtoRichTextBox frmCargando.status, "Liberando recursos...", 0, 0, 0, 0, 0, 1
    LiberarObjetosDX

'TODO : Esto debería ir en otro lado como al cambair a esta res
    If Not bNoResChange Then
        Dim typDevM As typDevMODE
        Dim lRes As Long
        
        lRes = EnumDisplaySettings(0, 0, typDevM)
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
        End With
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    End If

    'Destruimos los objetos públicos creados
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    
    Call UnloadAllForms
    
    'Actualizar tip
    Config_Inicio.tip = tipf
    Call EscribirGameIni(Config_Inicio)
End

ManejadorErrores:
    MsgBox "Ha ocurrido un error irreparable, el cliente se cerrará."
    LogError "Contexto:" & Err.HelpContext & " Desc:" & Err.Description & " Fuente:" & Err.Source
    End
End Sub
Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(Mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lorelativo a mapas, no tiene anda que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean

    HayAgua = MapData(X, Y).Graphic(1).GrhIndex >= 1505 And _
                MapData(X, Y).Graphic(1).GrhIndex <= 1520 And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus
    End If
End Sub
    
Public Sub LeerLineaComandos()
    Dim T() As String
    Dim I As Long
    
    'Parseo los comandos
    T = Split(Command, " ")
    
    For I = LBound(T) To UBound(T)
        Select Case UCase$(T(I))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
        End Select
    Next I
End Sub

Private Sub LoadRenderMode()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    Open App.Path & "\init\ao.dat" For Binary Access Read Lock Write As fHandle
        Get fHandle, , RenderMod
    Close fHandle
    
    Musica = RenderMod.bNoMusic
    Fx = RenderMod.bNoSound
    
    Select Case RenderMod.iImageSize
        Case 4
            RenderMod.iImageSize = 0
        Case 3
            RenderMod.iImageSize = 1
        Case 2
            RenderMod.iImageSize = 2
        Case 1
            RenderMod.iImageSize = 3
        Case 0
            RenderMod.iImageSize = 4
    End Select
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(1) = "Ullathorpe"
    Ciudades(2) = "Nix"
    Ciudades(3) = "Banderbill"

    CityDesc(1) = "Ullathorpe está establecida en el medio de los grandes bosques de Argentum, es principalmente un pueblo de campesinos y leñadores. Su ubicación hace de Ullathorpe un punto de paso obligado para todos los aventureros ya que se encuentra cerca de los lugares más legendarios de este mundo."
    CityDesc(2) = "Nix es una gran ciudad. Edificada sobre la costa oeste del principal continente de Argentum."
    CityDesc(3) = "Banderbill se encuentra al norte de Ullathorpe y Nix, es una de las ciudades más importantes de todo el imperio."

    ListaRazas(1) = "Humano"
    ListaRazas(2) = "Elfo"
    ListaRazas(3) = "Elfo Oscuro"
    ListaRazas(4) = "Gnomo"
    ListaRazas(5) = "Enano"

    ListaClases(1) = "Mago"
    ListaClases(2) = "Clerigo"
    ListaClases(3) = "Guerrero"
    ListaClases(4) = "Asesino"
    ListaClases(5) = "Ladron"
    ListaClases(6) = "Bardo"
    ListaClases(7) = "Druida"
    ListaClases(8) = "Bandido"
    ListaClases(9) = "Paladin"
    ListaClases(10) = "Cazador"
    ListaClases(11) = "Pescador"
    ListaClases(12) = "Herrero"
    ListaClases(13) = "Leñador"
    ListaClases(14) = "Minero"
    ListaClases(15) = "Carpintero"
    ListaClases(16) = "Pirata"

    SkillsNames(Skills.Suerte) = "Suerte"
    SkillsNames(Skills.Magia) = "Magia"
    SkillsNames(Skills.Robar) = "Robar"
    SkillsNames(Skills.Tacticas) = "Tacticas de combate"
    SkillsNames(Skills.Armas) = "Combate con armas"
    SkillsNames(Skills.Meditar) = "Meditar"
    SkillsNames(Skills.Apuñalar) = "Apuñalar"
    SkillsNames(Skills.Ocultarse) = "Ocultarse"
    SkillsNames(Skills.Supervivencia) = "Supervivencia"
    SkillsNames(Skills.Talar) = "Talar árboles"
    SkillsNames(Skills.Comerciar) = "Comercio"
    SkillsNames(Skills.Defensa) = "Defensa con escudos"
    SkillsNames(Skills.Pesca) = "Pesca"
    SkillsNames(Skills.Mineria) = "Mineria"
    SkillsNames(Skills.Carpinteria) = "Carpinteria"
    SkillsNames(Skills.Herreria) = "Herreria"
    SkillsNames(Skills.Liderazgo) = "Liderazgo"
    SkillsNames(Skills.Domar) = "Domar animales"
    SkillsNames(Skills.Proyectiles) = "Armas de proyectiles"
    SkillsNames(Skills.Wresterling) = "Wresterling"
    SkillsNames(Skills.Navegacion) = "Navegacion"

    AtributosNames(1) = "Fuerza"
    AtributosNames(2) = "Agilidad"
    AtributosNames(3) = "Inteligencia"
    AtributosNames(4) = "Carisma"
    AtributosNames(5) = "Constitucion"
End Sub
