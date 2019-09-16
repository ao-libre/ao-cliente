Attribute VB_Name = "Mod_General"
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
'Public cKeys As Collection
Public bFogata As Boolean

Public Type tRedditPost
    Title As String
    URL As String
End Type

Public Posts() As tRedditPost

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Private lFrameTimer As Long

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Function GetRawName(ByRef sName As String) As String
'***************************************************
'Author: ZaMa
'Last Modify Date: 13/01/2010
'Last Modified By: -
'Returns the char name without the clan name (if it has it).
'***************************************************

    Dim Pos As Integer
    
    Pos = InStr(1, sName, "<")
    
    If Pos > 0 Then
        GetRawName = Trim$(Left$(sName, Pos - 1))
    Else
        GetRawName = sName
    End If

End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, _
                    ByVal Text As String, _
                    Optional ByVal Red As Integer = -1, _
                    Optional ByVal Green As Integer, _
                    Optional ByVal Blue As Integer, _
                    Optional ByVal bold As Boolean = False, _
                    Optional ByVal italic As Boolean = False, _
                    Optional ByVal bCrLf As Boolean = True, _
                    Optional ByVal Alignment As Byte = rtfLeft)
    
'****************************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D apperance!
'****************************************************
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martin Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'Jopi 17/08/2019 : Consola transparente.
'Jopi 17/08/2019 : Ahora podes especificar el alineamiento del texto.
'****************************************************
    With RichTextBox
        
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        ' 0 = Left
        ' 1 = Center
        ' 2 = Right
        .SelAlignment = Alignment

        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        
        .SelText = Text

        ' Esto arregla el bug de las letras superponiendose la consola del frmMain
        If Not RichTextBox = frmMain.RecTxt Then RichTextBox.Refresh

    End With
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).active = 1 Then
            MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).CharIndex = loopc
        End If
    Next loopc
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    Dim Len_cad As Long
    
    cad = LCase$(cad)
    Len_cad = Len(cad)
    
    For i = 1 To Len_cad
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData() As Boolean
    
    'Validamos los datos del user
    
    Dim loopc As Long
    Dim CharAscii As Integer
    Dim Len_accountName As Long, Len_accountPassword As Long

    If LenB(AccountPassword) = 0 Then
        MsgBox JsonLanguage.item("VALIDACION_PASSWORD").item("TEXTO")
        Exit Function
    End If
    
    Len_accountPassword = Len(AccountPassword)
    
    For loopc = 1 To Len_accountPassword
        CharAscii = Asc(mid$(AccountPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox Replace$(JsonLanguage.item("VALIDACION_BAD_PASSWORD").item("TEXTO").item(2), "VAR_CHAR_INVALIDO", Chr$(CharAscii))
            Exit Function
        End If
    Next loopc

    If Len(AccountName) > 30 Then
        MsgBox JsonLanguage.item("VALIDACION_BAD_EMAIL").item("TEXTO").item(2)
        Exit Function
    End If
        
    Len_accountName = Len(AccountName)
    
    For loopc = 1 To Len_accountName
        CharAscii = Asc(mid$(AccountName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox Replace$(JsonLanguage.item("VALIDACION_BAD_PASSWORD").item("TEXTO").item(4), "VAR_CHAR_INVALIDO", Chr$(CharAscii))
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

    'Unload the connect form
    Unload frmCrearPersonaje
    Unload frmConnect
    Unload frmPanelAccount
    
    frmMain.lblName.Caption = UserName
    'Load main form
    frmMain.Visible = True
    lastKeys.Clear
    Call frmMain.ControlSM(eSMType.mWork, False)
    Call frmMain.ControlSM(eSMType.mSpells, False)
    FPSFLAG = True

End Sub

Sub CargarTip()
    Dim N As Integer
    N = RandomNumber(1, UBound(Tips))
    
    frmtip.tip.Caption = Tips(N)
End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saque lo que impedia que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    If LegalOk And Not UserParalizado Then
        Call WriteWalk(Direccion)
        If Not UserDescansar And Not UserMeditar Then
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    If frmMain.trainingMacro.Enabled Then Call frmMain.DesactivarMacroHechizos
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
    
    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.X, UserPos.Y)
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    Call Map_MoveTo(RandomNumber(NORTH, WEST))
End Sub

Private Sub CheckKeys()
     '*****************************************************************
    'Checks keys and respond
    '*****************************************************************
    'No input allowed while Argentum is not the active window
    If Not Application.IsAppActive() Then Exit Sub
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If MirandoForo Then Exit Sub
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'TODO: Deberia informarle por consola?
    If Traveling Then Exit Sub
    
    If UserMoving = 0 Then
        If Not UserEstupido Then
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyUp)) = False Then lastKeys.Add (CustomKeys.BindedKey(eKeyType.mKeyUp)) ' Agrega la tecla al arraylist
            Else
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyUp)) Then lastKeys.Remove (CustomKeys.BindedKey(eKeyType.mKeyUp)) ' Remueve la tecla que teniamos presionada
            End If
            
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyDown)) = False Then lastKeys.Add (CustomKeys.BindedKey(eKeyType.mKeyDown)) ' Agrega la tecla al arraylist
            Else
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyDown)) Then lastKeys.Remove (CustomKeys.BindedKey(eKeyType.mKeyDown)) ' Remueve la tecla que teniamos presionada
            End If
            
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyLeft)) = False Then lastKeys.Add (CustomKeys.BindedKey(eKeyType.mKeyLeft)) ' Agrega la tecla al arraylist
            Else
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyLeft)) Then lastKeys.Remove (CustomKeys.BindedKey(eKeyType.mKeyLeft)) ' Remueve la tecla que teniamos presionada
            End If
            
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyRight)) = False Then lastKeys.Add (CustomKeys.BindedKey(eKeyType.mKeyRight)) ' Agrega la tecla al arraylist
            Else
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyRight)) Then lastKeys.Remove (CustomKeys.BindedKey(eKeyType.mKeyRight)) ' Remueve la tecla que teniamos presionada
            End If
            'Move Up
            If lastKeys.Count() = 38 Then
                Debug.Print ("[" + CStr(lastKeys.item(1)) + "," + CStr(lastKeys.item(2)) + "," + CStr(lastKeys.item(3)) + "," + CStr(lastKeys.item(4)) + "]")
                Call Map_MoveTo(NORTH)
                Call Char_UserPos
                Exit Sub
            End If
            'Move Right
            If lastKeys.Count = 39 Then
                Debug.Print ("[" + CStr(lastKeys.item(1)) + "," + CStr(lastKeys.item(2)) + "," + CStr(lastKeys.item(3)) + "," + CStr(lastKeys.item(4)) + "]")
                Call Map_MoveTo(EAST)
                Call Char_UserPos
                Exit Sub
            End If
            'Move down
            If lastKeys.Count = 40 Then
                Debug.Print ("[" + CStr(lastKeys.item(1)) + "," + CStr(lastKeys.item(2)) + "," + CStr(lastKeys.item(3)) + "," + CStr(lastKeys.item(4)) + "]")
                Call Map_MoveTo(SOUTH)
                Call Char_UserPos
                Exit Sub
            End If
            'Move left
            If lastKeys.Count = 37 Then
                Debug.Print ("[" + CStr(lastKeys.item(1)) + "," + CStr(lastKeys.item(2)) + "," + CStr(lastKeys.item(3)) + "," + CStr(lastKeys.item(4)) + "]")
                Call Map_MoveTo(WEST)
                Call Char_UserPos
                Exit Sub
            End If
            ' We haven't moved - Update 3D sounds!
            Call Audio.MoveListener(UserPos.X, UserPos.Y)
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            Else
                ' We haven't moved - Update 3D sounds!
                Call Audio.MoveListener(UserPos.X, UserPos.Y)
            End If

            Call Char_UserPos
        End If
    Else

        Call frmMain.ActualizarMiniMapa   'integrado por ReyarB
        
    End If
End Sub

Sub SwitchMap(ByVal Map As Integer)
    '**********************************************************************************
    'Disenado y creado por Juan Martin Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
    '**********************************************************************************
    
    '**********************************************************************************
    'Formato de mapas optimizado para reducir el espacio que ocupan.
    'Nueva carga de mapas desde la memoria (clsByteBuffer)
    '[ https://www.gs-zone.org/temas/carga-de-mapas-desde-la-memoria-cliente.91444/ ]
    '**********************************************************************************

    Dim Y        As Long
    Dim X        As Long
    
    Dim ByFlags  As Byte
    Dim handle   As Integer
    Dim fileBuff As clsByteBuffer
   
    Dim dData()  As Byte
    Dim dLen     As Long
   
    Set fileBuff = New clsByteBuffer
    
    'Limpieza adicional del mapa. PARCHE: Solucion a bug de clones. [Gracias Yhunja]
    'EDIT: cambio el rango de valores en x y para solucionar otro bug con respecto al cambio de mapas
    Call Char_CleanAll
    
    dLen = FileLen(Game.path(Mapas) & "Mapa" & Map & ".map")
    ReDim dData(dLen - 1)
    
    handle = FreeFile()
    
    Open Game.path(Mapas) & "Mapa" & Map & ".map" For Binary As handle
        Get handle, , dData
    Close handle
     
    fileBuff.initializeReader dData
    
    mapInfo.MapVersion = fileBuff.getInteger
   
    With MiCabecera
        .Desc = fileBuff.getString(Len(.Desc))
        .CRC = fileBuff.getLong
        .MagicWord = fileBuff.getLong
    End With
    
    fileBuff.getDouble
   
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            ByFlags = fileBuff.getByte()

            With MapData(X, Y)
            
                .Blocked = (ByFlags And 1)
                .Graphic(1).GrhIndex = fileBuff.getInteger()
                InitGrh .Graphic(1), .Graphic(1).GrhIndex
           
                'Layer 2 used?
                If ByFlags And 2 Then
                    .Graphic(2).GrhIndex = fileBuff.getInteger()
                    InitGrh .Graphic(2), .Graphic(2).GrhIndex
                Else
                    .Graphic(2).GrhIndex = 0
                End If
               
                'Layer 3 used?
                If ByFlags And 4 Then
                    .Graphic(3).GrhIndex = fileBuff.getInteger()
                    InitGrh .Graphic(3), .Graphic(3).GrhIndex
                Else
                    .Graphic(3).GrhIndex = 0
                End If
               
                'Layer 4 used?
                If ByFlags And 8 Then
                    .Graphic(4).GrhIndex = fileBuff.getInteger()
                    InitGrh .Graphic(4), .Graphic(4).GrhIndex
                Else
                    .Graphic(4).GrhIndex = 0
                End If
           
                'Trigger used?
                If ByFlags And 16 Then
                    .Trigger = fileBuff.getInteger()
                Else
                    .Trigger = 0
                End If
           
                'Erase NPCs
                If .CharIndex > 0 Then
                    .CharIndex = 0
                End If
           
                'Erase OBJs
                If .ObjGrh.GrhIndex > 0 Then
                    .ObjGrh.GrhIndex = 0
                End If
            
                'Erase Lights
                Call Engine_D3DColor_To_RGB_List(.Engine_Light(), Estado_Actual) 'Standelf, Light & Meteo Engine
            
            End With
        Next X
    Next Y
    
    Call LightRemoveAll
    
    'Erase particle effects
    ReDim Effect(1 To NumEffects)
    Call Particle_Group_Remove_All
    
    'Limpiamos el buffer
    Set fileBuff = Nothing
    
    With mapInfo
        .name = vbNullString
        .Music = vbNullString
    End With
    
    'Dibujamos el Mini-Mapa
    If FileExist(Game.path(Graficos) & "MiniMapa\" & Map & ".bmp", vbArchive) Then
        frmMain.MiniMapa.Picture = LoadPicture(Game.path(Graficos) & "MiniMapa\" & Map & ".bmp")
    Else
        frmMain.MiniMapa.Visible = False
        frmMain.RecTxt.Width = frmMain.RecTxt.Width + 100
    End If
    
    CurMap = Map
    
    Call Init_Ambient(Map)
    
    'Carga las particulas especificas del mapa.
    Call Load_Map_Particles(Map)
    
    'Resetear el mensaje en render con el nombre del mapa.
    renderText = nameMap
    renderFont = 2
    colorRender = 240
    
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")
End Function

Public Function IsIp(ByVal Ip As String) As Boolean
    Dim i As Long
    Dim Upper_serversLst As Long
        Upper_serversLst = UBound(ServersLst)
    
    For i = 1 To Upper_serversLst
        If ServersLst(i).Ip = Ip Then
            IsIp = True
            Exit Function
        End If
    Next i
    
End Function

Private Function GetCountryFromIp(ByVal Ip As String) As String
'********************************
'Author: Recox
'Last Modification: 08/12/2018
'Added endpoint to obtain the country of the server.
'********************************
On Error Resume Next
    Dim URL As String
    Dim Endpoint As String
    Dim JsonObject As Object
    Dim Response As String
    
    Set Inet = New clsInet
    
    URL = GetVar(Game.path(INIT) & "Config.ini", "Parameters", "IpApiEndpoint")
    Endpoint = URL & Ip & "/json/"
    
    Response = Inet.OpenRequest(Endpoint, "GET")
    Response = Inet.Execute
    Response = Inet.GetResponseAsString
    
    Set JsonObject = JSON.parse(Response)
    
    GetCountryFromIp = JsonObject.item("country")
    
    Set Inet = Nothing
End Function

Public Sub CargarServidores()
'********************************
'Author: Unknown
'Last Modification: 21/12/2019
'Last Modified by: Recox
'Added Instruction "CloseClient" before End so the mutex is cleared (Rapsodius)
'Added IP Api to get the country of the IP. (Recox)
'Get ping from server (Recox)
'********************************
On Error GoTo errorH
    Dim File As String
    Dim Quantity As Integer
    Dim i As Integer
    Dim CountryCode As String
    Dim IpApiEnabled As Boolean
    Dim DoPingsEnabled As Boolean
    
    File = Game.path(INIT) & "sinfo.dat"
    Quantity = Val(GetVar(File, "INIT", "Cant"))
    IpApiEnabled = GetVar(Game.path(INIT) & "Config.ini", "Parameters", "IpApiEnabled")
    DoPingsEnabled = GetVar(Game.path(INIT) & "Config.ini", "Parameters", "DoPingsEnabled")
    
    frmConnect.lstServers.Clear
    
    ReDim ServersLst(1 To Quantity) As tServerInfo
    For i = 1 To Quantity
        Dim CurrentIp As String
        CurrentIp = Trim$(GetVar(File, "S" & i, "Ip"))
        
        If IpApiEnabled Then

           'If is not numeric do a url transformation
            If CheckIfIpIsNumeric(CurrentIp) = False Then
                CurrentIp = GetIPFromHostName(CurrentIp)
            End If

            CountryCode = GetCountryCode(CurrentIp)
            ServersLst(i).Desc = CountryCode & " - " & GetVar(File, "S" & i, "Desc")
        Else
            ServersLst(i).Desc = GetVar(File, "S" & i, "Desc")
        End If
        
        ServersLst(i).Ip = GetVar(File, "S" & i, "Ip")
        ServersLst(i).Puerto = CInt(GetVar(File, "S" & i, "PJ"))
        ServersLst(i).Mundo = GetVar(File, "S" & i, "MUNDO")
        'ServersLst(i).Ping = PingAddress(CurrentIp, "SomeRandomText")
        'ServersLst(i).Country = CountryCode

        'We should delete this validations and append text to the desc when we start working in something more suitable
        'in the UI to show the Pings, Country, Desc, etc.
        'All this functions are in the CODIGO/modPing.bas
        If DoPingsEnabled Then
            ServersLst(i).Desc = PingAddress(CurrentIp, "SomeRandomText") & " " & ServersLst(i).Desc
        End If

        frmConnect.lstServers.AddItem (ServersLst(i).Desc)
    Next i
    
    If CurServer = 0 Then CurServer = 1

Exit Sub

errorH:
    Call MsgBox("Error cargando los servidores, actualicelos de la web", vbCritical + vbOKOnly, "Argentum Online")
    
    'Call CloseClient
End Sub


Private Function CheckIfIpIsNumeric(CurrentIp As String) As String
    If IsNumeric(mid$(CurrentIp, 1, 1)) Then
        CheckIfIpIsNumeric = True
    Else
        CheckIfIpIsNumeric = False
    End If
End Function

Private Function GetCountryCode(CurrentIp As String) As String
    Dim CountryCode As String
    CountryCode = GetCountryFromIp(CurrentIp)

    If LenB(CountryCode) > 0 Then
        GetCountryCode = CountryCode
    Else
        GetCountryCode = "??"
    End If

End Function

Public Function CurServerIp() As String
    CurServerIp = frmConnect.IPTxt
End Function

Public Function CurServerPort() As Integer
    CurServerPort = Val(frmConnect.PortTxt)
End Function

Sub Main()
    ' Detecta el idioma del sistema (TRUE) y carga las traducciones
    Call SetLanguageApplication
    
    'Load client configurations.
    Call Game.LeerConfiguracion

    Call modCompression.GenerateContra(vbNullString, 0) ' 0 = Graficos.AO
    
    Call CargarHechizos

    ' Map Sounds
    Set Sonidos = New clsSoundMapas
    Call Sonidos.LoadSoundMapInfo
    
    
    'Comento esto ya que nosotros si permitimos abrir mas de un cliente a la ves.
    '#If Testeo = 0 Then
    '    If Application.FindPreviousInstance Then
    '        Call MsgBox(JsonLanguage.Item("OTRO_CLIENTE_ABIERTO").Item("TEXTO"), vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
    '        End
    '    End If
    '#End If

    'Read command line. Do it AFTER config file is loaded to prevent this from
    'canceling the effects of "/nores" option.
    Call LeerLineaComandos
    
    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.path & "\")
    
    ChDrive App.path
    ChDir App.path

    'Set resolution BEFORE the loading form is displayed, therefore it will be centered.
    Call Resolution.SetResolution(800, 600)

    ' Load constants, classes, flags, graphics..
    Call LoadInitialConfig
    
    If GetVar(Game.path(INIT) & "Config.ini", "Parameters", "TestMode") <> 1 Then
        With frmPres
            .Picture = LoadPicture(Game.path(Interfaces) & "ImagenPresentacion.jpg")
            .Show vbModal    'Es modal, asi que se detiene la ejecucionn de Main hasta que se desaparece
        End With
    End If

    frmConnect.Visible = True
    
    'Inicializacion de variables globales
    PrimeraVez = True
    prgRun = True
    pausa = False
    
    ' Intervals
    LoadTimerIntervals
        
    'Set the dialog's font
    Set DialogosClanes = New clsGuildDlg
    DialogosClanes.Activo = ClientSetup.bGldMsgConsole
    DialogosClanes.CantidadDialogos = ClientSetup.bCantMsgs
    Dialogos.Font = frmMain.Font
    DialogosClanes.Font = frmMain.Font
    
    lFrameTimer = GetTickCount
    
    ' Load the form for screenshots
    Call Load(frmScreenshots)
        
    Do While prgRun

        'Solo dibujamos si la ventana no esta minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
            'Play ambient sounds
            Call RenderSounds
            
            Call CheckKeys
        End If
        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            If FPSFLAG Then frmMain.lblFPS.Caption = Mod_TileEngine.FPS
            
            lFrameTimer = GetTickCount
        End If
        
        ' If there is anything to be sent, we send it
        Call FlushBuffer
        
        DoEvents
        
    Loop
    
    Call CloseClient
End Sub

Public Function GetVersionOfTheGame() As String
    GetVersionOfTheGame = GetVar(Game.path(INIT) & "Config.ini", "Cliente", "VersionTagRelease")
End Function

Private Sub LoadInitialConfig()
'***************************************************
'Author: ZaMa
'Last Modification: 15/03/2011
'15/03/2011: ZaMa - Initialize classes lazy way.
'***************************************************

    frmCargando.Show
    frmCargando.Refresh

    frmConnect.version = GetVersionOfTheGame()
    
    '#######
    ' CLASES
    Call AddtoRichTextBox(frmCargando.status, _
                            JsonLanguage.item("INICIA_CLASES").item("TEXTO"), _
                            JsonLanguage.item("INICIA_CLASES").item("COLOR").item(1), _
                            JsonLanguage.item("INICIA_CLASES").item("COLOR").item(2), _
                            JsonLanguage.item("INICIA_CLASES").item("COLOR").item(3), _
                            True, False, True, rtfCenter)
                            
    Set Dialogos = New clsDialogs
    Set Audio = New clsAudio
    Set lastKeys = New clsArrayList
    Set Inventario = New clsGraphicalInventory
    Set CustomKeys = New clsCustomKeys
    Set CustomMessages = New clsCustomMessages
    Set incomingData = New clsByteQueue
    Set outgoingData = New clsByteQueue
    Set MainTimer = New clsTimer
    Set clsForos = New clsForum
    Set frmMain.Client = New clsSocket
    Call lastKeys.Initialize(1, 4)
    Call AddtoRichTextBox(frmCargando.status, _
                            "   " & JsonLanguage.item("HECHO").item("TEXTO"), _
                            JsonLanguage.item("HECHO").item("COLOR").item(1), _
                            JsonLanguage.item("HECHO").item("COLOR").item(2), _
                            JsonLanguage.item("HECHO").item("COLOR").item(3), _
                            True, False, False, rtfLeft)
    
    '#############
    ' DIRECT SOUND
    Call AddtoRichTextBox(frmCargando.status, _
                            JsonLanguage.item("INICIA_SONIDO").item("TEXTO"), _
                            JsonLanguage.item("INICIA_SONIDO").item("COLOR").item(1), _
                            JsonLanguage.item("INICIA_SONIDO").item("COLOR").item(2), _
                            JsonLanguage.item("INICIA_SONIDO").item("COLOR").item(3), _
                            True, False, True, rtfCenter)
                            
    'Inicializamos el sonido
    Call Audio.Initialize(DirectX, frmMain.hWnd, Game.path(Sounds), Game.path(Musica))

    'Enable / Disable audio
    Audio.MusicActivated = ClientSetup.bMusic
    Audio.SoundActivated = ClientSetup.bSound
    Audio.SoundEffectsActivated = ClientSetup.bSoundEffects
    Call Audio.PlayMIDI("6.mid")
    
    Call AddtoRichTextBox(frmCargando.status, _
                            "   " & JsonLanguage.item("HECHO").item("TEXTO"), _
                            JsonLanguage.item("HECHO").item("COLOR").item(1), _
                            JsonLanguage.item("HECHO").item("COLOR").item(2), _
                            JsonLanguage.item("HECHO").item("COLOR").item(3), _
                            True, False, False, rtfLeft)
    
    '###########
    ' CONSTANTES
    Call AddtoRichTextBox(frmCargando.status, _
                            JsonLanguage.item("INICIA_CONSTANTES").item("TEXTO"), _
                            JsonLanguage.item("INICIA_CONSTANTES").item("COLOR").item(1), _
                            JsonLanguage.item("INICIA_CONSTANTES").item("COLOR").item(2), _
                            JsonLanguage.item("INICIA_CONSTANTES").item("COLOR").item(3), _
                            True, False, True, rtfCenter)
                            
    Call InicializarNombres
    
    ' Initialize FONTTYPES
    Call Protocol.InitFonts
 
    UserMap = 1
    
    ' Mouse Pointer (Loaded before opening any form with buttons in it)
    If FileExist(Game.path(Extras) & "Hand.ico", vbArchive) Then _
        Set picMouseIcon = LoadPicture(Game.path(Extras) & "Hand.ico")
    
    Call AddtoRichTextBox(frmCargando.status, _
                            "   " & JsonLanguage.item("HECHO").item("TEXTO"), _
                            JsonLanguage.item("HECHO").item("COLOR").item(1), _
                            JsonLanguage.item("HECHO").item("COLOR").item(2), _
                            JsonLanguage.item("HECHO").item("COLOR").item(3), _
                            True, False, False, rtfLeft)
    

    '##############
    ' MOTOR GRAFICO
    Call AddtoRichTextBox(frmCargando.status, _
                            JsonLanguage.item("INICIA_MOTOR_GRAFICO").item("TEXTO"), _
                            JsonLanguage.item("INICIA_MOTOR_GRAFICO").item("COLOR").item(1), _
                            JsonLanguage.item("INICIA_MOTOR_GRAFICO").item("COLOR").item(2), _
                            JsonLanguage.item("INICIA_MOTOR_GRAFICO").item("COLOR").item(3), _
                            True, False, True, rtfCenter)
    
    '     Iniciamos el Engine de DirectX 8
    If Not Engine_DirectX8_Init Then
        Call CloseClient
    End If
          
    '     Tile Engine
    If Not InitTileEngine(frmMain.hWnd, 32, 32, 8, 8) Then
        Call CloseClient
    End If
    
    Call Engine_DirectX8_Aditional_Init

    Call AddtoRichTextBox(frmCargando.status, _
                            "   " & JsonLanguage.item("HECHO").item("TEXTO"), _
                            JsonLanguage.item("HECHO").item("COLOR").item(1), _
                            JsonLanguage.item("HECHO").item("COLOR").item(2), _
                            JsonLanguage.item("HECHO").item("COLOR").item(3), _
                            True, False, False, rtfLeft)
    
    '###################
    ' ANIMACIONES EXTRAS
    Call AddtoRichTextBox(frmCargando.status, _
                            JsonLanguage.item("INICIA_FXS").item("TEXTO"), _
                            JsonLanguage.item("INICIA_FXS").item("COLOR").item(1), _
                            JsonLanguage.item("INICIA_FXS").item("COLOR").item(2), _
                            JsonLanguage.item("INICIA_FXS").item("COLOR").item(3), _
                            True, False, True, rtfCenter)
                            
    Call CargarTips
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    
    Call AddtoRichTextBox(frmCargando.status, _
                            "   " & JsonLanguage.item("HECHO").item("TEXTO"), _
                            JsonLanguage.item("HECHO").item("COLOR").item(1), _
                            JsonLanguage.item("HECHO").item("COLOR").item(2), _
                            JsonLanguage.item("HECHO").item("COLOR").item(3), _
                            True, False, False, rtfLeft)
    
    'Inicializamos el inventario grafico
    Call Inventario.Initialize(DirectD3D8, frmMain.PicInv, MAX_INVENTORY_SLOTS, , , , , , , , True)
    'Set cKeys = New Collection
    Call AddtoRichTextBox(frmCargando.status, _
                            JsonLanguage.item("BIENVENIDO").item("TEXTO"), _
                            JsonLanguage.item("BIENVENIDO").item("COLOR").item(1), _
                            JsonLanguage.item("BIENVENIDO").item("COLOR").item(2), _
                            JsonLanguage.item("BIENVENIDO").item("COLOR").item(3), _
                            True, False, True, rtfCenter)
                            
    '###################
    ' PETICIONES API
    Call GetPostsFromReddit

    Unload frmCargando
    
End Sub

Private Sub LoadTimerIntervals()
    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/03/2011
    'Set the intervals of timers
    '***************************************************
    
    With MainTimer
    
        Call .SetInterval(TimersIndex.Attack, eIntervalos.INT_ATTACK)
        Call .SetInterval(TimersIndex.Work, eIntervalos.INT_WORK)
        Call .SetInterval(TimersIndex.UseItemWithU, eIntervalos.INT_USEITEMU)
        Call .SetInterval(TimersIndex.UseItemWithDblClick, eIntervalos.INT_USEITEMDCK)
        Call .SetInterval(TimersIndex.SendRPU, eIntervalos.INT_SENTRPU)
        Call .SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
        Call .SetInterval(TimersIndex.Arrows, eIntervalos.INT_ARROWS)
        Call .SetInterval(TimersIndex.CastAttack, eIntervalos.INT_CAST_ATTACK)
        
        With frmMain.macrotrabajo
            .Interval = eIntervalos.INT_MACRO_TRABAJO
            .Enabled = False
        End With
    
        'Init timers
        Call .Start(TimersIndex.Attack)
        Call .Start(TimersIndex.Work)
        Call .Start(TimersIndex.UseItemWithU)
        Call .Start(TimersIndex.UseItemWithDblClick)
        Call .Start(TimersIndex.SendRPU)
        Call .Start(TimersIndex.CastSpell)
        Call .Start(TimersIndex.Arrows)
        Call .Start(TimersIndex.CastAttack)
    
    End With

End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, Value, File
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Funcion para chequear el email
'
'  Corregida por Maraxus para que reconozca como validas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    Dim Len_sString As Long
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . despues de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        'pre-calculo la cantidad de caracteres para mejorar el rendimiento
        Len_sString = Len(sString) - 1
        
        '3er test: Recorre todos los caracteres y los valida
        For lX = 0 To Len_sString
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como validas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer aca....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And _
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

''
' Checks the command line parameters, if you are running Ao with /nores command
'
'

Public Sub LeerLineaComandos()
'*************************************************
'Author: Unknown
'Last modified: 25/11/2008 (BrianPr)
'
'*************************************************
    
    Dim i As Long, T() As String, Upper_t As Long, Lower_t As Long
    
    'Parseo los comandos
    T = Split(Command, " ")
    Lower_t = LBound(T)
    Upper_t = UBound(T)
    
    For i = Lower_t To Upper_t
        Select Case UCase$(T(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
        End Select
    Next i

End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Lindos"
    Ciudades(eCiudad.cArghal) = "Arghal"
    
    ListaRazas(eRaza.Humano) = JsonLanguage.item("RAZAS").item("HUMANO")
    ListaRazas(eRaza.Elfo) = JsonLanguage.item("RAZAS").item("ELFO")
    ListaRazas(eRaza.ElfoOscuro) = JsonLanguage.item("RAZAS").item("ELFO_OSCURO")
    ListaRazas(eRaza.Gnomo) = JsonLanguage.item("RAZAS").item("GNOMO")
    ListaRazas(eRaza.Enano) = JsonLanguage.item("RAZAS").item("ENANO")


    ' No uso las traducciones ya que muchas cosas estan hardcodeadas en castellano
    ' ListaClases(eClass.Mage) = JsonLanguage.Item("CLASES").Item("MAGO")
    ' ListaClases(eClass.Cleric) = JsonLanguage.Item("CLASES").Item("CLERIGO")
    ' ListaClases(eClass.Warrior) = JsonLanguage.Item("CLASES").Item("GUERRERO")
    ' ListaClases(eClass.Assasin) = JsonLanguage.Item("CLASES").Item("ASESINO")
    ' ListaClases(eClass.Thief) = JsonLanguage.Item("CLASES").Item("LADRON")
    ' ListaClases(eClass.Bard) = JsonLanguage.Item("CLASES").Item("BARDO")
    ' ListaClases(eClass.Druid) = JsonLanguage.Item("CLASES").Item("DRUIDA")
    ' ListaClases(eClass.Bandit) = JsonLanguage.Item("CLASES").Item("BANDIDO")
    ' ListaClases(eClass.Paladin) = JsonLanguage.Item("CLASES").Item("PALADIN")
    ' ListaClases(eClass.Hunter) = JsonLanguage.Item("CLASES").Item("CAZADOR")
    ' ListaClases(eClass.Worker) = JsonLanguage.Item("CLASES").Item("TRABAJADOR")
    ' ListaClases(eClass.Pirat) = JsonLanguage.Item("CLASES").Item("PIRATA")

    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"
    
    SkillsNames(eSkill.Magia) = JsonLanguage.item("HABILIDADES").item("MAGIA").item("TEXTO")
    SkillsNames(eSkill.Robar) = JsonLanguage.item("HABILIDADES").item("ROBAR").item("TEXTO")
    SkillsNames(eSkill.Tacticas) = JsonLanguage.item("HABILIDADES").item("EVASION_EN_COMBATE").item("TEXTO")
    SkillsNames(eSkill.Armas) = JsonLanguage.item("HABILIDADES").item("COMBATE_CON_ARMAS").item("TEXTO")
    SkillsNames(eSkill.Meditar) = JsonLanguage.item("HABILIDADES").item("MEDITAR").item("TEXTO")
    SkillsNames(eSkill.Apunalar) = JsonLanguage.item("HABILIDADES").item("APUNALAR").item("TEXTO")
    SkillsNames(eSkill.Ocultarse) = JsonLanguage.item("HABILIDADES").item("OCULTARSE").item("TEXTO")
    SkillsNames(eSkill.Supervivencia) = JsonLanguage.item("HABILIDADES").item("SUPERVIVENCIA").item("TEXTO")
    SkillsNames(eSkill.Talar) = JsonLanguage.item("HABILIDADES").item("TALAR").item("TEXTO")
    SkillsNames(eSkill.Comerciar) = JsonLanguage.item("HABILIDADES").item("COMERCIO").item("TEXTO")
    SkillsNames(eSkill.Defensa) = JsonLanguage.item("HABILIDADES").item("DEFENSA_CON_ESCUDOS").item("TEXTO")
    SkillsNames(eSkill.Pesca) = JsonLanguage.item("HABILIDADES").item("PESCA").item("TEXTO")
    SkillsNames(eSkill.Mineria) = JsonLanguage.item("HABILIDADES").item("MINERIA").item("TEXTO")
    SkillsNames(eSkill.Carpinteria) = JsonLanguage.item("HABILIDADES").item("CARPINTERIA").item("TEXTO")
    SkillsNames(eSkill.Herreria) = JsonLanguage.item("HABILIDADES").item("HERRERIA").item("TEXTO")
    SkillsNames(eSkill.Liderazgo) = JsonLanguage.item("HABILIDADES").item("LIDERAZGO").item("TEXTO")
    SkillsNames(eSkill.Domar) = JsonLanguage.item("HABILIDADES").item("DOMAR_ANIMALES").item("TEXTO")
    SkillsNames(eSkill.Proyectiles) = JsonLanguage.item("HABILIDADES").item("COMBATE_A_DISTANCIA").item("TEXTO")
    SkillsNames(eSkill.Wrestling) = JsonLanguage.item("HABILIDADES").item("COMBATE_CUERPO_A_CUERPO").item("TEXTO")
    SkillsNames(eSkill.Navegacion) = JsonLanguage.item("HABILIDADES").item("NAVEGACION").item("TEXTO")

    AtributosNames(eAtributos.Fuerza) = JsonLanguage.item("ATRIBUTOS").item("FUERZA")
    AtributosNames(eAtributos.Agilidad) = JsonLanguage.item("ATRIBUTOS").item("AGILIDAD")
    AtributosNames(eAtributos.Inteligencia) = JsonLanguage.item("ATRIBUTOS").item("INTELIGENCIA")
    AtributosNames(eAtributos.Carisma) = JsonLanguage.item("ATRIBUTOS").item("CARISMA")
    AtributosNames(eAtributos.Constitucion) = JsonLanguage.item("ATRIBUTOS").item("CONSTITUCION")
End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    frmMain.RecTxt.Text = vbNullString
    
    Call DialogosClanes.RemoveDialogs
    
    Call Dialogos.RemoveAllDialogs
End Sub

Public Sub CloseClient()
    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 8/14/2007
    'Frees all used resources, cleans up and leaves
    '**************************************************************
    
    ' Allow new instances of the client to be opened
    Call Application.ReleaseInstance
    
    EngineRun = False
    
    'Cerramos Sockets/Winsocks/WindowsAPI
    frmMain.Client.CloseSck
    
    'Stop tile engine
    Call Engine_DirectX8_End

    'Destruimos los objetos publicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    Set JsonLanguage = Nothing
    Set frmMain.Client = Nothing
    
    Call UnloadAllForms
    
    'Si se cambio la resolucion, la reseteamos.
    If ResolucionCambiada Then Resolution.ResetResolution
    
    End
    
End Sub


Public Function esGM(CharIndex As Integer) As Boolean
esGM = False
If charlist(CharIndex).priv >= 1 And charlist(CharIndex).priv <= 5 Or charlist(CharIndex).priv = 25 Then _
    esGM = True

End Function

Public Function getTagPosition(ByVal Nick As String) As Integer
Dim buf As Integer
buf = InStr(Nick, "<")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
buf = InStr(Nick, "[")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
getTagPosition = Len(Nick) + 2
End Function

Public Sub checkText(ByVal Text As String)
Dim Nivel As Integer
If Right$(Text, Len(JsonLanguage.item("MENSAJE_FRAGSHOOTER_TE_HA_MATADO").item("TEXTO"))) = JsonLanguage.item("MENSAJE_FRAGSHOOTER_TE_HA_MATADO").item("TEXTO") Then
    Call ScreenCapture(True)
    Exit Sub
End If
If Left$(Text, Len(JsonLanguage.item("MENSAJE_FRAGSHOOTER_HAS_MATADO").item("TEXTO"))) = JsonLanguage.item("MENSAJE_FRAGSHOOTER_HAS_MATADO").item("TEXTO") Then
    EsperandoLevel = True
    Exit Sub
End If
If EsperandoLevel Then
    If Right$(Text, Len(JsonLanguage.item("MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA").item("TEXTO"))) = JsonLanguage.item("MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA").item("TEXTO") Then
        If CInt(mid$(Text, Len(JsonLanguage.item("MENSAJE_FRAGSHOOTER_HAS_GANADO").item("TEXTO")), (Len(Text) - (Len(JsonLanguage.item("MENSAJE_FRAGSHOOTER_HAS_GANADO").item("TEXTO")))))) / 2 > ClientSetup.byMurderedLevel Then
            Call ScreenCapture(True)
        End If
    End If
End If
EsperandoLevel = False
End Sub

Public Function getStrenghtColor() As Long
Dim m As Long
m = 255 / MAXATRIBUTOS
getStrenghtColor = RGB(255 - (m * UserFuerza), (m * UserFuerza), 0)
End Function
Public Function getDexterityColor() As Long
Dim m As Long
m = 255 / MAXATRIBUTOS
getDexterityColor = RGB(255, m * UserAgilidad, 0)
End Function

Public Function getCharIndexByName(ByVal name As String) As Integer
Dim i As Long
For i = 1 To LastChar
    If charlist(i).Nombre = name Then
        getCharIndexByName = i
        Exit Function
    End If
Next i
End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Returns true if the post is sticky.
'***************************************************
    Select Case ForumType
        Case eForumMsgType.ieCAOS_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieGENERAL_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieREAL_STICKY
            EsAnuncio = True
            
    End Select
    
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte
'***************************************************
'Author: ZaMa
'Last Modification: 01/03/2010
'Returns the forum alignment.
'***************************************************
    Select Case yForumType
        Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
            ForumAlignment = eForumType.ieCAOS
            
        Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
            ForumAlignment = eForumType.ieGeneral
            
        Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
            ForumAlignment = eForumType.ieREAL
            
    End Select
    
End Function

Public Sub ResetAllInfo()

    ' Disable timers
    frmMain.Second.Enabled = False
    frmMain.macrotrabajo.Enabled = False
    Connected = False
    
    'Unload all forms except frmMain, frmConnect and frmCrearPersonaje
    Dim frm As Form
    For Each frm In Forms
        If frm.name <> frmMain.name And frm.name <> frmConnect.name And _
            frm.name <> frmCrearPersonaje.name Then
            
            Unload frm
        End If
    Next
    
    On Local Error GoTo 0
    
    ' Return to connection screen
    frmConnect.MousePointer = vbNormal
    If Not frmCrearPersonaje.Visible Then frmConnect.Visible = True
    frmMain.Visible = False
    
    'Stop audio
    Call Audio.StopWave
    frmMain.IsPlaying = PlayLoop.plNone
    
    ' Reset flags
    pausa = False
    UserMeditar = False
    UserEstupido = False
    UserCiego = False
    UserDescansar = False
    UserParalizado = False
    Traveling = False
    UserNavegando = False
    UserEvento = False
    bFogata = False
    bRain = False
    bFogata = False
    Comerciando = False
    bShowTutorial = False
    
    MirandoAsignarSkills = False
    MirandoCarpinteria = False
    MirandoEstadisticas = False
    MirandoForo = False
    MirandoHerreria = False
    MirandoParty = False
    
    'Delete all kind of dialogs
    Call CleanDialogs

    'Reset some char variables...
    Dim i As Long
    For i = 1 To LastChar
        charlist(i).invisible = False
    Next i

    ' Reset stats
    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = vbNullString
    SkillPoints = 0
    Alocados = 0
    
    ' Reset skills
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    ' Reset attributes
    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    ' Clear inventory slots
    Inventario.ClearAllSlots

    ' Connection screen midi
    Call Audio.PlayMIDI("2.mid")

End Sub

Public Function DevolverNombreHechizo(ByVal Index As Byte) As String
Dim i As Long
 
    For i = 1 To NumHechizos
        If i = Index Then
            DevolverNombreHechizo = Hechizos(i).Nombre
            Exit Function
        End If
    Next i
End Function
Public Function DevolverIndexHechizo(ByVal Nombre As String) As Byte
Dim i As Long
 
    For i = 1 To NumHechizos
        If Hechizos(i).Nombre = Nombre Then
            DevolverIndexHechizo = i
            Exit Function
        End If
    Next i
End Function

Sub DownloadServersFile(myURL As String)
'**********************************************************
'Downloads the sinfo.dat file from a given url
'Last change: 01/11/2018
'Implemented by Cucsifae
'Check content of strData to avoid clean the file sinfo.ini if there is no response from Github by Recox
'**********************************************************
On Error Resume Next
    Dim strData As String
    Dim f As Integer
    
    Set Inet = New clsInet
    
    strData = Inet.OpenRequest(myURL, "GET")
    strData = Inet.Execute
    strData = Inet.GetResponseAsString
    
    f = FreeFile
    
    If LenB(strData) <> 0 Then
        Open Game.path(INIT) & "sinfo.dat" For Output As #f
            Print #f, strData
        Close #f
    End If
    
    Exit Sub
End Sub

Function EaseOutCubic(Time As Double)
    Time = Time - 1
    EaseOutCubic = Time * Time * Time + 1
End Function

Public Sub GetPostsFromReddit()
On Error Resume Next
    If UBound(Posts) = 0 Then
        Set Inet = New clsInet
        
        Dim ResponseReddit As String
        Dim JsonObject As Object
        Dim Endpoint As String
        
        Endpoint = GetVar(Game.path(INIT) & "Config.ini", "Parameters", "SubRedditEndpoint")
    
        ResponseReddit = Inet.OpenRequest(Endpoint, "GET")
        ResponseReddit = Inet.Execute
        ResponseReddit = Inet.GetResponseAsString
        
        
        Set JsonObject = JSON.parse(ResponseReddit)
        
        Dim qtyPostsOnReddit As Integer: qtyPostsOnReddit = JsonObject.item("data").item("children").Count
        
        ReDim Preserve Posts(qtyPostsOnReddit)
        
        'Clear lstRedditPosts before populate it again to prevent repeated values.
        frmConnect.lstRedditPosts.Clear
        'Long funciona mas rapido en los loops que Integer
        Dim i As Long
        i = 1
        Do While i <= qtyPostsOnReddit
            Posts(i).Title = JsonObject.item("data").item("children").item(i).item("data").item("title")
            Posts(i).URL = JsonObject.item("data").item("children").item(i).item("data").item("url")
            
            frmConnect.lstRedditPosts.AddItem JsonObject.item("data").item("children").item(i).item("data").item("title")
            
            i = i + 1
        Loop
        
        Set Inet = Nothing
    Else
        Dim ia As Long
        For ia = 1 To UBound(Posts)
            frmConnect.lstRedditPosts.AddItem Posts(ia).Title
        Next ia
    End If
    
End Sub
