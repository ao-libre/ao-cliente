Attribute VB_Name = "ProtocolCmdParse"
'Argentum Online
'
'Copyright (C) 2006 Juan Martín Sotuyo Dodero (Maraxus)
'Copyright (C) 2006 Alejandro Santos (AlejoLp)

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

Option Explicit

Public Enum eNumber_Types
    ent_Byte
    ent_Integer
    ent_Long
    ent_Trigger
End Enum

Public Sub AuxWriteWhisper(ByVal UserName As String, ByVal Mensaje As String)
    If LenB(UserName) = 0 Then Exit Sub
    
    Dim i As Long
    Dim nameLength As Long
    
    If (InStrB(UserName, "+") <> 0) Then
        UserName = Replace$(UserName, "+", " ")
    End If
    
    UserName = UCase$(UserName)
    nameLength = Len(UserName)
    
    i = 1
    Do While i <= LastChar
        If UCase$(charlist(i).Nombre) = UserName Or UCase$(Left$(charlist(i).Nombre, nameLength + 2)) = UserName & " <" Then
            Exit Do
        Else
            i = i + 1
        End If
    Loop
    
    If i <= LastChar Then
        Call WriteWhisper(i, Mensaje)
    End If
End Sub

''
' Interpreta, valida y ejecuta el comando ingresado .
'
' @param    RawCommand El comando en version String
' @remarks  None Known.

Public Sub ParseUserCommand(ByVal RawCommand As String)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modification: 16/11/2009
'Interpreta, valida y ejecuta el comando ingresado
'26/03/2009: ZaMa - Flexibilizo la cantidad de parametros de /nene,  /onlinemap y /telep
'16/11/2009: ZaMa - Ahora el /ct admite radio
'***************************************************
    Dim TmpArgos() As String
    
    Dim Comando As String
    Dim ArgumentosAll() As String
    Dim ArgumentosRaw As String
    Dim Argumentos2() As String
    Dim Argumentos3() As String
    Dim Argumentos4() As String
    Dim CantidadArgumentos As Long
    Dim notNullArguments As Boolean
    
    Dim tmpArr() As String
    Dim tmpInt As Integer
    
    ' TmpArgs: Un array de a lo sumo dos elementos,
    ' el primero es el comando (hasta el primer espacio)
    ' y el segundo elemento es el resto. Si no hay argumentos
    ' devuelve un array de un solo elemento
    TmpArgos = Split(RawCommand, " ", 2)
    
    Comando = Trim$(UCase$(TmpArgos(0)))
    
    If UBound(TmpArgos) > 0 Then
        ' El string en crudo que este despues del primer espacio
        ArgumentosRaw = TmpArgos(1)
        
        'veo que los argumentos no sean nulos
        notNullArguments = LenB(Trim$(ArgumentosRaw))
        
        ' Un array separado por blancos, con tantos elementos como
        ' se pueda
        ArgumentosAll = Split(TmpArgos(1), " ")
        
        ' Cantidad de argumentos. En ESTE PUNTO el minimo es 1
        CantidadArgumentos = UBound(ArgumentosAll) + 1
        
        ' Los siguientes arrays tienen A LO SUMO, COMO MAXIMO
        ' 2, 3 y 4 elementos respectivamente. Eso significa
        ' que pueden tener menos, por lo que es imperativo
        ' preguntar por CantidadArgumentos.
        
        Argumentos2 = Split(TmpArgos(1), " ", 2)
        Argumentos3 = Split(TmpArgos(1), " ", 3)
        Argumentos4 = Split(TmpArgos(1), " ", 4)
    Else
        CantidadArgumentos = 0
    End If
    
    ' Sacar cartel APESTA!! (y es ilógico, estás diciendo una pausa/espacio  :rolleyes: )
    If Comando = "" Then Comando = " "
    
    If Left$(Comando, 1) = "/" Then
        ' Comando normal
        
        Select Case Comando
            Case "/ONLINE"
                Call WriteOnline
                
            Case "/SALIR"
                If UserParalizado Then 'Inmo
                    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
                        Call ShowConsoleMsg("No puedes salir estando paralizado.", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
                Call WriteQuit
                
            Case "/SALIRCLAN"
                Call WriteGuildLeave
                
            Case "/BALANCE"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteRequestAccountState
                
            Case "/QUIETO"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WritePetStand
                
            Case "/ACOMPAÑAR"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WritePetFollow
                
            Case "/LIBERAR"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteReleasePet
                
            Case "/ENTRENAR"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteTrainList
                
            Case "/DESCANSAR"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteRest
                
            Case "/MEDITAR"
                If UserMinMAN = UserMaxMAN Then Exit Sub
                
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteMeditate
        
            Case "/CONSULTA"
                Call WriteConsulta
            
            Case "/RESUCITAR"
                Call WriteResucitate
                
            Case "/CURAR"
                Call WriteHeal
                              
            Case "/EST"
                Call WriteRequestStats
            
            Case "/AYUDA"
                Call WriteHelp
                
            Case "/COMERCIAR"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                
                ElseIf Comerciando Then 'Comerciando
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Ya estás comerciando", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteCommerceStart
                
            Case "/BOVEDA"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteBankStart
                
            Case "/ENLISTAR"
                Call WriteEnlist
                    
            Case "/INFORMACION"
                Call WriteInformation
                
            Case "/RECOMPENSA"
                Call WriteReward
                
            Case "/MOTD"
                Call WriteRequestMOTD
                
            Case "/UPTIME"
                Call WriteUpTime
                
            Case "/SALIRPARTY"
                Call WritePartyLeave
                
            Case "/CREARPARTY"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WritePartyCreate
                
            Case "/PARTY"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WritePartyJoin
            
            Case "/COMPARTIRNPC"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                
                Call WriteShareNpc
                
            Case "/NOCOMPARTIRNPC"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                
                Call WriteStopSharingNpc
                
            Case "/ENCUESTA"
                If CantidadArgumentos = 0 Then
                    ' Version sin argumentos: Inquiry
                    Call WriteInquiry
                Else
                    ' Version con argumentos: InquiryVote
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Byte) Then
                        Call WriteInquiryVote(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Para votar una opcion, escribe /encuesta NUMERODEOPCION, por ejemplo para votar la opcion 1, escribe /encuesta 1.")
                    End If
                End If
        
            Case "/CMSG"
                'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
                If CantidadArgumentos > 0 Then
                    Call WriteGuildMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")
                End If
        
            Case "/PMSG"
                'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
                If CantidadArgumentos > 0 Then
                    Call WritePartyMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")
                End If
            
            Case "/CENTINELA"
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteCentinelReport(CInt(ArgumentosRaw))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("El código de verificación debe ser numerico. Utilice /centinela X, siendo X el código de verificación.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /centinela X, siendo X el código de verificación.")
                End If
        
            Case "/ONLINECLAN"
                Call WriteGuildOnline
                
            Case "/ONLINEPARTY"
                Call WritePartyOnline
                
            Case "/BMSG"
                If notNullArguments Then
                    Call WriteCouncilMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")
                End If
                
            Case "/ROL"
                If notNullArguments Then
                    Call WriteRoleMasterRequest(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba una pregunta.")
                End If
                
            Case "/GM"
                Call WriteGMRequest
                
            Case "/_BUG"
                If notNullArguments Then
                    Call WriteBugReport(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba una descripción del bug.")
                End If
            
            Case "/DESC"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                
                Call WriteChangeDescription(ArgumentosRaw)
            
            Case "/VOTO"
                If notNullArguments Then
                    Call WriteGuildVote(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /voto NICKNAME.")
                End If
               
            Case "/PENAS"
                If notNullArguments Then
                    Call WritePunishments(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /penas NICKNAME.")
                End If
                
            Case "/CONTRASEÑA"
                Call frmNewPassword.Show(vbModal, frmMain)
            
            Case "/APOSTAR"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteGamble(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Cantidad incorrecta. Utilice /apostar CANTIDAD.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /apostar CANTIDAD.")
                End If
                
            Case "/RETIRARFACCION"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                
                Call WriteLeaveFaction
                
            Case "/RETIRAR"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                
                If notNullArguments Then
                    ' Version con argumentos: BankExtractGold
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                        Call WriteBankExtractGold(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Cantidad incorrecta. Utilice /retirar CANTIDAD.")
                    End If
                End If

            Case "/DEPOSITAR"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If

                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                        Call WriteBankDepositGold(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Cantidad incorecta. Utilice /depositar CANTIDAD.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan paramtetros. Utilice /depositar CANTIDAD.")
                End If
                
            Case "/DENUNCIAR"
                If notNullArguments Then
                    Call WriteDenounce(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Formule su denuncia.")
                End If
                
            Case "/FUNDARCLAN"
                If UserLvl >= 25 Then
                    Call WriteGuildFundate
                Else
                    Call ShowConsoleMsg("Para fundar un clan tenés que ser nivel 25 y tener 90 skills en liderazgo.")
                End If
            
            Case "/FUNDARCLANGM"
                Call WriteGuildFundation(eClanType.ct_GM)
            
            Case "/ECHARPARTY"
                If notNullArguments Then
                    Call WritePartyKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /echarparty NICKNAME.")
                End If
                
            Case "/PARTYLIDER"
                If notNullArguments Then
                    Call WritePartySetLeader(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /partylider NICKNAME.")
                End If
                
            Case "/ACCEPTPARTY"
                If notNullArguments Then
                    Call WritePartyAcceptMember(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /acceptparty NICKNAME.")
                End If

            '
            ' BEGIN GM COMMANDS
            '
            
            Case "/GMSG"
                If notNullArguments Then
                    Call WriteGMMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")
                End If
                
            Case "/SHOWNAME"
                Call WriteShowName
                
            Case "/ONLINEREAL"
                Call WriteOnlineRoyalArmy
                
            Case "/ONLINECAOS"
                Call WriteOnlineChaosLegion
                
            Case "/IRCERCA"
                If notNullArguments Then
                    Call WriteGoNearby(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ircerca NICKNAME.")
                End If
                
            Case "/REM"
                If notNullArguments Then
                    Call WriteComment(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un comentario.")
                End If
            
            Case "/HORA"
                Call Protocol.WriteServerTime
            
            Case "/DONDE"
                If notNullArguments Then
                    Call WriteWhere(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /donde NICKNAME.")
                End If
                
            Case "/NENE"
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteCreaturesInMap(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Mapa incorrecto. Utilice /nene MAPA.")
                    End If
                Else
                    'Por default, toma el mapa en el que esta
                    Call WriteCreaturesInMap(UserMap)
                End If
                
            Case "/TELEPLOC"
                Call WriteWarpMeToTarget
                
            Case "/TELEP"
                If notNullArguments And CantidadArgumentos >= 4 Then
                    If ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                        Call WriteWarpChar(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /telep NICKNAME MAPA X Y.")
                    End If
                ElseIf CantidadArgumentos = 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        'Por defecto, si no se indica el nombre, se teletransporta el mismo usuario
                        Call WriteWarpChar("YO", ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    ElseIf ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        'Por defecto, si no se indica el mapa, se teletransporta al mismo donde esta el usuario
                        Call WriteWarpChar(ArgumentosAll(0), UserMap, ArgumentosAll(1), ArgumentosAll(2))
                    Else
                        'No uso ningun formato por defecto
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /telep NICKNAME MAPA X Y.")
                    End If
                ElseIf CantidadArgumentos = 2 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) Then
                        ' Por defecto, se considera que se quiere unicamente cambiar las coordenadas del usuario, en el mismo mapa
                        Call WriteWarpChar("YO", UserMap, ArgumentosAll(0), ArgumentosAll(1))
                    Else
                        'No uso ningun formato por defecto
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /telep NICKNAME MAPA X Y.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /telep NICKNAME MAPA X Y.")
                End If
                
            Case "/SILENCIAR"
                If notNullArguments Then
                    Call WriteSilence(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /silenciar NICKNAME.")
                End If
                
            Case "/SHOW"
                If notNullArguments Then
                    Select Case UCase$(ArgumentosAll(0))
                        Case "SOS"
                            Call WriteSOSShowList
                            
                        Case "INT"
                            Call WriteShowServerForm
                            
                    End Select
                End If
                
            Case "/IRA"
                If notNullArguments Then
                    Call WriteGoToChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ira NICKNAME.")
                End If
        
            Case "/INVISIBLE"
                Call WriteInvisible
                
            Case "/PANELGM"
                Call WriteGMPanel
                
            Case "/TRABAJANDO"
                Call WriteWorking
                
            Case "/OCULTANDO"
                Call WriteHiding
                
            Case "/CARCEL"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@")
                    If UBound(tmpArr) = 2 Then
                        If ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Then
                            Call WriteJail(tmpArr(0), tmpArr(1), tmpArr(2))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Tiempo incorrecto. Utilice /carcel NICKNAME@MOTIVO@TIEMPO.")
                        End If
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /carcel NICKNAME@MOTIVO@TIEMPO.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /carcel NICKNAME@MOTIVO@TIEMPO.")
                End If
                
            Case "/RMATA"
                Call WriteKillNPC
                
            Case "/ADVERTENCIA"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        Call WriteWarnUser(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /advertencia NICKNAME@MOTIVO.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /advertencia NICKNAME@MOTIVO.")
                End If
                
            Case "/MOD"
                If notNullArguments And CantidadArgumentos >= 3 Then
                    Select Case UCase$(ArgumentosAll(1))
                        Case "BODY"
                            tmpInt = eEditOptions.eo_Body
                        
                        Case "HEAD"
                            tmpInt = eEditOptions.eo_Head
                        
                        Case "ORO"
                            tmpInt = eEditOptions.eo_Gold
                        
                        Case "LEVEL"
                            tmpInt = eEditOptions.eo_Level
                        
                        Case "SKILLS"
                            tmpInt = eEditOptions.eo_Skills
                        
                        Case "SKILLSLIBRES"
                            tmpInt = eEditOptions.eo_SkillPointsLeft
                        
                        Case "CLASE"
                            tmpInt = eEditOptions.eo_Class
                        
                        Case "EXP"
                            tmpInt = eEditOptions.eo_Experience
                        
                        Case "CRI"
                            tmpInt = eEditOptions.eo_CriminalsKilled
                        
                        Case "CIU"
                            tmpInt = eEditOptions.eo_CiticensKilled
                        
                        Case "NOB"
                            tmpInt = eEditOptions.eo_Nobleza
                        
                        Case "ASE"
                            tmpInt = eEditOptions.eo_Asesino
                        
                        Case "SEX"
                            tmpInt = eEditOptions.eo_Sex
                            
                        Case "RAZA"
                            tmpInt = eEditOptions.eo_Raza
                        
                        Case "AGREGAR"
                            tmpInt = eEditOptions.eo_addGold
                        
                        Case Else
                            tmpInt = -1
                    End Select
                    
                    If tmpInt > 0 Then
                        If CantidadArgumentos = 3 Then
                            Call WriteEditChar(ArgumentosAll(0), tmpInt, ArgumentosAll(2), "")
                        Else
                            Call WriteEditChar(ArgumentosAll(0), tmpInt, ArgumentosAll(2), ArgumentosAll(3))
                        End If
                    Else
                        'Avisar que no exite el comando
                        Call ShowConsoleMsg("Comando incorrecto.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros.")
                End If
            
            Case "/INFO"
                If notNullArguments Then
                    Call WriteRequestCharInfo(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /info NICKNAME.")
                End If
                
            Case "/STAT"
                If notNullArguments Then
                    Call WriteRequestCharStats(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /stat NICKNAME.")
                End If
                
            Case "/BAL"
                If notNullArguments Then
                    Call WriteRequestCharGold(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /bal NICKNAME.")
                End If
                
            Case "/INV"
                If notNullArguments Then
                    Call WriteRequestCharInventory(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /inv NICKNAME.")
                End If
                
            Case "/BOV"
                If notNullArguments Then
                    Call WriteRequestCharBank(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /bov NICKNAME.")
                End If
                
            Case "/SKILLS"
                If notNullArguments Then
                    Call WriteRequestCharSkills(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /skills NICKNAME.")
                End If
                
            Case "/REVIVIR"
                If notNullArguments Then
                    Call WriteReviveChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /revivir NICKNAME.")
                End If
                
            Case "/ONLINEGM"
                Call WriteOnlineGM
                
            Case "/ONLINEMAP"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteOnlineMap(ArgumentosAll(0))
                    Else
                        Call ShowConsoleMsg("Mapa incorrecto.")
                    End If
                Else
                    Call WriteOnlineMap(UserMap)
                End If
                
            Case "/PERDON"
                If notNullArguments Then
                    Call WriteForgive(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /perdon NICKNAME.")
                End If
                
            Case "/ECHAR"
                If notNullArguments Then
                    Call WriteKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /echar NICKNAME.")
                End If
                
            Case "/EJECUTAR"
                If notNullArguments Then
                    Call WriteExecute(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ejecutar NICKNAME.")
                End If
                
            Case "/BAN"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        Call WriteBanChar(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /ban NICKNAME@MOTIVO.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ban NICKNAME@MOTIVO.")
                End If
                
            Case "/UNBAN"
                If notNullArguments Then
                    Call WriteUnbanChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /unban NICKNAME.")
                End If
                
            Case "/SEGUIR"
                Call WriteNPCFollow
                
            Case "/SUM"
                If notNullArguments Then
                    Call WriteSummonChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /sum NICKNAME.")
                End If
                
            Case "/CC"
                Call WriteSpawnListRequest
                
            Case "/RESETINV"
                Call WriteResetNPCInventory
                
            Case "/LIMPIAR"
                Call WriteCleanWorld
                
            Case "/RMSG"
                If notNullArguments Then
                    Call WriteServerMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")
                End If
                
            Case "/NICK2IP"
                If notNullArguments Then
                    Call WriteNickToIP(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /nick2ip NICKNAME.")
                End If
                
            Case "/IP2NICK"
                If notNullArguments Then
                    If validipv4str(ArgumentosRaw) Then
                        Call WriteIPToNick(str2ipv4l(ArgumentosRaw))
                    Else
                        'No es una IP
                        Call ShowConsoleMsg("IP incorrecta. Utilice /ip2nick IP.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ip2nick IP.")
                End If
                
            Case "/ONCLAN"
                If notNullArguments Then
                    Call WriteGuildOnlineMembers(ArgumentosRaw)
                Else
                    'Avisar sintaxis incorrecta
                    Call ShowConsoleMsg("Utilice /onclan nombre del clan.")
                End If
                
            Case "/CT"
                If notNullArguments And CantidadArgumentos >= 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And _
                        ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        
                        If CantidadArgumentos = 3 Then
                            Call WriteTeleportCreate(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                        Else
                            If ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                                Call WriteTeleportCreate(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                            Else
                                'No es numerico
                                Call ShowConsoleMsg("Valor incorrecto. Utilice /ct MAPA X Y RADIO(Opcional).")
                            End If
                        End If
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /ct MAPA X Y RADIO(Opcional).")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ct MAPA X Y RADIO(Opcional).")
                End If
                
            Case "/DT"
                Call WriteTeleportDestroy
                
            Case "/LLUVIA"
                Call WriteRainToggle
                
            Case "/SETDESC"
                Call WriteSetCharDescription(ArgumentosRaw)
            
            Case "/FORCEMIDIMAP"
                If notNullArguments Then
                    'elegir el mapa es opcional
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                            'eviamos un mapa nulo para que tome el del usuario.
                            Call WriteForceMIDIToMap(ArgumentosAll(0), 0)
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Midi incorrecto. Utilice /forcemidimap MIDI MAPA, siendo el mapa opcional.")
                        End If
                    Else
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                            Call WriteForceMIDIToMap(ArgumentosAll(0), ArgumentosAll(1))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Valor incorrecto. Utilice /forcemidimap MIDI MAPA, siendo el mapa opcional.")
                        End If
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Utilice /forcemidimap MIDI MAPA, siendo el mapa opcional.")
                End If
                
            Case "/FORCEWAVMAP"
                If notNullArguments Then
                    'elegir la posicion es opcional
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                            'eviamos una posicion nula para que tome la del usuario.
                            Call WriteForceWAVEToMap(ArgumentosAll(0), 0, 0, 0)
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                        End If
                    ElseIf CantidadArgumentos = 4 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                            Call WriteForceWAVEToMap(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                        End If
                    Else
                        'Avisar que falta el parametro
                        Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                End If
                
            Case "/REALMSG"
                If notNullArguments Then
                    Call WriteRoyalArmyMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")
                End If
                 
            Case "/CAOSMSG"
                If notNullArguments Then
                    Call WriteChaosLegionMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")
                End If
                
            Case "/CIUMSG"
                If notNullArguments Then
                    Call WriteCitizenMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")
                End If
            
            Case "/CRIMSG"
                If notNullArguments Then
                    Call WriteCriminalMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")
                End If
            
            Case "/TALKAS"
                If notNullArguments Then
                    Call WriteTalkAsNPC(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")
                End If
        
            Case "/MASSDEST"
                Call WriteDestroyAllItemsInArea
    
            Case "/ACEPTCONSE"
                If notNullArguments Then
                    Call WriteAcceptRoyalCouncilMember(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /aceptconse NICKNAME.")
                End If
                
            Case "/ACEPTCONSECAOS"
                If notNullArguments Then
                    Call WriteAcceptChaosCouncilMember(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /aceptconsecaos NICKNAME.")
                End If
                
            Case "/PISO"
                Call WriteItemsInTheFloor
                
            Case "/ESTUPIDO"
                If notNullArguments Then
                    Call WriteMakeDumb(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /estupido NICKNAME.")
                End If
                
            Case "/NOESTUPIDO"
                If notNullArguments Then
                    Call WriteMakeDumbNoMore(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /noestupido NICKNAME.")
                End If
                
            Case "/DUMPSECURITY"
                Call WriteDumpIPTables
                
            Case "/KICKCONSE"
                If notNullArguments Then
                    Call WriteCouncilKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /kickconse NICKNAME.")
                End If
                
            Case "/TRIGGER"
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Trigger) Then
                        Call WriteSetTrigger(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Numero incorrecto. Utilice /trigger NUMERO.")
                    End If
                Else
                    'Version sin parametro
                    Call WriteAskTrigger
                End If
                
            Case "/BANIPLIST"
                Call WriteBannedIPList
                
            Case "/BANIPRELOAD"
                Call WriteBannedIPReload
                
            Case "/MIEMBROSCLAN"
                If notNullArguments Then
                    Call WriteGuildMemberList(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /miembrosclan GUILDNAME.")
                End If
                
            Case "/BANCLAN"
                If notNullArguments Then
                    Call WriteGuildBan(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /banclan GUILDNAME.")
                End If
                
            Case "/BANIP"
                If CantidadArgumentos >= 2 Then
                    If validipv4str(ArgumentosAll(0)) Then
                        Call WriteBanIP(True, str2ipv4l(ArgumentosAll(0)), vbNullString, Right$(ArgumentosRaw, Len(ArgumentosRaw) - Len(ArgumentosAll(0)) - 1))
                    Else
                        'No es una IP, es un nick
                        Call WriteBanIP(False, str2ipv4l("0.0.0.0"), ArgumentosAll(0), Right$(ArgumentosRaw, Len(ArgumentosRaw) - Len(ArgumentosAll(0)) - 1))
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /banip IP motivo o /banip nick motivo.")
                End If
                
            Case "/UNBANIP"
                If notNullArguments Then
                    If validipv4str(ArgumentosRaw) Then
                        Call WriteUnbanIP(str2ipv4l(ArgumentosRaw))
                    Else
                        'No es una IP
                        Call ShowConsoleMsg("IP incorrecta. Utilice /unbanip IP.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /unbanip IP.")
                End If
                
            Case "/CI"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Long) Then
                        Call WriteCreateItem(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Objeto incorrecto. Utilice /ci OBJETO.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ci OBJETO.")
                End If
                
            Case "/DEST"
                Call WriteDestroyItems
                
            Case "/NOCAOS"
                If notNullArguments Then
                    Call WriteChaosLegionKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /nocaos NICKNAME.")
                End If
    
            Case "/NOREAL"
                If notNullArguments Then
                    Call WriteRoyalArmyKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /noreal NICKNAME.")
                End If
    
            Case "/FORCEMIDI"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteForceMIDIAll(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Midi incorrecto. Utilice /forcemidi MIDI.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /forcemidi MIDI.")
                End If
    
            Case "/FORCEWAV"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteForceWAVEAll(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Wav incorrecto. Utilice /forcewav WAV.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /forcewav WAV.")
                End If
                
            Case "/BORRARPENA"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 3)
                    If UBound(tmpArr) = 2 Then
                        Call WriteRemovePunishment(tmpArr(0), tmpArr(1), tmpArr(2))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /borrarpena NICK@PENA@NuevaPena.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /borrarpena NICK@PENA@NuevaPena.")
                End If
                
            Case "/BLOQ"
                Call WriteTileBlockedToggle
                
            Case "/MATA"
                Call WriteKillNPCNoRespawn
        
            Case "/MASSKILL"
                Call WriteKillAllNearbyNPCs
                
            Case "/LASTIP"
                If notNullArguments Then
                    Call WriteLastIP(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /lastip NICKNAME.")
                End If
    
            Case "/MOTDCAMBIA"
                Call WriteChangeMOTD
                
            Case "/SMSG"
                If notNullArguments Then
                    Call WriteSystemMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")
                End If
                
            Case "/ACC"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteCreateNPC(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Npc incorrecto. Utilice /acc NPC.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /acc NPC.")
                End If
                
            Case "/RACC"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteCreateNPCWithRespawn(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Npc incorrecto. Utilice /racc NPC.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /racc NPC.")
                End If
        
            Case "/AI" ' 1 - 4
                If notNullArguments And CantidadArgumentos >= 2 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                        Call WriteImperialArmour(ArgumentosAll(0), ArgumentosAll(1))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /ai ARMADURA OBJETO.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ai ARMADURA OBJETO.")
                End If
                
            Case "/AC" ' 1 - 4
                If notNullArguments And CantidadArgumentos >= 2 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                        Call WriteChaosArmour(ArgumentosAll(0), ArgumentosAll(1))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /ac ARMADURA OBJETO.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ac ARMADURA OBJETO.")
                End If
                
            Case "/NAVE"
                Call WriteNavigateToggle
        
            Case "/HABILITAR"
                Call WriteServerOpenToUsersToggle
            
            Case "/APAGAR"
                Call WriteTurnOffServer
                
            Case "/CONDEN"
                If notNullArguments Then
                    Call WriteTurnCriminal(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /conden NICKNAME.")
                End If
                
            Case "/RAJAR"
                If notNullArguments Then
                    Call WriteResetFactions(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /rajar NICKNAME.")
                End If
                
            Case "/RAJARCLAN"
                If notNullArguments Then
                    Call WriteRemoveCharFromGuild(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /rajarclan NICKNAME.")
                End If
                
            Case "/LASTEMAIL"
                If notNullArguments Then
                    Call WriteRequestCharMail(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /lastemail NICKNAME.")
                End If
                
            Case "/APASS"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        Call WriteAlterPassword(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /apass PJSINPASS@PJCONPASS.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /apass PJSINPASS@PJCONPASS.")
                End If
                
            Case "/AEMAIL"
                If notNullArguments Then
                    tmpArr = AEMAILSplit(ArgumentosRaw)
                    If LenB(tmpArr(0)) = 0 Then
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /aemail NICKNAME-NUEVOMAIL.")
                    Else
                        Call WriteAlterMail(tmpArr(0), tmpArr(1))
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /aemail NICKNAME-NUEVOMAIL.")
                End If
                
            Case "/ANAME"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        Call WriteAlterName(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /aname ORIGEN@DESTINO.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /aname ORIGEN@DESTINO.")
                End If
                
            Case "/SLOT"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        If ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Then
                            Call WriteCheckSlot(tmpArr(0), tmpArr(1))
                        Else
                            'Faltan o sobran los parametros con el formato propio
                            Call ShowConsoleMsg("Formato incorrecto. Utilice /slot NICK@SLOT.")
                        End If
                    Else
                        'Faltan o sobran los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /slot NICK@SLOT.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /slot NICK@SLOT.")
                End If
                
            Case "/CENTINELAACTIVADO"
                Call WriteToggleCentinelActivated
                
            Case "/DOBACKUP"
                Call WriteDoBackup
                
            Case "/SHOWCMSG"
                If notNullArguments Then
                    Call WriteShowGuildMessages(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /showcmsg GUILDNAME.")
                End If
                
            Case "/GUARDAMAPA"
                Call WriteSaveMap
                
            Case "/MODMAPINFO" ' PK, BACKUP
                If CantidadArgumentos > 1 Then
                    Select Case UCase$(ArgumentosAll(0))
                        Case "PK" ' "/MODMAPINFO PK"
                            Call WriteChangeMapInfoPK(ArgumentosAll(1) = "1")
                        
                        Case "BACKUP" ' "/MODMAPINFO BACKUP"
                            Call WriteChangeMapInfoBackup(ArgumentosAll(1) = "1")
                        
                        Case "RESTRINGIR" '/MODMAPINFO RESTRINGIR
                            Call WriteChangeMapInfoRestricted(ArgumentosAll(1))
                        
                        Case "MAGIASINEFECTO" '/MODMAPINFO MAGIASINEFECTO
                            Call WriteChangeMapInfoNoMagic(ArgumentosAll(1))
                        
                        Case "INVISINEFECTO" '/MODMAPINFO INVISINEFECTO
                            Call WriteChangeMapInfoNoInvi(ArgumentosAll(1))
                        
                        Case "RESUSINEFECTO" '/MODMAPINFO RESUSINEFECTO
                            Call WriteChangeMapInfoNoResu(ArgumentosAll(1))
                        
                        Case "TERRENO" '/MODMAPINFO TERRENO
                            Call WriteChangeMapInfoLand(ArgumentosAll(1))
                        
                        Case "ZONA" '/MODMAPINFO ZONA
                            Call WriteChangeMapInfoZone(ArgumentosAll(1))
                    End Select
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parametros. Opciones: PK, BACKUP, RESTRINGIR, MAGIASINEFECTO, INVISINEFECTO, RESUSINEFECTO, TERRENO, ZONA")
                End If
                
            Case "/GRABAR"
                Call WriteSaveChars
                
            Case "/BORRAR"
                If notNullArguments Then
                    Select Case UCase(ArgumentosAll(0))
                        Case "SOS" ' "/BORRAR SOS"
                            Call WriteCleanSOS
                            
                    End Select
                End If
                
            Case "/NOCHE"
                Call WriteNight
                
            Case "/ECHARTODOSPJS"
                Call WriteKickAllChars
                
            Case "/RELOADNPCS"
                Call WriteReloadNPCs
                
            Case "/RELOADSINI"
                Call WriteReloadServerIni
                
            Case "/RELOADHECHIZOS"
                Call WriteReloadSpells
                
            Case "/RELOADOBJ"
                Call WriteReloadObjects
                 
            Case "/REINICIAR"
                Call WriteRestart
                
            Case "/AUTOUPDATE"
                Call WriteResetAutoUpdate
            
            Case "/CHATCOLOR"
                If notNullArguments And CantidadArgumentos >= 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        Call WriteChatColor(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /chatcolor R G B.")
                    End If
                ElseIf Not notNullArguments Then    'Go back to default!
                    Call WriteChatColor(0, 255, 0)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /chatcolor R G B.")
                End If
            
            Case "/IGNORADO"
                Call WriteIgnored
            
            Case "/PING"
                Call WritePing
                
            Case "/SETINIVAR"
                If CantidadArgumentos = 3 Then
                    ArgumentosAll(2) = Replace(ArgumentosAll(2), "+", " ")
                    Call WriteSetIniVar(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                Else
                    Call ShowConsoleMsg("Prámetros incorrectos. Utilice /SETINIVAR LLAVE CLAVE VALOR")
                End If
            
            Case "/HOGAR"
                Call WriteHome
                
#If SeguridadAlkon Then
            Case Else
                Call ParseUserCommandEx(Comando, CantidadArgumentos, ArgumentosAll, ArgumentosRaw)
#End If
        End Select
        
    ElseIf Left$(Comando, 1) = "\" Then
        If UserEstado = 1 Then 'Muerto
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        ' Mensaje Privado
        Call AuxWriteWhisper(mid$(Comando, 2), ArgumentosRaw)
        
    ElseIf Left$(Comando, 1) = "-" Then
        If UserEstado = 1 Then 'Muerto
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        ' Gritar
        Call WriteYell(mid$(RawCommand, 2))
        
    Else
        ' Hablar
        Call WriteTalk(RawCommand)
    End If
End Sub

''
' Show a console message.
'
' @param    Message The message to be written.
' @param    red Sets the font red color.
' @param    green Sets the font green color.
' @param    blue Sets the font blue color.
' @param    bold Sets the font bold style.
' @param    italic Sets the font italic style.

Public Sub ShowConsoleMsg(ByVal Message As String, Optional ByVal red As Integer = 255, Optional ByVal green As Integer = 255, Optional ByVal blue As Integer = 255, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/03/07
'
'***************************************************
    Call AddtoRichTextBox(frmMain.RecTxt, Message, red, green, blue, bold, italic)
End Sub

''
' Returns whether the number is correct.
'
' @param    Numero The number to be checked.
' @param    Tipo The acceptable type of number.

Public Function ValidNumber(ByVal Numero As String, ByVal TIPO As eNumber_Types) As Boolean
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/06/07
'
'***************************************************
    Dim Minimo As Long
    Dim Maximo As Long
    
    If Not IsNumeric(Numero) Then _
        Exit Function
    
    Select Case TIPO
        Case eNumber_Types.ent_Byte
            Minimo = 0
            Maximo = 255

        Case eNumber_Types.ent_Integer
            Minimo = -32768
            Maximo = 32767

        Case eNumber_Types.ent_Long
            Minimo = -2147483648#
            Maximo = 2147483647
        
        Case eNumber_Types.ent_Trigger
            Minimo = 0
            Maximo = 6
    End Select
    
    If Val(Numero) >= Minimo And Val(Numero) <= Maximo Then _
        ValidNumber = True
End Function

''
' Returns whether the ip format is correct.
'
' @param    IP The ip to be checked.

Private Function validipv4str(ByVal Ip As String) As Boolean
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/06/07
'
'***************************************************
    Dim tmpArr() As String
    
    tmpArr = Split(Ip, ".")
    
    If UBound(tmpArr) <> 3 Then _
        Exit Function

    If Not ValidNumber(tmpArr(0), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(3), eNumber_Types.ent_Byte) Then _
        Exit Function
    
    validipv4str = True
End Function

''
' Converts a string into the correct ip format.
'
' @param    IP The ip to be converted.

Private Function str2ipv4l(ByVal Ip As String) As Byte()
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/26/07
'Last Modified By: Rapsodius
'Specify Return Type as Array of Bytes
'Otherwise, the default is a Variant or Array of Variants, that slows down
'the function
'***************************************************
    Dim tmpArr() As String
    Dim bArr(3) As Byte
    
    tmpArr = Split(Ip, ".")
    
    bArr(0) = CByte(tmpArr(0))
    bArr(1) = CByte(tmpArr(1))
    bArr(2) = CByte(tmpArr(2))
    bArr(3) = CByte(tmpArr(3))

    str2ipv4l = bArr
End Function

''
' Do an Split() in the /AEMAIL in onother way
'
' @param text All the comand without the /aemail
' @return An bidimensional array with user and mail

Private Function AEMAILSplit(ByRef Text As String) As String()
'***************************************************
'Author: Lucas Tavolaro Ortuz (Tavo)
'Useful for AEMAIL BUG FIX
'Last Modification: 07/26/07
'Last Modified By: Rapsodius
'Specify Return Type as Array of Strings
'Otherwise, the default is a Variant or Array of Variants, that slows down
'the function
'***************************************************
    Dim tmpArr(0 To 1) As String
    Dim Pos As Byte
    
    Pos = InStr(1, Text, "-")
    
    If Pos <> 0 Then
        tmpArr(0) = mid$(Text, 1, Pos - 1)
        tmpArr(1) = mid$(Text, Pos + 1)
    Else
        tmpArr(0) = vbNullString
    End If
    
    AEMAILSplit = tmpArr
End Function
