Attribute VB_Name = "ProtocolCmdParse"
'Argentum Online
'
'Copyright (C) 2006 Juan Martín Sotuyo Dodero (Maraxus)
'Copyright (C) 2006 Alejandro Santos (AlejoLp)

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

Option Explicit

''
' Interpreta, valida y ejecuta el comando ingresado .
'
' @param    RawCommand El comando en version String
' @remarks  None Known.

Public Sub ParseUserCommand(ByVal RawCommand As String)
'***************************************************
'Autor: Alejandro Santos (AlejoLp)
'Last Modification: 12/20/06
'Interpreta, valida y ejecuta el comando ingresado
'***************************************************

Dim TmpArgos() As String

Dim Comando As String
Dim ArgumentosAll() As String
Dim ArgumentosRaw As String
Dim Argumentos2() As String
Dim Argumentos3() As String
Dim Argumentos4() As String
Dim CantidadArgumentos As Long

' TmpArgos: Un array de a lo sumo dos elementos,
' el primero es el comando (hasta el primer espacio)
' y el segundo elemento es el resto. Si no hay argumentos
' devuelve un array de un solo elemento
TmpArgos = Split(RawCommand, " ", 2)

Comando = Trim$(UCase$(TmpArgos(0)))

If UBound(TmpArgos) > 0 Then
    ' El string en crudo que este despues del primer espacio
    ArgumentosRaw = TmpArgos(1)
    
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

If Comando = "" Then _
    Exit Sub
    
If Left$(Comando, 1) = "/" Then
    ' Comando normal
    
    Select Case Comando
        Case "/SEG"
            Call WriteSafeToggle
            
        Case "/ONLINE"
            Call WriteOnline
            
        Case "/SALIR"
            Call WriteQuit
            
        Case "/SALIRCLAN"
            Call WriteGuildLeave
            
        Case "/BALANCE"
            Call WriteRequestAccountState
            
        Case "/QUIETO"
            Call WritePetStand
            
        Case "/ACOMPAÑAR"
            Call WritePetFollow
            
        Case "/ENTRENAR"
            Call WriteTrainList
            
        Case "/DESCANSAR"
            Call WriteRest
            
        Case "/MEDITAR"
            Call WriteMeditate
    
        Case "/RESUCITAR"
            Call WriteResucitate
            
        Case "/CURAR"
            Call WriteHeal
            
        Case "/AYUDA"
            Call WriteHelp
            
        Case "/EST"
            Call WriteRequestStats
            
        Case "/COMERCIAR"
            Call WriteCommerceStart
            
        Case "/BOVEDA"
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
            Call WritePartyCreate
            
        Case "/PARTY"
            Call WritePartyJoin
            
        Case "/ENCUESTA"

            If CantidadArgumentos = 0 Then
                ' Version sin argumentos: Inquiry
                Call WriteInquiry
            Else
                ' Version con argumentos: InquiryVote
                If IsNumeric(ArgumentosRaw) Then
                    Call WriteInquiryVote(ArgumentosRaw)
                Else
                    ' TODO: No es numerico
                End If
            End If
    
        Case "/CMSG"
            If CantidadArgumentos > 0 Then
                If IsNumeric(ArgumentosRaw) Then
                    Call WriteGuildMessage(ArgumentosRaw)
                Else
                    ' TODO: No es numerico
                End If
            Else
                ' TODO: Avisar que falta el parametro
            End If
    
        Case "/PMSG"
            If CantidadArgumentos > 0 Then
                Call WritePartyMessage(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
    
        Case "/CENTINELA"
            If CantidadArgumentos > 0 Then
                If IsNumeric(ArgumentosRaw) Then
                    Call WriteCentinelReport(ArgumentosRaw)
                Else
                    ' TODO: No es numerico
                End If
            Else
                ' TODO: Avisar que falta el parametro
            End If
    
        Case "/ONLINECLAN"
            Call WriteGuildOnline
            
        Case "/ONLINEPARTY"
            Call WritePartyOnline
            
        Case "/BMSG"
            If CantidadArgumentos > 0 Then
                Call WriteCouncilMessage(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/ROL"
            If CantidadArgumentos > 0 Then
                Call WriteRoleMasterRequest(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/GM"
            Call WriteGMRequest
            
        Case "/_BUG"
            If CantidadArgumentos > 0 Then
                Call WriteBugReport(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/VOTO"
            If CantidadArgumentos > 0 Then
                Call WriteGuildVote(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
           
        Case "/PENAS"
            If CantidadArgumentos > 0 Then
                Call WritePunishments(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/PASSWD"
            If CantidadArgumentos > 0 Then
                Call WriteChangePassword(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/APOSTAR"
            If CantidadArgumentos > 0 Then
                If IsNumeric(ArgumentosRaw) Then
                    Call WriteGamble(ArgumentosRaw)
                Else
                    ' TODO: No es numerico
                End If
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/RETIRAR"
            If CantidadArgumentos = 0 Then
                ' Version sin argumentos: LeaveFaction
                Call WriteLeaveFaction
            Else
                ' Version con argumentos: BankExtractGold
                If IsNumeric(ArgumentosRaw) Then
                    Call WriteBankExtractGold(ArgumentosRaw)
                Else
                    ' TODO: No es numerico
                End If
            End If

        Case "/DEPOSITAR"
            If CantidadArgumentos > 0 Then
                If IsNumeric(ArgumentosRaw) Then
                    Call WriteBankDepositGold(ArgumentosRaw)
                Else
                    ' TODO: No es numerico
                End If
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/DENUNCIAR"
            Call WriteDenounce ' TODO: No existe la funcion
            
        Case "/FUNDARCLAN"
            frmEligeAlineacion.Show vbModeless, Me
            
        Case "/ECHARPARTY"
            If CantidadArgumentos > 0 Then
                Call WritePartyKick(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/PARTYLIDER"
            If CantidadArgumentos > 0 Then
                Call WritePartySetLeader(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/ACCEPTPARTY"
            If CantidadArgumentos > 0 Then
                Call WritePartyAcceptMember(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
    
        Case "/MIEMBROSCLAN"
            If CantidadArgumentos > 0 Then
                Call WriteGuildMemeberList(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
    
        '
        ' BEGIN GM COMMANDS
        '
        
        Case "/GMSG"
            If CantidadArgumentos > 0 Then
                Call WriteGMMessage(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/SHOWNAME"
            Call WriteShowName
            
        Case "/ONLINEREAL"
            Call WriteOnlineRoyalArmy
            
        Case "/ONLINECAOS"
            Call WriteOnlineChaosLegion
            
        Case "/IRCERCA"
            If CantidadArgumentos > 0 Then
                Call WriteGoNearby(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/REM"
            If CantidadArgumentos > 0 Then
                Call WriteComment(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
    
        Case "/HORA"
            Call WriteTime
            
        Case "/DONDE"
            If CantidadArgumentos > 0 Then
                Call WriteWhere(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/NENE"
            If CantidadArgumentos > 0 Then
                Call WriteCreaturesInMap(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/TELEPLOC"
            Call WriteWarpMeToTarget
            
        Case "/TELEP"
            If CantidadArgumentos >= 4 Then
                If IsNumeric(ArgumentosAll(1)) And IsNumeric(ArgumentosAll(2)) And IsNumeric(ArgumentosAll(3)) Then
                    Call WriteWarpChar(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                Else
                    ' TODO: No es numerico
                End If
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/SILENCIAR"
            If CantidadArgumentos > 0 Then
                Call WriteSilence(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/SHOW"
            If CantidadArgumentos > 0 Then
                Select Case ArgumentosAll(0)
                    Case "SOS"
                        Call WriteSOSShowList
                        
                    Case "INT"
                        Call WriteShowSeverForm
                        
                End Select
            End If
            
        Case "/IRA"
            If CantidadArgumentos > 0 Then
                Call WriteGoToChar(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
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
            If CantidadArgumentos >= 3 Then
                If IsNumeric(ArgumentosAll(2)) Then
                    Call WriteJail(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                Else
                    ' TODO: No es numerico
                End If
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/RMATA"
            Call WriteKillNPC
            
        Case "/ADVERTENCIA"
            If CantidadArgumentos >= 3 Then
                Call WriteWarnUser(ArgumentosAll(0), ArgumentosAll(2))
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/MOD"
            If CantidadArgumentos >= 4 Then
                Call WriteEditChar(ArgumentosAll(0), , ArgumentosAll(2), ArgumentosAll(3))
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/INFO"
            If CantidadArgumentos > 0 Then
                Call WriteRequestCharInfo(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/STAT"
            If CantidadArgumentos > 0 Then
                Call WriteRequestCharStats(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/BAL"
            If CantidadArgumentos > 0 Then
                Call WriteRequestCharGold(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/INV"
            If CantidadArgumentos > 0 Then
                Call WriteRequestCharInventory(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/BOV"
            If CantidadArgumentos > 0 Then
                Call WriteRequestCharBank(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/SKILLS"
            If CantidadArgumentos > 0 Then
                Call WriteRequestCharSkills(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/REVIVIR"
            If CantidadArgumentos > 0 Then
                Call WriteReviveChar(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/ONLINEGM"
            Call WriteOnlineGM
            
        Case "/ONLINEMAP"
            Call WriteOnlineMap
            
        Case "/PERDON"
            If CantidadArgumentos > 0 Then
                Call WriteForgive(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/ECHAR"
            If CantidadArgumentos > 0 Then
                Call WriteKick(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/EJECUTAR"
            If CantidadArgumentos > 0 Then
                Call WriteExecute(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/BAN"
            If CantidadArgumentos > 0 Then
                Call WriteBanChar(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/UNBAN"
            If CantidadArgumentos > 0 Then
                Call WriteUnbanChar(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/SEGUIR"
            Call WriteNPCFollow
            
        Case "/SUM"
            If CantidadArgumentos > 0 Then
                Call WriteSummonChar(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/CC"
            Call WriteSpawnListRequest
            
        Case "/RESETINV"
            Call WriteResetNPCInventory
            
        Case "/LIMPIAR"
            Call WriteCleanWorld
            
        Case "/RMSG"
            If CantidadArgumentos > 0 Then
                Call WriteServerMessage(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/NICK2IP"
            If CantidadArgumentos > 0 Then
                Call WriteNickToIP(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/IP2NICK"
            If CantidadArgumentos > 0 Then
                Call WriteIPToNick(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/ONCLAN"
            Call WriteGuildOnline
            
        Case "/CT"
            If CantidadArgumentos >= 3 Then
                Call WriteTeleportCreate(ArgumentosRaw)
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/DT"
            Call WriteTeleportDestroy
            
        Case "/LLUVIA"
            Call WriteRainToggle
            
        Case "/SETDESC"
            If CantidadArgumentos >= 2 Then
                Call WriteSetCharDescription(ArgumentosAll(0), ArgumentosAll(1))
            Else
                ' TODO: Avisar que falta el parametro
            End If
            
        Case "/FORCEMIDIMAP"
    
        Case "/FORCEWAVMAP"
    
        Case "/REALMSG"
    
        Case "/CAOSMSG"
    
        Case "/CIUMSG"
    
        Case "/TALKAS"
    
        Case "/MASSDEST"
    
        Case "/ACEPTCONSE"
    
        Case "/ACEPTCONSECAOS"
    
        Case "/PISO"
    
        Case "/ESTUPIDO"
    
        Case "/NOESTUPIDO"
    
        Case "/DUMPSECURITY"
    
        Case "/KICKCONSE"
    
        Case "/TRIGGER"
    
        Case "/BANIPLIST"
    
        Case "/BANIPRELOAD"
    
        Case "/MIEMBROSCLAN"
    
        Case "/BANCLAN"
    
        Case "/BANIP"
    
        Case "/UNBANIP"
    
        Case "/CI"
    
        Case "/DEST"
    
        Case "/NOCAOS"
    
        Case "/NOREAL"
    
        Case "/FORCEMIDI"
    
        Case "/FORCEWAV"
    
        Case "/BORRARPENA"
    
        Case "/BLOQ"
    
        Case "/MATA"
    
        Case "/MASSKILL"
    
        Case "/LASTIP"
    
        Case "/MOTDCAMBIA"
    
        Case "/SMSG"
    
        Case "/ACC"
    
        Case "/RACC"
    
        Case "/AI" ' 1 - 4
    
        Case "/AC" ' 1 - 4
    
        Case "/NAVE"
    
        Case "/HABILITAR"
    
        Case "/APAGAR"
    
        Case "/CONDEN"
    
        Case "/RAJAR"
    
        Case "/RAJARCLAN"
    
        Case "/LASTEMAIL"
    
        Case "/APASS"
    
        Case "/AEMAIL"
    
        Case "/ANAME"
    
        Case "/CENTINELAACTIVADO"
    
        Case "/DOBACKUP"
    
        Case "/SHOWCMSG"
    
        Case "/GUARDAMAPA"
    
        Case "/MODMAPINFO" ' PK, BACKUP
    
        Case "/GRABAR"
    
        Case "/BORRAR"
        
            If CantidadArgumentos > 0 Then
                Select Case ArgumentosAll(0)
                    Case "SOS" ' "/BORRAR SOS"
                        
                End Select
            End If
            
        Case "/NOCHE"
    
        Case "/ECHARTODOSPJS"
    
        Case "/TCPESSTATS"
    
        Case "/RELOADNPCS"
    
        Case "/RELOADSINI"
    
        Case "/RELOADHECHIZOS"
    
        Case "/RELOADOBJ"
    
        Case "/REINICIAR"
        
        Case "/AUTOUPDATE"
        
        Case "/CHATCOLOR"
        
        Case "/IGNORADO"
        
    End Select
    
ElseIf Left$(Comando, 1) = "\" Then
    ' Mensaje Privado
    
ElseIf Left$(Comando, 1) = "-" Then
    ' Gritar
    
Else
    ' Hablar
    
End If

End Sub
