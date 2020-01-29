Attribute VB_Name = "mPooChar"
'---------------------------------------------------------------------------------------
' Module    : Mod_PooChar
' Author    :  Miqueas
' Date      : 02/02/2014
' Purpose   :  xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'---------------------------------------------------------------------------------------

Option Explicit
 
Public Sub Char_Erase(ByVal CharIndex As Integer)
        '*****************************************************************
        'Erases a character from CharList and map
        '*****************************************************************
 
        With charlist(CharIndex)
        
                If (CharIndex = 0) Then Exit Sub
                If (CharIndex > LastChar) Then Exit Sub
                
                If Map_InBounds(.Pos.X, .Pos.Y) Then  '// Posicion valida
                        MapData(.Pos.X, .Pos.Y).CharIndex = 0  '// Borramos el user
                End If
       
                'Update lastchar
                If CharIndex = LastChar Then
 
                        Do Until charlist(LastChar).Heading > 0
               
                                LastChar = LastChar - 1
 
                                If LastChar = 0 Then
                                
                                        NumChars = 0

                                        Exit Sub

                                End If
                       
                        Loop
 
                End If
   
                Call Char_ResetInfo(CharIndex)
                
                'Remove char's dialog
                Call Dialogos.RemoveDialog(CharIndex)
                
                'Remove particles
                Call Char_Particle_Group_Remove_All(CharIndex)
                
                'Update NumChars
                NumChars = NumChars - 1
 
                Exit Sub
 
        End With
 
End Sub
 
Private Sub Char_ResetInfo(ByVal CharIndex As Integer)

        '*****************************************************************
        'Author: Ao 13.0
        'Last Modify Date: 13/12/2013
        'Reset Info User
        '*****************************************************************

        With charlist(CharIndex)
            Call Delete_All_Auras(CharIndex)
            Call Char_Particle_Group_Remove_All(CharIndex)
            
            .active = 0
            .Criminal = 0
            .FxIndex = 0
            .invisible = False
            
            .Moving = 0
            .muerto = False
            .Nombre = vbNullString
            .pie = False
            .Pos.X = 0
            .Pos.Y = 0
            .UsandoArma = False
            .attacking = False
            
        End With
 
End Sub
 
Private Sub Char_MapPosGet(ByVal CharIndex As Long, ByRef X As Byte, ByRef Y As Byte)
                                
        '*****************************************************************
        'Author: Aaron Perkins
        'Last Modify Date: 13/12/2013
        '// By Miqueas150
        '
        '*****************************************************************
        
        'Make sure it's a legal char_index
      
        With charlist(CharIndex)
                  
                'Get map pos
                X = .Pos.X
                Y = .Pos.Y
        
        End With
 
End Sub
 
Public Sub Char_MapPosSet(ByVal X As Byte, ByVal Y As Byte)

        'Sets the user postion

        If (Map_InBounds(X, Y)) Then  '// Posicion valida
        
                UserPos.X = X
                UserPos.Y = Y
                        
                'Set char
                MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
                charlist(UserCharIndex).Pos = UserPos
        
                Exit Sub
 
        End If

End Sub
 
Public Function Char_Techo() As Boolean

        '// Autor : Marcos Zeni
        '// Nueva forma de establecer si el usuario esta bajo un techo

        Char_Techo = False
 
        With charlist(UserCharIndex)
      
                If (Map_InBounds(.Pos.X, .Pos.Y)) Then '// Posicion valida
                       
                        If (MapData(.Pos.X, .Pos.Y).Trigger = eTrigger.BAJOTECHO Or MapData(.Pos.X, .Pos.Y).Trigger = eTrigger.CASA) Then
                                Char_Techo = True
                        End If
                               
                End If
   
        End With

End Function
 
Public Function Char_MapPosExits(ByVal X As Byte, ByVal Y As Byte) As Integer
 
        '*****************************************************************
        'Checks to see if a tile position has a char_index and return it
        '*****************************************************************
   
        If (Map_InBounds(X, Y)) Then
                Char_MapPosExits = MapData(X, Y).CharIndex
        Else
                Char_MapPosExits = 0
        End If
  
End Function
 
Public Sub Char_UserPos()

        '// Author Miqueas
        '// Actualizamo el lbl de la posicion del usuario
 
        Dim X As Byte
        Dim Y As Byte
     
        If Char_Check(UserCharIndex) Then
        
                '// Damos valor a las variables asi sacamos la pos del usuario.
                Call Char_MapPosGet(UserCharIndex, X, Y)
                
                bTecho = Char_Techo '// Pos : Techo :P
               
                frmMain.Coord.Caption = "Map:" & UserMap & " X:" & X & " Y:" & Y

                Call frmMain.ActualizarMiniMapa
 
                Exit Sub
 
        End If
End Sub
 
Public Sub Char_UserIndexSet(ByVal CharIndex As Integer)
 
        UserCharIndex = CharIndex
 
        With charlist(UserCharIndex)
 
                'Nueva posicion para el usuario.
                UserPos = .Pos
         
                Exit Sub
 
        End With
         
End Sub
 
Public Function Char_Check(ByVal CharIndex As Integer) As Boolean
       
        '**************************************************************
        'Author: Aaron Perkins - Modified by Juan Martin Sotuyo Dodero
        'Last Modify by Miqueas150 Date: 24/02/2013
        'Chequeamos el Char
        '**************************************************************
       
        'check char_index
 
        If CharIndex > 0 And CharIndex <= LastChar Then
 
                With charlist(CharIndex) '// check char_index
                        Char_Check = (.Heading > 0)
                End With
 
        End If
   
End Function
 
Public Sub Char_SetInvisible(ByVal CharIndex As Integer, ByVal Value As Boolean)
       
        '**************************************************************
        'Author: Aaron Perkins - Modified by Juan Martin Sotuyo Dodero
        'Last Modify by Miqueas150 Date: 24/02/2013
 
        '**************************************************************
       
        If Char_Check(CharIndex) Then
 
                With charlist(CharIndex)
 
                        .invisible = Value '// User invisible o no ?
                        
                        Exit Sub
 
                End With
 
        End If
 
End Sub
 
Public Sub Char_SetBody(ByVal CharIndex As Integer, ByVal BodyIndex As Integer)
 
        '**************************************************************
        'Author: Aaron Perkins
        'Last Modify by Miqueas150 Date: 24/02/2013
        'Seteamos el CharBody
        '**************************************************************

        If BodyIndex < LBound(BodyData()) Or BodyIndex > UBound(BodyData()) Then
                charlist(CharIndex).Body = BodyData(0)
                charlist(CharIndex).iBody = 0

                Exit Sub

        End If

        If Char_Check(CharIndex) Then

                With charlist(CharIndex)
               
                        .Body = BodyData(BodyIndex)
                        .iBody = BodyIndex
                        
                        Exit Sub
 
                End With
 
        End If
 
End Sub
 
Public Sub Char_SetHead(ByVal CharIndex As Integer, ByVal HeadIndex As Integer)
 
        '**************************************************************
        'Author: Aaron Perkins
        'Last Modify by Miqueas150 Date: 24/02/2013
        'Seteamos el CharHead
        '**************************************************************
 
        If HeadIndex < LBound(HeadData()) Or HeadIndex > UBound(HeadData()) Then
                charlist(CharIndex).Head = HeadData(0)
                charlist(CharIndex).iHead = 0

                Exit Sub

        End If

        If Char_Check(CharIndex) Then
 
                With charlist(CharIndex)
                        .Head = HeadData(HeadIndex)
                        .iHead = HeadIndex
                               
                        .muerto = (HeadIndex = eCabezas.CASPER_HEAD)
                     
                        Exit Sub
 
                End With
 
        End If
 
End Sub
 
Public Sub Char_SetHeading(ByVal CharIndex As Long, ByVal Heading As Byte)
 
        '**************************************************************
        'Author: Aaron Perkins
        'Last Modify by Miqueas150 Date: 24/02/2013
        'Changes the character heading
        '*****************************************************************
    
        'Make sure it's a legal char_index
 
        If Char_Check(CharIndex) Then
 
                With charlist(CharIndex)
               
                        .Heading = Heading
 
                        Exit Sub
 
                End With
 
        End If
 
End Sub

Public Sub Char_SetName(ByVal CharIndex As Integer, ByVal Name As String)
 
        '**************************************************************
        'Author: Miqueas150
        'Last Modify Date: 04/12/2013
        '
        '**************************************************************
 
        If (Len(Name) = 0) Then

                Exit Sub

        End If

        If Char_Check(CharIndex) Then
 
                With charlist(CharIndex)
               
                        .Nombre = Name
 
                        Exit Sub
 
                End With
 
        End If
 
End Sub
 
Public Sub Char_SetWeapon(ByVal CharIndex As Integer, ByVal WeaponIndex As Integer)
 
        '**************************************************************
        'Author: Aaron Perkins
        'Last Modify by Miqueas150 Date: 24/02/2013
        '
        '**************************************************************
 
        If WeaponIndex > UBound(WeaponAnimData()) Or WeaponIndex < LBound(WeaponAnimData()) Then

                Exit Sub

        End If

        If Char_Check(CharIndex) Then
 
                With charlist(CharIndex)
               
                        .Arma = WeaponAnimData(WeaponIndex)
 
                        Exit Sub
 
                End With
 
        End If
 
End Sub
 
Public Sub Char_SetShield(ByVal CharIndex As Integer, ByVal ShieldIndex As Integer)
 
        '**************************************************************
        'Author: Aaron Perkins
        'Last Modify by Miqueas150 Date: 24/02/2013
        '
        '**************************************************************
 
        If ShieldIndex > UBound(ShieldAnimData()) Or ShieldIndex < LBound(ShieldAnimData()) Then

                Exit Sub

        End If

        If Char_Check(CharIndex) Then
 
                With charlist(CharIndex)
   
                        .Escudo = ShieldAnimData(ShieldIndex)
                        
                        Exit Sub
 
                End With
 
        End If
 
End Sub
 
Public Sub Char_SetCasco(ByVal CharIndex As Integer, ByVal CascoIndex As Integer)
 
        '**************************************************************
        'Author: Aaron Perkins
        'Last Modify by Miqueas150 Date: 24/02/2013
        '
        '**************************************************************
 
        If CascoIndex > UBound(CascoAnimData()) Or CascoIndex < LBound(CascoAnimData()) Then

                Exit Sub

        End If

        If Char_Check(CharIndex) Then
 
                With charlist(CharIndex)
               
                        .Casco = CascoAnimData(CascoIndex)
 
                        Exit Sub
 
                End With
 
        End If
     
End Sub
 
Public Sub Char_SetFx(ByVal CharIndex As Integer, _
                      ByVal fX As Integer, _
                      ByVal Loops As Integer)
 
        '***************************************************
        'Author: Juan Martin Sotuyo Dodero (Maraxus)
        'Last Modify Date: 12/03/04
        'Sets an FX to the character.
        '***************************************************
  
        If (Char_Check(CharIndex)) Then
        
                With charlist(CharIndex)

                        .FxIndex = fX
        
                        If .FxIndex > 0 Then
                        
                                Call InitGrh(.fX, FxData(fX).Animacion)
                                .fX.Loops = Loops
                                
                        End If

                End With
        
        End If
   
End Sub
 
Public Sub Char_Make(ByVal CharIndex As Integer, _
                     ByVal Body As Integer, _
                     ByVal Head As Integer, _
                     ByVal Heading As Byte, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal Arma As Integer, _
                     ByVal Escudo As Integer, _
                     ByVal Casco As Integer)
 
        'Apuntamos al ultimo Char
 
        If CharIndex > LastChar Then
                LastChar = CharIndex
        End If
 
        NumChars = NumChars + 1

        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        With charlist(CharIndex)
       
                'If the char wasn't allready active (we are rewritting it) don't increase char count
                
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
                
                'attack state
                .attacking = False
       
                'Update position
                .Pos.X = X
                .Pos.Y = Y
           
        End With
   
        'Plot on map
        MapData(X, Y).CharIndex = CharIndex
       
End Sub

Public Sub Char_RefreshAll()
        '*****************************************************************
        'Goes through the charlist and replots all the characters on the map
        'Used to make sure everyone is visible
        '*****************************************************************
 
        Dim LoopC As Long
   
        For LoopC = 1 To LastChar
 
                If (Char_Check(LoopC)) Then  '// Char valido

                        With charlist(LoopC)

                                If (Map_InBounds(.Pos.X, .Pos.Y)) Then
                                        MapData(.Pos.X, .Pos.Y).CharIndex = LoopC  '// Ahora si refrescamos sin error alguno :3
                                End If

                        End With

                End If

        Next LoopC
 
End Sub

Sub Char_MovebyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)

        Dim X        As Integer
        Dim Y        As Integer
        
        Dim addx     As Integer
        Dim addy     As Integer
        
        Dim nHeading As E_Heading
    
        If (CharIndex <= 0) Then
                Exit Sub
    
        End If
    
        With charlist(CharIndex)
                X = .Pos.X
                Y = .Pos.Y
                
                '// Miqueas : Agrego este parchesito para evitar un run time
                
                If Not (Map_InBounds(X, Y)) Then
                        Exit Sub
                End If

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

Sub Char_MoveScreen(ByVal nHeading As E_Heading)
        '******************************************
        'Starts the screen moving in a direction
        '******************************************

        Dim X  As Integer
        Dim Y  As Integer
        
        Dim TX As Integer
        Dim TY As Integer
    
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
        TX = UserPos.X + X
        TY = UserPos.Y + Y

        'Check to see if its out of bounds

        If (TX < MinXBorder) Or (TX > MaxXBorder) Or (TY < MinYBorder) Or (TY > MaxYBorder) Then

                Exit Sub

        Else
                'Start moving... MainLoop does the rest
                AddtoUserPos.X = X
                UserPos.X = TX
                AddtoUserPos.Y = Y
                UserPos.Y = TY
                UserMoving = 1
                
                bTecho = Char_Techo
               
        End If

End Sub

Sub Char_MovebyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
        '*****************************************************************
        'Starts the movement of a character in nHeading direction
        '*****************************************************************

        Dim addx As Integer
        Dim addy As Integer
        
        Dim X    As Integer
        Dim Y    As Integer
        
        Dim nX   As Integer
        Dim nY   As Integer
    
        If (CharIndex <= 0) Then
                Exit Sub
    
        End If

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
                
                '// Miqueas : Agrego este parchesito para evitar un run time
               
                If Not (Map_InBounds(nX, nY)) Then
                        Exit Sub

                End If

                MapData(nX, nY).CharIndex = CharIndex
                .Pos.X = nX
                .Pos.Y = nY
                MapData(X, Y).CharIndex = 0
         
                .MoveOffsetX = -1 * (TilePixelWidth * addx)
                .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
                .Moving = 1
                .Heading = nHeading
        
                .scrollDirectionX = addx
                .scrollDirectionY = addy
        End With
    
        If (UserEstado = 0) Then
                Call DoPasosFx(CharIndex)
        End If
      
        'areas viejos

        If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
                If CharIndex <> UserCharIndex Then
                        Call Char_Erase(CharIndex)
                End If
        End If

End Sub

Sub Char_CleanAll()
    '// Borramos los obj y char que esten

    Dim X         As Long, Y As Long
    Dim CharIndex As Integer, Obj As Integer
    
    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
          
            'Erase NPCs
            CharIndex = Char_MapPosExits(CByte(X), CByte(Y))
 
            If (CharIndex > 0) Then
                Call Char_Erase(CharIndex)
            End If
                        
            'Erase OBJs
            Obj = Map_PosExitsObject(CByte(X), CByte(Y))

            If (Obj > 0) Then
                Call Map_DestroyObject(CByte(X), CByte(Y))
            End If

        Next Y
    Next X

End Sub

