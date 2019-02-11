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
    
    On Error GoTo Char_Erase_Err
    
 
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
                
        'Update NumChars
        NumChars = NumChars - 1
 
        Exit Sub
 
    End With
 
    
    Exit Sub

Char_Erase_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_Erase"
    End If
Resume Next
    
End Sub
 
Private Sub Char_ResetInfo(ByVal CharIndex As Integer)
    
    On Error GoTo Char_ResetInfo_Err
    

    '*****************************************************************
    'Author: Ao 13.0
    'Last Modify Date: 13/12/2013
    'Reset Info User
    '*****************************************************************

    With charlist(CharIndex)
        Delete_All_Auras CharIndex
            
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
            
        If .ParticleIndex <> 0 Then
            Call Effect_Kill(.ParticleIndex, False)
            .ParticleIndex = 0

        End If

    End With
 
    
    Exit Sub

Char_ResetInfo_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_ResetInfo"
    End If
Resume Next
    
End Sub
 
Private Sub Char_MapPosGet(ByVal CharIndex As Long, ByRef X As Byte, ByRef Y As Byte)
    
    On Error GoTo Char_MapPosGet_Err
    
                                
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
 
    
    Exit Sub

Char_MapPosGet_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_MapPosGet"
    End If
Resume Next
    
End Sub
 
Public Sub Char_MapPosSet(ByVal X As Byte, ByVal Y As Byte)
    
    On Error GoTo Char_MapPosSet_Err
    

    'Sets the user postion

    If (Map_InBounds(X, Y)) Then  '// Posicion valida
        
        UserPos.X = X
        UserPos.Y = Y
                        
        'Set char
        MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
        charlist(UserCharIndex).Pos = UserPos
        
        Exit Sub
 
    End If

    
    Exit Sub

Char_MapPosSet_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_MapPosSet"
    End If
Resume Next
    
End Sub
 
Public Function Char_Techo() As Boolean
    
    On Error GoTo Char_Techo_Err
    

    '// Autor : Marcos Zeni
    '// Nueva forma de establecer si el usuario esta bajo un techo

    Char_Techo = False
 
    With charlist(UserCharIndex)
      
        If (Map_InBounds(.Pos.X, .Pos.Y)) Then '// Posicion valida
                       
            If (MapData(.Pos.X, .Pos.Y).Trigger = eTrigger.BAJOTECHO) Then
                Char_Techo = True

            End If
                               
        End If
   
    End With

    
    Exit Function

Char_Techo_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_Techo"
    End If
Resume Next
    
End Function
 
Public Function Char_MapPosExits(ByVal X As Byte, ByVal Y As Byte) As Integer
    
    On Error GoTo Char_MapPosExits_Err
    
 
    '*****************************************************************
    'Checks to see if a tile position has a char_index and return it
    '*****************************************************************
   
    If (Map_InBounds(X, Y)) Then
        Char_MapPosExits = MapData(X, Y).CharIndex
    Else
        Char_MapPosExits = 0

    End If
  
    
    Exit Function

Char_MapPosExits_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_MapPosExits"
    End If
Resume Next
    
End Function
 
Public Sub Char_UserPos()
    
    On Error GoTo Char_UserPos_Err
    

    '// Author Miqueas
    '// Actualizamo el lbl de la posicion del usuario
 
    Dim X As Byte
    Dim Y As Byte
     
    If Char_Check(UserCharIndex) Then
        
        '// Damos valor a las variables asi sacamos la pos del usuario.
        Call Char_MapPosGet(UserCharIndex, X, Y)
                
        bTecho = Char_Techo '// Pos : Techo :P
               
        frmMain.Coord.Caption = "Map:" & UserMap & " X:" & X & " Y:" & Y
 
        Exit Sub
 
    End If
 
    
    Exit Sub

Char_UserPos_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_UserPos"
    End If
Resume Next
    
End Sub
 
Public Sub Char_UserIndexSet(ByVal CharIndex As Integer)
    
    On Error GoTo Char_UserIndexSet_Err
    
 
    UserCharIndex = CharIndex
 
    With charlist(UserCharIndex)
 
        'Nueva posicion para el usuario.
        UserPos = .Pos
         
        Exit Sub
 
    End With
         
    
    Exit Sub

Char_UserIndexSet_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_UserIndexSet"
    End If
Resume Next
    
End Sub
 
Public Function Char_Check(ByVal CharIndex As Integer) As Boolean
    
    On Error GoTo Char_Check_Err
    
       
    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify by Miqueas150 Date: 24/02/2013
    'Chequeamos el Char
    '**************************************************************
       
    'check char_index
 
    If CharIndex > 0 And CharIndex <= LastChar Then
 
        With charlist(CharIndex) '// check char_index
            Char_Check = (.Heading > 0)

        End With
 
    End If
   
    
    Exit Function

Char_Check_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_Check"
    End If
Resume Next
    
End Function
 
Public Sub Char_SetInvisible(ByVal CharIndex As Integer, ByVal value As Boolean)
    
    On Error GoTo Char_SetInvisible_Err
    
       
    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify by Miqueas150 Date: 24/02/2013
 
    '**************************************************************
       
    If Char_Check(CharIndex) Then
 
        With charlist(CharIndex)
 
            .invisible = value '// User invisible o no ?
                        
            Exit Sub
 
        End With
 
    End If
 
    
    Exit Sub

Char_SetInvisible_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_SetInvisible"
    End If
Resume Next
    
End Sub
 
Public Sub Char_SetBody(ByVal CharIndex As Integer, ByVal BodyIndex As Integer)
    
    On Error GoTo Char_SetBody_Err
    
 
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
 
    
    Exit Sub

Char_SetBody_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_SetBody"
    End If
Resume Next
    
End Sub
 
Public Sub Char_SetHead(ByVal CharIndex As Integer, ByVal HeadIndex As Integer)
    
    On Error GoTo Char_SetHead_Err
    
 
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
                               
            .muerto = (HeadIndex = CASPER_HEAD)
                     
            Exit Sub
 
        End With
 
    End If
 
    
    Exit Sub

Char_SetHead_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_SetHead"
    End If
Resume Next
    
End Sub
 
Public Sub Char_SetHeading(ByVal CharIndex As Long, ByVal Heading As Byte)
    
    On Error GoTo Char_SetHeading_Err
    
 
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
 
    
    Exit Sub

Char_SetHeading_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_SetHeading"
    End If
Resume Next
    
End Sub

Public Sub Char_SetName(ByVal CharIndex As Integer, ByVal Name As String)
    
    On Error GoTo Char_SetName_Err
    
 
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
 
    
    Exit Sub

Char_SetName_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_SetName"
    End If
Resume Next
    
End Sub
 
Public Sub Char_SetWeapon(ByVal CharIndex As Integer, ByVal WeaponIndex As Integer)
    
    On Error GoTo Char_SetWeapon_Err
    
 
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
 
    
    Exit Sub

Char_SetWeapon_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_SetWeapon"
    End If
Resume Next
    
End Sub
 
Public Sub Char_SetShield(ByVal CharIndex As Integer, ByVal ShieldIndex As Integer)
    
    On Error GoTo Char_SetShield_Err
    
 
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
 
    
    Exit Sub

Char_SetShield_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_SetShield"
    End If
Resume Next
    
End Sub
 
Public Sub Char_SetCasco(ByVal CharIndex As Integer, ByVal CascoIndex As Integer)
    
    On Error GoTo Char_SetCasco_Err
    
 
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
     
    
    Exit Sub

Char_SetCasco_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_SetCasco"
    End If
Resume Next
    
End Sub
 
Public Sub Char_SetFx(ByVal CharIndex As Integer, _
                      ByVal fX As Integer, _
                      ByVal Loops As Integer)
    
    On Error GoTo Char_SetFx_Err
    
 
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
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
   
    
    Exit Sub

Char_SetFx_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_SetFx"
    End If
Resume Next
    
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
    
    On Error GoTo Char_Make_Err
    
 
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
       
    
    Exit Sub

Char_Make_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_Make"
    End If
Resume Next
    
End Sub

Public Sub Char_RefreshAll()
    '*****************************************************************
    'Goes through the charlist and replots all the characters on the map
    'Used to make sure everyone is visible
    '*****************************************************************
    
    On Error GoTo Char_RefreshAll_Err
    
 
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
 
    
    Exit Sub

Char_RefreshAll_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_RefreshAll"
    End If
Resume Next
    
End Sub

Sub Char_MovebyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
    
    On Error GoTo Char_MovebyPos_Err
    

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

    
    Exit Sub

Char_MovebyPos_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_MovebyPos"
    End If
Resume Next
    
End Sub

Sub Char_MoveScreen(ByVal nHeading As E_Heading)
    '******************************************
    'Starts the screen moving in a direction
    '******************************************
    
    On Error GoTo Char_MoveScreen_Err
    

    Dim X  As Integer
    Dim Y  As Integer
        
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

    If (tX < MinXBorder) Or (tX > MaxXBorder) Or (tY < MinYBorder) Or (tY > MaxYBorder) Then

        Exit Sub

    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
                
        bTecho = Char_Techo
               
    End If

    
    Exit Sub

Char_MoveScreen_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_MoveScreen"
    End If
Resume Next
    
End Sub

Sub Char_MovebyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
    '*****************************************************************
    'Starts the movement of a character in nHeading direction
    '*****************************************************************
    
    On Error GoTo Char_MovebyHead_Err
    

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

    
    Exit Sub

Char_MovebyHead_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_MovebyHead"
    End If
Resume Next
    
End Sub

Sub Char_CleanAll()
    '// Borramos los obj y char que esten
    
    On Error GoTo Char_CleanAll_Err
    

    Dim X         As Long, Y As Long
    Dim CharIndex As Integer, obj As Integer
    
    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
          
            'Erase NPCs
            CharIndex = Char_MapPosExits(CByte(X), CByte(Y))
 
            If (CharIndex > 0) Then
                Call Char_Erase(CharIndex)

            End If
                        
            'Erase OBJs
            obj = Map_PosExitsObject(CByte(X), CByte(Y))

            If (obj > 0) Then
                Call Map_DestroyObject(CByte(X), CByte(Y))

            End If

        Next Y
    Next X

    
    Exit Sub

Char_CleanAll_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mPooChar" & "->" & "Char_CleanAll"
    End If
Resume Next
    
End Sub

