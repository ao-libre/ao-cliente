Attribute VB_Name = "mPooMap"
'---------------------------------------------------------------------------------------
' Module    : Mod_PooMap
' Author    :  Miqueas
' Date      : 02/02/2014
' Purpose   :  xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'---------------------------------------------------------------------------------------

Option Explicit

Private Const GrhFogata As Long = 1521

Public Sub Map_RemoveOldUser()

      With MapData(UserPos.X, UserPos.Y)

            If (.CharIndex = UserCharIndex) Then
                  .CharIndex = 0
            End If

      End With

End Sub

Public Sub Map_CreateObject(ByVal X As Byte, ByVal Y As Byte, ByVal GrhIndex As Long)

      If Not GrhCheck(GrhIndex) Then Exit Sub
                        
      If (Map_InBounds(X, Y)) Then
            Call InitGrh(MapData(X, Y).ObjGrh, GrhIndex)
      End If

End Sub

Public Sub Map_DestroyObject(ByVal X As Byte, ByVal Y As Byte)

      If (Map_InBounds(X, Y)) Then

            With MapData(X, Y)
                
                .OBJInfo.objindex = 0
                .OBJInfo.Amount = 0
                  
                Call GrhUninitialize(.ObjGrh)
        
            End With

      End If

End Sub

Public Function Map_PosExitsObject(ByVal X As Byte, ByVal Y As Byte) As Integer
 
      '*****************************************************************
      'Checks to see if a tile position has a char_index and return it
      '*****************************************************************

      If (Map_InBounds(X, Y)) Then
            Map_PosExitsObject = MapData(X, Y).ObjGrh.GrhIndex
      Else
            Map_PosExitsObject = 0
      End If
 
End Function

Public Function Map_GetBlocked(ByVal X As Integer, ByVal Y As Integer) As Boolean
      '*****************************************************************
      'Author: Aaron Perkins - Modified by Juan Martin Sotuyo Dodero
      'Last Modify Date: 10/07/2002
      'Checks to see if a tile position is blocked
      '*****************************************************************

      If (Map_InBounds(X, Y)) Then
            Map_GetBlocked = (MapData(X, Y).Blocked)
      End If

End Function

Public Sub Map_SetBlocked(ByVal X As Byte, ByVal Y As Byte, ByVal block As Byte)

      If (Map_InBounds(X, Y)) Then
            MapData(X, Y).Blocked = block
      End If

End Sub

Sub Map_MoveTo(ByVal Direccion As E_Heading)
      '***************************************************
      'Author: Alejandro Santos (AlejoLp)
      'Last Modify Date: 06/28/2008
      'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
      ' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
      ' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
      ' 06/28/2008: NicoNZ - Saque lo que impedia que si el usuario estaba paralizado se ejecute el sub.
      '***************************************************

      Dim LegalOk As Boolean
      Static lastmovement As Long
      
      If Cartel Then Cartel = False
    
      Select Case Direccion

            Case E_Heading.NORTH
                  LegalOk = Map_LegalPos(UserPos.X, UserPos.Y - 1)

            Case E_Heading.EAST
                  LegalOk = Map_LegalPos(UserPos.X + 1, UserPos.Y)

            Case E_Heading.SOUTH
                  LegalOk = Map_LegalPos(UserPos.X, UserPos.Y + 1)

            Case E_Heading.WEST
                  LegalOk = Map_LegalPos(UserPos.X - 1, UserPos.Y)
                        
      End Select

      If LegalOk And Not UserParalizado And Not UserDescansar And Not UserMeditar Then
          Call WriteWalk(Direccion)
          Call frmMain.ActualizarMiniMapa   'integrado por ReyarB

          Call Char_MovebyHead(UserCharIndex, Direccion)
          Call Char_MoveScreen(Direccion)
      
      Else
      
        If (charlist(UserCharIndex).Heading <> Direccion) Then
            If MainTimer.Check(TimersIndex.ChangeHeading) Then
                Call WriteChangeHeading(Direccion)
                Call Char_SetHeading(UserCharIndex, Direccion)
            End If
        End If
                
      End If
    
      If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
      If frmMain.trainingMacro.Enabled Then Call frmMain.DesactivarMacroHechizos

      ' Update 3D sounds!
      Call Audio.MoveListener(UserPos.X, UserPos.Y)
  
      ' Esto es un parche por que por alguna razon si el pj esta meditando y nos movemos el juego explota por eso cambie
      ' Las validaciones en la linea 131 y agregue esto para arreglarlo (Recox)
      If UserMeditar Then
        UserMeditar = Not UserMeditar
      End If

      If UserDescansar Then
        UserDescansar = Not UserDescansar
      End If
        
End Sub

Function Map_LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
      '*****************************************************************
      'Author: ZaMa
      'Last Modification: 06/04/2020
      'Checks to see if a tile position is legal, including if there is a casper in the tile
      '10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
      '01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
      '12/01/2020: Recox - Now we manage monturas.
      '06/04/2020: FrankoH298 - Si estamos montados, no nos deja ingresar a las casas.
      '*****************************************************************

      Dim CharIndex As Integer
    
      'Limites del mapa

      If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then

            Exit Function

      End If
    
      'Tile Bloqueado?

      If (Map_GetBlocked(X, Y)) Then
         
            Exit Function

      End If
    
      CharIndex = (Char_MapPosExits(CByte(X), CByte(Y)))
        
      'Hay un personaje?

      If (CharIndex > 0) Then
    
            If (Map_GetBlocked(UserPos.X, UserPos.Y)) Then
                
                  Exit Function

            End If
        
            With charlist(CharIndex)
                  ' Si no es casper, no puede pasar

                  If .iHead <> eCabezas.CASPER_HEAD And .iBody <> eCabezas.FRAGATA_FANTASMAL Then
                              
                        Exit Function

                  Else
                        ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)

                        If (Map_CheckWater(UserPos.X, UserPos.Y)) Then
                              If Not (Map_CheckWater(X, Y)) Then
                                            
                                    Exit Function

                              End If

                        Else
                              ' No puedo intercambiar con un casper que este en la orilla (Lado agua)

                              If (Map_CheckWater(X, Y)) Then
                                             
                                    Exit Function

                              End If
                                        
                        End If
                
                        ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles

                        If (EsGM(UserCharIndex)) Then

                              If (charlist(UserCharIndex).invisible) Then
                                             
                                    Exit Function

                              End If
                                        
                        End If
                  End If

            End With

      End If
   
      If (UserNavegando <> Map_CheckWater(X, Y)) Then
               
            Exit Function

      End If
      
      'Esta el usuario Equitando bajo un techo?
      If UserEquitando And MapData(X, Y).Trigger = eTrigger.BAJOTECHO Or MapData(X, Y).Trigger = eTrigger.CASA Then
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_MONTURA_SALIR").item("TEXTO"))
            Exit Function
      End If
      
      If UserEvento Then Exit Function
      
    
      Map_LegalPos = True
End Function

Function Map_InBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
      '*****************************************************************
      'Checks to see if a tile position is in the maps bounds
      '*****************************************************************

      If (X < XMinMapSize) Or (X > XMaxMapSize) Or (Y < YMinMapSize) Or (Y > YMaxMapSize) Then
            Map_InBounds = False

            Exit Function

      End If
    
      Map_InBounds = True
End Function

Public Function Map_CheckBonfire(ByRef Location As Position) As Boolean

      Dim J As Long
      Dim k As Long
    
      For J = UserPos.X - 8 To UserPos.X + 8
            For k = UserPos.Y - 6 To UserPos.Y + 6

                  If Map_InBounds(J, k) Then
                        If MapData(X, Y).ObjGrh.GrhIndex = GrhFogata Then
                              Location.X = J
                              Location.Y = k
                              Map_CheckBonfire = True

                              Exit Function

                        End If
                  End If

            Next k
      Next J

End Function

Function Map_CheckWater(ByVal X As Integer, ByVal Y As Integer) As Boolean

      If Map_InBounds(X, Y) Then

            With MapData(X, Y)

                  If ((.Graphic(1).GrhIndex >= 1505 And .Graphic(1).GrhIndex <= 1520) Or (.Graphic(1).GrhIndex >= 5665 And .Graphic(1).GrhIndex <= 5680) Or (.Graphic(1).GrhIndex >= 13547 And .Graphic(1).GrhIndex <= 13562)) And .Graphic(2).GrhIndex = 0 Then
                        Map_CheckWater = True
                  Else
                        Map_CheckWater = False
                  End If

            End With

      End If
                  
End Function

