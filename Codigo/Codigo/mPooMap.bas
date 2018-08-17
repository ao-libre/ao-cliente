Attribute VB_Name = "mPooMap"
'---------------------------------------------------------------------------------------
' Module    : Mod_PooMap
' Author    :  Miqueas
' Date      : 02/02/2014
' Purpose   :  xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'---------------------------------------------------------------------------------------

Option Explicit

Private Const GrhFogata As Integer = 1521

Public Sub Map_RemoveOldUser()

      With MapData(UserPos.x, UserPos.y)

            If (.CharIndex = UserCharIndex) Then
                  .CharIndex = 0
            End If

      End With

End Sub
Public Sub Map_CreateObject(ByVal x As Byte, ByVal y As Byte, ByVal GrhIndex As Integer)

      'Dim objgrh As Integer
        
      If Not GrhCheck(GrhIndex) Then
            Exit Sub

      End If
                        
      If (Map_InBounds(x, y)) Then

            With MapData(x, y)

                  'If (Map_PosExitsObject(x, y) > 0) Then
                  '      Call Map_DestroyObject(x, y)
                  'End If

                  '.objgrh.GrhIndex = GrhIndex
                  Call InitGrh(.ObjGrh, GrhIndex)
            End With

      End If

End Sub

Public Sub Map_DestroyObject(ByVal x As Byte, ByVal y As Byte)

      If (Map_InBounds(x, y)) Then

            With MapData(x, y)
                  '.objgrh.GrhIndex = 0
                  .OBJInfo.ObjIndex = 0
                  .OBJInfo.Amount = 0
                  Call GrhUninitialize(.ObjGrh)
        
            End With

      End If

End Sub

Public Function Map_PosExitsObject(ByVal x As Byte, ByVal y As Byte) As Integer
 
      '*****************************************************************
      'Checks to see if a tile position has a char_index and return it
      '*****************************************************************

      If (Map_InBounds(x, y)) Then
            Map_PosExitsObject = MapData(x, y).ObjGrh.GrhIndex
      Else
            Map_PosExitsObject = 0
      End If
 
End Function

Public Function Map_GetBlocked(ByVal x As Integer, ByVal y As Integer) As Boolean
      '*****************************************************************
      'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
      'Last Modify Date: 10/07/2002
      'Checks to see if a tile position is blocked
      '*****************************************************************

      If (Map_InBounds(x, y)) Then
            Map_GetBlocked = (MapData(x, y).Blocked)
      End If

End Function

Public Sub Map_SetBlocked(ByVal x As Byte, ByVal y As Byte, ByVal block As Byte)

      If (Map_InBounds(x, y)) Then
            MapData(x, y).Blocked = block
      End If

End Sub

Sub Map_MoveTo(ByVal Direccion As E_Heading)
      '***************************************************
      'Author: Alejandro Santos (AlejoLp)
      'Last Modify Date: 06/28/2008
      'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
      ' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
      ' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
      ' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
      '***************************************************

      Dim LegalOk As Boolean
    
      If Cartel Then Cartel = False
    
      Select Case Direccion

            Case E_Heading.NORTH
                  LegalOk = Map_LegalPos(UserPos.x, UserPos.y - 1)

            Case E_Heading.EAST
                  LegalOk = Map_LegalPos(UserPos.x + 1, UserPos.y)

            Case E_Heading.SOUTH
                  LegalOk = Map_LegalPos(UserPos.x, UserPos.y + 1)

            Case E_Heading.WEST
                  LegalOk = Map_LegalPos(UserPos.x - 1, UserPos.y)
                        
      End Select
    
      If LegalOk And Not UserParalizado Then
        
            Call WriteWalk(Direccion)

            If Not UserDescansar And Not UserMeditar Then
                  Call Char_MovebyHead(UserCharIndex, Direccion)
                  Call Char_MoveScreen(Direccion)
            End If

      Else

            If (charlist(UserCharIndex).Heading <> Direccion) Then
                  Call WriteChangeHeading(Direccion)
            End If
                
      End If
    
      If (frmMain.macrotrabajo.Enabled) Then
            Call frmMain.DesactivarMacroTrabajo
      End If

      ' Update 3D sounds!
      Call Audio.MoveListener(UserPos.x, UserPos.y)
        
End Sub

Function Map_LegalPos(ByVal x As Integer, ByVal y As Integer) As Boolean
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

      If (Map_GetBlocked(x, y)) Then
         
            Exit Function

      End If
    
      CharIndex = (Char_MapPosExits(CByte(x), CByte(y)))
        
      '¿Hay un personaje?

      If (CharIndex > 0) Then
    
            If (Map_GetBlocked(UserPos.x, UserPos.y)) Then
                
                  Exit Function

            End If
        
            With charlist(CharIndex)
                  ' Si no es casper, no puede pasar

                  If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                              
                        Exit Function

                  Else
                        ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)

                        If (Map_CheckWater(UserPos.x, UserPos.y)) Then
                              If Not (Map_CheckWater(x, y)) Then
                                            
                                    Exit Function

                              End If

                        Else
                              ' No puedo intercambiar con un casper que este en la orilla (Lado agua)

                              If (Map_CheckWater(x, y)) Then
                                             
                                    Exit Function

                              End If
                                        
                        End If
                
                        ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles

                        If (esGM(UserCharIndex)) Then

                              If (charlist(UserCharIndex).invisible) Then
                                             
                                    Exit Function

                              End If
                                        
                        End If
                  End If

            End With

      End If
   
      If (UserNavegando <> Map_CheckWater(x, y)) Then
               
            Exit Function

      End If
    
      Map_LegalPos = True
End Function

Function Map_InBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
      '*****************************************************************
      'Checks to see if a tile position is in the maps bounds
      '*****************************************************************

      If (x < XMinMapSize) Or (x > XMaxMapSize) Or (y < YMinMapSize) Or (y > YMaxMapSize) Then
            Map_InBounds = False

            Exit Function

      End If
    
      Map_InBounds = True
End Function

Public Function Map_CheckBonfire(ByRef Location As Position) As Boolean

      Dim j As Long
      Dim k As Long
    
      For j = UserPos.x - 8 To UserPos.x + 8
            For k = UserPos.y - 6 To UserPos.y + 6

                  If Map_InBounds(j, k) Then
                        If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                              Location.x = j
                              Location.y = k
                    
                              Map_CheckBonfire = True

                              Exit Function

                        End If
                  End If

            Next k
      Next j

End Function

Function Map_CheckWater(ByVal x As Integer, ByVal y As Integer) As Boolean

      If Map_InBounds(x, y) Then

            With MapData(x, y)

                  If ((.Graphic(1).GrhIndex >= 1505 And .Graphic(1).GrhIndex <= 1520) Or (.Graphic(1).GrhIndex >= 5665 And .Graphic(1).GrhIndex <= 5680) Or (.Graphic(1).GrhIndex >= 13547 And .Graphic(1).GrhIndex <= 13562)) And .Graphic(2).GrhIndex = 0 Then
                        Map_CheckWater = True
                  Else
                        Map_CheckWater = False
                  End If

            End With

      End If
                  
End Function

