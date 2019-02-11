Attribute VB_Name = "mDx8_Auras"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 26/05/10
'Blisse-AO | Sistema de Auras
'***************************************************

Option Explicit

Public Type Aura

    Grh As Integer '   GrhIndex
    
    Rotation As Byte '   Rotate or Not
    Angle As Single '   Angle
    Speed As Single '   Speed
    TickCount As Long '   TickCount from Speed Controls
    
    Color(0 To 3) As Long '   Color
    
    OffsetX As Integer '   PixelOffset X
    OffsetY As Integer '   PixelOffset Y

End Type

Public Auras() As Aura '   List of Aura's

Public Sub SetCharacterAura(ByVal CharIndex As Integer, _
                            ByVal AuraIndex As Byte, _
                            ByVal slot As Byte)
    
    On Error GoTo SetCharacterAura_Err
    

    '***************************************************
    'Author: Standelf
    'Last Modify Date: 27/05/2010
    '***************************************************
    If slot <= 0 Or slot >= 5 Then Exit Sub
    Set_Aura CharIndex, slot, AuraIndex

    
    Exit Sub

SetCharacterAura_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Auras" & "->" & "SetCharacterAura"
    End If
Resume Next
    
End Sub

Public Sub Load_Auras()
    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 26/05/10
    'Load Auras
    '***************************************************
    
    On Error GoTo Load_Auras_Err
    
    Dim i As Integer, AurasTotales As Integer, Leer As New ClsIniReader
    Leer.Initialize App.path & "\init\auras.ini"

    AurasTotales = Val(Leer.GetValue("Auras", "NumAuras"))
    
    ReDim Preserve Auras(1 To AurasTotales)
    
    For i = 1 To AurasTotales
        Auras(i).Grh = Val(Leer.GetValue(i, "GrhIndex"))
                
        Auras(i).Rotation = Val(Leer.GetValue(i, "Rotate"))
        Auras(i).Angle = 0
        Auras(i).Speed = Leer.GetValue(i, "Speed")
                
        Auras(i).OffsetX = Val(Leer.GetValue(i, "OffsetX"))
        Auras(i).OffsetY = Val(Leer.GetValue(i, "OffsetY"))

        Dim ColorSet As Byte, TempSet As String
            
        For ColorSet = 0 To 3
            TempSet = Leer.GetValue(Val(i), "Color" & ColorSet)
            Auras(i).Color(ColorSet) = D3DColorXRGB(ReadField(1, TempSet, Asc(",")), ReadField(2, TempSet, Asc(",")), ReadField(3, TempSet, Asc(",")))
        Next ColorSet
                
        Auras(i).TickCount = 0
    Next i
                                                         
    Set Leer = Nothing

    
    Exit Sub

Load_Auras_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Auras" & "->" & "Load_Auras"
    End If
Resume Next
    
End Sub

Public Sub DeInit_Auras()
    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 26/05/10
    'DeInit Auras
    '***************************************************
    '   Erase Data
    
    On Error GoTo DeInit_Auras_Err
    
    Erase Auras()
    
    '   Finish
    Exit Sub

    
    Exit Sub

DeInit_Auras_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Auras" & "->" & "DeInit_Auras"
    End If
Resume Next
    
End Sub

Public Sub Set_Aura(ByVal CharIndex As Integer, slot As Byte, Aura As Byte)
    
    On Error GoTo Set_Aura_Err
    

    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 26/05/10
    'Set Aura to Char
    '***************************************************
    If slot <= 0 Or slot >= 5 Then Exit Sub
    
    With charlist(CharIndex).Aura(slot)
        .Grh = Auras(Aura).Grh
            
        .Angle = Auras(Aura).Angle
        .Rotation = Auras(Aura).Rotation
        .Speed = Auras(Aura).Speed
        
        .OffsetX = Auras(Aura).OffsetX
        .OffsetY = Auras(Aura).OffsetY
        
        .Color(0) = Auras(Aura).Color(0)
        .Color(1) = Auras(Aura).Color(1)
        .Color(2) = Auras(Aura).Color(2)
        .Color(3) = Auras(Aura).Color(3)
        
        .TickCount = GetTickCount

    End With

    
    Exit Sub

Set_Aura_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Auras" & "->" & "Set_Aura"
    End If
Resume Next
    
End Sub

Public Sub Delete_All_Auras(ByVal CharIndex As Integer)
    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 26/05/10
    'Kill all of aura´s from Char
    '***************************************************
    
    On Error GoTo Delete_All_Auras_Err
    
    Delete_Aura CharIndex, 1
    Delete_Aura CharIndex, 2
    Delete_Aura CharIndex, 3
    Delete_Aura CharIndex, 4

    
    Exit Sub

Delete_All_Auras_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Auras" & "->" & "Delete_All_Auras"
    End If
Resume Next
    
End Sub
    
Public Sub Delete_Aura(ByVal CharIndex As Integer, slot As Byte)
    
    On Error GoTo Delete_Aura_Err
    

    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 26/05/10
    'Kill Aura from Char
    '***************************************************
    If slot <= 0 Or slot >= 5 Then Exit Sub
    
    charlist(CharIndex).Aura(slot) = Auras(1) '1 = Fake Aura

    
    Exit Sub

Delete_Aura_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Auras" & "->" & "Delete_Aura"
    End If
Resume Next
    
End Sub

Public Sub Update_Aura(ByVal CharIndex As Integer, slot As Byte)
    
    On Error GoTo Update_Aura_Err
    

    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 26/05/10
    'Update Angle of Aura
    '***************************************************
    If slot <= 0 Or slot >= 5 Then Exit Sub
    
    With charlist(CharIndex).Aura(slot)

        If GetTickCount - .TickCount > FPS Then
            .Angle = .Angle + .Speed

            If .Angle >= 360 Then .Angle = 0
            .TickCount = GetTickCount

        End If

    End With

    
    Exit Sub

Update_Aura_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Auras" & "->" & "Update_Aura"
    End If
Resume Next
    
End Sub

Public Sub Render_Auras(ByVal CharIndex As Integer, X As Integer, Y As Integer)

    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 26/05/10
    'Render the Auras from a Char
    '***************************************************
    On Error GoTo handle

    Dim i As Byte

    For i = 1 To 4

        With charlist(CharIndex).Aura(i)

            If .Grh <> 0 Then
                If .Rotation = 1 Then Update_Aura CharIndex, i
                Call DDrawTransGrhIndextoSurface(.Grh, X + .OffsetX, Y + .OffsetY, 1, .Color(), .Angle, True)

            End If

        End With

    Next i

handle:
    Exit Sub

End Sub

